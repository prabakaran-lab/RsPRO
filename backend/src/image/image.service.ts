import { Injectable, Logger, ConflictException, NotFoundException } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { StorageService } from '../storage/storage.service';
import * as path from 'path';
import * as fs from 'fs';

@Injectable()
export class ImageService {
  private readonly logger = new Logger(ImageService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly storage: StorageService,
  ) {}

  private cleanFileName(fileName: string): string {
    return path
      .basename(fileName)
      .trim()
      .replace(/[\/\\]+/g, '_')
      .replace(/\s+/g, '_')
      .replace(/[^a-zA-Z0-9._-]/g, '_');
  }

  private imageMimeType(fileName: string): string | null {
    const extension = path.extname(fileName).toLowerCase();
    const mimeTypes: Record<string, string> = {
      '.avif': 'image/avif',
      '.bmp': 'image/bmp',
      '.gif': 'image/gif',
      '.jpeg': 'image/jpeg',
      '.jpg': 'image/jpeg',
      '.png': 'image/png',
      '.svg': 'image/svg+xml',
      '.tif': 'image/tiff',
      '.tiff': 'image/tiff',
      '.webp': 'image/webp',
    };

    return mimeTypes[extension] || null;
  }

  private async syncStoredImages(): Promise<number> {
    const imagesDir = path.join(this.storage.getStaticDir(), 'images');
    const directoryEntries = await fs.promises.readdir(imagesDir, { withFileTypes: true });
    const imageFiles = directoryEntries
      .filter(entry => entry.isFile() && this.imageMimeType(entry.name))
      .map(entry => entry.name);

    if (!imageFiles.length) return 0;

    const pathKeys = imageFiles.map(fileName => `/images/${fileName}`);
    const existing = await this.prisma.imageAsset.findMany({
      where: { path: { in: pathKeys } },
      select: { path: true },
    });
    const existingPaths = new Set(existing.map(asset => asset.path));
    const host = process.env.HOST_URL || 'http://localhost:8001';
    const missingAssets = await Promise.all(
      imageFiles
        .filter(fileName => !existingPaths.has(`/images/${fileName}`))
        .map(async fileName => {
          const stats = await fs.promises.stat(path.join(imagesDir, fileName));
          return {
            filename: fileName,
            path: `/images/${fileName}`,
            mimeType: this.imageMimeType(fileName)!,
            sizeBytes: stats.size,
            url: `${host}/static/images/${encodeURIComponent(fileName)}`,
          };
        }),
    );

    if (!missingAssets.length) return 0;

    await this.prisma.imageAsset.createMany({ data: missingAssets });
    this.logger.log(`Indexed ${missingAssets.length} existing image file(s) from storage.`);
    return missingAssets.length;
  }

  private enrichAsset(asset: any) {
    const host = process.env.HOST_URL || 'http://localhost:8001';
    const encodedFileName = encodeURIComponent(asset.filename);
    const url = asset.url || `${host}/static/images/${encodedFileName}`;

    return {
      ...asset,
      url,
      copyLink: url,
      downloadUrl: `${host}/api/image/download/${encodedFileName}`,
    };
  }

  async handleUpload(file: Express.Multer.File): Promise<any> {
    const originalName = file.originalname;
    const cleanName = this.cleanFileName(originalName);
    if (!cleanName) {
      throw new ConflictException('Invalid image filename');
    }
    const pathKey = `/images/${cleanName}`;

    // Check if asset already exists in DB
    const existing = await this.prisma.imageAsset.findUnique({
      where: { path: pathKey },
    });

    if (existing) {
      // Overwrite file on disk and update DB
      const storageResult = await this.storage.saveFile(file.buffer, 'images', cleanName);
      const updated = await this.prisma.imageAsset.update({
        where: { path: pathKey },
        data: {
          mimeType: file.mimetype,
          sizeBytes: file.size,
          url: storageResult.url,
        },
      });
      return this.enrichAsset(updated);
    }

    // Save to storage
    const storageResult = await this.storage.saveFile(file.buffer, 'images', cleanName);

    // Save metadata in database
    const asset = await this.prisma.imageAsset.create({
      data: {
        filename: cleanName,
        path: pathKey,
        mimeType: file.mimetype,
        sizeBytes: file.size,
        url: storageResult.url,
      },
    });

    return this.enrichAsset(asset);
  }

  async listImages(query: { search?: string; type?: string; page?: number; limit?: number }) {
    await this.syncStoredImages();

    const search = query.search || '';
    const type = query.type || '';
    const page = Math.max(1, Number(query.page) || 1);
    const limit = Math.min(100, Math.max(1, Number(query.limit) || 20));
    const skip = (page - 1) * limit;

    const where: any = {};

    if (search) {
      where.OR = [
        { filename: { contains: search } },
        { path: { contains: search } },
      ];
    }

    if (type) {
      where.mimeType = { contains: type };
    }

    const [items, total] = await Promise.all([
      this.prisma.imageAsset.findMany({
        where,
        orderBy: { createdAt: 'desc' },
        skip,
        take: limit,
      }),
      this.prisma.imageAsset.count({ where }),
    ]);

    return {
      items: items.map(item => this.enrichAsset(item)),
      data: items.map(item => this.enrichAsset(item)),
      total,
      page,
      limit,
      totalPages: Math.ceil(total / limit),
    };
  }

  async renameImage(id: string, requestedName: string) {
    const asset = await this.prisma.imageAsset.findUnique({
      where: { id },
    });

    if (!asset) {
      throw new NotFoundException(`Image asset with ID ${id} not found`);
    }

    const trimmed = (requestedName || '').trim();
    if (!trimmed) {
      throw new ConflictException('Image name cannot be empty');
    }

    const currentExt = path.extname(asset.filename);
    const requestedExt = path.extname(trimmed);
    const nextFileName = this.cleanFileName(requestedExt ? trimmed : `${trimmed}${currentExt}`);
    if (!nextFileName) {
      throw new ConflictException('Invalid image name');
    }

    if (nextFileName === asset.filename) {
      return this.enrichAsset(asset);
    }

    const nextPath = `/images/${nextFileName}`;
    const existing = await this.prisma.imageAsset.findUnique({
      where: { path: nextPath },
    });
    if (existing && existing.id !== id) {
      throw new ConflictException(`Image named "${nextFileName}" already exists`);
    }

    const oldAbs = this.storage.getAbsolutePath('images', asset.filename);
    const nextAbs = this.storage.getAbsolutePath('images', nextFileName);
    if (fs.existsSync(nextAbs)) {
      throw new ConflictException(`File "${nextFileName}" already exists`);
    }
    if (fs.existsSync(oldAbs)) {
      await fs.promises.rename(oldAbs, nextAbs);
    }

    const host = process.env.HOST_URL || 'http://localhost:8001';
    const updated = await this.prisma.imageAsset.update({
      where: { id },
      data: {
        filename: nextFileName,
        path: nextPath,
        url: `${host}/static/images/${encodeURIComponent(nextFileName)}`,
      },
    });

    return this.enrichAsset(updated);
  }

  async deleteImage(id: string) {
    const asset = await this.prisma.imageAsset.findUnique({
      where: { id },
    });

    if (!asset) {
      throw new NotFoundException(`Image asset with ID ${id} not found`);
    }

    // Delete from file system
    await this.storage.deleteFile('images', asset.filename);

    // Delete from DB
    await this.prisma.imageAsset.delete({
      where: { id },
    });

    return { success: true };
  }
}
