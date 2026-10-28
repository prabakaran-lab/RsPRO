import { Injectable, Logger } from '@nestjs/common';
import * as fs from 'fs';
import * as path from 'path';

@Injectable()
export class StorageService {
  private readonly logger = new Logger(StorageService.name);
  private readonly staticDir = path.join(process.cwd(), '..', 'static');

  constructor() {
    // Ensure static upload directories exist
    this.ensureDirExists(this.staticDir);
    this.ensureDirExists(path.join(this.staticDir, 'images'));
    this.ensureDirExists(path.join(this.staticDir, 'pdfs'));
  }

  private ensureDirExists(dirPath: string) {
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }
  }

  async saveFile(fileBuffer: Buffer, folder: 'images' | 'pdfs', fileName: string): Promise<{ url: string; relativePath: string; absolutePath: string }> {
    // Clean fileName to prevent folder traversal
    const safeName = path.basename(fileName).replace(/[\/\s]/g, '_');
    const relativePath = `/${folder}/${safeName}`;
    const absolutePath = path.join(this.staticDir, folder, safeName);

    // Save locally
    await fs.promises.writeFile(absolutePath, fileBuffer);
    this.logger.log(`Saved file locally: ${absolutePath}`);

    // If S3/R2 is configured (e.g. process.env.AWS_ACCESS_KEY_ID), we would upload here.
    // For local fallback: we return the localhost server link.
    const host = process.env.HOST_URL || 'http://localhost:8001';
    const url = `${host}/static${relativePath}`;

    return {
      url,
      relativePath,
      absolutePath,
    };
  }

  async deleteFile(folder: 'images' | 'pdfs', fileName: string): Promise<void> {
    const safeName = path.basename(fileName);
    const absolutePath = path.join(this.staticDir, folder, safeName);
    if (fs.existsSync(absolutePath)) {
      await fs.promises.unlink(absolutePath);
      this.logger.log(`Deleted file: ${absolutePath}`);
    }
  }

  getAbsolutePath(folder: 'images' | 'pdfs', fileName: string): string {
    return path.join(this.staticDir, folder, path.basename(fileName));
  }

  getStaticDir(): string {
    return this.staticDir;
  }
}
