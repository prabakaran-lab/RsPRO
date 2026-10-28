import { Body, Controller, Get, Post, Patch, Delete, Param, Query, UseInterceptors, UploadedFiles, Res, NotFoundException } from '@nestjs/common';
import { FilesInterceptor } from '@nestjs/platform-express';
import { ImageService } from './image.service';
import { StorageService } from '../storage/storage.service';
import { Response } from 'express';
import * as fs from 'fs';

@Controller('api/image')
export class ImageController {
  constructor(
    private readonly imageService: ImageService,
    private readonly storage: StorageService,
  ) {}

  @Post('upload')
  @UseInterceptors(FilesInterceptor('files', 500)) // Accept up to 500 files per batch
  async uploadImages(@UploadedFiles() files: Array<Express.Multer.File>) {
    if (!files?.length) {
      return { uploadedCount: 0, failedCount: 0, results: [] };
    }

    const results = [];
    for (const file of files) {
      try {
        const asset = await this.imageService.handleUpload(file);
        results.push({
          originalName: file.originalname,
          filename: asset.filename,
          status: 'success',
          asset,
          id: asset.id,
          path: asset.path,
          url: asset.url,
          copyLink: asset.url,
          downloadUrl: `/api/image/download/${encodeURIComponent(asset.filename)}`,
        });
      } catch (err) {
        results.push({ filename: file.originalname, status: 'error', message: err.message });
      }
    }
    return {
      uploadedCount: results.filter(r => r.status === 'success').length,
      failedCount: results.filter(r => r.status === 'error').length,
      results,
    };
  }

  @Get('list')
  async listImages(
    @Query('search') search?: string,
    @Query('type') type?: string,
    @Query('page') page?: string,
    @Query('limit') limit?: string,
  ) {
    return this.imageService.listImages({
      search,
      type,
      page: page ? parseInt(page, 10) : 1,
      limit: limit ? parseInt(limit, 10) : 20,
    });
  }

  @Delete(':id')
  async deleteImage(@Param('id') id: string) {
    return this.imageService.deleteImage(id);
  }

  @Patch(':id/rename')
  async renameImage(@Param('id') id: string, @Body('filename') filename: string) {
    return this.imageService.renameImage(id, filename);
  }

  @Get('download/:filename')
  async downloadImage(@Param('filename') filename: string, @Res() res: Response) {
    const filePath = this.storage.getAbsolutePath('images', filename);
    if (!fs.existsSync(filePath)) {
      throw new NotFoundException(`File ${filename} not found`);
    }
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.sendFile(filePath);
  }
}
