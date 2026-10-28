import { Controller, Get, Post, Delete, Param, Query, UseInterceptors, UploadedFile, Res, NotFoundException, BadRequestException, Body, HttpException } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { JobService } from './job.service';
import { Response } from 'express';
import { StorageService } from '../storage/storage.service';
import * as fs from 'fs';
import * as path from 'path';

@Controller('api')
export class JobController {
  constructor(
    private readonly jobService: JobService,
    private readonly storage: StorageService,
  ) { }

  @Post('job/upload')
  @UseInterceptors(FileInterceptor('file'))
  async uploadExcel(
    @UploadedFile() file: Express.Multer.File,
    @Body('headerMode') headerMode?: string,
    @Body('cleanupDuplicates') cleanupDuplicates?: string,
  ) {
    if (!file) {
      throw new BadRequestException('Excel file is required');
    }
    try {
      const job = await this.jobService.createJob(file, headerMode, cleanupDuplicates === 'true');
      return job;
    } catch (err) {
      if (err instanceof HttpException) throw err;
      const message = err instanceof Error ? err.message : 'PDF generation could not be started';
      throw new BadRequestException(message);
    }
  }

  @Post('job/validate')
  @UseInterceptors(FileInterceptor('file'))
  async validateExcel(@UploadedFile() file: Express.Multer.File) {
    if (!file) throw new BadRequestException('Excel file is required');
    try {
      return this.jobService.validateWorkbook(file);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Workbook validation could not be completed';
      throw new BadRequestException(message);
    }
  }

  @Post('pdf/preview')
  @UseInterceptors(FileInterceptor('file'))
  async previewPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
    @Body('headerMode') headerMode?: string,
  ) {
    if (!file) {
      throw new BadRequestException('Excel file is required');
    }
    try {
      const pdfBuffer = await this.jobService.generatePreview(file, headerMode);
      res.set({
        'Content-Type': 'application/pdf',
        'Content-Disposition': 'inline; filename="preview.pdf"',
        'Content-Length': pdfBuffer.length,
      });
      res.end(pdfBuffer);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Preview could not be created';
      throw new BadRequestException(message);
    }
  }

  /** Upload a reference PDF template */
  @Post('template/upload')
  @UseInterceptors(FileInterceptor('file'))
  async uploadTemplate(@UploadedFile() file: Express.Multer.File) {
    if (!file) throw new BadRequestException('PDF file is required');
    if (file.mimetype !== 'application/pdf') throw new BadRequestException('Only PDF files are allowed');
    const templatePath = path.join(this.storage.getStaticDir(), 'template.pdf');
    await fs.promises.writeFile(templatePath, file.buffer);
    return { message: 'Template saved successfully', size: file.size };
  }

  /** Check if a template exists */
  @Get('template/exists')
  async templateExists() {
    const templatePath = path.join(this.storage.getStaticDir(), 'template.pdf');
    return { exists: fs.existsSync(templatePath) };
  }

  /** Serve the reference PDF template inline */
  @Get('template/view')
  async viewTemplate(@Res() res: Response) {
    const templatePath = path.join(this.storage.getStaticDir(), 'template.pdf');
    if (!fs.existsSync(templatePath)) {
      throw new NotFoundException('No template uploaded yet');
    }
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': 'inline; filename="template.pdf"',
    });
    res.sendFile(templatePath);
  }

  @Get('job/list')
  async listJobs() {
    return this.jobService.listJobs();
  }

  @Get('job/:id')
  async getJob(@Param('id') id: string) {
    return this.jobService.getJob(id);
  }

  @Delete('job/:id')
  async deleteJob(@Param('id') id: string) {
    return this.jobService.deleteJob(id);
  }

  @Get('job/:id/zip')
  async downloadZip(@Param('id') id: string, @Res() res: Response) {
    try {
      await this.jobService.zipGeneratedPdfs(id, res);
    } catch (err) {
      throw new BadRequestException(err.message);
    }
  }

  @Get('pdf/list')
  async listPdfs(@Query('search') search?: string) {
    return this.jobService.listAllPdfs(search);
  }

  @Get('pdf/download/:filename')
  async downloadPdf(@Param('filename') filename: string, @Res() res: Response) {
    const filePath = this.storage.getAbsolutePath('pdfs', filename);
    if (!fs.existsSync(filePath)) {
      throw new NotFoundException(`File ${filename} not found`);
    }
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.sendFile(filePath);
  }
}
