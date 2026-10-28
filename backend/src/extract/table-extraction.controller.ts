import { BadRequestException, Body, Controller, Post, Res, UploadedFile, UploadedFiles, UseInterceptors } from '@nestjs/common';
import { FileInterceptor, FilesInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { TableExtractionService } from './table-extraction.service';

@Controller('api/extract')
export class TableExtractionController {
  constructor(private readonly extraction: TableExtractionService) {}

  @Post('table')
  @UseInterceptors(FileInterceptor('file', {
    limits: { fileSize: 25 * 1024 * 1024 },
  }))
  async extractTable(
    @UploadedFile() file: Express.Multer.File,
    @Body() selection: Record<string, string>,
  ) {
    if (!file) throw new BadRequestException('An image or PDF file is required.');
    return this.extraction.extract(file, selection);
  }

  @Post('preview')
  @UseInterceptors(FileInterceptor('file', {
    limits: { fileSize: 25 * 1024 * 1024 },
  }))
  async preview(
    @UploadedFile() file: Express.Multer.File,
    @Body('page') requestedPage: string,
    @Res() response: Response,
  ) {
    if (!file) throw new BadRequestException('An image or PDF file is required.');
    const preview = await this.extraction.renderPreview(file, Number(requestedPage) || 1);
    response.set({
      'Content-Type': 'image/png',
      'Content-Length': String(preview.buffer.length),
      'X-Page-Count': String(preview.pageCount),
      'X-Page-Number': String(preview.page),
      'Cache-Control': 'no-store',
    });
    response.end(preview.buffer);
  }

  @Post('bulk')
  @UseInterceptors(FilesInterceptor('files', 30, {
    limits: { fileSize: 80 * 1024 * 1024, files: 30 },
  }))
  async extractBulk(@UploadedFiles() files: Express.Multer.File[]) {
    if (!files?.length) throw new BadRequestException('Select one or more PDF files.');
    return this.extraction.extractBulk(files);
  }

  @Post('bulk/xlsx')
  createBulkWorkbook(@Body() body: { records?: unknown[] }, @Res() response: Response) {
    const output = this.extraction.createBulkWorkbook(body?.records || []);
    response.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="Page_by_Page_Table_JSON.xlsx"',
      'Content-Length': String(output.length),
      'Cache-Control': 'no-store',
    });
    response.end(output);
  }

  @Post('table/xlsx')
  createTableWorkbook(
    @Body() body: { json?: unknown; filename?: string },
    @Res() response: Response,
  ) {
    const output = this.extraction.createTableJsonWorkbook(body?.json);
    const baseName = String(body?.filename || 'Extracted_Table_JSON')
      .replace(/\.[^.]+$/, '')
      .replace(/[^a-z0-9_-]+/gi, '_')
      .replace(/^_+|_+$/g, '') || 'Extracted_Table_JSON';
    response.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${baseName}_Table_JSON.xlsx"`,
      'Content-Length': String(output.length),
      'Cache-Control': 'no-store',
    });
    response.end(output);
  }
}
