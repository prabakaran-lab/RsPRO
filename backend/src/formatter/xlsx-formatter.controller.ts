import { BadRequestException, Controller, Post, Res, UploadedFile, UseInterceptors } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { XlsxFormatterService } from './xlsx-formatter.service';

@Controller('api/format-xlsx')
export class XlsxFormatterController {
  constructor(private readonly formatter: XlsxFormatterService) {}

  @Post('analyze')
  @UseInterceptors(FileInterceptor('file', { limits: { fileSize: 60 * 1024 * 1024 } }))
  analyze(@UploadedFile() file: Express.Multer.File) {
    this.validate(file);
    return this.formatter.analyze(file);
  }

  @Post('convert')
  @UseInterceptors(FileInterceptor('file', { limits: { fileSize: 60 * 1024 * 1024 } }))
  convert(@UploadedFile() file: Express.Multer.File, @Res() response: Response) {
    this.validate(file);
    const output = this.formatter.convert(file);
    const baseName = file.originalname.replace(/\.(xlsx|xls)$/i, '').replace(/[^a-z0-9._-]+/gi, '_');
    response.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${baseName || 'workbook'}_PDF_Ready.xlsx"`,
      'Content-Length': String(output.length),
      'Cache-Control': 'no-store',
    });
    response.end(output);
  }

  private validate(file?: Express.Multer.File) {
    if (!file) throw new BadRequestException('An XLSX or XLS workbook is required.');
    if (!/\.(xlsx|xls)$/i.test(file.originalname)) throw new BadRequestException('Upload an XLSX or XLS file.');
  }
}
