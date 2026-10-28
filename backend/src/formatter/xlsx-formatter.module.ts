import { Module } from '@nestjs/common';
import { XlsxFormatterController } from './xlsx-formatter.controller';
import { XlsxFormatterService } from './xlsx-formatter.service';

@Module({
  controllers: [XlsxFormatterController],
  providers: [XlsxFormatterService],
})
export class XlsxFormatterModule {}
