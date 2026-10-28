import { Module } from '@nestjs/common';
import { TableExtractionController } from './table-extraction.controller';
import { TableExtractionService } from './table-extraction.service';

@Module({
  controllers: [TableExtractionController],
  providers: [TableExtractionService],
})
export class TableExtractionModule {}
