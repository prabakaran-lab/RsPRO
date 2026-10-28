import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { PrismaModule } from './prisma/prisma.module';
import { StorageModule } from './storage/storage.module';
import { ImageModule } from './image/image.module';
import { JobModule } from './job/job.module';
import { TableExtractionModule } from './extract/table-extraction.module';
import { XlsxFormatterModule } from './formatter/xlsx-formatter.module';

@Module({
  imports: [
    ConfigModule.forRoot({ isGlobal: true }),
    PrismaModule,
    StorageModule,
    ImageModule,
    JobModule,
    TableExtractionModule,
    XlsxFormatterModule,
  ],
})
export class AppModule {}
