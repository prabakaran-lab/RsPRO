import { Module } from '@nestjs/common';
import { JobController } from './job.controller';
import { JobService } from './job.service';
import { PdfCompilerService } from './pdf-compiler.service';

@Module({
  controllers: [JobController],
  providers: [JobService, PdfCompilerService],
})
export class JobModule {}
