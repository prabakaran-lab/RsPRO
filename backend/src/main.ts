import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import * as express from 'express';
import * as path from 'path';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);
  
  // Enable CORS for frontend connection
  app.enableCors({
    origin: '*',
    methods: 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS',
    credentials: true,
    exposedHeaders: ['X-Page-Count', 'X-Page-Number', 'Content-Disposition'],
  });

  app.use(express.json({ limit: '15mb' }));

  // Serve the static uploads directory under /static
  const staticPath = path.join(__dirname, '..', '..', 'static');
  app.use('/static', express.static(staticPath));
  console.log(`Serving static files from: ${staticPath}`);

  const port = process.env.PORT || 8001;
  await app.listen(port);
  console.log(`NestJS Backend server running on: http://localhost:${port}`);
}
bootstrap();
