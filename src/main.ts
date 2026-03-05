import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { NestExpressApplication } from '@nestjs/platform-express';
import { join } from 'path';

async function bootstrap() {
  const app = await NestFactory.create<NestExpressApplication>(AppModule);
  app.enableCors();
  app.useStaticAssets(join(process.cwd(), process.env.UPLOAD_DIR || 'uploads'), {
    prefix: '/files/',
  });
  await app.listen(process.env.API_PORT || 3000);
}
bootstrap();
