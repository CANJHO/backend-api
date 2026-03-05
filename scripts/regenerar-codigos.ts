import { NestFactory } from '@nestjs/core';
import { AppModule } from '../src/app.module';
import { UsuariosService } from '../src/usuarios/usuarios.service';

async function main() {
  const app = await NestFactory.createApplicationContext(AppModule, {
    logger: false,
  });

  const usuarios = app.get(UsuariosService);

  // ✅ IDs que faltan en uploads (los 10 que te salieron)
  const ids = [
    'e272252c-da45-41c9-9234-a79e03ed91c8',
    '06d14d69-7f06-4072-acb1-cae75580d5c9',
    '93a7f3a3-98d9-4895-b9a5-a03da2b53de9',
    'ca47eb20-24e6-4d74-a2d6-9e18984e399a',
    'afc2980b-b146-42f3-aa93-80120151587c',
    '861a4bba-20e7-408d-9ecd-6124ef52f072',
    '15c479d1-7a62-49c5-b845-cfc7e0671725',
    'ae788e15-d285-413e-b6ae-b7fe8fca1d9e',
    'daa1d2e0-c19e-48dc-9189-733a27a5c31d',
    '31f4d758-1575-4e62-8331-cb63a7d33316',
  ];

  for (const id of ids) {
    try {
      await usuarios.generarQR(id);
      await usuarios.generarBarcode(id);
      console.log('OK:', id);
    } catch (e: any) {
      console.error('FAIL:', id, e?.message ?? e);
    }
  }

  await app.close();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
