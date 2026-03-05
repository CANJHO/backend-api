import {
  BadRequestException,
  Controller,
  Get,
  Post,
  Query,
  Param,
  UseInterceptors,
  UploadedFile,
  ParseUUIDPipe,
  Res,
} from '@nestjs/common';
import { EmpleadosService } from './empleados.service';
import { FileInterceptor } from '@nestjs/platform-express';
import type { Response } from 'express';

@Controller('empleados')
export class EmpleadosController {
  constructor(private readonly svc: EmpleadosService) {}

  // ✅ NUEVO: Cumpleaños próximos (para modal + alertas)
  @Get('cumpleanos-proximos')
  async cumpleanosProximos(@Query('dias') dias = '5') {
    const n = Number(dias);
    const diasSafe = !n || n < 0 ? 5 : n > 30 ? 30 : n;
    return this.svc.listarCumpleanosProximos(diasSafe);
  }

  // ====== Listado de empleados (para el panel) ======
  @Get()
  async listar(
    @Query('pagina') pagina = '1',
    @Query('limite') limite = '20',
    @Query('buscar') buscar?: string,
  ) {
    const pag = Number(pagina) || 1;
    const lim = Number(limite) || 20;
    return this.svc.listarEmpleados(pag, lim, buscar);
  }

  // ====== Carnet PDF del empleado ======
  @Get(':id/carnet-pdf')
  async carnetPdf(
    @Param('id', new ParseUUIDPipe({ version: '4' })) id: string,
    @Res() res: Response,
  ) {
    const pdfBuffer = await this.svc.generarCarnetPdf(id);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      `inline; filename="carnet-empleado-${id}.pdf"`,
    );

    res.send(pdfBuffer);
  }

  // ====== Ficha del empleado ======
  @Get(':id/ficha')
  async ficha(
    @Param('id', new ParseUUIDPipe({ version: '4' })) id: string,
  ) {
    return this.svc.obtenerFichaEmpleado(id);
  }

  // ====== Kiosko: buscar empleado por código escaneable ======
  @Get('lookup')
  async lookup(@Query('code') code?: string) {
    if (!code) throw new BadRequestException('Falta el parámetro ?code');
    return this.svc.lookupByCode(code);
  }

  // ====== Subir / actualizar foto de empleado ======
  @Post(':id/foto')
  @UseInterceptors(FileInterceptor('foto'))
  async subirFoto(
    @Param('id', new ParseUUIDPipe({ version: '4' })) id: string,
    @UploadedFile() archivo?: Express.Multer.File,
  ) {
    if (!archivo) {
      throw new BadRequestException('No se recibió archivo "foto".');
    }

    return this.svc.actualizarFotoPerfil(id, archivo);
  }
}
