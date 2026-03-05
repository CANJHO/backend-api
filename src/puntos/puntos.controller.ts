import { Body, Controller, Delete, Get, Param, Post, Put, Query } from '@nestjs/common';
import { Roles } from '../common/roles.decorator';
import { PuntosService } from './puntos.service';

@Controller('puntos-trabajo')
export class PuntosController {
  constructor(private svc: PuntosService) {}

  @Get()
  @Roles('RRHH','Gerencia')
  list(@Query('sedeId') sedeId?: string) {
    return this.svc.list(sedeId);
  }

  @Get(':id')
  @Roles('RRHH','Gerencia')
  get(@Param('id') id: string) {
    return this.svc.get(id);
  }

  @Post()
  @Roles('RRHH')
  create(@Body() dto: any) {
    return this.svc.create(dto);
  }

  @Put(':id')
  @Roles('RRHH')
  update(@Param('id') id: string, @Body() dto: any) {
    return this.svc.update(id, dto);
  }

  @Delete(':id')
  @Roles('RRHH')
  delete(@Param('id') id: string) {
    return this.svc.delete(id);
  }

  // ───────────────────────────────────────────────
  // ASIGNACIONES
  // ───────────────────────────────────────────────

  @Get('vigentes/:usuarioId')
  @Roles('RRHH','Gerencia')
  vigentes(@Param('usuarioId') id: string) {
    return this.svc.vigentes(id);
  }

  @Post('asignar')
  @Roles('RRHH')
  asignar(@Body() dto: any) {
    return this.svc.asignar(dto);
  }

  @Put('asignacion/:id/cerrar')
  @Roles('RRHH')
  cerrar(@Param('id') id: string) {
    return this.svc.cambiarEstado(id, 'CERRADA');
  }

  @Put('asignacion/:id/anular')
  @Roles('RRHH')
  anular(@Param('id') id: string) {
    return this.svc.cambiarEstado(id, 'ANULADA');
  }

  @Delete('asignar/:id')
  @Roles('RRHH')
  quitar(@Param('id') id: string) {
    return this.svc.quitar(id);
  }
}
