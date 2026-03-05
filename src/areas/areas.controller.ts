import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  Put,
  Query,
} from '@nestjs/common';
import { AreasService } from './areas.service';
import { Roles } from '../common/roles.decorator';

@Controller('areas')
export class AreasController {
  constructor(private svc: AreasService) {}

  // Listado general (con filtro opcional)
  @Get()
  @Roles('RRHH', 'Gerencia')
  list(@Query('q') q?: string) {
    return this.svc.list(q);
  }

  // Solo activas → para combos
  @Get('activas')
  @Roles('RRHH', 'Gerencia')
  listActivas() {
    return this.svc.listActivas();
  }

  @Get(':id')
  @Roles('RRHH', 'Gerencia')
  get(@Param('id') id: string) {
    return this.svc.get(id);
  }

  @Post()
  @Roles('RRHH','Gerencia') // Solo Gerencia crea áreas
  create(@Body() dto: any) {
    return this.svc.create(dto);
  }

  @Put(':id')
  @Roles('Gerencia','RRHH')
  update(@Param('id') id: string, @Body() dto: any) {
    return this.svc.update(id, dto);
  }

  @Put(':id/desactivar')
  @Roles('Gerencia','RRHH')
  desactivar(@Param('id') id: string) {
    return this.svc.desactivar(id);
  }
}
