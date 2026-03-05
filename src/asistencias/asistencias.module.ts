import { Module } from '@nestjs/common';
import { AsistenciasController } from './asistencias.controller';
import { AsistenciasAdminController } from './asistencias.admin.controller';
import { AsistenciasService } from './asistencias.service';
import { HorariosModule } from '../horarios/horarios.module';
import { BitacoraService } from '../common/bitacora.service';

@Module({
  imports: [HorariosModule],
  controllers: [AsistenciasController, AsistenciasAdminController],
  providers: [AsistenciasService, BitacoraService],
})
export class AsistenciasModule {}


