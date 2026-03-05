import { Module } from '@nestjs/common';
import { HttpModule } from '@nestjs/axios';
import { FeriadosController } from './feriados.controller';
import { FeriadosService } from './feriados.service';

@Module({
  imports: [HttpModule],
  controllers: [FeriadosController],
  providers: [FeriadosService],
})
export class FeriadosModule {}
