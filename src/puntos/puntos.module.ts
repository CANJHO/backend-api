import { Module } from '@nestjs/common';
import { PuntosService } from './puntos.service';
import { PuntosController } from './puntos.controller';

@Module({ providers:[PuntosService], controllers:[PuntosController] })
export class PuntosModule {}
