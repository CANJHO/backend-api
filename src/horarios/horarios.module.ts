import { Module } from "@nestjs/common";
import { HorariosService } from "./horarios.service";
import { HorariosController } from "./horarios.controller";

@Module({
  providers: [HorariosService],
  controllers: [HorariosController],
  exports: [HorariosService],
})
export class HorariosModule {}
