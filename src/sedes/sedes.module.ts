import { Module } from "@nestjs/common";
import { SedesService } from "./sedes.service";
import { SedesController } from "./sedes.controller";

@Module({
  providers: [SedesService],
  controllers: [SedesController],
  exports: [SedesService],
})
export class SedesModule {}
