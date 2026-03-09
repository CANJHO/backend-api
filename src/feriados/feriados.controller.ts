import { Controller, Get, Param, Query } from "@nestjs/common";
import { Roles } from "../common/roles.decorator";
import { FeriadosService } from "./feriados.service";

@Controller("feriados")
export class FeriadosController {
  constructor(private feriados: FeriadosService) {}

  // Listar feriados (opcional)
  @Roles("Gerencia", "RRHH")
  @Get()
  listar(@Query("anio") anio?: string) {
    return this.feriados.listar(anio ? Number(anio) : undefined);
  }

  // Sync por año (ej: /feriados/sync/2026)
  @Roles("Gerencia", "RRHH")
  @Get("sync/:anio")
  sync(@Param("anio") anio: string) {
    return this.feriados.syncPeruYear(Number(anio));
  }
}
