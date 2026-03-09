import { Body, Controller, Get, Param, Post, Put, Query } from "@nestjs/common";
import { SedesService } from "./sedes.service";
import { Roles } from "../common/roles.decorator";

@Controller("sedes")
export class SedesController {
  constructor(private svc: SedesService) {}

  // Listado general (con búsqueda opcional)
  @Get()
  @Roles("RRHH", "Gerencia")
  list(@Query("q") q?: string) {
    return this.svc.list(q);
  }

  // Solo sedes activas: ideal para combos en el frontend
  @Get("activas")
  @Roles("RRHH", "Gerencia")
  listActivas() {
    return this.svc.listActivas();
  }

  @Get(":id")
  @Roles("RRHH", "Gerencia")
  get(@Param("id") id: string) {
    return this.svc.get(id);
  }

  @Post()
  @Roles("Gerencia", "RRHH")
  create(@Body() dto: any) {
    return this.svc.create(dto);
  }

  @Put(":id")
  @Roles("Gerencia", "RRHH")
  update(@Param("id") id: string, @Body() dto: any) {
    return this.svc.update(id, dto);
  }

  @Post(":id/desactivar")
  @Roles("Gerencia", "RRHH")
  desactivar(@Param("id") id: string) {
    return this.svc.desactivar(id);
  }
}
