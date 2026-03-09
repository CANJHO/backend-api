import { Body, Controller, Post } from "@nestjs/common";
import { AsistenciasService } from "./asistencias.service";
import { Public } from "../common/public.decorator";

@Controller("asistencias")
export class AsistenciasController {
  constructor(private readonly svc: AsistenciasService) {}

  @Public()
  @Post("validar-geo")
  validarGeo(@Body() dto: { usuarioId: string; lat?: number; lng?: number }) {
    return this.svc.validarGeo(dto.usuarioId, dto.lat, dto.lng);
  }

  @Public()
  @Post("marcar")
  marcar(
    @Body()
    dto: {
      usuarioId: string;
      tipo: "IN" | "OUT";
      metodo:
        | "scanner_barras"
        | "qr_fijo"
        | "qr_dinamico"
        | "manual_supervisor";
      lat?: number;
      lng?: number;
      evidenciaUrl?: string;
      deviceId?: string;
    },
  ) {
    return this.svc.marcar(dto);
  }
  // ✅ Automático (nuevo)
  @Public()
  @Post("marcar-auto")
  marcarAuto(@Body() dto: { identificador: string }) {
    return this.svc.marcarAutoDesdeKiosko(dto.identificador);
  }
}
