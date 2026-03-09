import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  Put,
  Patch,
  Query,
  UploadedFile,
  UseInterceptors,
} from "@nestjs/common";
import { FileInterceptor } from "@nestjs/platform-express";
import { Roles } from "../common/roles.decorator";
import { UsuariosService } from "./usuarios.service";

@Controller("usuarios")
export class UsuariosController {
  constructor(private svc: UsuariosService) {}

  @Get()
  @Roles("RRHH", "Gerencia")
  list(@Query("q") q?: string) {
    return this.svc.list(q);
  }

  @Get(":id")
  @Roles("RRHH", "Gerencia")
  get(@Param("id") id: string) {
    return this.svc.get(id);
  }

  @Post()
  @Roles("RRHH", "Gerencia")
  create(@Body() dto: any) {
    return this.svc.create(dto);
  }

  @Put(":id")
  @Roles("RRHH", "Gerencia")
  update(@Param("id") id: string, @Body() dto: any) {
    return this.svc.update(id, dto);
  }

  @Post(":id/foto")
  @Roles("RRHH", "Gerencia")
  @UseInterceptors(FileInterceptor("file"))
  upload(@Param("id") id: string, @UploadedFile() file: Express.Multer.File) {
    return this.svc.uploadFoto(id, file.buffer, file.originalname);
  }

  @Post(":id/generar-barcode")
  @Roles("RRHH", "Gerencia")
  genBarcode(@Param("id") id: string) {
    return this.svc.generarBarcode(id);
  }

  @Post(":id/generar-qr")
  @Roles("RRHH", "Gerencia")
  genQR(@Param("id") id: string) {
    return this.svc.generarQR(id);
  }

  @Post("generar-codigos-todos")
  @Roles("RRHH", "Gerencia")
  genCodigosTodos() {
    return this.svc.generarCodigosTodos();
  }
  @Patch(":id/estado")
  @Roles("RRHH", "Gerencia")
  cambiarEstado(@Param("id") id: string, @Body("activo") activo: boolean) {
    return this.svc.cambiarEstado(id, activo);
  }
}
