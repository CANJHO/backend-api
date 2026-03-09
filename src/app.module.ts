import { Module } from "@nestjs/common";
import { ScheduleModule } from "@nestjs/schedule";
import { ConfigModule } from "@nestjs/config";

import { AuthModule } from "./auth/auth.module";
import { EmpleadosModule } from "./empleados/empleados.module";
import { AsistenciasModule } from "./asistencias/asistencias.module";
import { UsuariosModule } from "./usuarios/usuarios.module";
import { PuntosModule } from "./puntos/puntos.module";
import { HorariosModule } from "./horarios/horarios.module";
import { ReportesModule } from "./reportes/reportes.module";

import { APP_GUARD } from "@nestjs/core";
import { JwtGuard } from "./common/jwt.guard";
import { RolesGuard } from "./common/roles.guard";

import { AppService } from "./app.service";
import { DatabaseModule } from "./database.module";
import { SedesModule } from "./sedes/sedes.module";
import { AreasModule } from "./areas/areas.module";
import { FeriadosModule } from "./feriados/feriados.module";

@Module({
  imports: [
    ScheduleModule.forRoot(),
    ConfigModule.forRoot({ isGlobal: true }),
    DatabaseModule,
    AuthModule,
    EmpleadosModule,
    AsistenciasModule,
    UsuariosModule,
    PuntosModule,
    HorariosModule,
    ReportesModule,
    SedesModule,
    AreasModule,
    FeriadosModule,
  ],
  providers: [
    AppService,

    // 🔐 JwtGuard global: TODAS las rutas requieren token,
    // excepto las marcadas con @Public()
    { provide: APP_GUARD, useClass: JwtGuard },

    // 🎭 RolesGuard global: valida roles donde uses @Roles()
    { provide: APP_GUARD, useClass: RolesGuard },
  ],
})
export class AppModule {}
