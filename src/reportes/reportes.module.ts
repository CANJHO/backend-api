import { Module } from "@nestjs/common";
import { ReportesController } from "./reportes.controller";
import { AuthModule } from "../auth/auth.module";
import { JwtGuard } from "../common/jwt.guard";
import { RolesGuard } from "../common/roles.guard";

@Module({
  imports: [AuthModule], // ⬅️ trae AuthService
  controllers: [ReportesController],
  providers: [JwtGuard, RolesGuard], // ⬅️ los guards con dependencias
})
export class ReportesModule {}
