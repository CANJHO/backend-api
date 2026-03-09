import { Global, Module } from "@nestjs/common";
import { JwtModule } from "@nestjs/jwt";
import { ConfigModule, ConfigService } from "@nestjs/config";
import { AuthService } from "./auth.service";
import { AuthController } from "./auth.controller";

@Global() // ⬅⬅⬅ IMPORTANTE: Hace el servicio global en toda la app
@Module({
  imports: [
    ConfigModule, // ⬅⬅⬅ Necesario porque JwtModule usa ConfigService
    JwtModule.registerAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (cfg: ConfigService) => ({
        secret: cfg.get<string>("JWT_SECRET") || "change_me_secret",
        signOptions: {
          expiresIn: parseInt(cfg.get<string>("JWT_EXPIRES_IN") ?? "3600"), // ⬅⬅⬅ 1 hora por defecto
        },
      }),
    }),
  ],
  providers: [AuthService],
  controllers: [AuthController],
  exports: [AuthService], // ⬅⬅⬅ JwtGuard puede usarlo globalmente
})
export class AuthModule {}
