import {
  CanActivate,
  ExecutionContext,
  Injectable,
  UnauthorizedException,
} from "@nestjs/common";
import { Reflector } from "@nestjs/core";
import { AuthService } from "../auth/auth.service";
import { IS_PUBLIC_KEY } from "./public.decorator";

@Injectable()
export class JwtGuard implements CanActivate {
  constructor(
    private auth: AuthService,
    private reflector: Reflector,
  ) {}

  async canActivate(ctx: ExecutionContext): Promise<boolean> {
    // 1️⃣ Rutas marcadas con @Public() se saltan el guard
    const isPublic = this.reflector.getAllAndOverride<boolean>(IS_PUBLIC_KEY, [
      ctx.getHandler(),
      ctx.getClass(),
    ]);
    if (isPublic) return true;

    const req = ctx.switchToHttp().getRequest();
    const authHeader = (req.headers["authorization"] || "") as string;

    if (!authHeader.startsWith("Bearer ")) {
      throw new UnauthorizedException("Falta token");
    }

    const token = authHeader.slice("Bearer ".length).trim();

    try {
      const payload = await this.auth.verify(token);
      // Guardamos el usuario en la request para usarlo en controladores
      req.user = payload;
      return true;
    } catch {
      throw new UnauthorizedException("Token inválido o expirado");
    }
  }
}
