import {
  CanActivate,
  ExecutionContext,
  ForbiddenException,
  Injectable,
} from "@nestjs/common";
import { Reflector } from "@nestjs/core";
import { ROLES_KEY } from "./roles.decorator";

@Injectable()
export class RolesGuard implements CanActivate {
  constructor(private reflector: Reflector) {}

  canActivate(ctx: ExecutionContext): boolean {
    const rolesRequeridos = this.reflector.getAllAndOverride<string[]>(
      ROLES_KEY,
      [ctx.getHandler(), ctx.getClass()],
    );

    // Si la ruta no pide roles específicos, pasa solo con estar autenticado
    if (!rolesRequeridos || rolesRequeridos.length === 0) return true;

    const req = ctx.switchToHttp().getRequest();
    const user = req.user;

    if (!user || !user.rol) {
      throw new ForbiddenException("No tiene rol asignado");
    }

    if (!rolesRequeridos.includes(user.rol)) {
      throw new ForbiddenException("No tiene permisos para esta operación");
    }

    return true;
  }
}
