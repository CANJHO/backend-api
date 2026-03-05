import { Injectable } from '@nestjs/common';
import { DataSource } from 'typeorm';
import { Request } from 'express';

type JwtUser = {
  sub: string;
  nombre: string;
  apellido_paterno?: string | null;
  rol?: string | null;
  doc?: string;
};

@Injectable()
export class BitacoraService {
  constructor(private ds: DataSource) {}

  private getIp(req: Request): string | null {
    const xfwd = (req.headers['x-forwarded-for'] as string) || '';
    const ip = xfwd.split(',')[0]?.trim() || (req.ip as string) || null;
    return ip || null;
  }

  private usuarioLabel(user: JwtUser): string {
    const nombre = `${user.nombre || ''} ${user.apellido_paterno || ''}`.trim();
    const rol = user.rol || 'SIN_ROL';
    const doc = user.doc || '';
    return `${rol}: ${nombre}${doc ? ' - ' + doc : ''}`.trim();
  }

  async log(req: Request, accion: string, detalle?: any) {
    const user = (req as any).user as JwtUser | undefined;

    const usuario_id = user?.sub || null;
    const usuario = user ? this.usuarioLabel(user) : 'DESCONOCIDO';
    const ip = this.getIp(req);

    const detalleJson = detalle ? JSON.stringify(detalle) : null;

    await this.ds.query(
      `
      INSERT INTO bitacora (usuario, accion, detalle, ip, usuario_id)
      VALUES ($1, $2, $3::jsonb, $4, $5)
      `,
      [usuario, accion, detalleJson, ip, usuario_id],
    );
  }
}
