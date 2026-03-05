import { Injectable, UnauthorizedException } from '@nestjs/common';
import { JwtService } from '@nestjs/jwt';
import { DataSource } from 'typeorm';
import * as bcrypt from 'bcrypt';
import { ConfigService } from '@nestjs/config';

export type RolSistema = 'Gerencia' | 'RRHH' | 'Empleado';

interface UsuarioRow {
  id: string;
  nombre: string;
  apellido_paterno: string;
  numero_documento: string;
  password_hash: string | null;
  rol: RolSistema | null;
  
}

export interface JwtUsuarioPayload {
  sub: string;
  nombre: string;
  apellido_paterno: string | null;
  rol: RolSistema | null;
  doc: string;
  
 
}

@Injectable()
export class AuthService {
  constructor(
    private ds: DataSource,
    private jwt: JwtService,
    private cfg: ConfigService,
  ) {}

  private getJwtSecret(): string {
    return this.cfg.get<string>('JWT_SECRET') || 'change_me_secret';
  }

    private getJwtExpiresIn(): number {
    const raw = this.cfg.get<string>('JWT_EXPIRES_IN');

    // Si no hay valor → 3600 por defecto (1 hora)
    if (!raw) return 3600;

    // Convertir SIEMPRE a número
    const n = Number(raw);

    // Si la conversión falla → 3600 por defecto
    if (isNaN(n) || n <= 0) return 3600;

    // Devolver en segundos
    return n;
  }


  async login(doc: string, pass: string) {
    const rows: UsuarioRow[] = await this.ds.query(
      `SELECT u.id,
              u.nombre,
              u.apellido_paterno,
              u.numero_documento,
              u.password_hash,
              r.nombre AS rol
         FROM usuarios u
         LEFT JOIN roles r ON r.id = u.rol_id
        WHERE u.numero_documento = $1
        LIMIT 1`,
      [doc],
    );

    const user = rows[0];
    if (!user) {
      throw new UnauthorizedException('Documento o contraseña inválidos');
    }

    if (user.password_hash) {
      const ok = await bcrypt.compare(pass, user.password_hash);
      if (!ok) {
        throw new UnauthorizedException('Documento o contraseña inválidos');
      }
    } else {
      // Usuarios importados sin hash: solo permitimos pass vacío
      if (pass && pass.trim()) {
        throw new UnauthorizedException('Documento o contraseña inválidos');
      }
    }

    const payload: JwtUsuarioPayload = {
      sub: user.id,
      nombre: user.nombre,
      apellido_paterno: user.apellido_paterno,
      rol: (user.rol as RolSistema) || null,
      doc: user.numero_documento,
     
    };

    const token = await this.jwt.signAsync(payload, {
      secret: this.getJwtSecret(),
      // 👇 Forzamos el tipo para que TS no moleste, pero usamos siempre el mismo valor
      expiresIn: this.getJwtExpiresIn() as any,
    });

    return { access_token: token, usuario: payload };
  }

  async verify(token: string) {
    return this.jwt.verifyAsync<JwtUsuarioPayload>(token, {
      secret: this.getJwtSecret(),
    });
  }
}
