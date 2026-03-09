import {
  Injectable,
  NotFoundException,
  ConflictException,
} from "@nestjs/common";
import { DataSource } from "typeorm";
import { join } from "path";
import * as fs from "fs";
import sharp from "sharp";
import * as QRCode from "qrcode";
import * as bcrypt from "bcrypt";
const bwipjs = require("bwip-js");

@Injectable()
export class UsuariosService {
  constructor(private ds: DataSource) {}

  async list(q?: string) {
    if (!q) {
      return this.ds.query(
        `SELECT
            u.id,
            u.nombre,
            u.apellido_paterno,
            u.apellido_materno,
            u.fecha_nacimiento,
            u.numero_documento,
            u.tipo_documento,
            u.sede_id,
            s.nombre  AS sede_nombre,
            u.area_id,
            a.nombre  AS area_nombre,
            u.rol_id,
            r.nombre  AS rol,
            u.activo,
            u.foto_perfil_url,
            u.barcode_url,
            u.qr_url,
            u.email_personal,
            u.email_institucional,
            u.telefono_celular
         FROM usuarios u
         LEFT JOIN roles  r ON r.id = u.rol_id
         LEFT JOIN sedes  s ON s.id = u.sede_id
         LEFT JOIN areas  a ON a.id = u.area_id
        ORDER BY u.created_at DESC`,
      );
    }

    return this.ds.query(
      `SELECT
          u.id,
          u.nombre,
          u.apellido_paterno,
          u.apellido_materno,
          u.fecha_nacimiento,
          u.numero_documento,
          u.tipo_documento,
          u.sede_id,
          s.nombre  AS sede_nombre,
          u.area_id,
          a.nombre  AS area_nombre,
          u.rol_id,
          r.nombre  AS rol,
          u.activo,
          u.foto_perfil_url,
          u.barcode_url,
          u.qr_url,
          u.email_personal,
          u.email_institucional,
          u.telefono_celular
       FROM usuarios u
       LEFT JOIN roles  r ON r.id = u.rol_id
       LEFT JOIN sedes  s ON s.id = u.sede_id
       LEFT JOIN areas  a ON a.id = u.area_id
      WHERE unaccent(
              COALESCE(u.nombre,'') || ' ' ||
              COALESCE(u.apellido_paterno,'') || ' ' ||
              COALESCE(u.apellido_materno,'')
            ) ILIKE unaccent($1)
         OR u.numero_documento ILIKE $1
      ORDER BY u.created_at DESC`,
      [`%${q}%`],
    );
  }

  async get(id: string) {
    const r = await this.ds.query(`SELECT * FROM usuarios WHERE id=$1`, [id]);
    if (!r[0]) throw new NotFoundException("Usuario no existe");
    return r[0];
  }

  async create(dto: any) {
    const plainPassword =
      (dto.password && String(dto.password).trim()) ||
      dto.numero_documento ||
      dto.dni ||
      null;

    if (!plainPassword) {
      throw new Error(
        "No se pudo determinar la contraseña por defecto del usuario.",
      );
    }

    const passwordHash = await bcrypt.hash(String(plainPassword), 10);

    let r: any[];

    try {
      r = await this.ds.query(
        `INSERT INTO usuarios (
          dni,
          nombre,
          apellido_paterno,
          apellido_materno,
          fecha_nacimiento,
          rol_id,
          sede_id,
          activo,
          password_hash,
          tipo_documento,
          numero_documento,
          email_personal,
          email_institucional,
          telefono_celular,
          area_id
        )
        VALUES (
          $1, $2, $3, $4, $5,
          $6, $7,
          COALESCE($8, TRUE),
          $9,
          $10,
          $11,
          $12,
          $13,
          $14,
          $15
        )
        RETURNING id`,
        [
          dto.dni || dto.numero_documento,
          dto.nombre,
          dto.apellido_paterno || null,
          dto.apellido_materno || null,
          dto.fecha_nacimiento || null,
          dto.rol_id,
          dto.sede_id || null,
          dto.activo,
          passwordHash,
          dto.tipo_documento || "DNI",
          dto.numero_documento,
          dto.email_personal || null,
          dto.email_institucional || null,
          dto.telefono_celular || null,
          dto.area_id || null,
        ],
      );
    } catch (e: any) {
      // Unique constraint violada (tipo_documento, numero_documento)
      if (e?.code === "23505") {
        throw new ConflictException(
          "El número de documento ya está registrado.",
        );
      }
      throw e;
    }

    const idNuevo = r[0].id;

    try {
      await this.generarBarcode(idNuevo);
      await this.generarQR(idNuevo);
    } catch (e) {
      console.error("Error generando barcode/QR al crear usuario:", e);
    }

    return this.get(idNuevo);
  }

  async generarCodigosTodos() {
    // Solo los que tienen code_scannable (documento configurado)
    // y aún no tienen barcode o QR
    const usuarios = await this.ds.query(
      `SELECT id
         FROM usuarios
        WHERE code_scannable IS NOT NULL
          AND (barcode_url IS NULL OR qr_url IS NULL)
      `,
    );

    let procesados = 0;
    const errores: { id: string; error: any }[] = [];

    for (const u of usuarios) {
      try {
        await this.generarBarcode(u.id);
        await this.generarQR(u.id);
        procesados++;
      } catch (e) {
        console.error("Error generando códigos para usuario", u.id, e);
        errores.push({ id: u.id, error: String(e) });
      }
    }

    return {
      totalPendientes: usuarios.length,
      procesados,
      errores,
    };
  }

  async update(id: string, dto: any) {
    // 1. Obtener el usuario actual para comparar documento
    const actual = await this.get(id);

    const nuevoTipo =
      dto.tipo_documento !== undefined && dto.tipo_documento !== null
        ? dto.tipo_documento
        : actual.tipo_documento;

    const nuevoNumero =
      dto.numero_documento !== undefined && dto.numero_documento !== null
        ? dto.numero_documento
        : actual.numero_documento;

    // 2. Ver si cambió el documento
    const documentoCambio =
      nuevoTipo !== actual.tipo_documento ||
      nuevoNumero !== actual.numero_documento;

    let nuevoPasswordHash: string | null = null;

    if (documentoCambio) {
      if (!nuevoNumero) {
        throw new Error(
          "No se puede actualizar contraseña: número de documento vacío.",
        );
      }
      nuevoPasswordHash = await bcrypt.hash(String(nuevoNumero), 10);
    }

    // 3. Normalizar fecha_nacimiento: si viene "" -> null
    const fechaNacimientoParam =
      dto.fecha_nacimiento === "" || dto.fecha_nacimiento === undefined
        ? null
        : dto.fecha_nacimiento;

    // 4. Ejecutar el UPDATE
    await this.ds.query(
      `UPDATE usuarios
        SET nombre              = COALESCE($2,  nombre),
            apellido_paterno    = COALESCE($3,  apellido_paterno),
            apellido_materno    = COALESCE($4,  apellido_materno),
            fecha_nacimiento    = COALESCE($5,  fecha_nacimiento),
            rol_id              = COALESCE($6,  rol_id),
            sede_id             = COALESCE($7,  sede_id),
            area_id             = COALESCE($8,  area_id),
            activo              = COALESCE($9,  activo),
            tipo_documento      = COALESCE($10, tipo_documento),
            numero_documento    = COALESCE($11, numero_documento),	
            email_personal      = COALESCE($12, email_personal),
            email_institucional = COALESCE($13, email_institucional),
            telefono_celular    = COALESCE($14, telefono_celular),
            password_hash       = COALESCE($15, password_hash)
      WHERE id = $1`,
      [
        id,
        dto.nombre,
        dto.apellido_paterno,
        dto.apellido_materno,
        fechaNacimientoParam, // 👈 aquí ya NO va ""
        dto.rol_id,
        dto.sede_id,
        dto.area_id,
        dto.activo,
        nuevoTipo,
        nuevoNumero,
        dto.email_personal,
        dto.email_institucional,
        dto.telefono_celular,
        nuevoPasswordHash, // solo si cambió documento
      ],
    );

    // 5. Si cambió el documento, regeneramos códigos
    if (documentoCambio) {
      await this.generarBarcode(id);
      await this.generarQR(id);
    }

    return this.get(id);
  }

  async uploadFoto(id: string, buffer: Buffer, filename: string) {
    const dir = process.env.UPLOAD_DIR || "uploads";
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const out = join(dir, `foto_${id}.jpg`);
    await sharp(buffer)
      .resize(600, 600, { fit: "cover" })
      .jpeg({ quality: 82 })
      .toFile(out);
    const url = `${
      process.env.PUBLIC_BASE_URL || ""
    }/files/${out.split(/[\\/]/).pop()}`;
    await this.ds.query(`UPDATE usuarios SET foto_perfil_url=$2 WHERE id=$1`, [
      id,
      url,
    ]);
    return { foto_perfil_url: url };
  }

  async generarBarcode(id: string) {
    const u = await this.get(id);

    // code_scannable ya lo calcula Postgres
    const code =
      u.code_scannable ||
      (u.tipo_documento === "CE" ? "C" : "D") + u.numero_documento;

    const png = await bwipToBuffer({
      bcid: "code128",
      text: code,
      scale: 3,
      height: 12,
      includetext: true,
    });

    const dir = process.env.UPLOAD_DIR || "uploads";
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    const out = join(dir, `barcode_${id}.png`);
    fs.writeFileSync(out, png);

    const url = `${
      process.env.PUBLIC_BASE_URL || ""
    }/files/${out.split(/[\\/]/).pop()}`;

    // 👇 IMPORTANTE: ya NO intentamos actualizar code_scannable
    await this.ds.query(`UPDATE usuarios SET barcode_url=$2 WHERE id=$1`, [
      id,
      url,
    ]);

    return { barcode_url: url, code_scannable: u.code_scannable || code };
  }

  async generarQR(id: string) {
    const u = await this.get(id);

    const code =
      u.code_scannable ||
      (u.tipo_documento === "CE" ? "C" : "D") + u.numero_documento;

    const dir = process.env.UPLOAD_DIR || "uploads";
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    const out = join(dir, `qr_${id}.png`);
    await QRCode.toFile(out, code, { margin: 1, width: 512 });

    const url = `${
      process.env.PUBLIC_BASE_URL || ""
    }/files/${out.split(/[\\/]/).pop()}`;

    // 👇 Igual: solo actualizamos la URL, no el code_scannable
    await this.ds.query(`UPDATE usuarios SET qr_url=$2 WHERE id=$1`, [id, url]);

    return { qr_url: url, code_scannable: u.code_scannable || code };
  }
  async cambiarEstado(id: string, activo: boolean) {
    if (activo) {
      // Reactivar usuario
      await this.ds.query(
        `UPDATE usuarios
           SET activo = TRUE,
               fecha_baja = NULL,
               motivo_baja = NULL
         WHERE id = $1`,
        [id],
      );
    } else {
      // Dar de baja usuario: guardamos la fecha de baja
      await this.ds.query(
        `UPDATE usuarios
           SET activo = FALSE,
               fecha_baja = now()
         WHERE id = $1`,
        [id],
      );
    }

    // Devolvemos el usuario actualizado
    return this.get(id);
  }
}

function bwipToBuffer(opts: any): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    bwipjs.toBuffer(opts, (err: Error, png: Buffer) => {
      if (err) reject(err);
      else resolve(png);
    });
  });
}
