import {
  Injectable,
  BadRequestException,
  NotFoundException,
} from "@nestjs/common";
import { DataSource } from "typeorm";

@Injectable()
export class PuntosService {
  constructor(private ds: DataSource) {}

  // ───────────────────────────────────────────────
  // CRUD DE PUNTOS
  // ───────────────────────────────────────────────

  async list(sedeId?: string) {
    if (!sedeId) {
      return this.ds.query(
        `SELECT * FROM puntos_trabajo ORDER BY created_at DESC`,
      );
    }
    return this.ds.query(
      `SELECT * FROM puntos_trabajo 
        WHERE sede_id = $1
        ORDER BY created_at DESC`,
      [sedeId],
    );
  }

  async get(id: string) {
    const r = await this.ds.query(`SELECT * FROM puntos_trabajo WHERE id=$1`, [
      id,
    ]);
    if (!r[0]) throw new NotFoundException("Punto no existe");
    return r[0];
  }

  async create(dto: any) {
    if (!dto.nombre || !dto.lat || !dto.lng) {
      throw new BadRequestException("Faltan datos obligatorios");
    }

    const res = await this.ds.query(
      `INSERT INTO puntos_trabajo (nombre, lat, lng, radio_m, activo, sede_id)
       VALUES ($1,$2,$3,COALESCE($4,120),COALESCE($5,true),$6)
       RETURNING id`,
      [
        dto.nombre,
        dto.lat,
        dto.lng,
        dto.radio_m,
        dto.activo,
        dto.sede_id || null,
      ],
    );

    return this.get(res[0].id);
  }

  async update(id: string, dto: any) {
    await this.ds.query(
      `UPDATE puntos_trabajo SET
        nombre = COALESCE($2,nombre),
        lat    = COALESCE($3,lat),
        lng    = COALESCE($4,lng),
        radio_m = COALESCE($5,radio_m),
        activo = COALESCE($6,activo),
        sede_id = COALESCE($7,sede_id)
       WHERE id=$1`,
      [id, dto.nombre, dto.lat, dto.lng, dto.radio_m, dto.activo, dto.sede_id],
    );
    return this.get(id);
  }

  async delete(id: string) {
    await this.ds.query(`DELETE FROM puntos_trabajo WHERE id=$1`, [id]);
    return { ok: true };
  }

  // ───────────────────────────────────────────────
  // ASIGNACIONES
  // ───────────────────────────────────────────────

  async asignar(dto: any) {
    const { punto_id, usuario_id, fecha_inicio, fecha_fin, supervisor_id } =
      dto;

    if (!punto_id || !usuario_id || !fecha_inicio || !fecha_fin) {
      throw new BadRequestException("Datos incompletos");
    }

    // validar fechas
    if (new Date(fecha_inicio) >= new Date(fecha_fin)) {
      throw new BadRequestException("Rango de fechas inválido");
    }

    // validar que el punto exista y esté activo
    const punto = await this.get(punto_id);
    if (!punto.activo) {
      throw new BadRequestException("Punto inactivo");
    }

    // validar que no haya solapamientos vigentes
    const solap = await this.ds.query(
      `SELECT id
         FROM asignaciones_punto
        WHERE usuario_id = $1
          AND estado = 'VIGENTE'
          AND (
              ($2 BETWEEN fecha_inicio AND fecha_fin)
              OR
              ($3 BETWEEN fecha_inicio AND fecha_fin)
              OR
              (fecha_inicio BETWEEN $2 AND $3)
          )`,
      [usuario_id, fecha_inicio, fecha_fin],
    );

    if (solap.length > 0) {
      throw new BadRequestException(
        "El usuario ya tiene un punto asignado en este rango",
      );
    }

    // insertar asignación
    return this.ds.query(
      `INSERT INTO asignaciones_punto (punto_id, usuario_id, fecha_inicio, fecha_fin, supervisor_id, estado)
       VALUES ($1,$2,$3,$4,$5,'VIGENTE')
       RETURNING id`,
      [punto_id, usuario_id, fecha_inicio, fecha_fin, supervisor_id || null],
    );
  }

  async quitar(id: string) {
    await this.ds.query(`DELETE FROM asignaciones_punto WHERE id=$1`, [id]);
    return { ok: true };
  }

  // Cambia estado de asignación
  async cambiarEstado(id: string, estado: "VIGENTE" | "CERRADA" | "ANULADA") {
    await this.ds.query(`UPDATE asignaciones_punto SET estado=$2 WHERE id=$1`, [
      id,
      estado,
    ]);
    return { ok: true };
  }

  // lista vigentes correctamente
  async vigentes(usuarioId: string) {
    return this.ds.query(
      `SELECT ap.id, ap.punto_id, pt.nombre, pt.lat, pt.lng, pt.radio_m, 
              ap.fecha_inicio, ap.fecha_fin
         FROM asignaciones_punto ap
         JOIN puntos_trabajo pt ON pt.id=ap.punto_id
        WHERE ap.usuario_id=$1
          AND ap.estado='VIGENTE'
          AND NOW() BETWEEN ap.fecha_inicio AND ap.fecha_fin
        ORDER BY ap.fecha_inicio DESC`,
      [usuarioId],
    );
  }
}
