import { Injectable, NotFoundException } from "@nestjs/common";
import { DataSource } from "typeorm";

@Injectable()
export class SedesService {
  constructor(private ds: DataSource) {}

  async list(q?: string) {
    if (!q) {
      return this.ds.query(
        `SELECT id, nombre, lat, lng, activo, radio_m
           FROM sedes
          ORDER BY nombre ASC`,
      );
    }

    return this.ds.query(
      `SELECT id, nombre, lat, lng, activo, radio_m
         FROM sedes
        WHERE unaccent(nombre) ILIKE unaccent($1)
        ORDER BY nombre ASC`,
      [`%${q}%`],
    );
  }

  // Para combos en el frontend (solo activas, solo lo necesario)
  async listActivas() {
    return this.ds.query(
      `SELECT id, nombre
         FROM sedes
        WHERE activo = true
        ORDER BY nombre ASC`,
    );
  }

  async get(id: string) {
    const rows = await this.ds.query(
      `SELECT id, nombre, lat, lng, activo, radio_m
         FROM sedes
        WHERE id = $1
        LIMIT 1`,
      [id],
    );

    if (!rows[0]) {
      throw new NotFoundException("Sede no existe");
    }

    return rows[0];
  }

  async create(dto: any) {
    const rows = await this.ds.query(
      `INSERT INTO sedes (nombre, lat, lng, activo, radio_m)
       VALUES ($1, $2, $3, COALESCE($4, true), COALESCE($5, 120))
       RETURNING id`,
      [dto.nombre, dto.lat ?? null, dto.lng ?? null, dto.activo, dto.radio_m],
    );

    return this.get(rows[0].id);
  }

  async update(id: string, dto: any) {
    await this.ds.query(
      `UPDATE sedes
          SET nombre  = COALESCE($2, nombre),
              lat     = COALESCE($3, lat),
              lng     = COALESCE($4, lng),
              activo  = COALESCE($5, activo),
              radio_m = COALESCE($6, radio_m)
        WHERE id = $1`,
      [id, dto.nombre, dto.lat, dto.lng, dto.activo, dto.radio_m],
    );

    return this.get(id);
  }

  // Baja lógica: no borramos la sede, solo la desactivamos
  async desactivar(id: string) {
    const r = await this.ds.query(
      `UPDATE sedes
          SET activo = FALSE
        WHERE id = $1
        RETURNING *`,
      [id],
    );

    if (!r[0]) {
      throw new NotFoundException("Sede no existe");
    }

    return r[0];
  }
}
