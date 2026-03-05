import { Injectable, NotFoundException } from '@nestjs/common';
import { DataSource } from 'typeorm';

@Injectable()
export class AreasService {
  constructor(private ds: DataSource) {}

  async list(q?: string) {
    if (!q) {
      return this.ds.query(
        `SELECT id, nombre, activo
           FROM areas
          ORDER BY nombre ASC`,
      );
    }

    return this.ds.query(
      `SELECT id, nombre, activo
         FROM areas
        WHERE unaccent(nombre) ILIKE unaccent($1)
        ORDER BY nombre ASC`,
      [`%${q}%`],
    );
  }

  // Para combos en el frontend (solo activas)
  async listActivas() {
    return this.ds.query(
      `SELECT id, nombre
         FROM areas
        WHERE activo = true
        ORDER BY nombre ASC`,
    );
  }

  async get(id: string) {
    const rows = await this.ds.query(
      `SELECT id, nombre, activo
         FROM areas
        WHERE id = $1
        LIMIT 1`,
      [id],
    );

    if (!rows[0]) {
      throw new NotFoundException('Área no existe');
    }

    return rows[0];
  }

  async create(dto: any) {
    const rows = await this.ds.query(
      `INSERT INTO areas (nombre, activo)
       VALUES ($1, COALESCE($2, true))
       RETURNING id`,
      [
        dto.nombre,
        dto.activo,
      ],
    );

    return this.get(rows[0].id);
  }

  async update(id: string, dto: any) {
    await this.ds.query(
      `UPDATE areas
          SET nombre = COALESCE($2, nombre),
              activo = COALESCE($3, activo)
        WHERE id = $1`,
      [
        id,
        dto.nombre,
        dto.activo,
      ],
    );

    return this.get(id);
  }

  // Baja lógica
  async desactivar(id: string) {
    await this.ds.query(
      `UPDATE areas
          SET activo = false
        WHERE id = $1`,
      [id],
    );

    return { id, activo: false };
  }
}
