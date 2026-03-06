import { Injectable, BadRequestException, NotFoundException } from '@nestjs/common';
import { DataSource } from 'typeorm';

function fechaLimaISO(d: Date = new Date()) {
  return d.toLocaleDateString('en-CA', { timeZone: 'America/Lima' }); // YYYY-MM-DD
}

@Injectable()
export class HorariosService {
  constructor(private ds: DataSource) {}

  // Normaliza fecha
  private toDate(fecha?: string) {
      if (!fecha) return new Date();

  // Evita interpretación UTC
  const [y, m, d] = fecha.split('-').map(Number);
  return new Date(y, m - 1, d); // fecha local real

  }

  // ───────────────────────────────────────────────
  // 1. OBTENER HORARIO DEL DÍA (Para tardanzas)
  // ───────────────────────────────────────────────
  async getHorarioDelDia(usuarioId: string, fecha?: string) {
    const f = this.toDate(fecha);
    const fechaStr = fecha || fechaLimaISO(f); // "YYYY-MM-DD"
	const d = new Date(`${fechaStr}T00:00:00-05:00`);
	const jsDow = d.getDay(); // 0=Dom,1=Lun,...6=Sab
	const diaSemana = ((jsDow + 6) % 7) + 1; // 1=Lun ... 7=Dom

	const exc = await this.ds.query(
      `SELECT * 
         FROM usuario_excepciones
        WHERE usuario_id = $1 AND fecha = $2`,
      [usuarioId, fechaLimaISO(f)]
    );

    const excepcion = exc[0] || null;

    // HORARIO VIGENTE
    const rows = await this.ds.query(
      `SELECT *
         FROM usuario_horarios
        WHERE usuario_id = $1
          AND dia_semana = $2
          AND fecha_inicio <= $3::date
          AND (fecha_fin IS NULL OR fecha_fin >= $3::date)
        ORDER BY creado_en DESC
        LIMIT 1`,
      [usuarioId, diaSemana, fechaLimaISO(f)]
    );

    const horario = rows[0] || null;

    return {
      fecha: fechaLimaISO(f),
      dia_semana: diaSemana,
      horario,
      excepcion
    };
  }

  // ───────────────────────────────────────────────
  // 2. HORARIOS VIGENTES EN UNA FECHA
  // ───────────────────────────────────────────────
  async getVigentes(usuarioId: string, fecha?: string) {
    const f = this.toDate(fecha);
    return this.ds.query(
      `SELECT *
         FROM usuario_horarios
        WHERE usuario_id = $1
          AND fecha_inicio <= $2::date
          AND (fecha_fin IS NULL OR fecha_fin >= $2::date)
        ORDER BY dia_semana`,
      [usuarioId, fechaLimaISO(f)]
    );
  }

  // ───────────────────────────────────────────────
  // 3. HISTORIAL COMPLETO DE HORARIOS
  // ───────────────────────────────────────────────
  historial(usuarioId: string) {
    return this.ds.query(
      `SELECT *
         FROM usuario_horarios
        WHERE usuario_id = $1
        ORDER BY fecha_inicio DESC, dia_semana`,
      [usuarioId]
    );
  }

// ───────────────────────────────────────────────
// 4. NUEVA SEMANA (CREA 7 DIAS DE HORARIO)
    // ───────────────────────────────────────────────
    async setSemana(usuarioId: string, dto: any) {
      const fi = dto.fecha_inicio || fechaLimaISO();
      const items = dto.items || [];

      if (items.length !== 7) {
        throw new BadRequestException('Debes enviar 7 días de horario.');
      }

      let diasLaborables = 0;

      for (const it of items) {
        const {
          dia,
          hora_inicio,
          hora_fin,
          hora_inicio_2,
          hora_fin_2,
          es_descanso,
        } = it;

        if (dia == null) {
          throw new BadRequestException('Cada item debe indicar el día (1..7).');
        }

        if (!es_descanso) {
          diasLaborables++;

          const t1i = hora_inicio;
          const t1f = hora_fin;
          const t2i = hora_inicio_2;
          const t2f = hora_fin_2;

          if ((t1i && !t1f) || (!t1i && t1f)) {
            throw new BadRequestException(
              `En el día ${dia}, si defines el Turno 1 debes indicar hora de inicio y fin.`,
            );
          }

          if ((t2i && !t2f) || (!t2i && t2f)) {
            throw new BadRequestException(
              `En el día ${dia}, si defines el Turno 2 debes indicar hora de inicio y fin.`,
            );
          }

          if (t1i && t1f && t1i >= t1f) {
            throw new BadRequestException(
              `En el día ${dia}, la hora de inicio del Turno 1 debe ser menor que la hora de fin.`,
            );
          }

          if (t2i && t2f && t2i >= t2f) {
            throw new BadRequestException(
              `En el día ${dia}, la hora de inicio del Turno 2 debe ser menor que la hora de fin.`,
            );
          }

          if (!t1i && !t1f && !t2i && !t2f) {
            throw new BadRequestException(
              `En el día ${dia}, configura al menos un turno o márcalo como descanso.`,
            );
          }
        }
      }

      if (diasLaborables === 0) {
        throw new BadRequestException(
          'El horario no puede ser solo descansos. Configura al menos un día laborable.',
        );
      }

      // 🔎 Verificar si ya existe un bloque con esa misma fecha_inicio
      const existeMismaFecha = await this.ds.query(
        `SELECT id
          FROM usuario_horarios
          WHERE usuario_id = $1
            AND fecha_inicio = $2::date`,
        [usuarioId, fi],
      );

      if (existeMismaFecha.length > 0) {
        // 🧹 Si existe, eliminar ese bloque completo
        await this.ds.query(
          `DELETE FROM usuario_horarios
            WHERE usuario_id = $1
              AND fecha_inicio = $2::date`,
          [usuarioId, fi],
        );
      } else {
        // 🔒 Si es una nueva vigencia, cerrar la anterior correctamente
        await this.ds.query(
          `UPDATE usuario_horarios
              SET fecha_fin = ($2::date - INTERVAL '1 day')::date
            WHERE usuario_id = $1
              AND fecha_fin IS NULL`,
          [usuarioId, fi],
        );
      }

      // ➜ Insertar nueva semana limpia
      for (const it of items) {
        const {
          dia,
          hora_inicio,
          hora_fin,
          hora_inicio_2,
          hora_fin_2,
          es_descanso,
          tolerancia_min,
        } = it;

        await this.ds.query(
          `INSERT INTO usuario_horarios
            (usuario_id, dia_semana, hora_inicio, hora_fin,
            hora_inicio_2, hora_fin_2,
            es_descanso, tolerancia_min, fecha_inicio)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
          [
            usuarioId,
            dia,
            es_descanso ? null : hora_inicio,
            es_descanso ? null : hora_fin,
            es_descanso ? null : (hora_inicio_2 || null),
            es_descanso ? null : (hora_fin_2 || null),
            !!es_descanso,
            tolerancia_min || 15,
            fi,
          ],
        );
      }

      return { ok: true };
    }



  // ───────────────────────────────────────────────
  // 5. CERRAR HORARIO
  // ───────────────────────────────────────────────
  async cerrarVigencia(usuarioId: string, fecha_fin: string) {
    await this.ds.query(
      `UPDATE usuario_horarios
          SET fecha_fin = $2
        WHERE usuario_id = $1 AND fecha_fin IS NULL`,
      [usuarioId, fecha_fin]
    );
    return { ok: true };
  }

  // ───────────────────────────────────────────────
  // 6. EXCEPCIONES
  // ───────────────────────────────────────────────
  async addExcepcion(usuarioId: string, e: any) {
    if (!e.fecha || !e.tipo) {
      throw new BadRequestException('Falta fecha o tipo');
    }

    // evitar duplicados
    const exists = await this.ds.query(
      `SELECT id
         FROM usuario_excepciones
        WHERE usuario_id = $1 AND fecha = $2`,
      [usuarioId, e.fecha]
    );

    if (exists.length) {
      throw new BadRequestException('Ya existe excepción para esta fecha');
    }

    await this.ds.query(
      `INSERT INTO usuario_excepciones
        (usuario_id, fecha, tipo, es_laborable, hora_inicio, hora_fin, observacion)
       VALUES ($1,$2,$3,$4,$5,$6,$7)`,
      [
        usuarioId,
        e.fecha,
        e.tipo,
        e.es_laborable,
        e.hora_inicio || null,
        e.hora_fin || null,
        e.observacion || null
      ]
    );

    return { ok: true };
  }

  async eliminarExcepcion(id: string) {
    await this.ds.query(
      `DELETE FROM usuario_excepciones WHERE id=$1`,
      [id]
    );
    return { ok: true };
  }
}
