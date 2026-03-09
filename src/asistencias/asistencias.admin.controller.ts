import {
  BadRequestException,
  Body,
  Controller,
  Get,
  Param,
  Post,
  Put,
  Query,
  Req,
  UseGuards,
} from "@nestjs/common";
import { DataSource } from "typeorm";
import type { Request } from "express";
import { JwtGuard } from "../common/jwt.guard";
import { Roles } from "../common/roles.decorator";
import { RolesGuard } from "../common/roles.guard";
import { BitacoraService } from "../common/bitacora.service";

type Evento = "JORNADA_IN" | "REFRIGERIO_OUT" | "REFRIGERIO_IN" | "JORNADA_OUT";

@UseGuards(JwtGuard, RolesGuard)
@Controller("asistencias-admin")
export class AsistenciasAdminController {
  constructor(
    private ds: DataSource,
    private bitacora: BitacoraService,
  ) {}

  private tipoPorEvento(evento: Evento): "IN" | "OUT" {
    if (evento === "JORNADA_IN" || evento === "REFRIGERIO_IN") return "IN";
    return "OUT";
  }

  private ensureMotivo(motivo?: string) {
    if (!motivo || !motivo.trim()) {
      throw new BadRequestException("Motivo es obligatorio");
    }
  }

  private buildFechaHora(fecha: string, hora: string): string {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
      throw new BadRequestException("Fecha inválida. Use YYYY-MM-DD");
    }
    if (!/^\d{2}:\d{2}(:\d{2})?$/.test(hora)) {
      throw new BadRequestException("Hora inválida. Use HH:mm o HH:mm:ss");
    }
    const hh = hora.length === 5 ? `${hora}:00` : hora;
    return `${fecha} ${hh}`;
  }

  @Get()
  @Roles("RRHH", "Gerencia")
  list(
    @Query()
    q: {
      usuarioId?: string;
      desde?: string;
      hasta?: string;
      estado?: string;
    },
  ) {
    const wh: string[] = [];
    const p: any[] = [];
    if (q.usuarioId) {
      p.push(q.usuarioId);
      wh.push(`a.usuario_id = $${p.length}`);
    }
    if (q.desde) {
      p.push(q.desde);
      wh.push(`a.fecha_hora >= $${p.length}::timestamp`);
    }
    if (q.hasta) {
      p.push(q.hasta);
      wh.push(`a.fecha_hora <  ($${p.length}::timestamp + interval '1 day')`);
    }
    if (q.estado) {
      p.push(q.estado);
      wh.push(`a.estado_validacion = $${p.length}`);
    }
    const where = wh.length ? `WHERE ${wh.join(" AND ")}` : "";
    return this.ds.query(
      `SELECT a.*,
        u.nombre,
        u.apellido_paterno,
        u.apellido_materno,
        u.numero_documento,
        (u.nombre || ' ' ||
         COALESCE(u.apellido_paterno,'') || ' ' ||
         COALESCE(u.apellido_materno,'')) AS nombre_completo
          FROM asistencias a
          JOIN usuarios u ON u.id = a.usuario_id
          ${where}
          ORDER BY a.fecha_hora DESC`,
      p,
    );
  }

  @Get("timeline")
  @Roles("RRHH", "Gerencia")
  async timeline(
    @Query()
    q: {
      usuarioId?: string;
      fecha?: string;
    },
  ) {
    if (!q.usuarioId) throw new BadRequestException("usuarioId es obligatorio");
    if (!q.fecha)
      throw new BadRequestException("fecha es obligatorio (YYYY-MM-DD)");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(q.fecha)) {
      throw new BadRequestException("fecha inválida. Use YYYY-MM-DD");
    }

    const emp = await this.ds.query(
      `
      SELECT
        u.id,
        u.nombre,
        u.apellido_paterno,
        u.apellido_materno,
        u.numero_documento,
        (u.nombre || ' ' ||
         COALESCE(u.apellido_paterno,'') || ' ' ||
         COALESCE(u.apellido_materno,'')) AS nombre_completo
      FROM usuarios u
      WHERE u.id = $1
      LIMIT 1
      `,
      [q.usuarioId],
    );

    if (!emp || emp.length === 0) {
      throw new BadRequestException("No existe usuario con ese usuarioId");
    }

    const rows = await this.ds.query(
      `
      SELECT a.*
      FROM asistencias a
      WHERE a.usuario_id = $1
        AND a.fecha_hora >= ($2::date)::timestamp
        AND a.fecha_hora <  (($2::date + interval '1 day'))::timestamp
      ORDER BY a.fecha_hora ASC
      `,
      [q.usuarioId, q.fecha],
    );

    return {
      empleado: emp[0],
      timeline: rows,
    };
  }

  @Post("manual")
  @Roles("RRHH")
  async manual(
    @Req() req: Request,
    @Body()
    body: {
      usuarioId: string;
      fecha: string;
      hora: string;
      evento: Evento;
      motivo: string;
      evidencia?: any;
    },
  ) {
    if (!body?.usuarioId)
      throw new BadRequestException("usuarioId es obligatorio");
    if (!body?.evento) throw new BadRequestException("evento es obligatorio");
    this.ensureMotivo(body?.motivo);

    const fecha_hora = this.buildFechaHora(body.fecha, body.hora);
    const tipo = this.tipoPorEvento(body.evento);
    const metodo = "manual_supervisor";
    const estado_validacion = "aprobado";

    try {
      const inserted = await this.ds.query(
        `
        INSERT INTO asistencias (usuario_id, fecha_hora, evento, tipo, metodo, estado_validacion)
        VALUES ($1, $2::timestamp, $3, $4, $5, $6)
        RETURNING id
        `,
        [
          body.usuarioId,
          fecha_hora,
          body.evento,
          tipo,
          metodo,
          estado_validacion,
        ],
      );

      const asistenciaId = inserted?.[0]?.id;

      await this.bitacora.log(req, "ASISTENCIA_MANUAL_CREAR", {
        asistenciaId,
        usuarioId: body.usuarioId,
        fecha_hora,
        evento: body.evento,
        tipo,
        metodo,
        estado_validacion,
        motivo: body.motivo,
        evidencia: body.evidencia ?? null,
      });

      return { ok: true, id: asistenciaId };
    } catch (e: any) {
      if (e?.code === "23505") {
        throw new BadRequestException(
          "Ya existe un marcaje para ese usuario, fecha/hora y evento (duplicado).",
        );
      }
      throw e;
    }
  }

  @Put(":id/anular")
  @Roles("RRHH")
  async anular(
    @Req() req: Request,
    @Param("id") id: string,
    @Body() body: { motivo: string; evidencia?: any },
  ) {
    this.ensureMotivo(body?.motivo);

    const updated = await this.ds.query(
      `
      UPDATE asistencias
      SET estado_validacion = 'rechazado'
      WHERE id = $1
      RETURNING id, usuario_id, fecha_hora, evento, tipo, metodo, estado_validacion
      `,
      [id],
    );

    if (!updated || updated.length === 0) {
      throw new BadRequestException("No existe asistencia con ese id");
    }

    await this.bitacora.log(req, "ASISTENCIA_ANULAR", {
      asistencia: updated[0],
      motivo: body.motivo,
      evidencia: body.evidencia ?? null,
    });

    return { ok: true };
  }

  @Put(":id/aprobar")
  @Roles("RRHH")
  async aprobar(
    @Req() req: Request,
    @Param("id") id: string,
    @Body() body: { motivo?: string },
  ) {
    const updated = await this.ds.query(
      `
      UPDATE asistencias
      SET estado_validacion='aprobado'
      WHERE id=$1
      RETURNING id, usuario_id, fecha_hora, evento, tipo, metodo, estado_validacion
      `,
      [id],
    );

    if (!updated || updated.length === 0) {
      throw new BadRequestException("No existe asistencia con ese id");
    }

    await this.bitacora.log(req, "ASISTENCIA_APROBAR", {
      asistencia: updated[0],
      motivo: body?.motivo?.trim() || null,
    });

    return { ok: true };
  }

  @Put(":id/rechazar")
  @Roles("RRHH")
  async rechazar(
    @Req() req: Request,
    @Param("id") id: string,
    @Body() body: { motivo?: string },
  ) {
    this.ensureMotivo(body?.motivo);

    const updated = await this.ds.query(
      `
      UPDATE asistencias
      SET estado_validacion='rechazado'
      WHERE id=$1
      RETURNING id, usuario_id, fecha_hora, evento, tipo, metodo, estado_validacion
      `,
      [id],
    );

    if (!updated || updated.length === 0) {
      throw new BadRequestException("No existe asistencia con ese id");
    }

    await this.bitacora.log(req, "ASISTENCIA_RECHAZAR", {
      asistencia: updated[0],
      motivo: body.motivo,
    });

    return { ok: true };
  }

  // ✅ NUEVO: pendiente más antiguo (1 usuario) - SOLO días anteriores (bloquea kiosko)
  @Get("pendiente")
  @Roles("RRHH", "Gerencia")
  async pendiente(@Query() q: { usuarioId?: string }) {
    if (!q.usuarioId) throw new BadRequestException("usuarioId es obligatorio");

    const rows = await this.ds.query(
      `
    WITH dias AS (
      SELECT
        a.usuario_id,
        (a.fecha_hora::date) AS fecha,
        BOOL_OR(a.evento = 'JORNADA_IN'  AND a.estado_validacion <> 'rechazado') AS tiene_in,
        BOOL_OR(a.evento = 'JORNADA_OUT' AND a.estado_validacion <> 'rechazado') AS tiene_out
      FROM asistencias a
      WHERE a.usuario_id = $1::uuid
      GROUP BY a.usuario_id, (a.fecha_hora::date)
    )
    SELECT usuario_id, fecha AS fecha_pendiente
    FROM dias
    WHERE tiene_in = true
      AND tiene_out = false
      AND fecha < CURRENT_DATE          -- ✅ EXCLUYE HOY
    ORDER BY fecha ASC
    LIMIT 1
    `,
      [q.usuarioId],
    );

    return {
      usuario_id: q.usuarioId,
      fecha_pendiente: rows?.[0]?.fecha_pendiente ?? null,
    };
  }

  // ✅ NUEVO: pendientes más antiguos (muchos usuarios) - SOLO días anteriores (bloquea kiosko)
  @Get("pendientes")
  @Roles("RRHH", "Gerencia")
  async pendientes(@Query() q: { usuarioIds?: string }) {
    const raw = (q.usuarioIds || "").trim();
    if (!raw) return [];

    const ids = raw
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    if (ids.length === 0) return [];

    const rows = await this.ds.query(
      `
    WITH dias AS (
      SELECT
        a.usuario_id,
        (a.fecha_hora::date) AS fecha,
        BOOL_OR(a.evento = 'JORNADA_IN'  AND a.estado_validacion <> 'rechazado') AS tiene_in,
        BOOL_OR(a.evento = 'JORNADA_OUT' AND a.estado_validacion <> 'rechazado') AS tiene_out
      FROM asistencias a
      WHERE a.usuario_id = ANY($1::uuid[])
      GROUP BY a.usuario_id, (a.fecha_hora::date)
    ),
    pendientes AS (
      SELECT
        usuario_id,
        MIN(fecha) AS fecha_pendiente
      FROM dias
      WHERE tiene_in = true
        AND tiene_out = false
        AND fecha < CURRENT_DATE         -- ✅ EXCLUYE HOY
      GROUP BY usuario_id
    )
    SELECT usuario_id, fecha_pendiente
    FROM pendientes
    ORDER BY fecha_pendiente ASC
    `,
      [ids],
    );

    return rows;
  }

  // ✅ NUEVO: Resumen del día para dashboard (Inicio)
  // 🔧 FIX: KPI "pendientes" = SOLO pendientes de días anteriores a la fecha seleccionada
  @Get("resumen-dia")
  @Roles("RRHH", "Gerencia")
  async resumenDia(
    @Query()
    q: {
      fecha?: string; // YYYY-MM-DD
      usuarioIds?: string; // CSV
    },
  ) {
    const fecha = (q.fecha || "").trim();
    if (!fecha)
      throw new BadRequestException("fecha es obligatorio (YYYY-MM-DD)");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
      throw new BadRequestException("fecha inválida. Use YYYY-MM-DD");
    }

    const raw = (q.usuarioIds || "").trim();
    if (!raw) {
      return {
        fecha,
        total_empleados: 0,
        marcaron_ingreso: 0,
        no_marcaron_ingreso: 0,
        tardanzas: 0,
        pendientes: 0,
        ingresos: [],
        top_tardanzas: [],
      };
    }

    const ids = raw
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    if (ids.length === 0) {
      return {
        fecha,
        total_empleados: 0,
        marcaron_ingreso: 0,
        no_marcaron_ingreso: 0,
        tardanzas: 0,
        pendientes: 0,
        ingresos: [],
        top_tardanzas: [],
      };
    }

    // 1) Ingresos (primer JORNADA_IN por empleado en el día)
    const ingresos = await this.ds.query(
      `
    WITH first_in AS (
      SELECT
        a.usuario_id,
        MIN(a.fecha_hora) AS fecha_hora_in
      FROM asistencias a
      WHERE a.usuario_id = ANY($1::uuid[])
        AND a.evento = 'JORNADA_IN'
        AND a.estado_validacion <> 'rechazado'
        AND a.fecha_hora >= ($2::date)::timestamp
        AND a.fecha_hora <  (($2::date + interval '1 day'))::timestamp
      GROUP BY a.usuario_id
    )
    SELECT
      u.id AS usuario_id,
      (u.nombre || ' ' || COALESCE(u.apellido_paterno,'') || ' ' || COALESCE(u.apellido_materno,'')) AS nombre_completo,
      u.numero_documento,
      fi.fecha_hora_in,
      COALESCE(a.minutos_tarde, 0) AS minutos_tarde
    FROM first_in fi
    JOIN usuarios u ON u.id = fi.usuario_id
    LEFT JOIN asistencias a
      ON a.usuario_id = fi.usuario_id
     AND a.evento = 'JORNADA_IN'
     AND a.fecha_hora = fi.fecha_hora_in
    ORDER BY fi.fecha_hora_in ASC
    `,
      [ids, fecha],
    );

    const marcaron_ingreso = ingresos.length;

    // 2) Tardanzas del día
    const tardanzas = ingresos.filter(
      (r: any) => Number(r?.minutos_tarde || 0) > 0,
    ).length;

    // 3) Pendientes BLOQUEANTES: días anteriores a la fecha seleccionada (tiene IN y no OUT)
    const pendientesRows = await this.ds.query(
      `
    WITH dias AS (
      SELECT
        a.usuario_id,
        (a.fecha_hora::date) AS fecha,
        BOOL_OR(a.evento = 'JORNADA_IN'  AND a.estado_validacion <> 'rechazado') AS tiene_in,
        BOOL_OR(a.evento = 'JORNADA_OUT' AND a.estado_validacion <> 'rechazado') AS tiene_out
      FROM asistencias a
      WHERE a.usuario_id = ANY($1::uuid[])
        AND a.fecha_hora::date < $2::date          -- ✅ SOLO ANTERIORES
      GROUP BY a.usuario_id, (a.fecha_hora::date)
    ),
    pendientes AS (
      SELECT usuario_id, MIN(fecha) AS fecha_pendiente
      FROM dias
      WHERE tiene_in = true AND tiene_out = false
      GROUP BY usuario_id
    )
    SELECT usuario_id, fecha_pendiente
    FROM pendientes
    `,
      [ids, fecha],
    );

    const pendientes = pendientesRows.length;

    // 4) Totales
    const total_empleados = ids.length;
    const no_marcaron_ingreso = Math.max(0, total_empleados - marcaron_ingreso);

    // 5) Top tardanzas
    const top_tardanzas = ingresos
      .filter((r: any) => Number(r?.minutos_tarde || 0) > 0)
      .sort(
        (a: any, b: any) =>
          Number(b.minutos_tarde || 0) - Number(a.minutos_tarde || 0),
      )
      .slice(0, 8);

    return {
      fecha,
      total_empleados,
      marcaron_ingreso,
      no_marcaron_ingreso,
      tardanzas,
      pendientes,
      ingresos,
      top_tardanzas,
    };
  }
}
