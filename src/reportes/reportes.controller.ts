// reportes.controller.ts (BACKEND - NestJS)
import {
  BadRequestException,
  Controller,
  Get,
  Query,
  Res,
} from "@nestjs/common";
import { DataSource } from "typeorm";
import { Roles } from "../common/roles.decorator";
import type { Response } from "express";
import * as ExcelJS from "exceljs";

// PDF
import PDFDocument from "pdfkit";
import * as path from "path";
import * as fs from "fs";

type ResumenRow = {
  usuario_id: string;
  usuario: string;

  marcas_in: number;
  marcas_out: number;

  tardanzas_jornada_in: number;
  tardanzas_refrigerio_in: number;
  minutos_tarde_jornada_in: number;
  minutos_tarde_refrigerio_in: number;

  minutos_extra_salida: number;

  tardanzas: number;
  minutos_tarde_total: number;

  primer_ingreso: string | null;
  ultima_salida: string | null;

  dias_laborables: number;
  dias_feriados: number;
  dias_con_asistencia: number;
  ausencias_justificadas: number;
  ausencias_injustificadas: number;

  horario_vigente_desde: string | null;
  total_excepciones: number;

  ranking?: number;
};

type DetalleAnaliticoRow = {
  fecha: string;
  hora: string;
  empleado: string;
  dni: string;
  sede: string;
  area: string;
  tipo: string;
  evento: string;
  minutos_tarde: number;
  metodo: string;
  estado_validacion: string;

  excepcion_tipo: string;
  excepcion_observacion: string;

  min_tarde_jornada_in: number;
  min_tarde_refrigerio_in: number;
  minutos_tarde_total: number;
  horas_tarde_total: number;

  tardanzas_jornada_in: number;
  tardanzas_refrigerio_in: number;
  tardanzas_dia: number;

  minutos_acumulados: number;
  horas_acumuladas: number;
};

type FiltrosQuery = {
  period?: string;
  ref?: string;
  desde?: string;
  hasta?: string;
  usuarioId?: string;
  sedeId?: string;
};

type SqlBuildResult = {
  params: any[];
  where: string;
};

@Controller("reportes")
export class ReportesController {
  constructor(private ds: DataSource) {}

  // ==========================
  // Filtros globales reportes
  // ==========================
  private readonly DNI_EXCLUIDO = "44823948";

  private filtroNoRechazado(aliasAsistencia = "a") {
    return `COALESCE(${aliasAsistencia}.estado_validacion,'') <> 'rechazado'`;
  }

  private filtroNoDniExcluido(aliasUsuario = "u") {
    return `COALESCE(${aliasUsuario}.numero_documento,'') <> '${this.DNI_EXCLUIDO}'`;
  }

  // ==========================
  // Helpers base
  // ==========================
  private pad2(n: number) {
    return String(n).padStart(2, "0");
  }

  private toDateOnlyLocal(d: Date): string {
    return `${d.getFullYear()}-${this.pad2(d.getMonth() + 1)}-${this.pad2(
      d.getDate(),
    )}`;
  }

  private formatDatePEFromDateOnly(yyyyMmDd: string): string {
    const [y, m, d] = yyyyMmDd.split("-");
    return `${d}/${m}/${y}`;
  }

  private minutosToHHMM(minutos: number): string {
    const m = Math.max(0, Number(minutos) || 0);
    const hh = Math.floor(m / 60);
    const mm = m % 60;
    return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
  }

  private resolverRango(params: {
    period?: string;
    ref?: string;
    desde?: string;
    hasta?: string;
  }) {
    const { period, ref, desde, hasta } = params;

    let start: Date, end: Date;

    const toDate = (s: string) =>
      new Date(s + (s.length === 10 ? "T00:00:00" : ""));

    if (desde && hasta) {
      start = toDate(desde);
      end = toDate(hasta);
      if (isNaN(+start) || isNaN(+end) || start > end) {
        throw new BadRequestException("Rango inválido");
      }
      end = new Date(end.getTime() + 24 * 60 * 60 * 1000 - 1);
    } else {
      const base = ref ? toDate(ref) : new Date();
      if (isNaN(+base)) {
        throw new BadRequestException("Fecha de referencia inválida");
      }

      const y = base.getFullYear();
      const m = base.getMonth();
      const set00 = (d: Date) => {
        d.setHours(0, 0, 0, 0);
        return d;
      };

      switch ((period || "mes").toLowerCase()) {
        case "semana": {
          const day = base.getDay();
          const diff = day === 0 ? -6 : 1 - day;
          start = set00(new Date(base));
          start.setDate(base.getDate() + diff);
          end = new Date(start);
          end.setDate(start.getDate() + 6);
          end.setHours(23, 59, 59, 999);
          break;
        }
        case "quincena": {
          const d = base.getDate();
          if (d <= 15) {
            start = set00(new Date(y, m, 1));
            end = set00(new Date(y, m, 15));
            end.setHours(23, 59, 59, 999);
          } else {
            start = set00(new Date(y, m, 16));
            end = set00(new Date(y, m + 1, 0));
            end.setHours(23, 59, 59, 999);
          }
          break;
        }
        case "bimestre": {
          const b = Math.floor(m / 2) * 2;
          start = set00(new Date(y, b, 1));
          end = set00(new Date(y, b + 2, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case "trimestre": {
          const q = Math.floor(m / 3) * 3;
          start = set00(new Date(y, q, 1));
          end = set00(new Date(y, q + 3, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case "semestre": {
          const s = m < 6 ? 0 : 6;
          start = set00(new Date(y, s, 1));
          end = set00(new Date(y, s + 6, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case "anual": {
          start = set00(new Date(y, 0, 1));
          end = set00(new Date(y, 12, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case "mes":
        default: {
          start = set00(new Date(y, m, 1));
          end = set00(new Date(y, m + 1, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
      }
    }

    const startDate = this.toDateOnlyLocal(start);
    const endDate = this.toDateOnlyLocal(end);

    return { start, end, startDate, endDate };
  }

  private validarRangoDetalle(desde?: string, hasta?: string) {
    if (!desde || !hasta) {
      throw new BadRequestException("Debe indicar desde y hasta");
    }
  }

  private validarRangoMaximoSinFiltros(
    desde?: string,
    hasta?: string,
    usuarioId?: string,
    sedeId?: string,
  ) {
    if (usuarioId || sedeId) return;
    if (!desde || !hasta) return;

    const d1 = new Date(desde + "T00:00:00");
    const d2 = new Date(hasta + "T00:00:00");
    const diffDays = Math.floor((+d2 - +d1) / (1000 * 60 * 60 * 24)) + 1;

    if (!isFinite(diffDays) || diffDays <= 0) {
      throw new BadRequestException("Rango inválido");
    }

    if (diffDays > 31) {
      throw new BadRequestException(
        "Para descargar sin filtros (usuario/sede) el rango máximo permitido es 31 días.",
      );
    }
  }

  private buildWhereAsistencias(params: {
    desde: string;
    hasta: string;
    usuarioId?: string;
    sedeId?: string;
    aliasAsistencia?: string;
    aliasUsuario?: string;
    fechaModo?: "datetime" | "date";
    usarAliasUsuarioEnUuid?: boolean;
  }): SqlBuildResult {
    const {
      desde,
      hasta,
      usuarioId,
      sedeId,
      aliasAsistencia = "a",
      aliasUsuario = "u",
      fechaModo = "date",
      usarAliasUsuarioEnUuid = false,
    } = params;

    const sqlParams: any[] = [desde, hasta];
    const conds: string[] = [
      fechaModo === "datetime"
        ? `${aliasAsistencia}.fecha_hora >= $1::date AND ${aliasAsistencia}.fecha_hora < ($2::date + interval '1 day')`
        : `${aliasAsistencia}.fecha_hora >= $1::date AND ${aliasAsistencia}.fecha_hora < ($2::date + interval '1 day')`,
      this.filtroNoRechazado(aliasAsistencia),
      this.filtroNoDniExcluido(aliasUsuario),
    ];

    let p = 3;

    if (usuarioId) {
      sqlParams.push(usuarioId);
      conds.push(
        usarAliasUsuarioEnUuid
          ? `${aliasUsuario}.id = $${p}::uuid`
          : `${aliasAsistencia}.usuario_id = $${p}::uuid`,
      );
      p++;
    }

    if (sedeId) {
      sqlParams.push(sedeId);
      conds.push(`${aliasUsuario}.sede_id = $${p}::uuid`);
      p++;
    }

    return {
      params: sqlParams,
      where: conds.join(" AND "),
    };
  }

  private applyExcelHeaderStyle(ws: ExcelJS.Worksheet) {
    const header = ws.getRow(1);
    header.font = { bold: true };
    header.alignment = { vertical: "middle", horizontal: "center" };
    ws.views = [{ state: "frozen", ySplit: 1 }];
  }

  private applyExcelHeaderStyleDark(ws: ExcelJS.Worksheet) {
    const header = ws.getRow(1);
    header.height = 22;

    header.eachCell((cell) => {
      cell.font = {
        bold: true,
        color: { argb: "FFFFFFFF" },
      };
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "1F4E78" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "D9D9D9" } },
        left: { style: "thin", color: { argb: "D9D9D9" } },
        bottom: { style: "thin", color: { argb: "D9D9D9" } },
        right: { style: "thin", color: { argb: "D9D9D9" } },
      };
    });

    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: ws.columnCount },
    };
  }

  private applyExcelBodyStyle(ws: ExcelJS.Worksheet) {
    for (let i = 2; i <= ws.rowCount; i++) {
      const row = ws.getRow(i);

      row.eachCell((cell) => {
        cell.alignment = {
          vertical: "middle",
          horizontal: "center",
        };
        cell.border = {
          top: { style: "thin", color: { argb: "EDEDED" } },
          left: { style: "thin", color: { argb: "EDEDED" } },
          bottom: { style: "thin", color: { argb: "EDEDED" } },
          right: { style: "thin", color: { argb: "EDEDED" } },
        };
      });

      if (i % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "F8FAFC" },
          };
        });
      }
    }
  }

  private async sendExcel(
    res: Response,
    wb: ExcelJS.Workbook,
    filename: string,
  ) {
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    await wb.xlsx.write(res);
    res.end();
  }

  private getLogoPath() {
    return path.join(process.cwd(), "public", "logo_negro.png");
  }

  private addLogoIfExists(
    doc: PDFKit.PDFDocument,
    width: number,
    x?: number,
    y?: number,
  ) {
    const logoPath = this.getLogoPath();
    if (fs.existsSync(logoPath)) {
      doc.image(
        logoPath,
        x ?? doc.page.margins.left,
        y ?? doc.page.margins.top - 8,
        { width },
      );
    }
  }

  private async resolverEtiquetasFiltros(
    usuarioId?: string,
    sedeId?: string,
  ): Promise<{ usuarioLabel: string; sedeLabel: string }> {
    let usuarioLabel = "-";
    let sedeLabel = "-";

    if (usuarioId) {
      const usuarioRows = await this.ds.query(
        `
        SELECT TRIM(
          COALESCE(u.nombre, '') || ' ' ||
          COALESCE(u.apellido_paterno, '') || ' ' ||
          COALESCE(u.apellido_materno, '')
        ) AS nombre
        FROM usuarios u
        WHERE u.id = $1::uuid
        LIMIT 1
        `,
        [usuarioId],
      );

      if (usuarioRows?.length) {
        usuarioLabel = usuarioRows[0].nombre || usuarioId;
      }
    }

    if (sedeId) {
      const sedeRows = await this.ds.query(
        `
        SELECT COALESCE(s.nombre, '-') AS nombre
        FROM sedes s
        WHERE s.id = $1::uuid
        LIMIT 1
        `,
        [sedeId],
      );

      if (sedeRows?.length) {
        sedeLabel = sedeRows[0].nombre || sedeId;
      }
    }

    return { usuarioLabel, sedeLabel };
  }

  private async obtenerDetalleAnaliticoRows(params: {
    desde: string;
    hasta: string;
    usuarioId?: string;
    sedeId?: string;
  }): Promise<DetalleAnaliticoRow[]> {
    const { desde, hasta, usuarioId, sedeId } = params;

    const build = this.buildWhereAsistencias({
      desde,
      hasta,
      usuarioId,
      sedeId,
      aliasAsistencia: "a",
      aliasUsuario: "u",
      usarAliasUsuarioEnUuid: true,
    });

    const rows = await this.ds.query(
      `
      WITH base AS (
        SELECT
          a.usuario_id,
          a.fecha_hora::date AS fecha,
          a.fecha_hora,
          TO_CHAR(a.fecha_hora, 'HH24:MI') AS hora,

          TRIM(
            COALESCE(u.nombre,'') || ' ' ||
            COALESCE(u.apellido_paterno,'') || ' ' ||
            COALESCE(u.apellido_materno,'')
          ) AS empleado,

          COALESCE(u.numero_documento, '') AS dni,
          COALESCE(s.nombre, '') AS sede,
          COALESCE(ar.nombre, '') AS area,

          CASE a.tipo
            WHEN 'IN'  THEN 'ENTRADA'
            WHEN 'OUT' THEN 'SALIDA'
            ELSE COALESCE(a.tipo, '')
          END AS tipo,

          COALESCE(a.evento, '') AS evento_codigo,

          CASE a.evento
            WHEN 'JORNADA_IN'     THEN 'Inicio de jornada'
            WHEN 'JORNADA_OUT'    THEN 'Fin de jornada'
            WHEN 'REFRIGERIO_IN'  THEN 'Entrada de refrigerio'
            WHEN 'REFRIGERIO_OUT' THEN 'Salida a refrigerio'
            WHEN 'ALMUERZO_IN'    THEN 'Entrada de almuerzo'
            WHEN 'ALMUERZO_OUT'   THEN 'Salida a almuerzo'
            WHEN 'BREAK_IN'       THEN 'Entrada de break'
            WHEN 'BREAK_OUT'      THEN 'Salida a break'
            ELSE COALESCE(a.evento, '')
          END AS evento,

          COALESCE(a.minutos_tarde, 0) AS minutos_tarde,
          COALESCE(a.metodo, '') AS metodo,
          COALESCE(a.estado_validacion, '') AS estado_validacion,

          COALESCE(ue.tipo, '') AS excepcion_tipo,
          COALESCE(ue.observacion, '') AS excepcion_observacion,

          h.hora_salida_programada
        FROM asistencias a
        JOIN usuarios u ON u.id = a.usuario_id
        LEFT JOIN sedes s ON s.id = u.sede_id
        LEFT JOIN areas ar ON ar.id = u.area_id
        LEFT JOIN usuario_excepciones ue
          ON ue.usuario_id = a.usuario_id
         AND ue.fecha = a.fecha_hora::date
        LEFT JOIN LATERAL (
          SELECT
            CASE
              WHEN uh.hora_fin_2 IS NOT NULL THEN uh.hora_fin_2
              ELSE uh.hora_fin
            END AS hora_salida_programada
          FROM usuario_horarios uh
          WHERE uh.usuario_id = a.usuario_id
            AND uh.dia_semana = EXTRACT(ISODOW FROM a.fecha_hora)::int
            AND uh.es_descanso = FALSE
            AND uh.fecha_inicio <= a.fecha_hora::date
            AND COALESCE(uh.fecha_fin, '9999-12-31'::date) >= a.fecha_hora::date
          ORDER BY uh.fecha_inicio DESC, uh.creado_en DESC
          LIMIT 1
        ) h ON TRUE
        WHERE ${build.where}
      ),
      diario AS (
        SELECT
          b.usuario_id,
          b.fecha,

          COALESCE(SUM(b.minutos_tarde) FILTER (WHERE b.evento_codigo = 'JORNADA_IN'), 0) AS min_tarde_jornada_in,
          COALESCE(SUM(b.minutos_tarde) FILTER (WHERE b.evento_codigo = 'REFRIGERIO_IN'), 0) AS min_tarde_refrigerio_in,
          COALESCE(SUM(b.minutos_tarde), 0) AS minutos_tarde_total,

          COUNT(*) FILTER (WHERE b.minutos_tarde > 0 AND b.evento_codigo = 'JORNADA_IN') AS tardanzas_jornada_in,
          COUNT(*) FILTER (WHERE b.minutos_tarde > 0 AND b.evento_codigo = 'REFRIGERIO_IN') AS tardanzas_refrigerio_in,
          COUNT(*) FILTER (WHERE b.minutos_tarde > 0) AS tardanzas_dia,

          COALESCE(SUM(
            CASE
              WHEN b.evento_codigo = 'JORNADA_OUT'
               AND b.hora_salida_programada IS NOT NULL
              THEN GREATEST(
                (
                  EXTRACT(HOUR FROM b.fecha_hora)::int * 60
                  + EXTRACT(MINUTE FROM b.fecha_hora)::int
                )
                -
                (
                  EXTRACT(HOUR FROM b.hora_salida_programada)::int * 60
                  + EXTRACT(MINUTE FROM b.hora_salida_programada)::int
                ),
                0
              )
              ELSE 0
            END
          ), 0) AS minutos_acumulados
        FROM base b
        GROUP BY b.usuario_id, b.fecha
      )
      SELECT
        TO_CHAR(b.fecha, 'DD/MM/YYYY') AS fecha,
        b.hora,
        b.empleado,
        b.dni,
        b.sede,
        b.area,
        b.tipo,
        b.evento,
        b.minutos_tarde,
        b.metodo,
        b.estado_validacion,
        b.excepcion_tipo,
        b.excepcion_observacion,

        COALESCE(d.min_tarde_jornada_in, 0) AS min_tarde_jornada_in,
        COALESCE(d.min_tarde_refrigerio_in, 0) AS min_tarde_refrigerio_in,
        COALESCE(d.minutos_tarde_total, 0) AS minutos_tarde_total,
        ROUND(COALESCE(d.minutos_tarde_total, 0)::numeric / 60.0, 2) AS horas_tarde_total,

        COALESCE(d.tardanzas_jornada_in, 0) AS tardanzas_jornada_in,
        COALESCE(d.tardanzas_refrigerio_in, 0) AS tardanzas_refrigerio_in,
        COALESCE(d.tardanzas_dia, 0) AS tardanzas_dia,

        COALESCE(d.minutos_acumulados, 0) AS minutos_acumulados,
        ROUND(COALESCE(d.minutos_acumulados, 0)::numeric / 60.0, 2) AS horas_acumuladas
      FROM base b
      JOIN diario d
        ON d.usuario_id = b.usuario_id
       AND d.fecha = b.fecha
      ORDER BY b.empleado, b.fecha, b.hora, b.fecha_hora
      `,
      build.params,
    );

    return rows.map((r: any) => ({
      fecha: r.fecha,
      hora: r.hora,
      empleado: r.empleado,
      dni: r.dni,
      sede: r.sede,
      area: r.area,
      tipo: r.tipo,
      evento: r.evento,
      minutos_tarde: Number(r.minutos_tarde) || 0,
      metodo: r.metodo || "",
      estado_validacion: r.estado_validacion || "",

      excepcion_tipo: r.excepcion_tipo || "",
      excepcion_observacion: r.excepcion_observacion || "",

      min_tarde_jornada_in: Number(r.min_tarde_jornada_in) || 0,
      min_tarde_refrigerio_in: Number(r.min_tarde_refrigerio_in) || 0,
      minutos_tarde_total: Number(r.minutos_tarde_total) || 0,
      horas_tarde_total: Number(r.horas_tarde_total) || 0,

      tardanzas_jornada_in: Number(r.tardanzas_jornada_in) || 0,
      tardanzas_refrigerio_in: Number(r.tardanzas_refrigerio_in) || 0,
      tardanzas_dia: Number(r.tardanzas_dia) || 0,

      minutos_acumulados: Number(r.minutos_acumulados) || 0,
      horas_acumuladas: Number(r.horas_acumuladas) || 0,
    }));
  }

  // ==========================
  // DATA RESUMEN
  // ==========================
  private async obtenerResumenData(params: FiltrosQuery) {
    const { startDate, endDate } = this.resolverRango(params);
    const { usuarioId, sedeId } = params;

    const resumenBuild = this.buildWhereAsistencias({
      desde: startDate,
      hasta: endDate,
      usuarioId,
      sedeId,
      aliasAsistencia: "a",
      aliasUsuario: "u",
    });

    const resumenRows = await this.ds.query(
      `
      WITH asist_base AS (
        SELECT
          a.id,
          a.usuario_id,
          a.fecha_hora,
          a.tipo,
          a.evento,
          COALESCE(a.minutos_tarde, 0) AS minutos_tarde
        FROM asistencias a
        JOIN usuarios u ON u.id = a.usuario_id
        WHERE ${resumenBuild.where}
      ),
      asist_con_horario AS (
        SELECT
          ab.*,
          h.hora_fin,
          h.hora_fin_2,
          CASE
            WHEN h.hora_fin_2 IS NOT NULL THEN h.hora_fin_2
            ELSE h.hora_fin
          END AS hora_salida_programada
        FROM asist_base ab
        LEFT JOIN LATERAL (
          SELECT
            uh.hora_fin,
            uh.hora_fin_2,
            uh.fecha_inicio,
            uh.creado_en
          FROM usuario_horarios uh
          WHERE uh.usuario_id = ab.usuario_id
            AND uh.dia_semana = EXTRACT(ISODOW FROM ab.fecha_hora)::int
            AND uh.es_descanso = FALSE
            AND uh.fecha_inicio <= ab.fecha_hora::date
            AND COALESCE(uh.fecha_fin, '9999-12-31'::date) >= ab.fecha_hora::date
          ORDER BY uh.fecha_inicio DESC, uh.creado_en DESC
          LIMIT 1
        ) h ON TRUE
      )
      SELECT
        u.id AS usuario_id,
        (u.nombre || ' ' || COALESCE(u.apellido_paterno,'') || ' ' || COALESCE(u.apellido_materno,'')) AS usuario,

        COUNT(*) FILTER (WHERE a.tipo = 'IN')  AS marcas_in,
        COUNT(*) FILTER (WHERE a.tipo = 'OUT') AS marcas_out,

        COUNT(*) FILTER (WHERE a.minutos_tarde > 0 AND a.evento = 'JORNADA_IN')    AS tardanzas_jornada_in,
        COUNT(*) FILTER (WHERE a.minutos_tarde > 0 AND a.evento = 'REFRIGERIO_IN') AS tardanzas_refrigerio_in,

        COALESCE(SUM(a.minutos_tarde) FILTER (WHERE a.evento = 'JORNADA_IN'), 0)    AS minutos_tarde_jornada_in,
        COALESCE(SUM(a.minutos_tarde) FILTER (WHERE a.evento = 'REFRIGERIO_IN'), 0) AS minutos_tarde_refrigerio_in,

        COALESCE(SUM(
          CASE
            WHEN a.evento = 'JORNADA_OUT'
             AND a.hora_salida_programada IS NOT NULL
            THEN GREATEST(
              (
                EXTRACT(HOUR FROM a.fecha_hora)::int * 60
                + EXTRACT(MINUTE FROM a.fecha_hora)::int
              )
              -
              (
                EXTRACT(HOUR FROM a.hora_salida_programada)::int * 60
                + EXTRACT(MINUTE FROM a.hora_salida_programada)::int
              ),
              0
            )
            ELSE 0
          END
        ), 0) AS minutos_extra_salida,

        COUNT(*) FILTER (WHERE a.minutos_tarde > 0) AS tardanzas,
        COALESCE(SUM(a.minutos_tarde), 0) AS minutos_tarde_total,

        MIN(a.fecha_hora) AS primer_ingreso,
        MAX(a.fecha_hora) AS ultima_salida
      FROM asist_con_horario a
      JOIN usuarios u ON u.id = a.usuario_id
      GROUP BY u.id, u.nombre, u.apellido_paterno, u.apellido_materno
      `,
      resumenBuild.params,
    );

    const ausParams: any[] = [startDate, endDate];
    let usuariosFiltro = `WHERE u.activo = TRUE AND COALESCE(u.numero_documento,'') <> '${this.DNI_EXCLUIDO}'`;

    let aidx = 3;
    if (usuarioId) {
      ausParams.push(usuarioId);
      usuariosFiltro += ` AND u.id = $${aidx}::uuid`;
      aidx++;
    }
    if (sedeId) {
      ausParams.push(sedeId);
      usuariosFiltro += ` AND u.sede_id = $${aidx}::uuid`;
      aidx++;
    }

    const ausRows = await this.ds.query(
      `
      WITH fechas AS (
        SELECT generate_series($1::date, $2::date, interval '1 day')::date AS fecha
      ),
      usuarios_filtrados AS (
        SELECT
          u.id,
          (u.nombre || ' ' || COALESCE(u.apellido_paterno,'') || ' ' || COALESCE(u.apellido_materno,'')) AS nombre
        FROM usuarios u
        ${usuariosFiltro}
      ),
      calendario AS (
        SELECT uf.id AS usuario_id, uf.nombre, f.fecha
        FROM usuarios_filtrados uf
        CROSS JOIN fechas f
      ),
      horario_vigente AS (
        SELECT
          uh.usuario_id,
          MIN(uh.fecha_inicio) AS horario_vigente_desde
        FROM usuario_horarios uh
        JOIN usuarios_filtrados uf ON uf.id = uh.usuario_id
        WHERE uh.es_descanso = FALSE
          AND uh.fecha_inicio <= $2::date
          AND COALESCE(uh.fecha_fin, '9999-12-31'::date) >= $1::date
        GROUP BY uh.usuario_id
      ),
      cal_hor AS (
        SELECT
          c.usuario_id,
          c.nombre,
          c.fecha,
          hv.horario_vigente_desde,
          (hv.horario_vigente_desde IS NOT NULL AND c.fecha >= hv.horario_vigente_desde) AS aplica_por_vigencia,
          CASE
            WHEN h.id IS NOT NULL AND h.es_descanso = FALSE THEN TRUE
            ELSE FALSE
          END AS laborable_por_horario
        FROM calendario c
        LEFT JOIN horario_vigente hv
          ON hv.usuario_id = c.usuario_id
        LEFT JOIN LATERAL (
          SELECT
            uh.id,
            uh.es_descanso,
            uh.fecha_inicio,
            uh.creado_en
          FROM usuario_horarios uh
          WHERE uh.usuario_id = c.usuario_id
            AND uh.dia_semana = EXTRACT(ISODOW FROM c.fecha)::int
            AND uh.fecha_inicio <= c.fecha
            AND COALESCE(uh.fecha_fin, '9999-12-31'::date) >= c.fecha
          ORDER BY uh.fecha_inicio DESC, uh.creado_en DESC
          LIMIT 1
        ) h ON TRUE
      ),
      exc AS (
        SELECT usuario_id, fecha, es_laborable, tipo
        FROM usuario_excepciones
      ),
      exc_count AS (
        SELECT
          e.usuario_id,
          COUNT(*) AS total_excepciones
        FROM usuario_excepciones e
        JOIN usuarios_filtrados uf ON uf.id = e.usuario_id
        WHERE e.fecha >= $1::date
          AND e.fecha <= $2::date
        GROUP BY e.usuario_id
      ),
      asis_dia AS (
        SELECT
          a.usuario_id,
          a.fecha_hora::date AS fecha,
          COUNT(*) AS marcas
        FROM asistencias a
        JOIN usuarios_filtrados uf ON uf.id = a.usuario_id
        WHERE a.fecha_hora >= $1::date
          AND a.fecha_hora < ($2::date + interval '1 day')
          AND COALESCE(a.estado_validacion,'') <> 'rechazado'
        GROUP BY a.usuario_id, a.fecha_hora::date
      ),
      cal_final AS (
        SELECT
          c.usuario_id,
          c.nombre,
          c.fecha,
          c.horario_vigente_desde,
          c.aplica_por_vigencia,
          c.laborable_por_horario,
          (
            c.horario_vigente_desde IS NOT NULL
            AND c.fecha >= c.horario_vigente_desde
            AND EXISTS (
              SELECT 1
              FROM public.feriados f
              WHERE f.fecha = c.fecha
            )
          ) AS es_feriado,
          e.es_laborable AS exc_es_laborable,
          e.tipo AS exc_tipo,
          COALESCE(ad.marcas, 0) > 0 AS tiene_asistencia
        FROM cal_hor c
        LEFT JOIN exc e
          ON e.usuario_id = c.usuario_id
         AND e.fecha = c.fecha
        LEFT JOIN asis_dia ad
          ON ad.usuario_id = c.usuario_id
         AND ad.fecha = c.fecha
      )
      SELECT
        cf.usuario_id,
        MIN(cf.nombre) AS usuario,

        COUNT(*) FILTER (
          WHERE cf.aplica_por_vigencia = TRUE
            AND cf.laborable_por_horario = TRUE
            AND cf.es_feriado = FALSE
        ) AS dias_laborables,

        COUNT(*) FILTER (
          WHERE cf.aplica_por_vigencia = TRUE
            AND cf.laborable_por_horario = TRUE
            AND cf.es_feriado = TRUE
        ) AS dias_feriados,

        COUNT(*) FILTER (
          WHERE cf.aplica_por_vigencia = TRUE
            AND cf.laborable_por_horario = TRUE
            AND cf.es_feriado = FALSE
            AND cf.tiene_asistencia = TRUE
        ) AS dias_con_asistencia,

        COUNT(*) FILTER (
          WHERE cf.aplica_por_vigencia = TRUE
            AND cf.laborable_por_horario = TRUE
            AND cf.es_feriado = FALSE
            AND cf.tiene_asistencia = FALSE
            AND cf.exc_es_laborable = FALSE
        ) AS ausencias_justificadas,

        COUNT(*) FILTER (
          WHERE cf.aplica_por_vigencia = TRUE
            AND cf.laborable_por_horario = TRUE
            AND cf.es_feriado = FALSE
            AND cf.tiene_asistencia = FALSE
            AND (cf.exc_es_laborable IS NULL OR cf.exc_es_laborable = TRUE)
        ) AS ausencias_injustificadas,

        MIN(cf.horario_vigente_desde)::text AS horario_vigente_desde,
        COALESCE(MAX(ec.total_excepciones), 0) AS total_excepciones
      FROM cal_final cf
      LEFT JOIN exc_count ec
        ON ec.usuario_id = cf.usuario_id
      GROUP BY cf.usuario_id
      `,
      ausParams,
    );

    const map = new Map<string, ResumenRow>();

    for (const r of resumenRows) {
      map.set(r.usuario_id, {
        usuario_id: r.usuario_id,
        usuario: r.usuario,

        marcas_in: Number(r.marcas_in) || 0,
        marcas_out: Number(r.marcas_out) || 0,

        tardanzas_jornada_in: Number(r.tardanzas_jornada_in) || 0,
        tardanzas_refrigerio_in: Number(r.tardanzas_refrigerio_in) || 0,
        minutos_tarde_jornada_in: Number(r.minutos_tarde_jornada_in) || 0,
        minutos_tarde_refrigerio_in: Number(r.minutos_tarde_refrigerio_in) || 0,
        minutos_extra_salida: Number(r.minutos_extra_salida) || 0,

        tardanzas: Number(r.tardanzas) || 0,
        minutos_tarde_total: Number(r.minutos_tarde_total) || 0,

        primer_ingreso: r.primer_ingreso ?? null,
        ultima_salida: r.ultima_salida ?? null,

        dias_laborables: 0,
        dias_feriados: 0,
        dias_con_asistencia: 0,
        ausencias_justificadas: 0,
        ausencias_injustificadas: 0,

        horario_vigente_desde: null,
        total_excepciones: 0,
      });
    }

    for (const a of ausRows) {
      const existing =
        map.get(a.usuario_id) ||
        ({
          usuario_id: a.usuario_id,
          usuario: a.usuario,

          marcas_in: 0,
          marcas_out: 0,

          tardanzas_jornada_in: 0,
          tardanzas_refrigerio_in: 0,
          minutos_tarde_jornada_in: 0,
          minutos_tarde_refrigerio_in: 0,
          minutos_extra_salida: 0,

          tardanzas: 0,
          minutos_tarde_total: 0,

          primer_ingreso: null,
          ultima_salida: null,

          dias_laborables: 0,
          dias_feriados: 0,
          dias_con_asistencia: 0,
          ausencias_justificadas: 0,
          ausencias_injustificadas: 0,

          horario_vigente_desde: null,
          total_excepciones: 0,
        } as ResumenRow);

      existing.dias_laborables = Number(a.dias_laborables) || 0;
      existing.dias_feriados = Number(a.dias_feriados) || 0;
      existing.dias_con_asistencia = Number(a.dias_con_asistencia) || 0;
      existing.ausencias_justificadas = Number(a.ausencias_justificadas) || 0;
      existing.ausencias_injustificadas =
        Number(a.ausencias_injustificadas) || 0;
      existing.horario_vigente_desde = a.horario_vigente_desde ?? null;
      existing.total_excepciones = Number(a.total_excepciones) || 0;

      map.set(a.usuario_id, existing);
    }

    const data = Array.from(map.values());

    data.sort((a, b) => {
      if (
        (b.ausencias_injustificadas || 0) !== (a.ausencias_injustificadas || 0)
      ) {
        return (
          (b.ausencias_injustificadas || 0) - (a.ausencias_injustificadas || 0)
        );
      }
      if ((b.minutos_tarde_total || 0) !== (a.minutos_tarde_total || 0)) {
        return (b.minutos_tarde_total || 0) - (a.minutos_tarde_total || 0);
      }
      return (b.tardanzas || 0) - (a.tardanzas || 0);
    });

    data.forEach((row, i) => (row.ranking = i + 1));

    return {
      periodo: { desde: startDate, hasta: endDate },
      filtros: { usuarioId: usuarioId || null, sedeId: sedeId || null },
      data,
    };
  }

  // ==========================
  // JSON
  // ==========================
  @Roles("Gerencia", "RRHH")
  @Get("resumen")
  async resumen(
    @Query("period") period?: string,
    @Query("ref") ref?: string,
    @Query("desde") desde?: string,
    @Query("hasta") hasta?: string,
    @Query("usuarioId") usuarioId?: string,
    @Query("sedeId") sedeId?: string,
  ) {
    return this.obtenerResumenData({
      period,
      ref,
      desde,
      hasta,
      usuarioId,
      sedeId,
    });
  }

  // ==========================
  // Excel (Resumen)
  // ==========================
  @Roles("Gerencia", "RRHH")
  @Get("resumen-excel")
  async resumenExcel(
    @Res() res: Response,
    @Query("period") period?: string,
    @Query("ref") ref?: string,
    @Query("desde") desde?: string,
    @Query("hasta") hasta?: string,
    @Query("usuarioId") usuarioId?: string,
    @Query("sedeId") sedeId?: string,
  ) {
    const result = await this.obtenerResumenData({
      period,
      ref,
      desde,
      hasta,
      usuarioId,
      sedeId,
    });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Resumen");

    ws.columns = [
      { header: "Ranking", key: "ranking", width: 10 },
      { header: "Usuario", key: "usuario", width: 34 },

      { header: "Marcas IN", key: "marcas_in", width: 11 },
      { header: "Marcas OUT", key: "marcas_out", width: 12 },

      { header: "Tard. Ingreso", key: "tard_jornada", width: 14 },
      { header: "Min. Ingreso", key: "min_jornada", width: 12 },
      { header: "HH:MM Ingreso", key: "hhmm_jornada", width: 14 },

      { header: "Tard. Refrig.", key: "tard_ref", width: 14 },
      { header: "Min. Refrig.", key: "min_ref", width: 12 },
      { header: "HH:MM Refrig.", key: "hhmm_ref", width: 14 },

      { header: "Min. Salida", key: "min_salida", width: 12 },
      { header: "HH:MM Salida", key: "hhmm_salida", width: 14 },

      { header: "Tardanzas Total", key: "tardanzas", width: 14 },
      { header: "Minutos Total", key: "minutos_tarde_total", width: 12 },
      { header: "HH:MM Total", key: "hhmm_total", width: 12 },

      { header: "Días laborables", key: "dias_laborables", width: 14 },
      { header: "Días feriados", key: "dias_feriados", width: 12 },
      { header: "Días con asistencia", key: "dias_con_asistencia", width: 16 },
      { header: "Ausencias just.", key: "ausencias_justificadas", width: 14 },
      { header: "Ausencias injust.", key: "ausencias_injustificadas", width: 16 },
      { header: "Excepciones", key: "total_excepciones", width: 12 },

      {
        header: "Horario vigente desde",
        key: "horario_vigente_desde",
        width: 18,
      },
      { header: "Primer ingreso", key: "primer_ingreso", width: 22 },
      { header: "Última salida", key: "ultima_salida", width: 22 },
    ];

    for (const r of result.data as ResumenRow[]) {
      ws.addRow({
        ranking: r.ranking ?? null,
        usuario: r.usuario,

        marcas_in: r.marcas_in,
        marcas_out: r.marcas_out,

        tard_jornada: r.tardanzas_jornada_in,
        min_jornada: r.minutos_tarde_jornada_in,
        hhmm_jornada: this.minutosToHHMM(r.minutos_tarde_jornada_in),

        tard_ref: r.tardanzas_refrigerio_in,
        min_ref: r.minutos_tarde_refrigerio_in,
        hhmm_ref: this.minutosToHHMM(r.minutos_tarde_refrigerio_in),

        min_salida: r.minutos_extra_salida,
        hhmm_salida: this.minutosToHHMM(r.minutos_extra_salida),

        tardanzas: r.tardanzas,
        minutos_tarde_total: r.minutos_tarde_total,
        hhmm_total: this.minutosToHHMM(r.minutos_tarde_total),

        dias_laborables: r.dias_laborables,
        dias_feriados: r.dias_feriados,
        dias_con_asistencia: r.dias_con_asistencia,
        ausencias_justificadas: r.ausencias_justificadas,
        ausencias_injustificadas: r.ausencias_injustificadas,
        total_excepciones: r.total_excepciones,

        horario_vigente_desde: r.horario_vigente_desde,
        primer_ingreso: r.primer_ingreso,
        ultima_salida: r.ultima_salida,
      });
    }

    this.applyExcelHeaderStyle(ws);
    await this.sendExcel(res, wb, "reporte_asistencias_resumen.xlsx");
  }

  // ==========================
  // RESUMEN DÍA - EXCEL
  // ==========================
  @Roles("Gerencia", "RRHH")
  @Get("resumen-dia-excel")
  async resumenDiaExcel(
    @Res() res: Response,
    @Query("desde") desde?: string,
    @Query("hasta") hasta?: string,
    @Query("usuarioId") usuarioId?: string,
    @Query("sedeId") sedeId?: string,
  ) {
    this.validarRangoDetalle(desde, hasta);

    const build = this.buildWhereAsistencias({
      desde: desde!,
      hasta: hasta!,
      usuarioId,
      sedeId,
      aliasAsistencia: "a",
      aliasUsuario: "u",
    });

    const rows = await this.ds.query(
      `
      WITH base AS (
        SELECT
          a.usuario_id,
          a.fecha_hora::date AS fecha,
          a.fecha_hora,
          a.evento,
          COALESCE(a.minutos_tarde, 0) AS minutos_tarde,
          (u.nombre || ' ' || COALESCE(u.apellido_paterno,'') || ' ' || COALESCE(u.apellido_materno,'')) AS usuario,
          COALESCE(s.nombre, '') AS sede,
          COALESCE(ar.nombre, '') AS area
        FROM asistencias a
        JOIN usuarios u ON u.id = a.usuario_id
        LEFT JOIN sedes s ON s.id = u.sede_id
        LEFT JOIN areas ar ON ar.id = u.area_id
        WHERE ${build.where}
      ),
      salida_calc AS (
        SELECT
          b.usuario_id,
          b.fecha,
          COALESCE(SUM(
            CASE
              WHEN b.evento = 'JORNADA_OUT' AND h.hora_salida_programada IS NOT NULL
              THEN GREATEST(
                (
                  EXTRACT(HOUR FROM b.fecha_hora)::int * 60
                  + EXTRACT(MINUTE FROM b.fecha_hora)::int
                )
                -
                (
                  EXTRACT(HOUR FROM h.hora_salida_programada)::int * 60
                  + EXTRACT(MINUTE FROM h.hora_salida_programada)::int
                ),
                0
              )
              ELSE 0
            END
          ), 0) AS minutos_extra_salida
        FROM base b
        LEFT JOIN LATERAL (
          SELECT
            CASE
              WHEN uh.hora_fin_2 IS NOT NULL THEN uh.hora_fin_2
              ELSE uh.hora_fin
            END AS hora_salida_programada
          FROM usuario_horarios uh
          WHERE uh.usuario_id = b.usuario_id
            AND uh.dia_semana = EXTRACT(ISODOW FROM b.fecha)::int
            AND uh.es_descanso = FALSE
            AND uh.fecha_inicio <= b.fecha
            AND COALESCE(uh.fecha_fin, '9999-12-31'::date) >= b.fecha
          ORDER BY uh.fecha_inicio DESC, uh.creado_en DESC
          LIMIT 1
        ) h ON TRUE
        GROUP BY b.usuario_id, b.fecha
      )
      SELECT
        b.fecha,
        b.usuario,
        b.sede,
        b.area,
        MIN(TO_CHAR(b.fecha_hora, 'HH24:MI')) FILTER (WHERE b.evento = 'JORNADA_IN') AS jornada_in,
        MIN(TO_CHAR(b.fecha_hora, 'HH24:MI')) FILTER (WHERE b.evento = 'REFRIGERIO_OUT') AS refrigerio_out,
        MIN(TO_CHAR(b.fecha_hora, 'HH24:MI')) FILTER (WHERE b.evento = 'REFRIGERIO_IN') AS refrigerio_in,
        MAX(TO_CHAR(b.fecha_hora, 'HH24:MI')) FILTER (WHERE b.evento = 'JORNADA_OUT') AS jornada_out,
        COALESCE(SUM(b.minutos_tarde) FILTER (WHERE b.evento = 'JORNADA_IN'), 0) AS min_ingreso,
        COALESCE(SUM(b.minutos_tarde) FILTER (WHERE b.evento = 'REFRIGERIO_IN'), 0) AS min_refrigerio,
        COALESCE(sc.minutos_extra_salida, 0) AS min_salida
      FROM base b
      LEFT JOIN salida_calc sc
        ON sc.usuario_id = b.usuario_id
       AND sc.fecha = b.fecha
      GROUP BY b.fecha, b.usuario, b.sede, b.area, sc.minutos_extra_salida
      ORDER BY b.fecha, b.usuario
      `,
      build.params,
    );

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Resumen Día");

    ws.columns = [
      { header: "Fecha", key: "fecha", width: 12 },
      { header: "Usuario", key: "usuario", width: 32 },
      { header: "Sede", key: "sede", width: 18 },
      { header: "Área", key: "area", width: 18 },
      { header: "Jor. In", key: "jornada_in", width: 10 },
      { header: "Ref. Out", key: "refrigerio_out", width: 10 },
      { header: "Ref. In", key: "refrigerio_in", width: 10 },
      { header: "Jor. Out", key: "jornada_out", width: 10 },
      { header: "Min. Ing.", key: "min_ingreso", width: 10 },
      { header: "Min. Ref.", key: "min_refrigerio", width: 10 },
      { header: "Min. Sal.", key: "min_salida", width: 10 },
      { header: "HH:MM Sal.", key: "hhmm_salida", width: 12 },
    ];

    rows.forEach((r: any) => {
      ws.addRow({
        ...r,
        hhmm_salida: this.minutosToHHMM(Number(r.min_salida) || 0),
      });
    });

    this.applyExcelHeaderStyle(ws);
    await this.sendExcel(res, wb, "reporte_asistencias_resumen_dia.xlsx");
  }

  // ==========================
  // PDF (Resumen)
  // ==========================
  @Roles("Gerencia", "RRHH")
  @Get("resumen-pdf")
  async resumenPdf(
    @Res() res: Response,
    @Query("period") period?: string,
    @Query("ref") ref?: string,
    @Query("desde") desde?: string,
    @Query("hasta") hasta?: string,
    @Query("usuarioId") usuarioId?: string,
    @Query("sedeId") sedeId?: string,
  ) {
    const { startDate, endDate } = this.resolverRango({
      period,
      ref,
      desde,
      hasta,
    });

    const result = await this.obtenerResumenData({
      period,
      ref,
      desde,
      hasta,
      usuarioId,
      sedeId,
    });

    const { usuarioLabel, sedeLabel } = await this.resolverEtiquetasFiltros(
      usuarioId,
      sedeId,
    );

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="reporte_asistencias_resumen.pdf"`,
    );

    const pageConfig = {
      margin: 26,
      size: "A4" as const,
      layout: "landscape" as const,
    };

    const doc = new PDFDocument(pageConfig);
    doc.pipe(res);

    const rows: ResumenRow[] = result.data as ResumenRow[];

    const toNumber = (n: any) => Number(n) || 0;
    const fmtNum = (n: any) => String(toNumber(n));
    const fmtHHMM = (n: any) => this.minutosToHHMM(toNumber(n));

    const drawPageHeader = (pageNumber: number) => {
      this.addLogoIfExists(doc, 82, doc.page.margins.left, 14);

      doc
        .font("Helvetica-Bold")
        .fontSize(17)
        .fillColor("#111111")
        .text("Reporte de Asistencias - Resumen", 0, 20, {
          align: "center",
        });

      doc
        .font("Helvetica")
        .fontSize(10)
        .fillColor("#333333")
        .text(
          `Periodo: ${this.formatDatePEFromDateOnly(
            startDate,
          )} a ${this.formatDatePEFromDateOnly(endDate)}`,
          0,
          48,
          { align: "center" },
        );

      doc
        .font("Helvetica")
        .fontSize(9)
        .fillColor("#444444")
        .text(`Filtros: usuario=${usuarioLabel} | sede=${sedeLabel}`, 0, 64, {
          align: "center",
        });

      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor("#666666")
        .text(`Página ${pageNumber}`, 0, 20, {
          align: "right",
        });
    };

    const col = {
      rk: 28,
      usuario: 228,
      lab: 42,
      asis: 46,
      ausi: 48,
      exc: 40,
      tardIng: 48,
      tardRef: 48,
      tardTot: 52,
      hhTard: 60,
      hhAcum: 60,
    };

    const tableWidth =
      col.rk +
      col.usuario +
      col.lab +
      col.asis +
      col.ausi +
      col.exc +
      col.tardIng +
      col.tardRef +
      col.tardTot +
      col.hhTard +
      col.hhAcum;

    const contentWidth =
      doc.page.width - doc.page.margins.left - doc.page.margins.right;

    const startX =
      doc.page.margins.left + Math.max(0, (contentWidth - tableWidth) / 2);

    const paddingX = 4;
    const paddingY = 4;
    const minRowH = 22;

    let pageNumber = 1;
    drawPageHeader(pageNumber);

    let y = 100;

    const getTextHeight = (
      text: any,
      width: number,
      opts?: {
        bold?: boolean;
        fontSize?: number;
        align?: "left" | "center";
      },
    ) => {
      doc
        .font(opts?.bold ? "Helvetica-Bold" : "Helvetica")
        .fontSize(opts?.fontSize ?? 9);

      return doc.heightOfString(String(text ?? ""), {
        width: width - paddingX * 2,
        align: opts?.align ?? "left",
      });
    };

    const drawCell = (
      text: any,
      x: number,
      width: number,
      yPos: number,
      height: number,
      opts?: {
        bold?: boolean;
        align?: "left" | "center";
        bgColor?: string;
        textColor?: string;
        fontSize?: number;
        borderColor?: string;
      },
    ) => {
      const bgColor = opts?.bgColor;
      const textColor = opts?.textColor ?? "#111111";
      const borderColor = opts?.borderColor ?? "#BFBFBF";
      const align = opts?.align ?? "left";
      const fontSize = opts?.fontSize ?? 9;

      if (bgColor) {
        doc.save();
        doc.rect(x, yPos, width, height).fill(bgColor);
        doc.restore();
      }

      doc.save();
      doc
        .lineWidth(0.5)
        .strokeColor(borderColor)
        .rect(x, yPos, width, height)
        .stroke();
      doc.restore();

      doc
        .font(opts?.bold ? "Helvetica-Bold" : "Helvetica")
        .fontSize(fontSize)
        .fillColor(textColor)
        .text(String(text ?? ""), x + paddingX, yPos + paddingY, {
          width: width - paddingX * 2,
          align,
        });
    };

    const drawTableHeader = () => {
      const headerH = 24;
      let x = startX;

      drawCell("Rk", x, col.rk, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.rk;

      drawCell("Usuario", x, col.usuario, y, headerH, {
        bold: true,
        align: "left",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.usuario;

      drawCell("Lab.", x, col.lab, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.lab;

      drawCell("Asis.", x, col.asis, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.asis;

      drawCell("Aus.I", x, col.ausi, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.ausi;

      drawCell("Exc.", x, col.exc, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.exc;

      drawCell("T.Ing", x, col.tardIng, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.tardIng;

      drawCell("T.Ref", x, col.tardRef, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.tardRef;

      drawCell("T.Total", x, col.tardTot, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.tardTot;

      drawCell("HH:Tard", x, col.hhTard, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.hhTard;

      drawCell("HH:Acum", x, col.hhAcum, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });

      y += headerH;
    };

    drawTableHeader();

    for (const [index, r] of rows.entries()) {
      const tardanzasTotales = toNumber(r.tardanzas);
      const rowH = Math.max(
        minRowH,
        getTextHeight(fmtNum(r.ranking), col.rk, { align: "center" }) +
          paddingY * 2,
        getTextHeight(r.usuario, col.usuario) + paddingY * 2,
        getTextHeight(fmtNum(r.dias_laborables), col.lab, { align: "center" }) +
          paddingY * 2,
        getTextHeight(fmtNum(r.dias_con_asistencia), col.asis, {
          align: "center",
        }) +
          paddingY * 2,
        getTextHeight(fmtNum(r.ausencias_injustificadas), col.ausi, {
          align: "center",
        }) +
          paddingY * 2,
        getTextHeight(fmtNum(r.total_excepciones), col.exc, {
          align: "center",
        }) +
          paddingY * 2,
        getTextHeight(fmtNum(r.tardanzas_jornada_in), col.tardIng, {
          align: "center",
        }) +
          paddingY * 2,
        getTextHeight(fmtNum(r.tardanzas_refrigerio_in), col.tardRef, {
          align: "center",
        }) +
          paddingY * 2,
        getTextHeight(fmtNum(tardanzasTotales), col.tardTot, {
          align: "center",
        }) +
          paddingY * 2,
        getTextHeight(fmtHHMM(r.minutos_tarde_total), col.hhTard, {
          align: "center",
        }) +
          paddingY * 2,
        getTextHeight(fmtHHMM(r.minutos_extra_salida), col.hhAcum, {
          align: "center",
        }) +
          paddingY * 2,
      );

      if (y + rowH > doc.page.height - doc.page.margins.bottom - 18) {
        doc.addPage(pageConfig);
        pageNumber += 1;
        drawPageHeader(pageNumber);
        y = 100;
        drawTableHeader();
      }

      const rowBg = index % 2 === 0 ? "#FFFFFF" : "#F9FBFD";

      let x = startX;

      drawCell(fmtNum(r.ranking), x, col.rk, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.rk;

      drawCell(r.usuario, x, col.usuario, y, rowH, {
        align: "left",
        bgColor: rowBg,
      });
      x += col.usuario;

      drawCell(fmtNum(r.dias_laborables), x, col.lab, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.lab;

      drawCell(fmtNum(r.dias_con_asistencia), x, col.asis, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.asis;

      drawCell(fmtNum(r.ausencias_injustificadas), x, col.ausi, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor:
          toNumber(r.ausencias_injustificadas) > 0 ? "#C00000" : "#111111",
        bold: toNumber(r.ausencias_injustificadas) > 0,
      });
      x += col.ausi;

      drawCell(fmtNum(r.total_excepciones), x, col.exc, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: toNumber(r.total_excepciones) > 0 ? "#7F6000" : "#111111",
        bold: toNumber(r.total_excepciones) > 0,
      });
      x += col.exc;

      drawCell(fmtNum(r.tardanzas_jornada_in), x, col.tardIng, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: toNumber(r.tardanzas_jornada_in) > 0 ? "#C00000" : "#111111",
        bold: toNumber(r.tardanzas_jornada_in) > 0,
      });
      x += col.tardIng;

      drawCell(fmtNum(r.tardanzas_refrigerio_in), x, col.tardRef, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor:
          toNumber(r.tardanzas_refrigerio_in) > 0 ? "#C00000" : "#111111",
        bold: toNumber(r.tardanzas_refrigerio_in) > 0,
      });
      x += col.tardRef;

      let tardColor = "#16A34A";
      if (tardanzasTotales > 0 && tardanzasTotales <= 2) tardColor = "#D97706";
      if (tardanzasTotales > 2) tardColor = "#DC2626";

      drawCell(fmtNum(tardanzasTotales), x, col.tardTot, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: tardColor,
        bold: tardanzasTotales > 0,
      });
      x += col.tardTot;

      drawCell(fmtHHMM(r.minutos_tarde_total), x, col.hhTard, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: toNumber(r.minutos_tarde_total) > 0 ? "#C00000" : "#111111",
        bold: toNumber(r.minutos_tarde_total) > 0,
      });
      x += col.hhTard;

      drawCell(fmtHHMM(r.minutos_extra_salida), x, col.hhAcum, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: toNumber(r.minutos_extra_salida) > 0 ? "#1F4E78" : "#111111",
        bold: toNumber(r.minutos_extra_salida) > 0,
      });

      y += rowH;
    }

    doc.end();
  }

  // ==========================
  // DETALLE - EXCEL
  // ==========================
  @Roles("Gerencia", "RRHH")
  @Get("detalle-excel")
  async detalleExcel(
    @Res() res: Response,
    @Query("desde") desde?: string,
    @Query("hasta") hasta?: string,
    @Query("usuarioId") usuarioId?: string,
    @Query("sedeId") sedeId?: string,
  ) {
    this.validarRangoDetalle(desde, hasta);
    this.validarRangoMaximoSinFiltros(desde, hasta, usuarioId, sedeId);

    const rows = await this.obtenerDetalleAnaliticoRows({
      desde: desde!,
      hasta: hasta!,
      usuarioId,
      sedeId,
    });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Detalle");

    ws.columns = [
      { header: "Fecha", key: "fecha", width: 12 },
      { header: "Hora", key: "hora", width: 9 },
      { header: "Empleado", key: "empleado", width: 32 },
      { header: "Área", key: "area", width: 24 },
      { header: "DNI", key: "dni", width: 14 },
      { header: "Sede", key: "sede", width: 18 },
      { header: "Evento", key: "evento", width: 22 },
      { header: "Tipo", key: "tipo", width: 10 },
      { header: "Excepción", key: "excepcion_tipo", width: 22 },
      { header: "Observación excepción", key: "excepcion_observacion", width: 32 },

      { header: "Min. tarde marca", key: "minutos_tarde", width: 14 },
      { header: "Min. tard. ingreso", key: "min_tarde_jornada_in", width: 16 },
      {
        header: "Min. tard. refrigerio",
        key: "min_tarde_refrigerio_in",
        width: 18,
      },
      { header: "Min. tarde total", key: "minutos_tarde_total", width: 15 },
      { header: "Horas tarde total", key: "horas_tarde_total", width: 15 },

      { header: "Tard. ingreso", key: "tardanzas_jornada_in", width: 13 },
      { header: "Tard. refrigerio", key: "tardanzas_refrigerio_in", width: 15 },
      { header: "Tardanzas día", key: "tardanzas_dia", width: 13 },

      { header: "Min. acumulados", key: "minutos_acumulados", width: 15 },
      { header: "Horas acumuladas", key: "horas_acumuladas", width: 15 },

      { header: "Estado", key: "estado_validacion", width: 14 },
      { header: "Método", key: "metodo", width: 16 },
    ];

    rows.forEach((r) => {
      ws.addRow(r);
    });

    this.applyExcelHeaderStyleDark(ws);
    this.applyExcelBodyStyle(ws);

    ws.getColumn("empleado").alignment = {
      horizontal: "left",
      vertical: "middle",
    };
    ws.getColumn("area").alignment = {
      horizontal: "left",
      vertical: "middle",
    };
    ws.getColumn("evento").alignment = {
      horizontal: "left",
      vertical: "middle",
    };
    ws.getColumn("excepcion_tipo").alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    ws.getColumn("excepcion_observacion").alignment = {
      horizontal: "left",
      vertical: "middle",
    };
    ws.getColumn("sede").alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    ws.getColumn("estado_validacion").alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    ws.getColumn("metodo").alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    ws.getColumn("horas_tarde_total").numFmt = "0.00";
    ws.getColumn("horas_acumuladas").numFmt = "0.00";

    for (let i = 2; i <= ws.rowCount; i++) {
      const row = ws.getRow(i);

      const minMarca = Number(row.getCell("minutos_tarde").value || 0);
      const tardDia = Number(row.getCell("tardanzas_dia").value || 0);
      const minAcum = Number(row.getCell("minutos_acumulados").value || 0);
      const excTipo = String(row.getCell("excepcion_tipo").value || "");

      if (minMarca > 0) {
        row.getCell("minutos_tarde").font = {
          color: { argb: "C00000" },
          bold: true,
        };
      }

      if (tardDia > 0) {
        row.getCell("tardanzas_dia").font = {
          color: { argb: "C00000" },
          bold: true,
        };
      }

      if (minAcum > 0) {
        row.getCell("minutos_acumulados").font = {
          color: { argb: "1F4E78" },
          bold: true,
        };
        row.getCell("horas_acumuladas").font = {
          color: { argb: "1F4E78" },
          bold: true,
        };
      }

      if (excTipo) {
        row.getCell("excepcion_tipo").font = {
          color: { argb: "7F6000" },
          bold: true,
        };
      }
    }

    await this.sendExcel(res, wb, "reporte_asistencias_detalle.xlsx");
  }

  // ==========================
  // DETALLE - PDF
  // ==========================
  @Roles("Gerencia", "RRHH")
  @Get("detalle-pdf")
  async detallePdf(
    @Res() res: Response,
    @Query("desde") desde?: string,
    @Query("hasta") hasta?: string,
    @Query("usuarioId") usuarioId?: string,
    @Query("sedeId") sedeId?: string,
  ) {
    this.validarRangoDetalle(desde, hasta);
    this.validarRangoMaximoSinFiltros(desde, hasta, usuarioId, sedeId);

    const rows = await this.obtenerDetalleAnaliticoRows({
      desde: desde!,
      hasta: hasta!,
      usuarioId,
      sedeId,
    });

    const { usuarioLabel, sedeLabel } = await this.resolverEtiquetasFiltros(
      usuarioId,
      sedeId,
    );

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="reporte_asistencias_detalle.pdf"`,
    );

    const pageConfig = {
      margin: 26,
      size: "A4" as const,
      layout: "landscape" as const,
    };

    const doc = new PDFDocument(pageConfig);
    doc.pipe(res);

    const fmtNum = (n: any) => String(Number(n) || 0);

    const fmtMetodo = (metodo: string) => {
      switch ((metodo || "").toLowerCase()) {
        case "manual_supervisor":
          return "MANUAL";
        case "scanner_barras":
          return "SCANNER";
        case "qr_fijo":
          return "QR FIJO";
        case "qr_dinamico":
          return "QR DINÁMICO";
        default:
          return metodo || "-";
      }
    };

    const fmtExcepcion = (tipo: string) => {
      switch ((tipo || "").toUpperCase()) {
        case "HORARIO_ESPECIAL":
          return "HORARIO ESPECIAL";
        case "DESCANSO_ESPECIAL":
          return "DESCANSO ESPECIAL";
        case "LABORABLE_EN_DESCANSO":
          return "LABORABLE EN DESCANSO";
        default:
          return tipo || "-";
      }
    };

    const drawPageHeader = (pageNumber: number) => {
      this.addLogoIfExists(doc, 82, doc.page.margins.left, 14);

      doc
        .font("Helvetica-Bold")
        .fontSize(17)
        .fillColor("#111111")
        .text("Asistencia - Detalle de marcajes", 0, 20, {
          align: "center",
        });

      doc
        .font("Helvetica")
        .fontSize(10)
        .fillColor("#333333")
        .text(
          `Periodo: ${this.formatDatePEFromDateOnly(
            desde!,
          )} a ${this.formatDatePEFromDateOnly(hasta!)}`,
          0,
          48,
          { align: "center" },
        );

      doc
        .font("Helvetica")
        .fontSize(9)
        .fillColor("#444444")
        .text(`Filtros: usuario=${usuarioLabel} | sede=${sedeLabel}`, 0, 64, {
          align: "center",
        });

      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor("#666666")
        .text(`Página ${pageNumber}`, 0, 20, {
          align: "right",
        });
    };

    const col = {
      fecha: 70,
      hora: 38,
      empleado: 178,
      sede: 82,
      tipo: 52,
      evento: 115,
      metodo: 70,
      excepcion: 100,
      minTarde: 52,
    };

    const tableWidth =
      col.fecha +
      col.hora +
      col.empleado +
      col.sede +
      col.tipo +
      col.evento +
      col.metodo +
      col.excepcion +
      col.minTarde;

    const contentWidth =
      doc.page.width - doc.page.margins.left - doc.page.margins.right;

    const startX =
      doc.page.margins.left + Math.max(0, (contentWidth - tableWidth) / 2);

    const paddingX = 4;
    const paddingY = 4;
    const minRowH = 22;

    let pageNumber = 1;
    drawPageHeader(pageNumber);

    let y = 100;

    const getTextHeight = (
      text: any,
      width: number,
      opts?: {
        bold?: boolean;
        fontSize?: number;
        align?: "left" | "center";
      },
    ) => {
      doc
        .font(opts?.bold ? "Helvetica-Bold" : "Helvetica")
        .fontSize(opts?.fontSize ?? 9);

      return doc.heightOfString(String(text ?? ""), {
        width: width - paddingX * 2,
        align: opts?.align ?? "left",
      });
    };

       const drawCell = (
        text: any,
        x: number,
        width: number,
        yPos: number,
        height: number,
        opts?: {
          bold?: boolean;
          align?: "left" | "center";
          bgColor?: string;
          textColor?: string;
          fontSize?: number;
          borderColor?: string;
        },
      ) => {
        const bgColor = opts?.bgColor;
        const textColor = opts?.textColor ?? "#111111";
        const borderColor = opts?.borderColor ?? "#BFBFBF";
        const align = opts?.align ?? "left";
        const fontSize = opts?.fontSize ?? 9;
        const fontName = opts?.bold ? "Helvetica-Bold" : "Helvetica";

        if (bgColor) {
          doc.save();
          doc.rect(x, yPos, width, height).fill(bgColor);
          doc.restore();
        }

        doc.save();
        doc
          .lineWidth(0.5)
          .strokeColor(borderColor)
          .rect(x, yPos, width, height)
          .stroke();
        doc.restore();

        doc.font(fontName).fontSize(fontSize);

        const textHeight = doc.heightOfString(String(text ?? ""), {
          width: width - paddingX * 2,
          align,
        });

        const textY = yPos + Math.max(paddingY, (height - textHeight) / 2);

        doc
          .font(fontName)
          .fontSize(fontSize)
          .fillColor(textColor)
          .text(String(text ?? ""), x + paddingX, textY, {
            width: width - paddingX * 2,
            align,
          });
      };


    const drawTableHeader = () => {
      const headerH = 24;
      let x = startX;

      drawCell("Fecha", x, col.fecha, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.fecha;

      drawCell("Hora", x, col.hora, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.hora;

      drawCell("Empleado", x, col.empleado, y, headerH, {
        bold: true,
        align: "left",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.empleado;

      drawCell("Sede", x, col.sede, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.sede;

      drawCell("Tipo", x, col.tipo, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.tipo;

      drawCell("Evento", x, col.evento, y, headerH, {
        bold: true,
        align: "left",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.evento;

      drawCell("Método", x, col.metodo, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.metodo;

      drawCell("Excepción", x, col.excepcion, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.excepcion;

      drawCell("Min. tarde", x, col.minTarde, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });

      y += headerH;
    };

    drawTableHeader();

    rows.forEach((r, index) => {
      const minutosTardeMarca = Number(r.minutos_tarde) || 0;
      const metodoLabel = fmtMetodo(r.metodo);
      const esManual = (r.metodo || "").toLowerCase() === "manual_supervisor";
      const excepcionLabel = fmtExcepcion(r.excepcion_tipo);

      const rowH = Math.max(
        minRowH,
        getTextHeight(r.fecha, col.fecha, { align: "center" }) + paddingY * 2,
        getTextHeight(r.hora, col.hora, { align: "center" }) + paddingY * 2,
        getTextHeight(r.empleado, col.empleado) + paddingY * 2,
        getTextHeight(r.sede, col.sede, { align: "center" }) + paddingY * 2,
        getTextHeight(r.tipo, col.tipo, { align: "center" }) + paddingY * 2,
        getTextHeight(r.evento, col.evento) + paddingY * 2,
        getTextHeight(metodoLabel, col.metodo, { align: "center" }) +
          paddingY * 2,
        getTextHeight(excepcionLabel, col.excepcion, { align: "center" }) +
          paddingY * 2,
        getTextHeight(fmtNum(minutosTardeMarca), col.minTarde, {
          align: "center",
        }) +
          paddingY * 2,
      );

      if (y + rowH > doc.page.height - doc.page.margins.bottom - 18) {
        doc.addPage(pageConfig);
        pageNumber += 1;
        drawPageHeader(pageNumber);
        y = 100;
        drawTableHeader();
      }

      const rowBg = index % 2 === 0 ? "#FFFFFF" : "#F9FBFD";

      let x = startX;

      drawCell(r.fecha, x, col.fecha, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.fecha;

      drawCell(r.hora, x, col.hora, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.hora;

      drawCell(r.empleado, x, col.empleado, y, rowH, {
        align: "left",
        bgColor: rowBg,
      });
      x += col.empleado;

      drawCell(r.sede, x, col.sede, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.sede;

      drawCell(r.tipo, x, col.tipo, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.tipo;

      drawCell(r.evento, x, col.evento, y, rowH, {
        align: "left",
        bgColor: rowBg,
      });
      x += col.evento;

      drawCell(metodoLabel, x, col.metodo, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: esManual ? "#C00000" : "#111111",
        bold: esManual,
      });
      x += col.metodo;

      drawCell(excepcionLabel, x, col.excepcion, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: r.excepcion_tipo ? "#7F6000" : "#111111",
        bold: !!r.excepcion_tipo,
      });
      x += col.excepcion;

      drawCell(fmtNum(minutosTardeMarca), x, col.minTarde, y, rowH, {
        align: "center",
        bgColor: rowBg,
        textColor: minutosTardeMarca > 0 ? "#C00000" : "#111111",
        bold: minutosTardeMarca > 0,
      });

      y += rowH;
    });

    doc.end();
  }

  // ============================================
  // Reporte maestro usuarios
  // ============================================
  @Roles("Gerencia", "RRHH")
  @Get("usuarios-excel")
  async usuariosExcel(@Res() res: Response) {
    const rows = await this.ds.query(
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
          u.fecha_baja,
          u.email_personal,
          u.email_institucional,
          u.telefono_celular
       FROM usuarios u
       LEFT JOIN roles  r ON r.id = u.rol_id
       LEFT JOIN sedes  s ON s.id = u.sede_id
       LEFT JOIN areas  a ON a.id = u.area_id
      WHERE COALESCE(u.numero_documento,'') <> '${this.DNI_EXCLUIDO}'
      ORDER BY u.created_at DESC`,
    );

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Usuarios");

    ws.columns = [
      { header: "ID", key: "id", width: 36 },
      { header: "Nombre", key: "nombre", width: 20 },
      { header: "Apellido paterno", key: "apellido_paterno", width: 18 },
      { header: "Apellido materno", key: "apellido_materno", width: 18 },
      { header: "Tipo doc.", key: "tipo_documento", width: 10 },
      { header: "N° documento", key: "numero_documento", width: 16 },
      { header: "Fecha nacimiento", key: "fecha_nacimiento", width: 15 },
      { header: "Rol", key: "rol", width: 15 },
      { header: "Sede", key: "sede_nombre", width: 20 },
      { header: "Área", key: "area_nombre", width: 20 },
      { header: "Email personal", key: "email_personal", width: 28 },
      { header: "Email institucional", key: "email_institucional", width: 28 },
      { header: "Teléfono", key: "telefono_celular", width: 14 },
      { header: "Estado", key: "estado", width: 12 },
      { header: "Fecha baja", key: "fecha_baja", width: 18 },
    ];

    for (const u of rows) {
      ws.addRow({
        id: u.id,
        nombre: u.nombre,
        apellido_paterno: u.apellido_paterno,
        apellido_materno: u.apellido_materno,
        tipo_documento: u.tipo_documento,
        numero_documento: u.numero_documento,
        fecha_nacimiento: u.fecha_nacimiento
          ? new Date(u.fecha_nacimiento)
          : null,
        rol: u.rol,
        sede_nombre: u.sede_nombre,
        area_nombre: u.area_nombre,
        email_personal: u.email_personal,
        email_institucional: u.email_institucional,
        telefono_celular: u.telefono_celular,
        estado: u.activo ? "ACTIVO" : "INACTIVO",
        fecha_baja: u.fecha_baja ? new Date(u.fecha_baja) : null,
      });
    }

    this.applyExcelHeaderStyle(ws);
    await this.sendExcel(res, wb, "usuarios.xlsx");
  }

  // ============================================
  // Usuarios PDF
  // ============================================
  @Roles("Gerencia", "RRHH")
  @Get("usuarios-pdf")
  async usuariosPdf(@Res() res: Response) {
    const rows = await this.ds.query(
      `
      SELECT
        TRIM(
          COALESCE(u.nombre,'') || ' ' ||
          COALESCE(u.apellido_paterno,'') || ' ' ||
          COALESCE(u.apellido_materno,'')
        ) AS nombre_completo,
        COALESCE(u.tipo_documento,'-') AS tipo_doc,
        COALESCE(u.numero_documento,'-') AS numero_documento,
        COALESCE(s.nombre,'-') AS sede,
        COALESCE(a.nombre,'-') AS area,
        COALESCE(u.telefono_celular,'-') AS telefono
      FROM usuarios u
      LEFT JOIN sedes s ON s.id = u.sede_id
      LEFT JOIN areas a ON a.id = u.area_id
      WHERE COALESCE(u.numero_documento,'') <> '${this.DNI_EXCLUIDO}'
      ORDER BY u.created_at DESC
      `,
    );

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="reporte_usuarios.pdf"`,
    );

    const pageConfig = {
      margin: 26,
      size: "A4" as const,
      layout: "landscape" as const,
    };

    const doc = new PDFDocument(pageConfig);
    doc.pipe(res);

    const fechaGeneracion = new Date().toLocaleString("es-PE");

    const drawPageHeader = (pageNumber: number) => {
      this.addLogoIfExists(doc, 82, doc.page.margins.left, 14);

      doc
        .font("Helvetica-Bold")
        .fontSize(17)
        .fillColor("#111111")
        .text("Reporte de Usuarios", 0, 20, {
          align: "center",
        });

      doc
        .font("Helvetica")
        .fontSize(10)
        .fillColor("#333333")
        .text(`Generado: ${fechaGeneracion}`, 0, 48, {
          align: "center",
        });

      doc
        .font("Helvetica")
        .fontSize(9)
        .fillColor("#444444")
        .text(`Total de registros: ${rows.length}`, 0, 64, {
          align: "center",
        });

      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor("#666666")
        .text(`Página ${pageNumber}`, 0, 20, {
          align: "right",
        });
    };

    const col = {
      nombre: 245,
      tipo: 58,
      doc: 92,
      sede: 100,
      area: 165,
      tel: 86,
    };

    const tableWidth =
      col.nombre + col.tipo + col.doc + col.sede + col.area + col.tel;

    const contentWidth =
      doc.page.width - doc.page.margins.left - doc.page.margins.right;

    const startX =
      doc.page.margins.left + Math.max(0, (contentWidth - tableWidth) / 2);

    const paddingX = 4;
    const paddingY = 4;
    const minRowH = 22;

    let pageNumber = 1;
    drawPageHeader(pageNumber);

    let y = 100;

    const getTextHeight = (
      text: any,
      width: number,
      opts?: {
        bold?: boolean;
        fontSize?: number;
        align?: "left" | "center";
      },
    ) => {
      doc
        .font(opts?.bold ? "Helvetica-Bold" : "Helvetica")
        .fontSize(opts?.fontSize ?? 9);

      return doc.heightOfString(String(text ?? "-"), {
        width: width - paddingX * 2,
        align: opts?.align ?? "left",
      });
    };

    const drawCell = (
      text: any,
      x: number,
      width: number,
      yPos: number,
      height: number,
      opts?: {
        bold?: boolean;
        align?: "left" | "center";
        bgColor?: string;
        textColor?: string;
        fontSize?: number;
        borderColor?: string;
      },
    ) => {
      const bgColor = opts?.bgColor;
      const textColor = opts?.textColor ?? "#111111";
      const borderColor = opts?.borderColor ?? "#BFBFBF";
      const align = opts?.align ?? "left";
      const fontSize = opts?.fontSize ?? 9;

      if (bgColor) {
        doc.save();
        doc.rect(x, yPos, width, height).fill(bgColor);
        doc.restore();
      }

      doc.save();
      doc
        .lineWidth(0.5)
        .strokeColor(borderColor)
        .rect(x, yPos, width, height)
        .stroke();
      doc.restore();

      doc
        .font(opts?.bold ? "Helvetica-Bold" : "Helvetica")
        .fontSize(fontSize)
        .fillColor(textColor)
        .text(String(text ?? "-"), x + paddingX, yPos + paddingY, {
          width: width - paddingX * 2,
          align,
        });
    };

    const drawTableHeader = () => {
      const headerH = 24;
      let x = startX;

      drawCell("Nombre completo", x, col.nombre, y, headerH, {
        bold: true,
        align: "left",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.nombre;

      drawCell("Tipo doc.", x, col.tipo, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.tipo;

      drawCell("N° documento", x, col.doc, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.doc;

      drawCell("Sede", x, col.sede, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.sede;

      drawCell("Área", x, col.area, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });
      x += col.area;

      drawCell("Teléfono", x, col.tel, y, headerH, {
        bold: true,
        align: "center",
        bgColor: "#E9EFF7",
        fontSize: 9,
      });

      y += headerH;
    };

    drawTableHeader();

    for (const [index, r] of rows.entries()) {
      const rowH = Math.max(
        minRowH,
        getTextHeight(r.nombre_completo, col.nombre) + paddingY * 2,
        getTextHeight(r.tipo_doc, col.tipo, { align: "center" }) + paddingY * 2,
        getTextHeight(r.numero_documento, col.doc, { align: "center" }) +
          paddingY * 2,
        getTextHeight(r.sede, col.sede, { align: "center" }) + paddingY * 2,
        getTextHeight(r.area, col.area, { align: "center" }) + paddingY * 2,
        getTextHeight(r.telefono, col.tel, { align: "center" }) + paddingY * 2,
      );

      if (y + rowH > doc.page.height - doc.page.margins.bottom - 18) {
        doc.addPage(pageConfig);
        pageNumber += 1;
        drawPageHeader(pageNumber);
        y = 100;
        drawTableHeader();
      }

      const rowBg = index % 2 === 0 ? "#FFFFFF" : "#F9FBFD";

      let x = startX;

      drawCell(String(r.nombre_completo || "-"), x, col.nombre, y, rowH, {
        align: "left",
        bgColor: rowBg,
      });
      x += col.nombre;

      drawCell(String(r.tipo_doc || "-"), x, col.tipo, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.tipo;

      drawCell(String(r.numero_documento || "-"), x, col.doc, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.doc;

      drawCell(String(r.sede || "-"), x, col.sede, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.sede;

      drawCell(String(r.area || "-"), x, col.area, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });
      x += col.area;

      drawCell(String(r.telefono || "-"), x, col.tel, y, rowH, {
        align: "center",
        bgColor: rowBg,
      });

      y += rowH;
    }

    doc.end();
  }
}
