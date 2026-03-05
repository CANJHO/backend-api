// reportes.controller.ts (BACKEND - NestJS)
import {
  BadRequestException,
  Controller,
  Get,
  Query,
  Res,
} from '@nestjs/common';
import { DataSource } from 'typeorm';
import { Roles } from '../common/roles.decorator';
import type { Response } from 'express';
import * as ExcelJS from 'exceljs';

// PDF
import PDFDocument from 'pdfkit';
import * as path from 'path';
import * as fs from 'fs';

type ResumenRow = {
  usuario_id: string;
  usuario: string;

  marcas_in: number;
  marcas_out: number;

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

  ranking?: number;
};

@Controller('reportes')
export class ReportesController {
  constructor(private ds: DataSource) {}

  // ==========================
  // Helpers fechas / filtros
  // ==========================
  private pad2(n: number) {
    return String(n).padStart(2, '0');
  }

  // YYYY-MM-DD en HORA LOCAL del servidor
  private toDateOnlyLocal(d: Date): string {
    return `${d.getFullYear()}-${this.pad2(d.getMonth() + 1)}-${this.pad2(
      d.getDate(),
    )}`;
  }

  // dd/MM/yyyy para mostrar (Perú)
  private formatDatePEFromDateOnly(yyyyMmDd: string): string {
    const [y, m, d] = yyyyMmDd.split('-');
    return `${d}/${m}/${y}`;
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
      new Date(s + (s.length === 10 ? 'T00:00:00' : ''));

    if (desde && hasta) {
      start = toDate(desde);
      end = toDate(hasta);
      if (isNaN(+start) || isNaN(+end) || start > end) {
        throw new BadRequestException('Rango inválido');
      }
      end = new Date(end.getTime() + 24 * 60 * 60 * 1000 - 1);
    } else {
      const base = ref ? toDate(ref) : new Date();
      if (isNaN(+base)) {
        throw new BadRequestException('Fecha de referencia inválida');
      }

      const y = base.getFullYear();
      const m = base.getMonth();
      const set00 = (d: Date) => {
        d.setHours(0, 0, 0, 0);
        return d;
      };

      switch ((period || 'mes').toLowerCase()) {
        case 'semana': {
          const day = base.getDay();
          const diff = day === 0 ? -6 : 1 - day;
          start = set00(new Date(base));
          start.setDate(base.getDate() + diff);
          end = new Date(start);
          end.setDate(start.getDate() + 6);
          end.setHours(23, 59, 59, 999);
          break;
        }
        case 'quincena': {
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
        case 'bimestre': {
          const b = Math.floor(m / 2) * 2;
          start = set00(new Date(y, b, 1));
          end = set00(new Date(y, b + 2, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case 'trimestre': {
          const q = Math.floor(m / 3) * 3;
          start = set00(new Date(y, q, 1));
          end = set00(new Date(y, q + 3, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case 'semestre': {
          const s = m < 6 ? 0 : 6;
          start = set00(new Date(y, s, 1));
          end = set00(new Date(y, s + 6, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case 'anual': {
          start = set00(new Date(y, 0, 1));
          end = set00(new Date(y, 12, 0));
          end.setHours(23, 59, 59, 999);
          break;
        }
        case 'mes':
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

  // ==========================
  // Helper: Minutos -> HH:MM
  // ==========================
  private minutosToHHMM(minutos: number): string {
    const m = Math.max(0, Number(minutos) || 0);
    const hh = Math.floor(m / 60);
    const mm = m % 60;
    return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
  }

  private async obtenerResumenData(params: {
    period?: string;
    ref?: string;
    desde?: string;
    hasta?: string;
    usuarioId?: string;
    sedeId?: string;
  }) {
    const { startDate, endDate } = this.resolverRango(params);
    const { usuarioId, sedeId } = params;

    const resumenParams: any[] = [startDate, endDate];
    const resumenConds: string[] = [
      `a.fecha_hora >= $1::date AND a.fecha_hora < ($2::date + interval '1 day')`,
    ];

    let p = 3;
    if (usuarioId) {
      resumenParams.push(usuarioId);
      resumenConds.push(`a.usuario_id = $${p}::uuid`);
      p++;
    }
    if (sedeId) {
      resumenParams.push(sedeId);
      resumenConds.push(`u.sede_id = $${p}::uuid`);
      p++;
    }

    const where = resumenConds.join(' AND ');

    const resumenRows = await this.ds.query(
      `SELECT
          u.id AS usuario_id,
          (u.nombre || ' ' || COALESCE(u.apellido_paterno,'') || ' ' || COALESCE(u.apellido_materno,'')) AS usuario,
          COUNT(*) FILTER (WHERE a.tipo = 'IN')  AS marcas_in,
          COUNT(*) FILTER (WHERE a.tipo = 'OUT') AS marcas_out,
          COUNT(*) FILTER (WHERE a.minutos_tarde > 0) AS tardanzas,
          COALESCE(SUM(a.minutos_tarde), 0) AS minutos_tarde_total,
          MIN(a.fecha_hora) AS primer_ingreso,
          MAX(a.fecha_hora) AS ultima_salida
        FROM asistencias a
        JOIN usuarios u ON u.id = a.usuario_id
       WHERE ${where}
       GROUP BY u.id, u.nombre, u.apellido_paterno, u.apellido_materno`,
      resumenParams,
    );

    const ausParams: any[] = [startDate, endDate];
    let usuariosFiltro = 'WHERE u.activo = TRUE';

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
        SELECT u.id, (u.nombre || ' ' || COALESCE(u.apellido_paterno,'') || ' ' || COALESCE(u.apellido_materno,'')) AS nombre
          FROM usuarios u
         ${usuariosFiltro}
      ),
      calendario AS (
        SELECT uf.id AS usuario_id, uf.nombre, f.fecha
          FROM usuarios_filtrados uf
          CROSS JOIN fechas f
      ),
      horas AS (
        SELECT
          uh.usuario_id,
          uh.dia_semana,
          uh.es_descanso,
          uh.fecha_inicio,
          COALESCE(uh.fecha_fin, '9999-12-31'::date) AS fecha_fin
        FROM usuario_horarios uh
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
            WHEN h.usuario_id IS NOT NULL
              AND h.es_descanso = FALSE
              THEN TRUE
            ELSE FALSE
          END AS laborable_por_horario
        FROM calendario c
        LEFT JOIN horario_vigente hv
          ON hv.usuario_id = c.usuario_id
        LEFT JOIN horas h
          ON h.usuario_id = c.usuario_id
         AND h.dia_semana = EXTRACT(ISODOW FROM c.fecha)::int
         AND c.fecha BETWEEN h.fecha_inicio AND h.fecha_fin
      ),
      exc AS (
        SELECT usuario_id, fecha, es_laborable
          FROM usuario_excepciones
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

        MIN(cf.horario_vigente_desde)::text AS horario_vigente_desde
      FROM cal_final cf
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
        } as ResumenRow);

      existing.dias_laborables = Number(a.dias_laborables) || 0;
      existing.dias_feriados = Number(a.dias_feriados) || 0;
      existing.dias_con_asistencia = Number(a.dias_con_asistencia) || 0;
      existing.ausencias_justificadas = Number(a.ausencias_justificadas) || 0;
      existing.ausencias_injustificadas =
        Number(a.ausencias_injustificadas) || 0;
      existing.horario_vigente_desde = a.horario_vigente_desde ?? null;

      map.set(a.usuario_id, existing);
    }

    const data = Array.from(map.values());

    data.sort((a, b) => {
      if (
        (b.ausencias_injustificadas || 0) !==
        (a.ausencias_injustificadas || 0)
      ) {
        return (
          (b.ausencias_injustificadas || 0) -
          (a.ausencias_injustificadas || 0)
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
  @Roles('Gerencia', 'RRHH')
  @Get('resumen')
  async resumen(
    @Query('period') period?: string,
    @Query('ref') ref?: string,
    @Query('desde') desde?: string,
    @Query('hasta') hasta?: string,
    @Query('usuarioId') usuarioId?: string,
    @Query('sedeId') sedeId?: string,
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
  @Roles('Gerencia', 'RRHH')
  @Get('resumen-excel')
  async resumenExcel(
    @Res() res: Response,
    @Query('period') period?: string,
    @Query('ref') ref?: string,
    @Query('desde') desde?: string,
    @Query('hasta') hasta?: string,
    @Query('usuarioId') usuarioId?: string,
    @Query('sedeId') sedeId?: string,
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
    const ws = wb.addWorksheet('Resumen');

    ws.columns = [
      { header: 'Ranking', key: 'ranking', width: 10 },
      { header: 'Usuario', key: 'usuario', width: 34 },

      { header: 'Hora entrada', key: 'marcas_in', width: 12 },
      { header: 'Hora salida', key: 'marcas_out', width: 12 },

      { header: 'Tardanzas', key: 'tardanzas', width: 12 },
      { header: 'Minutos tarde', key: 'minutos_tarde_total', width: 14 },
      { header: 'Hora tarde', key: 'hora_tarde', width: 12 },

      { header: 'Días laborables', key: 'dias_laborables', width: 14 },
      { header: 'Días feriados', key: 'dias_feriados', width: 12 },

      { header: 'Días con asistencia', key: 'dias_con_asistencia', width: 16 },
      { header: 'Ausencias just.', key: 'ausencias_justificadas', width: 14 },
      { header: 'Ausencias injust.', key: 'ausencias_injustificadas', width: 16 },

      { header: 'Horario vigente desde', key: 'horario_vigente_desde', width: 18 },

      { header: 'Primer ingreso', key: 'primer_ingreso', width: 22 },
      { header: 'Última salida', key: 'ultima_salida', width: 22 },
    ];

    for (const r of result.data as ResumenRow[]) {
      ws.addRow({
        ranking: r.ranking ?? null,
        usuario: r.usuario,

        marcas_in: r.marcas_in,
        marcas_out: r.marcas_out,

        tardanzas: r.tardanzas,
        minutos_tarde_total: r.minutos_tarde_total,
        hora_tarde: this.minutosToHHMM(r.minutos_tarde_total),

        dias_laborables: r.dias_laborables,
        dias_feriados: r.dias_feriados,

        dias_con_asistencia: r.dias_con_asistencia,
        ausencias_justificadas: r.ausencias_justificadas,
        ausencias_injustificadas: r.ausencias_injustificadas,

        horario_vigente_desde: r.horario_vigente_desde,

        primer_ingreso: r.primer_ingreso,
        ultima_salida: r.ultima_salida,
      });
    }

    const header = ws.getRow(1);
    header.font = { bold: true };
    header.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="reporte_asistencias_resumen.xlsx"`,
    );

    await wb.xlsx.write(res);
    res.end();
  }

  // ==========================
  // PDF (Resumen)
  // ==========================
  @Roles('Gerencia', 'RRHH')
  @Get('resumen-pdf')
  async resumenPdf(
    @Res() res: Response,
    @Query('period') period?: string,
    @Query('ref') ref?: string,
    @Query('desde') desde?: string,
    @Query('hasta') hasta?: string,
    @Query('usuarioId') usuarioId?: string,
    @Query('sedeId') sedeId?: string,
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

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="reporte_asistencias_resumen.pdf"`,
    );

    const doc = new PDFDocument({ margin: 36, size: 'A4' });
    doc.pipe(res);

    const logoPath = path.join(process.cwd(), 'public', 'logo_negro.png');
    if (fs.existsSync(logoPath)) {
      doc.image(logoPath, doc.page.margins.left, 18, { width: 110 });
    }

    doc
      .font('Helvetica-Bold')
      .fontSize(16)
      .text('Reporte de Asistencias - Resumen', 0, 30, { align: 'center' });

    doc.moveDown(1.7);

    doc.font('Helvetica').fontSize(10);

    const periodoTxt = `Periodo: ${this.formatDatePEFromDateOnly(
      startDate,
    )} a ${this.formatDatePEFromDateOnly(endDate)}`;

    const filtrosTxt = `Filtros: usuarioId=${
      result.filtros.usuarioId ?? '-'
    } | sedeId=${result.filtros.sedeId ?? '-'}`;

    doc.text(periodoTxt, doc.page.margins.left, doc.y);
    doc.text(filtrosTxt, doc.page.margins.left, doc.y + 2);

    doc.moveDown(1.0);

    const rows: ResumenRow[] = result.data;

    const col = {
      rk: 26,
      usuario: 220,
      lab: 42,
      feri: 42,
      asis: 42,
      ausi: 50,
      tard: 42,
      minh: 52,
    };

    const tableWidth =
      col.rk +
      col.usuario +
      col.lab +
      col.feri +
      col.asis +
      col.ausi +
      col.tard +
      col.minh;

    const contentWidth =
      doc.page.width - doc.page.margins.left - doc.page.margins.right;

    const startX =
      doc.page.margins.left + Math.max(0, (contentWidth - tableWidth) / 2);

    let y = doc.y;

    const drawHeader = () => {
      doc.font('Helvetica-Bold').fontSize(9);

      doc.text('Rk', startX, y, { width: col.rk, align: 'center' });
      doc.text('Usuario', startX + col.rk, y, {
        width: col.usuario,
        align: 'left',
      });

      let x = startX + col.rk + col.usuario;
      doc.text('Lab.', x, y, { width: col.lab, align: 'center' });
      x += col.lab;
      doc.text('Feri.', x, y, { width: col.feri, align: 'center' });
      x += col.feri;
      doc.text('Asis.', x, y, { width: col.asis, align: 'center' });
      x += col.asis;
      doc.text('Aus.I', x, y, { width: col.ausi, align: 'center' });
      x += col.ausi;
      doc.text('Tard.', x, y, { width: col.tard, align: 'center' });
      x += col.tard;
      doc.text('Min.HH', x, y, { width: col.minh, align: 'center' });

      doc.moveTo(startX, y + 12).lineTo(startX + tableWidth, y + 12).stroke();

      y += 18;
      doc.font('Helvetica').fontSize(9);
    };

    const drawRow = (r: ResumenRow) => {
      doc.text(String(r.ranking ?? ''), startX, y, {
        width: col.rk,
        align: 'center',
      });
      doc.text(r.usuario, startX + col.rk, y, {
        width: col.usuario,
        align: 'left',
      });

      let x = startX + col.rk + col.usuario;
      doc.text(String(r.dias_laborables ?? 0), x, y, {
        width: col.lab,
        align: 'center',
      });
      x += col.lab;
      doc.text(String(r.dias_feriados ?? 0), x, y, {
        width: col.feri,
        align: 'center',
      });
      x += col.feri;
      doc.text(String(r.dias_con_asistencia ?? 0), x, y, {
        width: col.asis,
        align: 'center',
      });
      x += col.asis;
      doc.text(String(r.ausencias_injustificadas ?? 0), x, y, {
        width: col.ausi,
        align: 'center',
      });
      x += col.ausi;
      doc.text(String(r.tardanzas ?? 0), x, y, {
        width: col.tard,
        align: 'center',
      });
      x += col.tard;
      doc.text(this.minutosToHHMM(r.minutos_tarde_total ?? 0), x, y, {
        width: col.minh,
        align: 'center',
      });

      y += 14;
    };

    drawHeader();

    for (const r of rows) {
      if (y > doc.page.height - doc.page.margins.bottom - 18) {
        doc.addPage();
        y = doc.y;
        drawHeader();
      }
      drawRow(r);
    }

    doc.end();
  }

  // ==========================
  // DETALLE - EXCEL (ACTUALIZADO)
  // ==========================
  @Roles('Gerencia', 'RRHH')
  @Get('detalle-excel')
  async detalleExcel(
    @Res() res: Response,
    @Query('desde') desde?: string,
    @Query('hasta') hasta?: string,
    @Query('usuarioId') usuarioId?: string,
    @Query('sedeId') sedeId?: string,
  ) {
    if (!desde || !hasta) {
      throw new BadRequestException('Debe indicar desde y hasta');
    }

    if (!usuarioId && !sedeId) {
      const d1 = new Date(desde + 'T00:00:00');
      const d2 = new Date(hasta + 'T00:00:00');
      const diffDays = Math.floor((+d2 - +d1) / (1000 * 60 * 60 * 24)) + 1;
      if (!isFinite(diffDays) || diffDays <= 0) {
        throw new BadRequestException('Rango inválido');
      }
      if (diffDays > 31) {
        throw new BadRequestException(
          'Para descargar sin filtros (usuario/sede) el rango máximo permitido es 31 días.',
        );
      }
    }

    const params: any[] = [desde, hasta];
    const conds: string[] = [
      `a.fecha_hora >= $1::date`,
      `a.fecha_hora <  ($2::date + interval '1 day')`,
    ];

    let p = 3;
    if (usuarioId) {
      params.push(usuarioId);
      conds.push(`u.id = $${p}::uuid`);
      p++;
    }
    if (sedeId) {
      params.push(sedeId);
      conds.push(`u.sede_id = $${p}::uuid`);
      p++;
    }

    const where = conds.join(' AND ');

    const rows = await this.ds.query(
      `
      SELECT
        a.fecha_hora::date               AS fecha,
        TO_CHAR(a.fecha_hora, 'HH24:MI') AS hora,

        CASE a.tipo
          WHEN 'IN'  THEN 'ENTRADA'
          WHEN 'OUT' THEN 'SALIDA'
          ELSE COALESCE(a.tipo, '')
        END AS tipo,

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

        COALESCE(a.minutos_tarde, 0)      AS minutos_tarde,
        COALESCE(a.metodo, '')            AS metodo,
        COALESCE(a.estado_validacion, '') AS estado_validacion,

        (u.nombre || ' ' ||
         COALESCE(u.apellido_paterno,'') || ' ' ||
         COALESCE(u.apellido_materno,'')) AS empleado,
        u.numero_documento AS dni,

        COALESCE(s.nombre,'')  AS sede,
        COALESCE(ar.nombre,'') AS area
      FROM asistencias a
      JOIN usuarios u ON u.id = a.usuario_id
      LEFT JOIN sedes s ON s.id = u.sede_id
      LEFT JOIN areas ar ON ar.id = u.area_id
      WHERE ${where}
      ORDER BY empleado, fecha, hora
      `,
      params,
    );

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Detalle Asistencias');

    ws.columns = [
      { header: 'Fecha', key: 'fecha', width: 12 },
      { header: 'Hora', key: 'hora', width: 10 },
      { header: 'Empleado', key: 'empleado', width: 30 },
      { header: 'DNI', key: 'dni', width: 14 },
      { header: 'Sede', key: 'sede', width: 18 },
      { header: 'Área', key: 'area', width: 18 },
      { header: 'Tipo', key: 'tipo', width: 10 },
      { header: 'Evento', key: 'evento', width: 22 },
      { header: 'Min. tarde', key: 'minutos_tarde', width: 12 },
      { header: 'Método', key: 'metodo', width: 16 },
      { header: 'Estado', key: 'estado_validacion', width: 14 },
    ];

    rows.forEach((r: any) => ws.addRow(r));

    const header = ws.getRow(1);
    header.font = { bold: true };
    header.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="reporte_asistencias_detalle.xlsx"',
    );

    await wb.xlsx.write(res);
    res.end();
  }

  // ==========================
  // ✅ DETALLE - PDF (NUEVO)
  // ==========================
  @Roles('Gerencia', 'RRHH')
  @Get('detalle-pdf')
  async detallePdf(
    @Res() res: Response,
    @Query('desde') desde?: string,
    @Query('hasta') hasta?: string,
    @Query('usuarioId') usuarioId?: string,
    @Query('sedeId') sedeId?: string,
  ) {
    if (!desde || !hasta) {
      throw new BadRequestException('Debe indicar desde y hasta');
    }

    // performance igual que excel
    if (!usuarioId && !sedeId) {
      const d1 = new Date(desde + 'T00:00:00');
      const d2 = new Date(hasta + 'T00:00:00');
      const diffDays = Math.floor((+d2 - +d1) / (1000 * 60 * 60 * 24)) + 1;
      if (!isFinite(diffDays) || diffDays <= 0) {
        throw new BadRequestException('Rango inválido');
      }
      if (diffDays > 31) {
        throw new BadRequestException(
          'Para descargar sin filtros (usuario/sede) el rango máximo permitido es 31 días.',
        );
      }
    }

    const params: any[] = [desde, hasta];
    const conds: string[] = [
      `a.fecha_hora >= $1::date`,
      `a.fecha_hora <  ($2::date + interval '1 day')`,
    ];

    let p = 3;
    if (usuarioId) {
      params.push(usuarioId);
      conds.push(`u.id = $${p}::uuid`);
      p++;
    }
    if (sedeId) {
      params.push(sedeId);
      conds.push(`u.sede_id = $${p}::uuid`);
      p++;
    }

    const where = conds.join(' AND ');

    // ✅ EXACTO como en la imagen (solo columnas clave)
    const rows = await this.ds.query(
      `
      SELECT
        a.fecha_hora::date               AS fecha,
        TO_CHAR(a.fecha_hora, 'HH24:MI') AS hora,

        (u.nombre || ' ' ||
         COALESCE(u.apellido_paterno,'') || ' ' ||
         COALESCE(u.apellido_materno,'')) AS empleado,

        COALESCE(s.nombre,'')  AS sede,

        CASE a.tipo
          WHEN 'IN'  THEN 'ENTRADA'
          WHEN 'OUT' THEN 'SALIDA'
          ELSE COALESCE(a.tipo, '')
        END AS tipo,

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
        COALESCE(a.metodo, '')       AS metodo
      FROM asistencias a
      JOIN usuarios u ON u.id = a.usuario_id
      LEFT JOIN sedes s ON s.id = u.sede_id
      WHERE ${where}
      ORDER BY fecha, hora, empleado
      `,
      params,
    );

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="reporte_asistencias_detalle.pdf"`,
    );

    const doc = new PDFDocument({
      margin: 28,
      size: 'A4',
      layout: 'landscape',
    });
    doc.pipe(res);

    // Logo
    const logoPath = path.join(process.cwd(), 'public', 'logo_negro.png');
    if (fs.existsSync(logoPath)) {
      doc.image(logoPath, doc.page.margins.left, 14, { width: 95 });
    }

    doc
      .font('Helvetica-Bold')
      .fontSize(15)
      .fillColor('#111')
      .text('Asistencia - Detalle de marcajes', 0, 22, { align: 'center' });

    doc.moveDown(1.2);
    doc.font('Helvetica').fontSize(9).fillColor('#111');

    const periodoTxt = `Periodo: ${this.formatDatePEFromDateOnly(
      desde,
    )} a ${this.formatDatePEFromDateOnly(hasta)}`;
    const filtrosTxt = `Filtros: usuarioId=${usuarioId || '-'} | sedeId=${
      sedeId || '-'
    }`;

    doc.text(periodoTxt, { align: 'center' });
    doc.text(filtrosTxt, { align: 'center' });

    doc.moveDown(0.8);

    // ===== Tabla =====
    const col = {
      fecha: 70,
      hora: 46,
      empleado: 240,
      sede: 120,
      tipo: 70,
      evento: 190,
      tarde: 70,
      metodo: 110,
    };

    const tableWidth =
      col.fecha +
      col.hora +
      col.empleado +
      col.sede +
      col.tipo +
      col.evento +
      col.tarde +
      col.metodo;

    const contentWidth =
      doc.page.width - doc.page.margins.left - doc.page.margins.right;

    const startX =
      doc.page.margins.left + Math.max(0, (contentWidth - tableWidth) / 2);

    let y = doc.y;

    const paddingX = 4;
    const paddingY = 4;
    const minRowH = 20;

    const fmtFechaPE = (d: any) => {
      const dt = d instanceof Date ? d : new Date(d);
      const dd = String(dt.getDate()).padStart(2, '0');
      const mm = String(dt.getMonth() + 1).padStart(2, '0');
      const yy = dt.getFullYear();
      return `${dd}/${mm}/${yy}`;
    };

    const getTextHeight = (text: any, w: number, bold = false) => {
      doc.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(8.5);
      return doc.heightOfString(String(text ?? ''), {
        width: w - paddingX * 2,
        align: 'left',
      });
    };

    const drawCell = (
      text: any,
      x: number,
      w: number,
      yPos: number,
      h: number,
      opts?: { bold?: boolean; align?: 'left' | 'center'; header?: boolean },
    ) => {
      const bold = opts?.bold ?? false;
      const align = opts?.align ?? 'left';
      const header = opts?.header ?? false;

      if (header) {
        doc.save();
        doc.rect(x, yPos, w, h).fill('#f2f2f2');
        doc.restore();
      }

      doc
        .save()
        .lineWidth(0.6)
        .strokeColor('#888')
        .rect(x, yPos, w, h)
        .stroke()
        .restore();

      doc
        .font(bold ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(8.5)
        .fillColor('#111')
        .text(String(text ?? ''), x + paddingX, yPos + paddingY, {
          width: w - paddingX * 2,
          align,
        });
    };

    const drawHeader = () => {
      const h = 22;

      let x = startX;
      drawCell('Fecha', x, col.fecha, y, h, {
        bold: true,
        header: true,
        align: 'center',
      });
      x += col.fecha;

      drawCell('Hora', x, col.hora, y, h, {
        bold: true,
        header: true,
        align: 'center',
      });
      x += col.hora;

      drawCell('Empleado', x, col.empleado, y, h, {
        bold: true,
        header: true,
      });
      x += col.empleado;

      drawCell('Sede', x, col.sede, y, h, {
        bold: true,
        header: true,
        align: 'center',
      });
      x += col.sede;

      drawCell('Tipo', x, col.tipo, y, h, {
        bold: true,
        header: true,
        align: 'center',
      });
      x += col.tipo;

      drawCell('Evento', x, col.evento, y, h, {
        bold: true,
        header: true,
      });
      x += col.evento;

      drawCell('Min. tarde', x, col.tarde, y, h, {
        bold: true,
        header: true,
        align: 'center',
      });
      x += col.tarde;

      drawCell('Método', x, col.metodo, y, h, {
        bold: true,
        header: true,
        align: 'center',
      });

      y += h;
    };

    drawHeader();

    for (const r of rows) {
      const rowH = Math.max(
        minRowH,
        getTextHeight(fmtFechaPE(r.fecha), col.fecha) + paddingY * 2,
        getTextHeight(r.hora, col.hora) + paddingY * 2,
        getTextHeight(r.empleado, col.empleado) + paddingY * 2,
        getTextHeight(r.sede, col.sede) + paddingY * 2,
        getTextHeight(r.tipo, col.tipo) + paddingY * 2,
        getTextHeight(r.evento, col.evento) + paddingY * 2,
        getTextHeight(String(r.minutos_tarde ?? 0), col.tarde) + paddingY * 2,
        getTextHeight(r.metodo, col.metodo) + paddingY * 2,
      );

      if (y + rowH > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        y = doc.page.margins.top;
        drawHeader();
      }

      let x = startX;
      drawCell(fmtFechaPE(r.fecha), x, col.fecha, y, rowH, { align: 'center' });
      x += col.fecha;

      drawCell(r.hora, x, col.hora, y, rowH, { align: 'center' });
      x += col.hora;

      drawCell(r.empleado, x, col.empleado, y, rowH);
      x += col.empleado;

      drawCell(r.sede, x, col.sede, y, rowH, { align: 'center' });
      x += col.sede;

      drawCell(r.tipo, x, col.tipo, y, rowH, { align: 'center' });
      x += col.tipo;

      drawCell(r.evento, x, col.evento, y, rowH);
      x += col.evento;

      drawCell(String(r.minutos_tarde ?? 0), x, col.tarde, y, rowH, {
        align: 'center',
      });
      x += col.tarde;

      drawCell(r.metodo, x, col.metodo, y, rowH, { align: 'center' });

      y += rowH;
    }

    doc.end();
  }

  // ============================================
  // Reporte maestro usuarios (SIN TOCAR)
  // ============================================
  @Roles('Gerencia', 'RRHH')
  @Get('usuarios-excel')
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
      ORDER BY u.created_at DESC`,
    );

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Usuarios');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 36 },
      { header: 'Nombre', key: 'nombre', width: 20 },
      { header: 'Apellido paterno', key: 'apellido_paterno', width: 18 },
      { header: 'Apellido materno', key: 'apellido_materno', width: 18 },
      { header: 'Tipo doc.', key: 'tipo_documento', width: 10 },
      { header: 'N° documento', key: 'numero_documento', width: 16 },
      { header: 'Fecha nacimiento', key: 'fecha_nacimiento', width: 15 },
      { header: 'Rol', key: 'rol', width: 15 },
      { header: 'Sede', key: 'sede_nombre', width: 20 },
      { header: 'Área', key: 'area_nombre', width: 20 },
      { header: 'Email personal', key: 'email_personal', width: 28 },
      { header: 'Email institucional', key: 'email_institucional', width: 28 },
      { header: 'Teléfono', key: 'telefono_celular', width: 14 },
      { header: 'Estado', key: 'estado', width: 12 },
      { header: 'Fecha baja', key: 'fecha_baja', width: 18 },
    ];

    for (const u of rows) {
      worksheet.addRow({
        id: u.id,
        nombre: u.nombre,
        apellido_paterno: u.apellido_paterno,
        apellido_materno: u.apellido_materno,
        tipo_documento: u.tipo_documento,
        numero_documento: u.numero_documento,
        fecha_nacimiento: u.fecha_nacimiento ? new Date(u.fecha_nacimiento) : null,
        rol: u.rol,
        sede_nombre: u.sede_nombre,
        area_nombre: u.area_nombre,
        email_personal: u.email_personal,
        email_institucional: u.email_institucional,
        telefono_celular: u.telefono_celular,
        estado: u.activo ? 'ACTIVO' : 'INACTIVO',
        fecha_baja: u.fecha_baja ? new Date(u.fecha_baja) : null,
      });
    }

    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader('Content-Disposition', 'attachment; filename="usuarios.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  }

  // ============================================
  // ✅ Usuarios PDF (tal cual lo enviaste)
  // ============================================
  @Roles('Gerencia', 'RRHH')
  @Get('usuarios-pdf')
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
      ORDER BY u.created_at DESC
      `,
    );

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="reporte_usuarios.pdf"`,
    );

    const doc = new PDFDocument({
      margin: 36,
      size: 'A4',
      layout: 'landscape',
    });
    doc.pipe(res);

    const logoPath = path.join(process.cwd(), 'public', 'logo_negro.png');
    if (fs.existsSync(logoPath)) {
      doc.image(logoPath, doc.page.margins.left, 18, { width: 110 });
    }

    doc
      .font('Helvetica-Bold')
      .fontSize(16)
      .fillColor('#111')
      .text('Reporte de Usuarios', 0, 30, { align: 'center' });

    doc.moveDown(1.4);
    doc.font('Helvetica').fontSize(10).fillColor('#111');
    doc.text(`Generado: ${new Date().toLocaleString('es-PE')}`, {
      align: 'center',
    });
    doc.moveDown(0.8);

    const col = {
      nombre: 230,
      tipo: 55,
      doc: 90,
      sede: 120,
      area: 190,
      tel: 85,
    };

    const tableWidth =
      col.nombre + col.tipo + col.doc + col.sede + col.area + col.tel;

    const contentWidth =
      doc.page.width - doc.page.margins.left - doc.page.margins.right;

    const startX =
      doc.page.margins.left + Math.max(0, (contentWidth - tableWidth) / 2);

    let y = doc.y;

    const paddingX = 4;
    const paddingY = 5;
    const minRowH = 22;

    const getTextHeight = (text: string, w: number, bold = false) => {
      doc.font(bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(9);
      return doc.heightOfString(String(text ?? '-'), {
        width: w - paddingX * 2,
        align: 'left',
      });
    };

    const drawCell = (
      text: string,
      x: number,
      w: number,
      yPos: number,
      h: number,
      opts?: { bold?: boolean; align?: 'left' | 'center'; header?: boolean },
    ) => {
      const bold = opts?.bold ?? false;
      const align = opts?.align ?? 'left';
      const header = opts?.header ?? false;

      if (header) {
        doc.save();
        doc.rect(x, yPos, w, h).fill('#f2f2f2');
        doc.restore();
      }

      doc
        .save()
        .lineWidth(0.6)
        .strokeColor('#888')
        .rect(x, yPos, w, h)
        .stroke()
        .restore();

      doc
        .font(bold ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(9)
        .fillColor('#111')
        .text(String(text ?? '-'), x + paddingX, yPos + paddingY, {
          width: w - paddingX * 2,
          align,
        });
    };

    const drawHeader = () => {
      const h = 24;
      drawCell('Nombre completo', startX, col.nombre, y, h, {
        bold: true,
        header: true,
      });
      drawCell('Tipo doc.', startX + col.nombre, col.tipo, y, h, {
        bold: true,
        align: 'center',
        header: true,
      });
      drawCell('N° documento', startX + col.nombre + col.tipo, col.doc, y, h, {
        bold: true,
        align: 'center',
        header: true,
      });
      drawCell(
        'Sede',
        startX + col.nombre + col.tipo + col.doc,
        col.sede,
        y,
        h,
        { bold: true, align: 'center', header: true },
      );
      drawCell(
        'Área',
        startX + col.nombre + col.tipo + col.doc + col.sede,
        col.area,
        y,
        h,
        { bold: true, align: 'center', header: true },
      );
      drawCell(
        'Teléfono',
        startX + col.nombre + col.tipo + col.doc + col.sede + col.area,
        col.tel,
        y,
        h,
        { bold: true, align: 'center', header: true },
      );

      y += h;
    };

    drawHeader();

    for (const r of rows) {
      const hNombre = getTextHeight(r.nombre_completo, col.nombre);
      const hTipo = getTextHeight(r.tipo_doc, col.tipo);
      const hDoc = getTextHeight(r.numero_documento, col.doc);
      const hSede = getTextHeight(r.sede, col.sede);
      const hArea = getTextHeight(r.area, col.area);
      const hTel = getTextHeight(r.telefono, col.tel);

      const rowH = Math.max(
        minRowH,
        hNombre + paddingY * 2,
        hTipo + paddingY * 2,
        hDoc + paddingY * 2,
        hSede + paddingY * 2,
        hArea + paddingY * 2,
        hTel + paddingY * 2,
      );

      if (y + rowH > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        y = doc.page.margins.top;
        drawHeader();
      }

      drawCell(String(r.nombre_completo || '-'), startX, col.nombre, y, rowH);
      drawCell(String(r.tipo_doc || '-'), startX + col.nombre, col.tipo, y, rowH, {
        align: 'center',
      });
      drawCell(
        String(r.numero_documento || '-'),
        startX + col.nombre + col.tipo,
        col.doc,
        y,
        rowH,
        { align: 'center' },
      );
      drawCell(
        String(r.sede || '-'),
        startX + col.nombre + col.tipo + col.doc,
        col.sede,
        y,
        rowH,
        { align: 'center' },
      );
      drawCell(
        String(r.area || '-'),
        startX + col.nombre + col.tipo + col.doc + col.sede,
        col.area,
        y,
        rowH,
        { align: 'center' },
      );
      drawCell(
        String(r.telefono || '-'),
        startX + col.nombre + col.tipo + col.doc + col.sede + col.area,
        col.tel,
        y,
        rowH,
        { align: 'center' },
      );

      y += rowH;
    }

    doc.end();
  }
}
