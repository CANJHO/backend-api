// feriados.service.ts
import { Injectable, Logger, OnModuleInit } from '@nestjs/common';
import { HttpService } from '@nestjs/axios';
import { firstValueFrom } from 'rxjs';
import { DataSource } from 'typeorm';
import { Cron } from '@nestjs/schedule';

type NagerHoliday = {
  date: string;        // "YYYY-MM-DD"
  localName: string;
  name: string;
  countryCode: string;
  types: string[];     // incluye "Public"
};

@Injectable()
export class FeriadosService implements OnModuleInit {
  private readonly logger = new Logger(FeriadosService.name);

  constructor(private http: HttpService, private ds: DataSource) {}

  // ✅ Al iniciar el backend, asegura que existan feriados del año actual
  async onModuleInit() {
    try {
      const year = new Date().getFullYear();

      // Si ya hay feriados para el año actual, no hace nada
      const [{ count }] = await this.ds.query(
        `SELECT COUNT(*)::int AS count FROM public.feriados WHERE anio = $1`,
        [year],
      );

      if ((count || 0) > 0) {
        this.logger.log(`Feriados ${year} ya existen (${count}). No se ejecuta sync inicial.`);
        return;
      }

      this.logger.log(`Tabla feriados sin datos para ${year}. Ejecutando sync inicial...`);
      await this.syncPeruYear(year);
      // opcional: también traer el siguiente año
      await this.syncPeruYear(year + 1);
    } catch (e) {
      this.logger.error('Error en sync inicial de feriados', e as any);
    }
  }

  // ✅ Auto-sync anual: 1 de enero a las 05:00 (hora Perú)
  @Cron('0 5 1 1 *', { timeZone: 'America/Lima' })
  async syncAutomaticoAnual() {
    const year = new Date().getFullYear();
    try {
      this.logger.log(`Cron: sincronizando feriados ${year} y ${year + 1}...`);
      await this.syncPeruYear(year);
      await this.syncPeruYear(year + 1);
      this.logger.log(`Cron OK: feriados ${year}/${year + 1} sincronizados.`);
    } catch (e) {
      this.logger.error('Cron ERROR sincronizando feriados', e as any);
    }
  }

  async syncPeruYear(year: number) {
    const url = `https://date.nager.at/api/v3/PublicHolidays/${year}/PE`;

    const { data } = await firstValueFrom(
      this.http.get<NagerHoliday[]>(url, { timeout: 15000 }),
    );

    // solo feriados "Public" (nacionales)
    const feriados = (data || []).filter(h => (h.types || []).includes('Public'));

    for (const h of feriados) {
      await this.ds.query(
        `
        INSERT INTO public.feriados (fecha, nombre, nombre_local, fuente, actualizado_en)
        VALUES ($1::date, $2, $3, 'Nager.Date', now())
        ON CONFLICT (fecha)
        DO UPDATE SET
          nombre = EXCLUDED.nombre,
          nombre_local = EXCLUDED.nombre_local,
          actualizado_en = now()
        `,
        [h.date, h.name, h.localName],
      );
    }

    return { year, total_api: data?.length ?? 0, total_public: feriados.length };
  }

  async listar(anio?: number) {
    if (anio) {
      return this.ds.query(
        `SELECT fecha, nombre, nombre_local, fuente
           FROM public.feriados
          WHERE anio = $1
          ORDER BY fecha`,
        [anio],
      );
    }
    return this.ds.query(
      `SELECT fecha, nombre, nombre_local, fuente
         FROM public.feriados
        ORDER BY fecha`,
    );
  }
}
