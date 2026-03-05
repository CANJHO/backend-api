-- =========================================================
-- INIT SQL - Registro de Asistencia (basado en tu dump)
-- =========================================================

-- Extensiones necesarias
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
CREATE EXTENSION IF NOT EXISTS "pgcrypto";
CREATE EXTENSION IF NOT EXISTS unaccent;

-- Limpieza previa
DROP TABLE IF EXISTS "areas" CASCADE;
DROP TABLE IF EXISTS "asignaciones_punto" CASCADE;
DROP TABLE IF EXISTS "bitacora" CASCADE;
DROP TABLE IF EXISTS "dias_semana" CASCADE;
DROP TABLE IF EXISTS "puntos_trabajo" CASCADE;
DROP TABLE IF EXISTS "roles" CASCADE;
DROP TABLE IF EXISTS "sedes" CASCADE;
DROP TABLE IF EXISTS "usuario_excepciones" CASCADE;
DROP TABLE IF EXISTS "usuario_horarios" CASCADE;
DROP TABLE IF EXISTS "usuarios" CASCADE;
DROP TABLE IF EXISTS "asistencias" CASCADE;

-- ===================== AREAS =========================
CREATE TABLE "public"."areas" (
    "id" uuid DEFAULT gen_random_uuid() NOT NULL,
    "nombre" text NOT NULL,
    "activo" boolean DEFAULT true NOT NULL,
    "created_at" timestamptz DEFAULT now(),
    "updated_at" timestamptz DEFAULT now(),
    CONSTRAINT "areas_pkey" PRIMARY KEY ("id")
)
WITH (oids = false);

-- ===================== ASIGNACIONES_PUNTO =========================
CREATE TABLE "public"."asignaciones_punto" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "punto_id" uuid NOT NULL,
    "usuario_id" uuid NOT NULL,
    "fecha_inicio" timestamp NOT NULL,
    "fecha_fin" timestamp NOT NULL,
    "supervisor_id" uuid,
    "created_at" timestamp DEFAULT now() NOT NULL,
    "estado" text DEFAULT 'VIGENTE',
    CONSTRAINT "asignaciones_punto_pkey" PRIMARY KEY ("id"),
    CONSTRAINT "asignaciones_rango_chk" CHECK ((fecha_fin > fecha_inicio)),
    CONSTRAINT "chk_asignaciones_estado" CHECK ((estado = ANY (ARRAY['VIGENTE'::text, 'CERRADA'::text, 'ANULADA'::text])))
)
WITH (oids = false);

CREATE INDEX idx_asign_usuario_fecha ON public.asignaciones_punto USING btree (usuario_id, fecha_inicio, fecha_fin);
CREATE INDEX idx_asign_punto_fecha ON public.asignaciones_punto USING btree (punto_id, fecha_inicio, fecha_fin);
CREATE INDEX ix_asignaciones_usuario_rango ON public.asignaciones_punto USING btree (usuario_id, fecha_inicio, fecha_fin);
CREATE INDEX ix_asignaciones_punto ON public.asignaciones_punto USING btree (punto_id);

-- ===================== ASISTENCIAS =========================
CREATE TABLE "public"."asistencias" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "usuario_id" uuid NOT NULL,
    "fecha_hora" timestamp NOT NULL,
    "tipo" character varying(10) NOT NULL,
    "evidencia_url" text,
    "gps" jsonb,
    -- tolerancia_min ELIMINADO
    "metodo" character varying(30),
    "estado_validacion" character varying(20) DEFAULT 'aprobado',
    "device_id" character varying(100),
    "punto_id" uuid,
    "validacion_modo" text,
    "distancia_m" integer,
    "minutos_tarde" integer,
    CONSTRAINT "asistencias_pkey" PRIMARY KEY ("id"),
    CONSTRAINT "asistencias_tipo_check" CHECK (((tipo)::text = ANY ((ARRAY['IN'::character varying, 'OUT'::character varying])::text[]))),
    CONSTRAINT "asistencias_metodo_chk" CHECK (((metodo IS NULL) OR ((metodo)::text = ANY ((ARRAY['scanner_barras'::character varying, 'qr_fijo'::character varying, 'qr_dinamico'::character varying, 'manual_supervisor'::character varying])::text[])))),
    CONSTRAINT "asistencias_estado_chk" CHECK (((estado_validacion)::text = ANY ((ARRAY['aprobado'::character varying, 'pendiente'::character varying, 'rechazado'::character varying])::text[])))
)
WITH (oids = false);

CREATE INDEX idx_asistencias_usuario_fecha ON public.asistencias USING btree (usuario_id, fecha_hora);
CREATE INDEX ix_asistencias_usuario_fecha ON public.asistencias USING btree (usuario_id, fecha_hora);

-- ===================== BITACORA =========================
CREATE TABLE "public"."bitacora" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "usuario" character varying(150) NOT NULL,
    "accion" character varying(150) NOT NULL,
    "detalle" jsonb,
    "fecha_hora" timestamp DEFAULT now() NOT NULL,
    "ip" character varying(64),
    "usuario_id" uuid,
    CONSTRAINT "bitacora_pkey" PRIMARY KEY ("id")
)
WITH (oids = false);

-- ===================== DIAS_SEMANA =========================
CREATE TABLE "public"."dias_semana" (
    "codigo" smallint NOT NULL,
    "nombre" character varying(10) NOT NULL,
    CONSTRAINT "dias_semana_pkey" PRIMARY KEY ("codigo")
)
WITH (oids = false);

CREATE UNIQUE INDEX dias_semana_nombre_key ON public.dias_semana USING btree (nombre);

-- ===================== PUNTOS_TRABAJO =========================
CREATE TABLE "public"."puntos_trabajo" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "nombre" character varying(100) NOT NULL,
    "lat" numeric(10,6) NOT NULL,
    "lng" numeric(10,6) NOT NULL,
    "radio_m" integer DEFAULT '120' NOT NULL,
    "activo" boolean DEFAULT true NOT NULL,
    "created_at" timestamp DEFAULT now() NOT NULL,
    "sede_id" uuid,
    CONSTRAINT "puntos_trabajo_pkey" PRIMARY KEY ("id")
)
WITH (oids = false);

CREATE INDEX idx_puntos_trabajo_activo ON public.puntos_trabajo USING btree (activo);
CREATE INDEX idx_puntos_trabajo_sede ON public.puntos_trabajo USING btree (sede_id);

-- ===================== ROLES =========================
CREATE TABLE "public"."roles" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "nombre" character varying(50) NOT NULL,
    "permisos" jsonb DEFAULT '{}',
    CONSTRAINT "roles_pkey" PRIMARY KEY ("id")
)
WITH (oids = false);

CREATE UNIQUE INDEX roles_nombre_key ON public.roles USING btree (nombre);

-- ===================== SEDES =========================
CREATE TABLE "public"."sedes" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "nombre" character varying(100) NOT NULL,
    "lat" numeric(10,6),
    "lng" numeric(10,6),
    "activo" boolean DEFAULT true NOT NULL,
    "radio_m" integer DEFAULT '120' NOT NULL,
    CONSTRAINT "sedes_pkey" PRIMARY KEY ("id"),
    CONSTRAINT "sedes_radio_chk" CHECK ((radio_m > 0))
)
WITH (oids = false);

CREATE UNIQUE INDEX sedes_nombre_key ON public.sedes USING btree (nombre);

-- ===================== USUARIO_EXCEPCIONES =========================
CREATE TABLE "public"."usuario_excepciones" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "usuario_id" uuid NOT NULL,
    "fecha" date NOT NULL,
    "tipo" character varying(30) NOT NULL,
    "es_laborable" boolean NOT NULL,
    "hora_inicio" time without time zone,
    "hora_fin" time without time zone,
    "observacion" text,
    "creado_en" timestamp DEFAULT now() NOT NULL,
    CONSTRAINT "usuario_excepciones_pkey" PRIMARY KEY ("id")
)
WITH (oids = false);

CREATE INDEX ix_ue_usuario_fecha ON public.usuario_excepciones USING btree (usuario_id, fecha);

-- ===================== USUARIO_HORARIOS =========================
CREATE TABLE "public"."usuario_horarios" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "usuario_id" uuid NOT NULL,
    "dia_semana" smallint NOT NULL,
    "hora_inicio" time without time zone NOT NULL,
    "hora_fin" time without time zone NOT NULL,
    "hora_inicio_2" time without time zone,
    "hora_fin_2" time without time zone,
    "es_descanso" boolean DEFAULT false NOT NULL,
    "tolerancia_min" smallint DEFAULT '15' NOT NULL,
    "fecha_inicio" date DEFAULT CURRENT_DATE NOT NULL,
    "fecha_fin" date,
    "creado_en" timestamp DEFAULT now() NOT NULL,
    "actualizado_en" timestamp DEFAULT now() NOT NULL,
    CONSTRAINT "usuario_horarios_pkey" PRIMARY KEY ("id"),
    CONSTRAINT "chk_usuario_horarios_tramos" CHECK (
      (hora_inicio < hora_fin) AND (
        ((hora_inicio_2 IS NULL) AND (hora_fin_2 IS NULL)) OR
        ((hora_inicio_2 IS NOT NULL) AND (hora_fin_2 IS NOT NULL) AND (hora_inicio_2 < hora_fin_2))
      )
    )
)
WITH (oids = false);

CREATE INDEX ix_uh_usuario_dia ON public.usuario_horarios USING btree (usuario_id, dia_semana, fecha_inicio, COALESCE(fecha_fin, '9999-12-31'::date));
CREATE INDEX ix_uh_usuario_dia_vigencia ON public.usuario_horarios USING btree (usuario_id, dia_semana, fecha_inicio, COALESCE(fecha_fin, '9999-12-31'::date));

-- ===================== USUARIOS =========================
CREATE TABLE "public"."usuarios" (
    "id" uuid DEFAULT uuid_generate_v4() NOT NULL,
    "dni" character varying(20) NOT NULL,
    "nombre" character varying(150) NOT NULL,
    "rol_id" uuid NOT NULL,
    "sede_id" uuid,
    "activo" boolean DEFAULT true NOT NULL,
    "password_hash" character varying(255) NOT NULL,
    "meta" jsonb DEFAULT '{}',
    "created_at" timestamp DEFAULT now() NOT NULL,
    "tipo_documento" character varying(3),
    "numero_documento" character varying(20),
    "foto_perfil_url" text,
    "barcode_url" text,
    "qr_url" text,
    "code_scannable" text GENERATED ALWAYS AS (
        CASE
          WHEN ((tipo_documento)::text = 'DNI'::text) THEN ('D'::text || (numero_documento))
          WHEN ((tipo_documento)::text = 'CE'::text)  THEN ('C'::text || (numero_documento))
        END
    ) STORED,
    "area_id" uuid,
    "fecha_baja" timestamptz,
    "motivo_baja" text,
    CONSTRAINT "usuarios_pkey" PRIMARY KEY ("id"),
    CONSTRAINT "usuarios_tipo_documento_chk" CHECK (((tipo_documento)::text = ANY ((ARRAY['DNI'::character varying, 'CE'::character varying])::text[]))),
    CONSTRAINT "usuarios_numero_documento_chk" CHECK (
      (((tipo_documento)::text = 'DNI'::text) AND ((numero_documento)::text ~ '^[0-9]{8}$'::text)) OR
      (((tipo_documento)::text = 'CE'::text)  AND ((numero_documento)::text ~ '^[0-9]{9}$'::text))
    )
)
WITH (oids = false);

CREATE INDEX idx_usuarios_dni ON public.usuarios USING btree (dni);
CREATE UNIQUE INDEX usuarios_doc_unique ON public.usuarios USING btree (tipo_documento, numero_documento);
CREATE INDEX idx_usuarios_code_scannable ON public.usuarios USING btree (code_scannable);
CREATE UNIQUE INDEX ux_usuarios_code_scannable ON public.usuarios USING btree (code_scannable);
CREATE INDEX ix_usuarios_numdoc ON public.usuarios USING btree (tipo_documento, numero_documento);
CREATE INDEX ix_usuarios_activo ON public.usuarios USING btree (activo);

-- ===================== FOREIGN KEYS =========================

ALTER TABLE ONLY "public"."asignaciones_punto" 
  ADD CONSTRAINT "asignaciones_punto_punto_id_fkey"      FOREIGN KEY (punto_id)   REFERENCES puntos_trabajo(id) ON DELETE CASCADE NOT DEFERRABLE;

ALTER TABLE ONLY "public"."asignaciones_punto" 
  ADD CONSTRAINT "asignaciones_punto_supervisor_id_fkey" FOREIGN KEY (supervisor_id) REFERENCES usuarios(id) ON DELETE SET NULL NOT DEFERRABLE;

ALTER TABLE ONLY "public"."asignaciones_punto" 
  ADD CONSTRAINT "asignaciones_punto_usuario_id_fkey"    FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE NOT DEFERRABLE;

ALTER TABLE ONLY "public"."asistencias" 
  ADD CONSTRAINT "asistencias_punto_fk"                  FOREIGN KEY (punto_id)   REFERENCES puntos_trabajo(id) ON DELETE SET NULL NOT DEFERRABLE;

ALTER TABLE ONLY "public"."asistencias" 
  ADD CONSTRAINT "asistencias_usuario_id_fkey"           FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE NOT DEFERRABLE;

ALTER TABLE ONLY "public"."bitacora" 
  ADD CONSTRAINT "bitacora_usuario_fk"                   FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE SET NULL NOT DEFERRABLE;

ALTER TABLE ONLY "public"."puntos_trabajo" 
  ADD CONSTRAINT "puntos_trabajo_sede_fk"                FOREIGN KEY (sede_id)    REFERENCES sedes(id) ON DELETE SET NULL NOT DEFERRABLE;

ALTER TABLE ONLY "public"."usuario_excepciones" 
  ADD CONSTRAINT "usuario_excepciones_usuario_id_fkey"   FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE NOT DEFERRABLE;

ALTER TABLE ONLY "public"."usuario_horarios" 
  ADD CONSTRAINT "usuario_horarios_dia_semana_fkey"      FOREIGN KEY (dia_semana) REFERENCES dias_semana(codigo) NOT DEFERRABLE;

ALTER TABLE ONLY "public"."usuario_horarios" 
  ADD CONSTRAINT "usuario_horarios_usuario_id_fkey"      FOREIGN KEY (usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE NOT DEFERRABLE;

ALTER TABLE ONLY "public"."usuarios" 
  ADD CONSTRAINT "fk_usuarios_area"                      FOREIGN KEY (area_id)    REFERENCES areas(id) ON DELETE SET NULL NOT DEFERRABLE;

ALTER TABLE ONLY "public"."usuarios" 
  ADD CONSTRAINT "fk_usuarios_rol"                       FOREIGN KEY (rol_id)     REFERENCES roles(id) ON DELETE RESTRICT NOT DEFERRABLE;

ALTER TABLE ONLY "public"."usuarios" 
  ADD CONSTRAINT "fk_usuarios_sede"                      FOREIGN KEY (sede_id)    REFERENCES sedes(id) ON DELETE SET NULL NOT DEFERRABLE;

-- ===================== DATOS BÁSICOS =========================

INSERT INTO dias_semana (codigo, nombre) VALUES
  (1,'Lunes'),
  (2,'Martes'),
  (3,'Miércoles'),
  (4,'Jueves'),
  (5,'Viernes'),
  (6,'Sábado'),
  (7,'Domingo')
ON CONFLICT (codigo) DO NOTHING;

INSERT INTO roles (id, nombre, permisos)
VALUES
  (uuid_generate_v4(), 'Gerencia', '{}'::jsonb),
  (uuid_generate_v4(), 'RRHH', '{}'::jsonb),
  (uuid_generate_v4(), 'Empleado', '{}'::jsonb)
ON CONFLICT (nombre) DO NOTHING;

