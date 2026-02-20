-- =========================
-- Trayectorias: Supabase Setup
-- =========================
-- 1) Tablas
create table if not exists public.estudiantes (
  id_estudiante text primary key,
  dni text,
  apellido text,
  nombre text,
  anio_actual int,
  division text,
  turno text,
  activo boolean default true,
  observaciones text,
  orientacion text,
  egresado boolean default false,
  anio_egreso text,
  ciclo_egreso text,
  fecha_pase_egresados timestamptz
);

create table if not exists public.materias_catalogo (
  id_materia text primary key,
  nombre text not null,
  anio int,
  es_troncal boolean default false,
  orientacion text,
  egresado boolean default false,
  anio_egreso text
);

create table if not exists public.estado_por_ciclo (
  ciclo_lectivo text not null,
  id_estudiante text not null references public.estudiantes(id_estudiante) on delete cascade,
  id_materia text not null references public.materias_catalogo(id_materia),
  condicion_academica text,
  nunca_cursada boolean default false,
  situacion_actual text,
  motivo_no_cursa text,
  fecha_actualizacion timestamptz,
  usuario text,
  resultado_cierre text,
  ciclo_cerrado boolean default false,
  primary key (ciclo_lectivo, id_estudiante, id_materia)
);

create index if not exists idx_estado_ciclo on public.estado_por_ciclo(ciclo_lectivo);
create index if not exists idx_estado_ciclo_est on public.estado_por_ciclo(ciclo_lectivo, id_estudiante);

create table if not exists public.auditoria (
  id bigserial primary key,
  timestamp timestamptz default now(),
  ciclo_lectivo text,
  id_estudiante text,
  id_materia text,
  campo text,
  antes text,
  despues text,
  usuario text
);

create table if not exists public.egresados (
  id_estudiante text primary key,
  apellido text,
  nombre text,
  division text,
  turno text,
  ciclo_egreso text,
  fecha_pase_egresados timestamptz,
  observaciones text
);

create table if not exists public.materias_aprobadas_limpieza (
  ciclo_lectivo text not null,
  id_estudiante text not null,
  id_materia text not null,
  condicion_academica text,
  nunca_cursada boolean,
  situacion_actual text,
  motivo_no_cursa text,
  fecha_actualizacion timestamptz,
  usuario text,
  resultado_cierre text,
  ciclo_cerrado boolean,
  primary key (ciclo_lectivo, id_estudiante, id_materia)
);

-- 2) RPC: obtener ciclos (para no leer toda la tabla desde Apps Script)
create or replace function public.get_cycles()
returns table(ciclo_lectivo text)
language sql
stable
as $$
  select distinct ciclo_lectivo
  from public.estado_por_ciclo
  where coalesce(trim(ciclo_lectivo),'') <> ''
  order by
    case when ciclo_lectivo ~ '^\d+$' then lpad(ciclo_lectivo, 10, '0') else ciclo_lectivo end desc;
$$;

