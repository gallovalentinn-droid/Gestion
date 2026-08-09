-- ═══════════════════════════════════════════════════════════
-- Esquema de la base — pegar entero en el SQL Editor de Supabase
-- ═══════════════════════════════════════════════════════════

-- Saldos: se reemplaza entero cada vez que alguien importa el export.
create table if not exists saldos (
  id bigint generated always as identity primary key,
  socio int not null,
  nombre text not null,
  division text,
  categoria_social text,
  jefefam int,
  deporte text,
  periodo int,
  descri_concepto_liq text,
  deuda numeric not null default 0
);
create index if not exists idx_saldos_socio on saldos (socio);

-- Gestión: un registro por socio. Nunca se pisa al importar saldos.
create table if not exists gestion (
  socio int primary key,
  estado text,
  ultimo_reclamo timestamptz,
  pago numeric default 0,
  saldo_al_pagar numeric,
  actualizado_en timestamptz default now(),
  actualizado_por text
);

-- Historial: registro de cada cambio, no se borra nunca (salvo al "Borrar gestión").
create table if not exists gestion_log (
  id bigint generated always as identity primary key,
  socio int not null,
  fecha timestamptz default now(),
  estado text,
  por text,
  monto numeric
);
create index if not exists idx_log_socio on gestion_log (socio, fecha desc);

-- Comentarios de secretaría y comisión.
create table if not exists comentarios (
  id bigint generated always as identity primary key,
  socio int not null,
  fecha timestamptz default now(),
  por text,
  texto text not null
);
create index if not exists idx_com_socio on comentarios (socio, fecha desc);

-- ── Acceso ───────────────────────────────────────────────────
-- RLS activado con una política abierta: cualquiera con el link de la
-- página puede leer y escribir. No hay usuarios ni contraseñas todavía.
-- Es el mismo modelo de confianza que tenían con el Excel compartido.
-- Si más adelante quieren login real, se agrega Supabase Auth y se
-- cambian estas políticas para exigir sesión.

alter table saldos enable row level security;
alter table gestion enable row level security;
alter table gestion_log enable row level security;
alter table comentarios enable row level security;

create policy "acceso abierto saldos" on saldos for all using (true) with check (true);
create policy "acceso abierto gestion" on gestion for all using (true) with check (true);
create policy "acceso abierto log" on gestion_log for all using (true) with check (true);
create policy "acceso abierto comentarios" on comentarios for all using (true) with check (true);

-- ── Tiempo real ──────────────────────────────────────────────
-- Necesario para que un cambio hecho por una persona aparezca
-- en la pantalla de las demás sin que nadie recargue la página.
alter publication supabase_realtime add table gestion;
alter publication supabase_realtime add table gestion_log;
alter publication supabase_realtime add table comentarios;
alter publication supabase_realtime add table saldos;
