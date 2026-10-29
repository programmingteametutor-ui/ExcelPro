-- ExcelPro cloud schema (Phase 4). Not yet wired to the app: Phase 1 persists to IndexedDB
-- through the same WorkbookStore interface (src/persistence/store.ts).
--
-- Design notes
-- * One row per sheet (workbook_sheets.sheet_state) so an edit rewrites only the changed sheet,
--   never the whole workbook.
-- * workbooks.revision implements optimistic concurrency: save_workbook() rejects a save whose
--   base revision is stale, mirroring ConflictError in the client.
-- * All access goes through Row Level Security. The client only ever uses the anon key; the
--   service-role key must never ship to the browser.

create extension if not exists pgcrypto;

create table public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  display_name text,
  avatar_url text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table public.workbooks (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users (id) on delete cascade,
  name text not null check (char_length(name) between 1 and 255),
  description text,
  workbook_state jsonb not null default '{}'::jsonb, -- sheet order, active sheet, named ranges
  file_format text not null default 'excelpro-v1',
  revision bigint not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  deleted_at timestamptz
);
create index workbooks_owner_idx on public.workbooks (owner_id, updated_at desc);

create table public.workbook_sheets (
  id uuid primary key default gen_random_uuid(),
  workbook_id uuid not null references public.workbooks (id) on delete cascade,
  client_sheet_id text not null,           -- stable id used by the client model
  sheet_name text not null check (char_length(sheet_name) between 1 and 31),
  sheet_order int not null,
  sheet_state jsonb not null,              -- cells + styles + layout (SheetDoc)
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workbook_id, client_sheet_id)
);
create index workbook_sheets_wb_idx on public.workbook_sheets (workbook_id, sheet_order);

create type public.permission_role as enum ('viewer', 'editor');

create table public.workbook_permissions (
  id uuid primary key default gen_random_uuid(),
  workbook_id uuid not null references public.workbooks (id) on delete cascade,
  user_id uuid not null references auth.users (id) on delete cascade,
  permission_role public.permission_role not null,
  created_at timestamptz not null default now(),
  unique (workbook_id, user_id)
);

create table public.workbook_versions (
  id uuid primary key default gen_random_uuid(),
  workbook_id uuid not null references public.workbooks (id) on delete cascade,
  version_number bigint not null,
  workbook_snapshot jsonb not null,
  created_by uuid references auth.users (id),
  created_at timestamptz not null default now(),
  unique (workbook_id, version_number)
);

create table public.workbook_activity (
  id bigint generated always as identity primary key,
  workbook_id uuid not null references public.workbooks (id) on delete cascade,
  user_id uuid references auth.users (id),
  action_type text not null,
  action_metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);
create index workbook_activity_wb_idx on public.workbook_activity (workbook_id, created_at desc);

-- ---------------------------------------------------------------- access helpers
create or replace function public.can_read_workbook(wb uuid) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from workbooks w where w.id = wb and w.owner_id = auth.uid())
      or exists (select 1 from workbook_permissions p where p.workbook_id = wb and p.user_id = auth.uid());
$$;

create or replace function public.can_write_workbook(wb uuid) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from workbooks w where w.id = wb and w.owner_id = auth.uid())
      or exists (select 1 from workbook_permissions p where p.workbook_id = wb and p.user_id = auth.uid() and p.permission_role = 'editor');
$$;

-- ---------------------------------------------------------------- RLS
alter table public.profiles enable row level security;
alter table public.workbooks enable row level security;
alter table public.workbook_sheets enable row level security;
alter table public.workbook_permissions enable row level security;
alter table public.workbook_versions enable row level security;
alter table public.workbook_activity enable row level security;

create policy "profiles: read own and collaborators" on public.profiles for select
  using (id = auth.uid() or exists (
    select 1 from workbook_permissions p join workbooks w on w.id = p.workbook_id
    where (w.owner_id = auth.uid() and p.user_id = profiles.id) or (p.user_id = auth.uid() and w.owner_id = profiles.id)));
create policy "profiles: update own" on public.profiles for update using (id = auth.uid()) with check (id = auth.uid());
create policy "profiles: insert own" on public.profiles for insert with check (id = auth.uid());

create policy "workbooks: read" on public.workbooks for select using (public.can_read_workbook(id));
create policy "workbooks: create own" on public.workbooks for insert with check (owner_id = auth.uid());
create policy "workbooks: edit" on public.workbooks for update using (public.can_write_workbook(id))
  with check (public.can_write_workbook(id));
create policy "workbooks: delete own" on public.workbooks for delete using (owner_id = auth.uid());

create policy "sheets: read" on public.workbook_sheets for select using (public.can_read_workbook(workbook_id));
create policy "sheets: write" on public.workbook_sheets for all using (public.can_write_workbook(workbook_id))
  with check (public.can_write_workbook(workbook_id));

create policy "permissions: read" on public.workbook_permissions for select
  using (user_id = auth.uid() or exists (select 1 from workbooks w where w.id = workbook_id and w.owner_id = auth.uid()));
create policy "permissions: owner manages" on public.workbook_permissions for all
  using (exists (select 1 from workbooks w where w.id = workbook_id and w.owner_id = auth.uid()))
  with check (exists (select 1 from workbooks w where w.id = workbook_id and w.owner_id = auth.uid()));

create policy "versions: read" on public.workbook_versions for select using (public.can_read_workbook(workbook_id));
create policy "versions: create" on public.workbook_versions for insert with check (public.can_write_workbook(workbook_id) and created_by = auth.uid());

create policy "activity: read" on public.workbook_activity for select using (public.can_read_workbook(workbook_id));
create policy "activity: append" on public.workbook_activity for insert with check (public.can_write_workbook(workbook_id) and user_id = auth.uid());

-- ---------------------------------------------------------------- atomic save with conflict check
-- p_sheets: [{client_sheet_id, sheet_name, sheet_order, sheet_state}] for changed sheets only.
create or replace function public.save_workbook(
  p_workbook_id uuid, p_base_revision bigint, p_name text, p_state jsonb,
  p_sheets jsonb, p_removed_sheet_ids text[]
) returns bigint
language plpgsql security invoker set search_path = public as $$
declare
  new_rev bigint;
begin
  update workbooks set revision = revision + 1, name = p_name, workbook_state = p_state, updated_at = now()
   where id = p_workbook_id and revision = p_base_revision
  returning revision into new_rev;
  if new_rev is null then
    raise exception 'conflict: workbook % changed since revision %', p_workbook_id, p_base_revision using errcode = '40001';
  end if;
  insert into workbook_sheets (workbook_id, client_sheet_id, sheet_name, sheet_order, sheet_state)
  select p_workbook_id, s->>'client_sheet_id', s->>'sheet_name', (s->>'sheet_order')::int, s->'sheet_state'
    from jsonb_array_elements(p_sheets) s
  on conflict (workbook_id, client_sheet_id) do update
    set sheet_name = excluded.sheet_name, sheet_order = excluded.sheet_order,
        sheet_state = excluded.sheet_state, updated_at = now();
  delete from workbook_sheets where workbook_id = p_workbook_id and client_sheet_id = any (p_removed_sheet_ids);
  insert into workbook_activity (workbook_id, user_id, action_type, action_metadata)
  values (p_workbook_id, auth.uid(), 'save', jsonb_build_object('revision', new_rev, 'sheets', jsonb_array_length(p_sheets)));
  return new_rev;
end;
$$;
