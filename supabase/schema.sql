-- 발표 리허설 룸 데이터베이스 설정
-- Supabase 프로젝트의 SQL Editor에 이 파일 전체를 붙여넣고 Run을 누르세요. 여러 번 실행해도 됩니다.

-- 1. 테이블 ------------------------------------------------------------
create table if not exists public.projects (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  created_at timestamptz not null default now(),
  default_min int not null default 20,
  default_sec int not null default 0,
  self_name text not null default '제안사',
  competitor_memo text not null default '',
  deck jsonb,              -- {kind:'pptx'|'pdf', path, url, name, size, at}
  pptx_name text,
  notes_count int,
  slide_count int not null default 0,
  proposal jsonb           -- {path, url, name, pages, chars, at}
);

create table if not exists public.slides (
  project_id uuid not null references public.projects(id) on delete cascade,
  no int not null,
  title text not null default '',
  body text not null default '',
  notes text not null default '',
  src text,
  pptx_index int,
  primary key (project_id, no)
);

create table if not exists public.scripts (
  project_id uuid not null references public.projects(id) on delete cascade,
  no int not null,
  body text not null default '',
  author text not null default '',
  src text,
  updated_at timestamptz not null default now(),
  primary key (project_id, no)
);

create table if not exists public.qa (
  id uuid primary key default gen_random_uuid(),
  project_id uuid not null references public.projects(id) on delete cascade,
  cat text not null default '',
  q text not null,
  a text not null default '',
  page text not null default '',
  rfp text not null default '',
  kw text not null default '',
  author text not null default '',
  src text not null default 'manual',
  created_at timestamptz not null default now(),
  updated_at timestamptz
);
create index if not exists qa_project_idx on public.qa(project_id);

create table if not exists public.proposal_text (
  project_id uuid primary key references public.projects(id) on delete cascade,
  body text not null default ''
);

-- 2. 접근 규칙: 링크를 아는 누구나 읽고 쓸 수 있음 -------------------------
do $$
declare t text;
begin
  foreach t in array array['projects','slides','scripts','qa','proposal_text'] loop
    execute format('alter table public.%I enable row level security', t);
    execute format('drop policy if exists "open access" on public.%I', t);
    execute format('create policy "open access" on public.%I for all to anon, authenticated using (true) with check (true)', t);
  end loop;
end $$;

-- 3. 파일 저장소 (발표자료 PPTX/PDF, 정성제안서 원본) ------------------------
insert into storage.buckets (id, name, public)
values ('files', 'files', true)
on conflict (id) do update set public = true;

drop policy if exists "files read" on storage.objects;
drop policy if exists "files write" on storage.objects;
drop policy if exists "files update" on storage.objects;
drop policy if exists "files delete" on storage.objects;
create policy "files read"   on storage.objects for select to anon, authenticated using (bucket_id = 'files');
create policy "files write"  on storage.objects for insert to anon, authenticated with check (bucket_id = 'files');
create policy "files update" on storage.objects for update to anon, authenticated using (bucket_id = 'files') with check (bucket_id = 'files');
create policy "files delete" on storage.objects for delete to anon, authenticated using (bucket_id = 'files');

-- 4. 실시간 반영 (다른 사람이 쓴 스크립트·질의가 바로 보이게) -----------------
do $$
declare t text;
begin
  foreach t in array array['projects','slides','scripts','qa'] loop
    if not exists (select 1 from pg_publication_tables where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = t) then
      execute format('alter publication supabase_realtime add table public.%I', t);
    end if;
  end loop;
end $$;
