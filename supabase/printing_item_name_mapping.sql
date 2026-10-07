-- Run in the Supabase SQL editor. The app must read every name pair to identify
-- unambiguous mappings; edits belong in the dashboard/SQL editor.
create table if not exists public.printing_item_name_mapping (
    m4_item_name text not null check (length(btrim(m4_item_name)) > 0),
    mjp_item_name text not null check (length(btrim(mjp_item_name)) > 0),
    approved boolean not null default false,
    primary key (m4_item_name, mjp_item_name)
);

create index if not exists printing_item_name_mapping_approved_idx
    on public.printing_item_name_mapping (approved, m4_item_name);

alter table public.printing_item_name_mapping enable row level security;
revoke all on public.printing_item_name_mapping from anon, authenticated;
grant select (m4_item_name, mjp_item_name, approved)
    on public.printing_item_name_mapping to anon;

drop policy if exists "Read approved printing name pairs" on public.printing_item_name_mapping;
drop policy if exists "Read printing name pairs" on public.printing_item_name_mapping;
create policy "Read printing name pairs"
    on public.printing_item_name_mapping for select to anon
    using (true);
