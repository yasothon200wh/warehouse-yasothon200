-- ============================================================
-- ต่อยอดจาก supabase-fix-archive-race-condition.sql
-- เพิ่มรายละเอียดให้ wh_archive_audit ใช้เป็น "Log ภาพรวมการทำงาน" ได้ครบตามที่ต้องการ
-- (วันที่/เวลา, สถานะ, เวลาในการประมวลผล, Reference ID, จำนวนก่อน-หลัง) และเพิ่ม view
-- สำหรับ Data Health Check ในหน้า Tracking
--
-- รันสคริปต์นี้ครั้งเดียวใน Supabase SQL Editor ของ project จริง (ต้องรันหลัง
-- supabase-fix-archive-race-condition.sql ไปแล้วเท่านั้น)
-- ============================================================

-- 1) เพิ่มคอลัมน์ให้ wh_archive_audit — ใช้ ALTER แบบ additive ไม่กระทบของเดิม/ไม่ต้องย้ายข้อมูล
alter table wh_archive_audit add column if not exists status         text;   -- SUCCESS | WARNING | ERROR | NO_DATA
alter table wh_archive_audit add column if not exists processing_ms  int;    -- เวลาที่ใช้ประมวลผลทั้งหมด (ms)
alter table wh_archive_audit add column if not exists ref_id         text;   -- transaction id อ้างอิงต่อการเรียกหนึ่งครั้ง

-- 2) แก้ close_work_day_tx ให้บันทึกฟิลด์ใหม่ด้วย (คงพฤติกรรม/การ์ดเดิมทั้งหมดไว้)
create or replace function close_work_day_tx(
  p_archive_date date,
  p_new_queue    jsonb default null,
  p_source       text  default 'unknown'
) returns jsonb as $$
declare
  v_started        timestamptz := clock_timestamp();
  v_ref_id         text := encode(gen_random_bytes(8), 'hex');
  v_queue          jsonb;
  v_trucks         jsonb;
  v_queue_len      int;
  v_trucks_len     int;
  v_existing_queue_len  int;
  v_existing_trucks_len int;
  v_will_overwrite boolean := true;
  v_prev_total     int;
  v_warning        text := null;
  v_status         text;
begin
  perform pg_advisory_xact_lock(hashtext('wh_close_work_day'));

  select coalesce(jsonb_agg(data order by (data->>'seq')::numeric nulls last), '[]'::jsonb), count(*)
    into v_queue, v_queue_len
    from wh_queue;

  select coalesce(jsonb_agg(data), '[]'::jsonb), count(*)
    into v_trucks, v_trucks_len
    from wh_trucks;

  select coalesce(max(jsonb_array_length(trucks)), 0), coalesce(max(jsonb_array_length(queue)), 0)
    into v_existing_trucks_len, v_existing_queue_len
    from wh_archive where archive_date = p_archive_date;

  if (v_existing_trucks_len > 0 or v_existing_queue_len > 0)
     and v_trucks_len <= v_existing_trucks_len and v_queue_len <= v_existing_queue_len
     and (v_trucks_len < v_existing_trucks_len or v_queue_len < v_existing_queue_len) then
    v_will_overwrite := false;
  end if;

  select coalesce(jsonb_array_length(trucks), 0) + coalesce(jsonb_array_length(queue), 0)
    into v_prev_total
    from wh_archive where archive_date = p_archive_date - 1;
  if v_prev_total is not null and v_prev_total >= 5
     and (v_trucks_len + v_queue_len) < (v_prev_total * 0.2) then
    v_warning := format('พบข้อมูลลดลงผิดปกติ: วันก่อนหน้ารวม %s รายการ วันนี้เหลือ %s รายการ กรุณาตรวจสอบ',
                         v_prev_total, v_trucks_len + v_queue_len);
  end if;

  if v_will_overwrite then
    insert into wh_archive (archive_date, queue, trucks)
    values (p_archive_date, v_queue, v_trucks)
    on conflict (archive_date) do update
      set queue = excluded.queue, trucks = excluded.trucks;

    v_status := case
      when v_trucks_len = 0 and v_queue_len = 0 then 'NO_DATA'
      when v_warning is not null then 'WARNING'
      else 'SUCCESS'
    end;

    insert into wh_archive_audit (source, archive_date, trucks_before, queue_before, trucks_written, queue_written, action, detail, status, processing_ms, ref_id)
    values (p_source, p_archive_date, v_trucks_len, v_queue_len, v_trucks_len, v_queue_len, 'archived', v_warning, v_status,
            extract(milliseconds from clock_timestamp() - v_started)::int, v_ref_id);
  else
    insert into wh_archive_audit (source, archive_date, trucks_before, queue_before, trucks_written, queue_written, action, detail, status, processing_ms, ref_id)
    values (p_source, p_archive_date, v_trucks_len, v_queue_len, null, null, 'skipped_guard',
            format('มีข้อมูลเดิม trucks=%s queue=%s อยู่แล้ว ปฏิเสธการเขียนทับด้วยข้อมูลใหม่ที่น้อยกว่า trucks=%s queue=%s',
                   v_existing_trucks_len, v_existing_queue_len, v_trucks_len, v_queue_len),
            'PARTIAL', extract(milliseconds from clock_timestamp() - v_started)::int, v_ref_id);
  end if;

  delete from wh_trucks;
  delete from wh_queue;
  if p_new_queue is not null and jsonb_array_length(p_new_queue) > 0 then
    insert into wh_queue (id, data)
    select coalesce(elem->>'id', 'Q' || gen_random_uuid()), elem
    from jsonb_array_elements(p_new_queue) elem
    on conflict (id) do update set data = excluded.data;
  end if;

  return jsonb_build_object(
    'archived', v_will_overwrite,
    'archive_date', p_archive_date,
    'trucks_closed', v_trucks_len,
    'queue_closed', v_queue_len,
    'warning', v_warning,
    'ref_id', v_ref_id
  );
exception when others then
  insert into wh_archive_audit (source, archive_date, trucks_before, queue_before, action, detail, status, processing_ms, ref_id)
  values (p_source, p_archive_date, 0, 0, 'error', sqlerrm, 'ERROR',
          extract(milliseconds from clock_timestamp() - v_started)::int, v_ref_id);
  raise;
end;
$$ language plpgsql security definer;

-- 3) View สำหรับ Data Health Check ในหน้า Tracking — เทียบจำนวนที่ archive ไว้จริง (wh_archive)
--    กับ log ล่าสุดของวันนั้น (wh_archive_audit) วันไหนไม่มี audit แต่มี archive แปลว่าข้อมูลมาจาก
--    ก่อนติดตั้งระบบ log นี้ / วันไหนมี archive แต่ trucks+queue เป็น 0 ทั้งคู่ให้ตีเป็น NO_DATA
create or replace view wh_daily_health as
select
  a.archive_date,
  coalesce(jsonb_array_length(a.trucks), 0)      as trucks_count,
  coalesce(jsonb_array_length(a.queue), 0)       as queue_count,
  la.status                                       as last_status,
  la.source                                       as last_source,
  la.ts                                           as last_run_at,
  la.detail                                       as last_detail,
  (select count(*) from wh_archive_audit x where x.archive_date = a.archive_date) as run_count
from wh_archive a
left join lateral (
  select * from wh_archive_audit x
  where x.archive_date = a.archive_date
  order by x.ts desc limit 1
) la on true
order by a.archive_date desc;

grant select on wh_daily_health to anon, authenticated;

-- ============================================================
-- ตรวจสอบหลังรัน
-- ============================================================
--   select * from wh_daily_health limit 20;
--   select * from wh_archive_audit order by ts desc limit 20;
-- ============================================================
