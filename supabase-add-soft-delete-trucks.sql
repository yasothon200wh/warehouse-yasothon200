-- ============================================================
-- Soft delete สำหรับ wh_trucks — กันลบข้อมูลรถทิ้งถาวรโดยไม่ตั้งใจ
-- (เช่นตอน "Merge รถซ้ำ" ใน Admin ที่ลบรถต้นทางทิ้งจริงหลัง merge — ถ้า merge พลาด/merge
-- ไม่ครบทุก field ข้อมูลจะหายถาวรทันทีไม่มีทางกู้ ตามที่ข้อ 7B ขอให้เลี่ยง DELETE ตรงๆ)
--
-- รันสคริปต์นี้ครั้งเดียวใน Supabase SQL Editor ของ project จริง (รันหลัง
-- supabase-fix-archive-race-condition.sql และ supabase-add-work-log-and-health-check.sql แล้ว)
-- ============================================================

alter table wh_trucks add column if not exists is_deleted boolean not null default false;
alter table wh_trucks add column if not exists deleted_at timestamptz;
alter table wh_trucks add column if not exists deleted_by text;

-- แก้ close_work_day_tx ให้ไม่นับรถที่ถูก soft-delete ไว้เข้าไปใน archive/สถิติของวันนั้น
-- (โครงสร้างเดิมทั้งหมดจาก supabase-add-work-log-and-health-check.sql คงไว้ แก้แค่จุด select
-- v_trucks/v_trucks_len ให้กรอง is_deleted ออก)
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
    from wh_trucks where is_deleted = false;

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

  -- ล้างตารางสำหรับรอบใหม่ (รวมแถวที่ soft-delete ไว้ด้วย เพราะ archive ของวันนี้เก็บภาพรวมไปแล้ว)
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

-- ============================================================
-- ตรวจสอบหลังรัน
--   select id, is_deleted, deleted_at, deleted_by from wh_trucks where is_deleted = true;
-- กู้รถที่ลบผิดคืน (ก่อนวันทำงานนั้นจะปิด):
--   update wh_trucks set is_deleted = false, deleted_at = null, deleted_by = null where id = '<truck id>';
-- ============================================================
