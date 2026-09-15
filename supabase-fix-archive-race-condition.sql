-- ============================================================
-- แก้ Root Cause ของบั๊ก "ข้อมูล Tracking วันที่ 6-8 ก.ย. หายไป"
-- รันสคริปต์นี้ครั้งเดียวใน Supabase SQL Editor ของ project จริง (ต้องรันด้วยสิทธิ์ owner/admin)
--
-- ── สาเหตุที่แท้จริง ──────────────────────────────────────────
-- มีจุด "ปิดวันทำงาน" (archive แถวปัจจุบันของ wh_queue/wh_trucks ไปที่ wh_archive
-- แล้วลบ wh_queue/wh_trucks ให้ว่างสำหรับวันใหม่) อยู่ 3 จุดที่ทำงานแยกกันโดยไม่รู้จักกัน:
--   1) pg_cron close_work_day() ฝั่ง server (ทุกชั่วโมง ตรงชั่วโมงตัดรอบ)
--   2) promoteIfDue() ฝั่ง client ใน App.jsx (ทุกแท็บ/เครื่องที่เปิดค้างไว้ เช็คทุก 60 วิ)
--   3) ปุ่ม "ล้างวันใหม่" (handleReset ใน App.jsx)
-- ทั้ง 3 จุดเขียนลง wh_archive ด้วย `upsert ... on conflict (archive_date) do update`
-- ซึ่งเป็นการ REPLACE ทั้งแถวทับของเดิม โดยไม่มี lock และไม่เช็คว่าแถวเดิมมีข้อมูลจริงอยู่แล้ว
-- หรือไม่ — ถ้าสองจุดนี้ทำงานชนกันในช่วงเวลาตัดรอบเดียวกัน (เกิดได้ง่ายเพราะข้อ 1 รันทุกวันอัตโนมัติ
-- และข้อ 2 รันบนทุกแท็บที่เปิดค้างพร้อมกันในโกดังที่มีหลายสถานี) ตัวที่ "เขียนทีหลัง" จะไปอ่าน
-- wh_queue/wh_trucks ที่ตัวแรกเพิ่ง wipe+ใส่คิววันใหม่เข้าไปแล้ว แล้วเอาไป upsert ทับ archive_date
-- ของ "วันที่เพิ่งปิด" ซ้ำ ทำให้ข้อมูลจริงของวันนั้นหายไป กลายเป็นข้อมูลของวันถัดไปแทน
-- (ตรวจสอบแล้วพบหลักฐานตรงนี้ในข้อมูลจริง: แถว archive_date=2026-09-06 มีคิวที่ field
--  "date" เป็น "7/9/2026" และแถว archive_date=2026-09-08 มีคิวที่ field "date" เป็น "9/9/2026")
--
-- ── วิธีแก้ ──────────────────────────────────────────────────
-- รวมการ "ปิดวันทำงาน" ทั้ง 3 จุดให้เรียกผ่านฟังก์ชันเดียว (close_work_day_tx) ที่:
--   - ล็อกด้วย pg_advisory_xact_lock ให้ทำทีละคำสั่งเท่านั้น (กันชนกัน)
--   - ก่อนเขียนทับ wh_archive ของวันที่นั้น เช็คก่อนว่าแถวเดิมมีข้อมูลจริง (trucks/queue ไม่ว่าง)
--     อยู่แล้วหรือไม่ ถ้ามี และข้อมูลใหม่ที่กำลังจะเขียนทับ "น้อยกว่า/ว่างกว่า" ของเดิม
--     → "ปฏิเสธการเขียนทับ" (เก็บของเดิมไว้) แล้ว log เป็น WARNING แทน
--   - log ทุกครั้งที่ทำงาน (สำเร็จ/ถูกกันไว้/error) ลง wh_archive_audit ให้ตรวจสอบย้อนหลังได้เสมอ
--     ว่า "ข้อมูลหายไปตรงไหน เมื่อไหร่ ใครเรียก(source)"
-- ============================================================

-- 1) ตาราง audit log แบบ append-only — ห้ามมีการ UPDATE/DELETE จากแอป (ใช้ตรวจสอบย้อนหลังเท่านั้น)
create table if not exists wh_archive_audit (
  id             bigserial primary key,
  ts             timestamptz not null default now(),
  source         text        not null,           -- 'cron' | 'client_promote' | 'manual_reset'
  archive_date   date        not null,
  trucks_before  int         not null default 0, -- จำนวนแถว wh_trucks ก่อนทำงาน (ของวันที่กำลังจะปิด)
  queue_before   int         not null default 0,
  trucks_written int,                             -- จำนวนที่ "เขียนจริง" ลง wh_archive (null ถ้าถูกกันไว้)
  queue_written  int,
  action         text        not null,            -- 'archived' | 'skipped_guard' | 'error'
  detail         text
);
alter table wh_archive_audit enable row level security;
drop policy if exists "allow all" on wh_archive_audit;
create policy "allow all" on wh_archive_audit for all using (true) with check (true);

-- 2) ฟังก์ชันปิดวันทำงานแบบปลอดภัย — ใช้แทนโค้ดปิดวันทำงานทั้งหมด (cron / client / ปุ่มรีเซ็ต)
--    p_new_queue = คิวที่จะใส่แทนสำหรับวันใหม่ (จาก advance-queue) หรือ null ถ้าจะเริ่มวันใหม่แบบว่าง
create or replace function close_work_day_tx(
  p_archive_date date,
  p_new_queue    jsonb default null,
  p_source       text  default 'unknown'
) returns jsonb as $$
declare
  v_queue          jsonb;
  v_trucks         jsonb;
  v_queue_len      int;
  v_trucks_len     int;
  v_existing_queue_len  int;
  v_existing_trucks_len int;
  v_will_overwrite boolean := true;
  v_prev_total     int;
  v_warning        text := null;
begin
  -- กันชนกัน: ให้ทำได้ทีละคำสั่งเท่านั้นทั้งระบบ (คำสั่งอื่นที่มาพร้อมกันจะรอคิวจนกว่า transaction นี้จะจบ)
  perform pg_advisory_xact_lock(hashtext('wh_close_work_day'));

  select coalesce(jsonb_agg(data order by (data->>'seq')::numeric nulls last), '[]'::jsonb), count(*)
    into v_queue, v_queue_len
    from wh_queue;

  select coalesce(jsonb_agg(data), '[]'::jsonb), count(*)
    into v_trucks, v_trucks_len
    from wh_trucks;

  -- aggregate เพื่อให้ได้แถวเดียวเสมอแม้ไม่มี archive ของวันนี้มาก่อน (coalesce ไม่ทำงานถ้าไม่มีแถวเลย)
  select coalesce(max(jsonb_array_length(trucks)), 0), coalesce(max(jsonb_array_length(queue)), 0)
    into v_existing_trucks_len, v_existing_queue_len
    from wh_archive where archive_date = p_archive_date;

  -- การ์ดสำคัญ: ถ้าของเดิมมีข้อมูลจริงอยู่แล้ว (ไม่ว่างทั้งคู่) และของใหม่ "ด้อยกว่า" ของเดิม
  -- (รถน้อยกว่าเดิม และคิวก็น้อยกว่าเดิม) ห้ามเขียนทับ — ของเดิมน่าเชื่อถือกว่าเสมอในกรณีนี้
  if (v_existing_trucks_len > 0 or v_existing_queue_len > 0)
     and v_trucks_len <= v_existing_trucks_len and v_queue_len <= v_existing_queue_len
     and (v_trucks_len < v_existing_trucks_len or v_queue_len < v_existing_queue_len) then
    v_will_overwrite := false;
  end if;

  -- ตรวจความผิดปกติ (สำหรับแจ้งเตือน ไม่ block การทำงาน): เทียบกับยอดรวมของวันก่อนหน้า
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

    insert into wh_archive_audit (source, archive_date, trucks_before, queue_before, trucks_written, queue_written, action, detail)
    values (p_source, p_archive_date, v_trucks_len, v_queue_len, v_trucks_len, v_queue_len, 'archived', v_warning);
  else
    insert into wh_archive_audit (source, archive_date, trucks_before, queue_before, trucks_written, queue_written, action, detail)
    values (p_source, p_archive_date, v_trucks_len, v_queue_len, null, null, 'skipped_guard',
            format('มีข้อมูลเดิม trucks=%s queue=%s อยู่แล้ว ปฏิเสธการเขียนทับด้วยข้อมูลใหม่ที่น้อยกว่า trucks=%s queue=%s',
                   v_existing_trucks_len, v_existing_queue_len, v_trucks_len, v_queue_len));
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
    'warning', v_warning
  );
exception when others then
  insert into wh_archive_audit (source, archive_date, trucks_before, queue_before, action, detail)
  values (p_source, p_archive_date, 0, 0, 'error', sqlerrm);
  raise;
end;
$$ language plpgsql security definer;

grant execute on function close_work_day_tx(date, jsonb, text) to anon, authenticated;

-- 3) แก้ close_work_day() (ที่ pg_cron เรียกทุกชั่วโมง) ให้เรียกผ่านฟังก์ชันปลอดภัยตัวเดียวกัน
--    แทนที่จะเขียน insert/delete ของตัวเองแยกต่างหากเหมือนเดิม
create or replace function close_work_day() returns void as $$
declare
  v_cutoff_hour  int := coalesce((select (value::text)::int from wh_settings where id = 'work_day_cutoff_hour'), 10);
  v_now_bkk      timestamp := now() at time zone 'Asia/Bangkok';
  v_archive_date date;
begin
  if extract(hour from v_now_bkk)::int <> v_cutoff_hour then
    return;
  end if;
  v_archive_date := v_now_bkk::date - 1;
  perform close_work_day_tx(v_archive_date, null, 'cron');
end;
$$ language plpgsql security definer;

-- cron schedule เดิม (close-work-day-10am) ยังใช้ได้ตามปกติ ไม่ต้องตั้งใหม่ เพราะแก้แค่ตัวฟังก์ชัน
-- ที่มันเรียกอยู่ข้างใน (select cron.schedule(...) เดิมจาก supabase-auto-close-10am.sql ไม่ต้องรันซ้ำ)

-- ============================================================
-- ตรวจสอบหลังรัน
-- ============================================================
-- ดู log การปิดวันทำงานทั้งหมด (ต้องเห็นทุกครั้งที่ cron/client/ปุ่มรีเซ็ตทำงาน จากนี้ไป):
--   select * from wh_archive_audit order by ts desc limit 50;
-- ทดสอบเรียกฟังก์ชันตรงๆ (ปลอดภัย ไม่ทำอะไรถ้า archive_date นั้นมีข้อมูลดีอยู่แล้วและคุณส่งข้อมูลว่างไป):
--   select close_work_day_tx('2020-01-01'::date, null, 'manual_test');
-- ============================================================
