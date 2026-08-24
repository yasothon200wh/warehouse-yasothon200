-- ============================================================
-- Auto-close work day ตามเวลาที่ตั้งไว้ใน Master Setting (Asia/Bangkok)
-- รันสคริปต์นี้ครั้งเดียวใน Supabase SQL Editor ของ project จริง
-- (ต้องรันด้วยสิทธิ์ owner/admin — anon key ของแอปทำแบบนี้ไม่ได้)
--
-- ⚠️ ถ้าเคยรันไฟล์นี้เวอร์ชันเก่าไปแล้ว (ที่ฮาร์ดโค้ดเวลาไว้ที่ 10:00 ตายตัว ไม่อ่าน
-- wh_settings เลย) ต้องรันไฟล์นี้ซ้ำอีกครั้งเพื่ออัปเดตทั้ง function และ cron schedule —
-- ของเดิมถ้าไปแก้ "เวลาตัดรอบวันทำงาน" ในหน้าตั้งค่าระบบเป็นชั่วโมงอื่น job นี้จะยังปิดงาน
-- ที่ 10:00 เหมือนเดิม ไม่ตรงกับที่แอปคำนวณให้ผู้ใช้เห็น (ดูปัญหาที่แก้ในคอมมิตนี้)
--
-- ทำสิ่งเดียวกับปุ่ม "ล้างวันใหม่" ใน Dashboard (handleReset ใน App.jsx):
--   1. archive แถวปัจจุบันของ wh_queue + wh_trucks ไปที่ wh_archive
--   2. ลบ wh_queue และ wh_trucks ให้ว่างสำหรับรอบใหม่
-- ปุ่ม "ล้างวันใหม่" ในแอปยังใช้งานได้ตามปกติ (เผื่อกดปิดงานนอกรอบ/เร็วกว่าเวลาตัดรอบ)
-- ============================================================

-- 1) เปิด extension pg_cron (ถ้ายังไม่เปิด)
--    ถ้ารันแล้ว error เรื่องสิทธิ์ ให้ไปเปิดผ่าน
--    Dashboard → Database → Extensions → ค้นหา "pg_cron" → Enable
create extension if not exists pg_cron;

-- 2) ฟังก์ชันปิดงาน — job นี้ถูกตั้ง (ดูข้อ 3) ให้ pg_cron เรียกทุกชั่วโมงตรงนาที 0
--    แต่ฟังก์ชันจะปิดงานจริงเฉพาะชั่วโมง (เวลา Bangkok) ที่ตรงกับค่า work_day_cutoff_hour
--    ล่าสุดใน wh_settings เท่านั้น (ถ้ายังไม่เคยตั้งค่าไว้ fallback เป็น 10 ให้ตรงกับ
--    default ใน src/lib/settings.js) — เปลี่ยนค่าตัดรอบในหน้า "ตั้งค่าระบบ" แล้วมีผลกับ
--    job นี้อัตโนมัติในรอบชั่วโมงถัดไปทันที ไม่ต้องมาแก้ไฟล์ SQL/cron ซ้ำอีก
--    archive_date คำนวณจาก "เมื่อวาน" ของวันที่ Bangkok เสมอ เพราะฟังก์ชันรันตรงต้นชั่วโมง
--    ตัดรอบพอดี (เวลาปัจจุบัน == cutoff hour) ข้อมูลที่กำลังจะ archive ตอนนี้คือของวันทำงาน
--    ที่เพิ่งปิดไป (เมื่อวาน) เสมอ
create or replace function close_work_day() returns void as $$
declare
  v_cutoff_hour  int := coalesce((select (value::text)::int from wh_settings where id = 'work_day_cutoff_hour'), 10);
  v_now_bkk      timestamp := now() at time zone 'Asia/Bangkok';
  v_archive_date date;
  v_queue        jsonb;
  v_trucks       jsonb;
begin
  if extract(hour from v_now_bkk)::int <> v_cutoff_hour then
    return;
  end if;

  v_archive_date := v_now_bkk::date - 1;

  select coalesce(jsonb_agg(data order by (data->>'seq')::numeric nulls last), '[]'::jsonb)
    into v_queue
    from wh_queue;

  select coalesce(jsonb_agg(data), '[]'::jsonb)
    into v_trucks
    from wh_trucks;

  insert into wh_archive (archive_date, queue, trucks)
  values (v_archive_date, v_queue, v_trucks)
  on conflict (archive_date) do update
    set queue  = excluded.queue,
        trucks = excluded.trucks;

  delete from wh_queue;
  delete from wh_trucks;
end;
$$ language plpgsql security definer;

-- 3) ตั้งเวลา — เรียกทุกชั่วโมงตรงนาที 0 (UTC ก็ได้ ไม่สำคัญแล้วเพราะฟังก์ชันเช็คเวลา
--    Bangkok เองข้างใน) ให้ฟังก์ชันตัดสินใจเองว่าถึงชั่วโมงตัดรอบจริงหรือยัง
--    รันซ้ำได้ปลอดภัย: ถ้ามี job ชื่อนี้อยู่แล้ว cron.schedule จะอัปเดตให้ ไม่สร้างซ้ำ
--    (ชื่อ job ยังคงเดิมว่า close-work-day-10am เพื่อไม่ต้อง unschedule ของเก่าแยก แม้ตอนนี้
--    จะไม่ได้ตายตัวที่ 10 โมงแล้วก็ตาม)
select cron.schedule(
  'close-work-day-10am',
  '0 * * * *',
  $$ select close_work_day(); $$
);

-- ============================================================
-- ตรวจสอบหลังรัน
-- ============================================================
-- ดูว่า job ถูกตั้งไว้จริง (schedule ควรเป็น '0 * * * *'):
--   select * from cron.job;
-- ดู log การรันแต่ละครั้ง (รันทุกชั่วโมงแต่จะปิดงานจริงแค่ชั่วโมงเดียวที่ตรง cutoff):
--   select * from cron.job_run_details order by start_time desc limit 30;
-- เช็คค่าตัดรอบที่ฟังก์ชันจะใช้ตอนนี้:
--   select coalesce((value::text)::int, 10) from wh_settings where id = 'work_day_cutoff_hour';
--
-- ⚠️ ห้ามรัน `select close_work_day();` ทดสอบตรง ๆ ในเวลาทำงานจริงตอนที่ตรงชั่วโมงตัดรอบพอดี
--   เพราะมันจะ archive + ลบ wh_queue/wh_trucks ทันทีเหมือนกดปุ่ม "ล้างวันใหม่" จริง
--   (นอกชั่วโมงตัดรอบเรียกได้ปลอดภัย เพราะฟังก์ชันจะ return เฉยๆ ไม่ทำอะไร)
--
-- ยกเลิก automation (ถ้าต้องการ):
--   select cron.unschedule('close-work-day-10am');
-- ============================================================
