-- ============================================================
-- ทำเครื่องหมายวันที่ข้อมูลใน wh_archive เป็นข้อมูล "กู้คืนบางส่วน" (ไม่ใช่ของจริงต้นฉบับ)
-- ใช้กับกรณีกู้ข้อมูลวันที่ 6-8 ก.ย. 2026 จาก R2 photo timestamp หลัง archive เดิมถูกเขียนทับ
-- จากบั๊ก race condition — เพื่อให้หน้า Tracking โชว์ป้ายเตือนตลอดไปว่าวันนั้นข้อมูลไม่ครบ 100%
-- (ไม่ปิดบัง ตามที่ขอไว้ว่าห้ามซ่อน error/ทำเหมือนข้อมูลสมบูรณ์)
--
-- รันสคริปต์นี้ครั้งเดียวใน Supabase SQL Editor ของ project จริง
-- ============================================================

alter table wh_archive add column if not exists is_reconstructed boolean not null default false;
alter table wh_archive add column if not exists reconstructed_note text;

-- โผล่ในตาราง wh_daily_health ด้วย (ดู supabase-add-work-log-and-health-check.sql) ให้เห็นในโหมด
-- ช่วงวันที่/ทั้งหมดของหน้า Tracking เช่นกัน ไม่ใช่แค่ตอนเปิดดูรายวัน
-- ใช้ drop+create แทน create or replace เพราะคอลัมน์ใหม่แทรกกลาง ไม่ได้ต่อท้าย
-- (create or replace view เปลี่ยนตำแหน่ง/ชื่อคอลัมน์เดิมไม่ได้ ถือเป็น rename ไป error 42P16)
drop view if exists wh_daily_health;
create view wh_daily_health as
select
  a.archive_date,
  coalesce(jsonb_array_length(a.trucks), 0)      as trucks_count,
  coalesce(jsonb_array_length(a.queue), 0)       as queue_count,
  a.is_reconstructed,
  a.reconstructed_note,
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
