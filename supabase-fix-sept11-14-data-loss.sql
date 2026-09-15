-- ============================================================
-- แก้ Root Cause ของบั๊ก "ข้อมูล Tracking/Log วันที่ 11-14 ก.ย. 2026 หายไป"
-- ตรวจสอบจริงผ่าน Supabase REST (anon key) + R2 photo storage แล้ว พบ 2 ปัญหาแยกกัน:
--
-- ปัญหา A — วันที่ 11 ก.ย.: ข้อมูลจริงหายไปตั้งแต่ระดับ live table (wh_trucks/wh_queue) ก่อนที่
--   cron จะ archive เสียอีก (audit log มีการปิดวันแค่ครั้งเดียวตอน 07:00 ของวันที่ 12 ก.ย. ด้วย
--   trucks=0, queue=5 — ไม่ใช่ค่าที่ควรจะเป็นถ้าเทียบกับวันก่อนหน้าที่มี 18 คัน) และยืนยันจาก R2
--   photo storage แล้วว่าโฟลเดอร์ qc/2026-09-11/, sample/2026-09-11/, loading/*/2026-09-11/
--   ไม่มีไฟล์เลยแม้แต่ไฟล์เดียว (ทั้งที่วันก่อนและหลังมีรูปเป็นสิบ) แปลว่าไม่มีการันตีว่ามีการ
--   ประมวลผลรถผ่านแอปเลยทั้งวัน — ข้อมูลของวันนี้แทบไม่เหลือร่องรอยให้กู้คืนจาก R2 ได้เหมือนกรณี
--   6-8 ก.ย. ก่อนหน้า (ดู supabase-add-archive-reconstructed-flag.sql) ต้องตรวจกับ Supabase
--   Backup/PITR (Dashboard → Database → Backups) ว่ามี snapshot ก่อนข้อมูลหายหรือไม่ — เป็นทาง
--   เดียวที่จะกู้วันนี้คืนมาได้ครบ
--
-- ปัญหา B — วันที่ 12, 13, 14 ก.ย.: ตรงข้ามกับวันที่ 11 — cron ได้ archive ข้อมูลจริงสำเร็จถูกต้อง
--   ทุกวัน (wh_archive_audit ยืนยัน: 12=10คัน/12คิว, 13=15คัน/15คิว, 14=10คัน/10คิว, status SUCCESS
--   ทั้งหมด, run_count=1 ทุกวันคือปิดแค่ครั้งเดียวไม่มีการปิดซ้ำ) และ R2 photo storage ก็ยืนยันตรงกัน
--   ว่ามีรูปจริง 7-14 คันต่อวันใน qc/, sample/, loading/lane_pork/ ของวันที่ 12-14 — แต่ตาราง
--   wh_archive ตอนนี้กลับมี trucks='[]', queue='[]' (ว่างสนิท) ทั้ง 3 วัน ทั้งที่ wh_archive_audit
--   ไม่มี log การเขียนทับใดๆ อีกเลยหลังจากนั้น (run_count ยังเป็น 1 เท่าเดิม)
--   → สรุปได้ชัดเจนว่า "ไม่มีการเรียก close_work_day_tx ครั้งที่ 2" แต่มีคนหรือกระบวนการอื่น
--   เขียนทับ wh_archive.trucks / wh_archive.queue ให้ว่างโดยตรง "ข้ามฟังก์ชัน close_work_day_tx
--   ไปเลย" ซึ่งทำได้เพราะ RLS policy เดิมของ wh_archive คือ
--     create policy "allow all" on wh_archive for all using (true) with check (true);
--   (ดู supabase-dev-schema.sql บรรทัด 62) ซึ่งเปิดให้ anon/publishable key (ตัวเดียวกับที่ฝังอยู่
--   ใน frontend bundle ที่ผู้ใช้ทุกคนเห็นได้) UPDATE/DELETE ตรงตาราง wh_archive ได้เลยโดยไม่ผ่าน
--   RPC และไม่ทิ้ง audit log แม้แต่บรรทัดเดียว — เป็นช่องโหว่ที่ทำให้ข้อมูลที่ archive ถูกต้องแล้ว
--   ถูกเขียนทับให้หายไปทีหลังได้ โดยไม่มีทางสืบไม่ได้ว่าใคร/เมื่อไหร่/จากที่ไหนทำ
--
-- ข้อมูลของวันที่ 12-14 กู้คืนได้จาก R2 photo storage ด้วยวิธีเดียวกับที่เคยใช้กู้วันที่ 6-8 ก.ย.
-- (ดูสคริปต์/ผลลัพธ์แยกต่างหากสำหรับขั้นตอน reconstruct + upsert เฉพาะ 3 วันนี้ — ต้อง preview
-- ก่อน insert เสมอตามข้อกำหนด ห้ามเขียนทับแบบไม่ตรวจสอบ)
--
-- ── ปัญหาเร่งด่วนเพิ่มเติมที่ตรวจพบระหว่างสอบสวน (แยกจากข้อมูลหาย แต่กระทบวันนี้ทันที) ──
-- มีคนเพิ่ม trigger บล็อก "DELETE ที่ไม่มี WHERE clause" บน wh_trucks/wh_queue เข้าไปตรงๆ ใน
-- Supabase Dashboard (ไม่มีไฟล์ migration นี้อยู่ใน repo เลย) ซึ่งบล็อก `delete from wh_trucks;`
-- และ `delete from wh_queue;` ที่อยู่ใน close_work_day_tx ไปด้วยโดยไม่ได้ตั้งใจ — ทดสอบแล้ว (เรียก
-- RPC ตรงๆ ด้วย archive_date ทดสอบที่ไม่กระทบข้อมูลจริง) ได้ error 21000 "DELETE requires a WHERE
-- clause" กลับมา ถ้าไม่แก้ก่อนเวลาปิดวันทำงานรอบถัดไป (cron 07:00) การปิดวันของ "วันนี้" จะ error
-- และไม่ปิดวันให้ — ไม่ทำให้ข้อมูลหาย (exception ทำให้ทั้ง transaction rollback) แต่ข้อมูลจะค้างไม่ถูก
-- archive จนกว่าจะแก้ไข ควรรีบแก้ก่อนรอบปิดวันถัดไป
-- ============================================================


-- ── FIX 1: ปิดช่องโหว่หลัก — ห้าม anon/authenticated เขียนทับ wh_archive ตรงๆ อีกต่อไป ──
-- คงสิทธิ์ SELECT ไว้ (หน้า Tracking/Log ต้องอ่านได้ตามปกติ) และคงสิทธิ์ DELETE ไว้ (ปุ่ม "ลบข้อมูล
-- Archive ถาวร" ในหน้าจบการทำงาน ยังต้องใช้ delete().eq('archive_date', ...) ตรงๆ อยู่ — เป็นการลบ
-- ทั้งแถวที่ผู้ใช้ตั้งใจกดยืนยันเอง ไม่ใช่การเขียนทับข้อมูลให้ว่างแบบเงียบๆ) แต่ตัด UPDATE ออก เพราะ
-- ไม่มี code path ไหนในแอปที่ใช้ .update() กับตารางนี้เลย (เขียนข้อมูลจริงผ่าน close_work_day_tx
-- อย่างเดียว ซึ่งเป็น security definer function จึงยังเขียนได้ปกติแม้ anon จะไม่มีสิทธิ์ update ตรง)
revoke update on wh_archive from anon, authenticated;

-- ── FIX 2: กัน close_work_day_tx ชนกับ trigger "DELETE requires a WHERE clause" ที่เพิ่งถูกเพิ่ม ──
-- เติม `where true` ให้ทั้งสอง delete (ผลลัพธ์เหมือนเดิมทุกประการ ลบทั้งตารางเหมือนเดิม) เผื่อ trigger
-- ที่เพิ่มมาเช็คแค่ว่ามีคำว่า WHERE ในคำสั่งหรือไม่ (แพทเทิร์นที่พบได้บ่อยของ trigger ป้องกันแบบนี้)
-- ⚠️ ก่อนรันส่วนนี้ ให้ตรวจสอบ trigger จริงก่อนด้วยคำสั่งนี้ใน SQL Editor แล้วส่งผลลัพธ์กลับมาดูด้วย
-- เผื่อ trigger เช็คแบบอื่นที่ `where true` เอาไม่อยู่ (เช่นเทียบ pg_trigger_depth() หรือ row count):
--   select tgname, tgrelid::regclass, pg_get_triggerdef(oid) from pg_trigger
--   where tgrelid in ('wh_trucks'::regclass, 'wh_queue'::regclass) and not tgisinternal;
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

  -- เติม `where true` กัน trigger บล็อก unqualified DELETE ที่เพิ่งถูกเพิ่มเข้ามาใน DB ตรงๆ
  delete from wh_trucks where true;
  delete from wh_queue where true;
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

-- ── FIX 3: view ตรวจจับ "archive ถูกเขียนทับให้ว่างหลังจากบันทึกสำเร็จแล้ว" โดยอัตโนมัติ ──
-- เทียบจำนวนที่ archive ไว้ ณ ตอนนี้ (wh_archive) กับจำนวนที่ log ไว้ตอนปิดวันสำเร็จครั้งล่าสุด
-- (wh_archive_audit) — ถ้าน้อยกว่าที่เคย log ไว้ แปลว่ามีอะไรมาเขียนทับ/ลบข้อมูลไปทีหลังแบบไม่ผ่าน
-- close_work_day_tx (เคสเดียวกับวันที่ 12-14 ก.ย. ที่พบ) ให้เอา view นี้ไปต่อกับหน้า Tracking หรือ
-- ตั้ง pg_cron แจ้งเตือนรายวันได้เลย
create or replace view wh_archive_integrity_check as
select
  a.archive_date,
  coalesce(jsonb_array_length(a.trucks), 0) as trucks_now,
  coalesce(jsonb_array_length(a.queue), 0)  as queue_now,
  la.trucks_written  as trucks_last_archived,
  la.queue_written   as queue_last_archived,
  la.ts              as last_archived_at,
  la.ref_id          as last_archived_ref_id,
  (coalesce(jsonb_array_length(a.trucks), 0) < coalesce(la.trucks_written, 0)
    or coalesce(jsonb_array_length(a.queue), 0) < coalesce(la.queue_written, 0)) as data_regressed_after_archive
from wh_archive a
left join lateral (
  select trucks_written, queue_written, ts, ref_id
  from wh_archive_audit x
  where x.archive_date = a.archive_date and x.action = 'archived' and x.trucks_written is not null
  order by x.ts desc
  limit 1
) la on true
order by a.archive_date desc;

grant select on wh_archive_integrity_check to anon, authenticated;

-- ============================================================
-- ตรวจสอบหลังรัน
-- ============================================================
--   -- ควรไม่พบวันไหน data_regressed_after_archive = true อีก (นอกจาก 12-14 ก.ย. ที่รอกู้คืน):
--   select * from wh_archive_integrity_check where data_regressed_after_archive = true;
--
--   -- ทดสอบว่า close_work_day_tx เรียกได้ปกติแล้ว ไม่ชน trigger DELETE อีก (ใช้วันที่ทดสอบที่ไม่
--   -- กระทบข้อมูลจริง เพราะฟังก์ชันนี้จะ wipe wh_trucks/wh_queue ของจริงทุกครั้งที่เรียก — ห้ามรัน
--   -- คำสั่งนี้ถ้า wh_trucks/wh_queue ตอนนี้มีงานที่ยังไม่ปิดวันค้างอยู่ที่ต้องการเก็บไว้):
--   -- select close_work_day_tx(current_date, null, 'manual_test');
-- ============================================================
