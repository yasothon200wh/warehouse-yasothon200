-- ============================================================
-- ปิดช่องโหว่ "แก้ไข/อัพโหลดคิวใน LG แล้วคิวทั้งวันหายถาวร" — น่าจะเป็นสาเหตุจริงของ
-- ข้อมูลวันที่ 11 ก.ย. 2026 หาย (ดู supabase-fix-sept11-14-data-loss.sql สำหรับเคส 12-14
-- ที่เป็นคนละสาเหตุ)
--
-- ── สาเหตุเดิม ──────────────────────────────────────────────
-- handleSetQueue() ใน App.jsx (เรียกทุกครั้งที่: แก้ไข 1 รายการในคิว, ลบ 1 รายการ, เพิ่ม
-- manual, หรืออัพโหลดไฟล์ Excel คิวใหม่ทั้งไฟล์จากหน้า LG) ทำ 2 คำสั่งแยกกันจาก client:
--   1) delete().neq("id","")   -- ลบคิวเดิมทั้งหมด
--   2) upsert(newQueue)        -- ใส่คิวใหม่เข้าไป
-- ไม่มี transaction ครอบ, ไม่มี guard เทียบจำนวนเดิม/ใหม่, ไม่มี audit log — ถ้าขั้นตอนที่ 2
-- ล้มเหลวหรือได้ข้อมูลไม่ครบ (เน็ตหลุด/ปิดแท็บ/ไฟล์ Excel ผิด/parse ไม่ครบ) คิวทั้งวันจะหาย
-- ถาวรทันทีโดยไม่มี error ที่ผู้ใช้เห็นชัดเจนและไม่มีทางสืบว่าเกิดอะไรขึ้น
--
-- ── วิธีแก้ ──────────────────────────────────────────────────
-- รวม delete+insert เป็น RPC เดียว (set_queue_tx) ทำใน transaction เดียวฝั่ง DB (ถ้า insert
-- fail จะ rollback การ delete ให้อัตโนมัติ ไม่ใช่ 2 คำสั่งแยกจาก client เหมือนเดิม) +
-- guard ปฏิเสธการเขียนทับที่ดูเหมือนผิดพลาด (คิวหายไปเกินครึ่ง/หายหมด) เว้นแต่จะยืนยันซ้ำ
-- (p_force) + log ทุกครั้งลง wh_queue_audit ให้ตรวจสอบย้อนหลังได้
-- ============================================================

-- 1) audit log แบบ append-only สำหรับ wh_queue (เหมือน wh_archive_audit)
create table if not exists wh_queue_audit (
  id            bigserial primary key,
  ts            timestamptz not null default now(),
  source        text        not null,   -- 'lg_edit' | 'lg_delete_row' | 'lg_manual_add' | 'lg_upload' | ...
  queue_before  int         not null default 0,
  queue_after   int,                     -- null ถ้าถูกกันไว้ (blocked_guard) หรือ error
  action        text        not null,    -- 'applied' | 'blocked_guard' | 'error'
  detail        text,
  ref_id        text
);
alter table wh_queue_audit enable row level security;
drop policy if exists "allow all" on wh_queue_audit;
create policy "allow all" on wh_queue_audit for all using (true) with check (true);
grant select on wh_queue_audit to anon, authenticated;

-- 2) ฟังก์ชันแทนที่ delete()+upsert() แยกกันของเดิม
create or replace function set_queue_tx(
  p_new_queue jsonb,
  p_source    text default 'unknown',
  p_force     boolean default false
) returns jsonb as $$
declare
  v_ref_id  text := encode(gen_random_bytes(8), 'hex');
  v_before  int;
  v_after   int;
  v_blocked boolean := false;
  v_detail  text;
begin
  perform pg_advisory_xact_lock(hashtext('wh_set_queue'));

  select count(*) into v_before from wh_queue;
  v_after := coalesce(jsonb_array_length(p_new_queue), 0);

  -- กันเคส "คิวเดิมมีของอยู่ แต่ข้อมูลใหม่ว่างเปล่าทั้งหมด" — ปกติการลบทีละแถวจะไม่มีทางทำให้
  -- ว่างกะทันหันแบบนี้เว้นแต่เป็นแถวสุดท้ายจริงๆ (กรณีนั้น p_force จาก client จะยืนยันซ้ำได้)
  if v_before >= 3 and v_after = 0 and not p_force then
    v_blocked := true;
    v_detail := format('ปฏิเสธ: คิวเดิมมี %s รายการ แต่ข้อมูลใหม่ว่างเปล่า (0 รายการ) — อาจเกิดจากไฟล์ผิด/อัพโหลดไม่สำเร็จ/เน็ตหลุดระหว่างบันทึก', v_before);
  -- กันเคส "ข้อมูลใหม่หายไปเกินครึ่งของเดิม" (เช่น ไฟล์ที่อัพโหลดผิดวัน/parse ไม่ครบ)
  elsif v_before >= 5 and v_after > 0 and v_after < v_before * 0.5 and not p_force then
    v_blocked := true;
    v_detail := format('ปฏิเสธ: คิวเดิมมี %s รายการ ข้อมูลใหม่เหลือ %s รายการ (ลดลงเกิน 50%%) — อาจเป็นไฟล์ผิด/ไม่ครบ', v_before, v_after);
  end if;

  if v_blocked then
    insert into wh_queue_audit (source, queue_before, queue_after, action, detail, ref_id)
    values (p_source, v_before, null, 'blocked_guard', v_detail, v_ref_id);
    return jsonb_build_object('applied', false, 'blocked', true, 'reason', v_detail,
                               'queue_before', v_before, 'queue_after', v_after, 'ref_id', v_ref_id);
  end if;

  delete from wh_queue where true;
  if p_new_queue is not null and jsonb_array_length(p_new_queue) > 0 then
    insert into wh_queue (id, data)
    select coalesce(elem->>'id', 'Q' || gen_random_uuid()), elem
    from jsonb_array_elements(p_new_queue) elem;
  end if;

  insert into wh_queue_audit (source, queue_before, queue_after, action, detail, ref_id)
  values (p_source, v_before, v_after, 'applied', null, v_ref_id);

  return jsonb_build_object('applied', true, 'blocked', false,
                             'queue_before', v_before, 'queue_after', v_after, 'ref_id', v_ref_id);
exception when others then
  insert into wh_queue_audit (source, queue_before, queue_after, action, detail, ref_id)
  values (p_source, coalesce(v_before, 0), null, 'error', sqlerrm, v_ref_id);
  raise;
end;
$$ language plpgsql security definer;

grant execute on function set_queue_tx(jsonb, text, boolean) to anon, authenticated;

-- 3) ตัดสิทธิ์ DELETE ตรงของ anon บน wh_queue — บังคับให้การ "ล้างคิวทั้งหมดแล้วใส่ใหม่"
--    ต้องผ่าน set_queue_tx เท่านั้นจากนี้ไป (ยังคง INSERT/UPDATE ไว้ตามเดิม เพราะมี code path
--    ที่ upsert รายแถวปกติอยู่ — recomputeQueueExitTimes ใน App.jsx — ไม่ได้แตะของจุดนั้น)
revoke delete on wh_queue from anon, authenticated;

-- ============================================================
-- ตรวจสอบหลังรัน
-- ============================================================
--   select * from wh_queue_audit order by ts desc limit 20;
--   -- ทดสอบว่า RPC เรียกได้ปกติ (ใช้คิวปัจจุบันจริงส่งกลับเข้าไปเหมือนเดิม ไม่เสี่ยง):
--   -- select set_queue_tx((select coalesce(jsonb_agg(data), '[]'::jsonb) from wh_queue), 'manual_test', false);
-- ============================================================
