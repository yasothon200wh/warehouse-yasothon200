import React, { useState, useEffect, useRef, useCallback, useMemo } from "react";

const useIsMobile = () => {
  const [v, setV] = useState(() => window.innerWidth < 768);
  useEffect(() => {
    const h = () => setV(window.innerWidth < 768);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return v;
};

const useIsNarrow = () => {
  const [v, setV] = useState(() => window.innerWidth < 900);
  useEffect(() => {
    const h = () => setV(window.innerWidth < 900);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return v;
};
import * as XLSX from "xlsx";
import { supabase } from "./lib/supabase";
import { uploadToR2 } from "./lib/r2";
import { sendTeamsNotification } from "./utils/teams";
import { settings, saveSetting } from "./lib/settings";
import {
  laneAliases, saveLaneAlias, deleteLaneAlias,
  waitingReasons as waitingReasonPresets, addWaitingReason, deleteWaitingReason,
  basketTypes, saveBasketType, deleteBasketType,
  detailSources, saveDetailSource, deleteDetailSource, defaultDetailCols,
  lanes, saveLane,
  roles, saveRole,
  bays, addBay, saveBay, deleteBay,
} from "./lib/masterData";

// ─────────────────────────────────────────────────────────────────────────────
// FLOW:
// 1. LG Upload คิวรถ
// 2. คนขับสแกนเข้า  → status: "arrived"
// 3. Picking พิมพ์เบิกพัสดุ → status: "picking"
// 4. QC ตรวจอุณหภูมิก่อนเข้าแต่ละลาน (ทีละลาน ไม่ต้องครบ 3 พร้อมกัน)
// 5. ลานโหลด: QC ผ่านลานไหน → โหลดลานนั้นได้เลย (truck.qcLanes / truck.loadLanes)
// 6. Picking พิมพ์สรุปค่าย (โหลดแล้วอย่างน้อย 1 ลาน) → status: "summary_printed"
// 7. วางแผน ออก Invoice → status: "invoiced"
// ─────────────────────────────────────────────────────────────────────────────

// ลาน 3 ลานของระบบ (lane_parts/lane_head/lane_pork) — label/สี/emoji มาจาก wh_lanes
// (ดู src/lib/masterData.js) ค่า default ในโค้ดคือ 3 บรรทัดข้างบนเดิม id คงที่เสมอเพราะ
// ผูกกับ kiosk URL/routing ด้านล่างของไฟล์นี้โดยตรง
const STATUS_META = {
  arrived:         { label: "เข้าโรงงานแล้ว",   color: "#3b82f6", bg: "#dbeafe", step: 1 },
  picking:         { label: "กำลังโหลด",          color: "#f97316", bg: "#ffedd5", step: 2 },
  summary_printed: { label: "โหลดเสร็จ/สรุปแล้ว", color: "#10b981", bg: "#d1fae5", step: 3 },
  invoiced:        { label: "ออก Invoice แล้ว",   color: "#6b7280", bg: "#f3f4f6", step: 4 },
};

const FLOW_STEPS = [
  { key: "arrived",         label: "เข้าโรงงาน",  emoji: "🚛" },
  { key: "picking",         label: "โหลดสินค้า",  emoji: "📋" },
  { key: "summary_printed", label: "ใบสรุป",       emoji: "🖨️" },
  { key: "invoiced",        label: "Invoice",      emoji: "📄" },
];

const TODAY = new Date().toLocaleDateString("th-TH", { year: "numeric", month: "long", day: "numeric" });
// รูปแบบ D/M/YYYY (ไม่ padStart) ให้ตรงกับที่ parseQueueDateToISO/toDateStr ฝั่งอ่านคาดหวัง
// ก่อนหน้านี้เป็น const คำนวณครั้งเดียวตอนโหลดโมดูล ไม่มี cutoff เลย (ต่างจาก DATE_STR/
// cycleDateStr) — เปลี่ยนเป็นฟังก์ชันที่ตัดรอบด้วย settings.workDayCutoffHour เดียวกัน
const SHORT_DATE = () => {
  const d = new Date();
  if (d.getHours() < settings.workDayCutoffHour) d.setDate(d.getDate() - 1);
  return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
};
const TIME_NOW = () => new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit" });
const getStep = (status) => STATUS_META[status]?.step ?? 0;

const DATE_STR = () => {
  const d = new Date();
  if (d.getHours() < settings.workDayCutoffHour) d.setDate(d.getDate() - 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const safePlate = p => String(p).replace(/[^a-zA-Z0-9]/g, "") || "unknown";

const compressImage = file => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.readAsDataURL(file);
  reader.onerror = () => reject(new Error(`อ่านไฟล์รูปไม่สำเร็จ: ${file.name || ""}`));
  reader.onload = ev => {
    const img = new Image();
    img.onerror = () => reject(new Error(`ไฟล์รูปเสียหายหรือไม่รองรับ: ${file.name || ""}`));
    img.src = ev.target.result;
    img.onload = () => {
      const canvas = document.createElement("canvas");
      let { width, height } = img;
      const MAX_SIZE = 1200;
      if (width > height) {
        if (width > MAX_SIZE) { height *= MAX_SIZE / width; width = MAX_SIZE; }
      } else {
        if (height > MAX_SIZE) { width *= MAX_SIZE / height; height = MAX_SIZE; }
      }
      canvas.width = width; canvas.height = height;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL("image/jpeg", 0.5));
    };
  };
});

async function uploadPhotos(folder, plate, photos) {
  if (!photos || !photos.length) return [];
  const ts = Date.now();
  const urls = [];
  for (let i = 0; i < photos.length; i++) {
    const blob = await fetch(photos[i]).then(r => r.blob());
    const ext = blob.type.split("/")[1] || "jpg";
    const path = `${folder}/${DATE_STR()}/${safePlate(plate)}/${ts}_${i}.${ext}`;
    const url = await uploadToR2(path, blob);
    urls.push(url);
  }
  return urls;
}

// ─── GEOFENCE (พิกัด/รัศมีอ่านจาก settings.geofence) ──────────────────────────
function haversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function useGeofence() {
  const [state, setState] = useState({ status: "idle", distance: null, error: null });
  const watchRef = useRef(null);

  const start = useCallback(() => {
    if (!navigator.geolocation) {
      setState({ status: "error", distance: null, error: "เบราว์เซอร์ไม่รองรับ GPS" });
      return;
    }
    setState(s => ({ ...s, status: "loading" }));
    watchRef.current = navigator.geolocation.watchPosition(
      (pos) => {
        const d = haversineDistance(pos.coords.latitude, pos.coords.longitude, settings.geofence.lat, settings.geofence.lng);
        setState({ status: d <= settings.geofence.radiusM ? "inside" : "outside", distance: Math.round(d), error: null });
      },
      (err) => {
        const msgs = { 1: "กรุณาอนุญาตการเข้าถึงตำแหน่ง (Location)", 2: "ไม่สามารถหาตำแหน่งได้ กรุณาเปิด GPS", 3: "หมดเวลาหาตำแหน่ง กรุณาลองใหม่" };
        setState({ status: "error", distance: null, error: msgs[err.code] || err.message });
      },
      { enableHighAccuracy: true, maximumAge: 10000, timeout: 15000 }
    );
  }, []);

  useEffect(() => {
    return () => { if (watchRef.current !== null) navigator.geolocation.clearWatch(watchRef.current); };
  }, []);

  return { ...state, start };
}

// ─── QR CODE (SVG-based, no external lib) ────────────────────────────────────
// Minimal QR-like display using a Google Charts API image
const DRIVER_URL = typeof window !== "undefined"
  ? `${window.location.origin}${window.location.pathname}?mode=driver`
  : "";

const QRCodeDisplay = ({ url, size = 220 }) => (
  <div style={{ textAlign: "center" }}>
    <img
      src={`https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodeURIComponent(url)}`}
      alt="QR Code"
      width={size}
      height={size}
      style={{ borderRadius: 0, border: "3px solid #e5e7eb" }}
    />
  </div>
);

const QR_HOME_ITEM = { mode: "home", emoji: "🏠", label: "หน้ารวมการทำงาน", color: "#16a34a", bg: "#f0fdf4" };

const QR_ITEMS = [
  { mode: "driver",        emoji: "🚛", label: "คนขับ เช็คอิน",     color: "#111",    bg: "#f9fafb" },
  { mode: "dashboard_transport", emoji: "📊", label: "Dashboard ขนส่ง", color: "#0ea5e9", bg: "#f0f9ff" },
  { mode: "qc_parts",      emoji: "🌡️", label: "ลานโหลด ชิ้นส่วน",      color: "#0369a1", bg: "#f0f9ff" },
  { mode: "qc_head",       emoji: "🌡️", label: "ลานโหลด หัว/เครื่องใน", color: "#0369a1", bg: "#f0f9ff" },
  { mode: "qc_pork",       emoji: "🌡️", label: "ลานโหลด หมูซีก",       color: "#0369a1", bg: "#f0f9ff" },
  { mode: "loading_parts", emoji: "🥩", label: "Checker ชิ้นส่วน",      color: "#c2410c", bg: "#fff7ed" },
  { mode: "loading_head",  emoji: "🐷", label: "Checker หัว/เครื่องใน", color: "#7c3aed", bg: "#faf5ff" },
  { mode: "loading_pork",  emoji: "🐖", label: "Checker หมูซีก",       color: "#be123c", bg: "#fff1f2" },
  { mode: "sample_parts",  emoji: "📷", label: "QC ชิ้นส่วน",          color: "#0d9488", bg: "#f0fdfa" },
  { mode: "sample_head",   emoji: "📷", label: "QC หัว/เครื่องใน",     color: "#0d9488", bg: "#f0fdfa" },
  { mode: "sample_pork",   emoji: "📷", label: "QC หมูซีก",           color: "#0d9488", bg: "#f0fdfa" },
];

const saveImageToDevice = async (url, fileName) => {
  let blob;
  try {
    blob = await (await fetch(url)).blob();
  } catch {
    window.open(url, "_blank");
    return;
  }
  const file = new File([blob], fileName, { type: blob.type || "image/jpeg" });
  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file] }); } catch { /* user cancelled share sheet */ }
    return;
  }
  const blobUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = blobUrl; a.download = fileName;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(blobUrl);
};

const saveQrImage = async (url, mode) => {
  try {
    const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=600x600&data=${encodeURIComponent(url)}`;
    const blob = await (await fetch(qrUrl)).blob();
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = blobUrl; a.download = `qr-${mode}.png`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(blobUrl);
  } catch {
    window.open(`https://api.qrserver.com/v1/create-qr-code/?size=600x600&data=${encodeURIComponent(url)}`, "_blank");
  }
};

const escapeHtml = s => String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

// พิมพ์ QR "หน้ารวมการทำงาน" เป็นไฟล์/หน้าเดียว ตามเทมเพลตป้ายติดหน้าโรงงาน
// (ชื่อโรงงานดึงจาก settings.facilityName สด ๆ — ตั้งค่าที่ Master Setting)
const printHomeQrTemplate = (url) => {
  const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=${encodeURIComponent(url)}`;
  const facilityName = escapeHtml(settings.facilityName);
  const safeUrl = escapeHtml(url);
  const html = `<!doctype html>
<html lang="th"><head><meta charset="utf-8" /><title>QR หน้ารวมการทำงาน</title>
<style>
  @page { size: A4; margin: 0; }
  * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Sarabun', sans-serif; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
  html, body { width: 100%; height: 100%; }
  body { display: flex; background: #fdf3f9; }
  .card { flex: 1; display: flex; flex-direction: column; border: 16px solid #f3d9ea; box-sizing: border-box; position: relative; overflow: hidden; min-height: 0; }
  .header { flex-shrink: 0; background: linear-gradient(135deg, #5a1a5c 0%, #3d1140 100%); padding: 26px 36px; display: flex; align-items: center; gap: 16px; }
  .accent { flex-shrink: 0; width: 6px; height: 46px; border-radius: 4px; background: #f0c94a; }
  .header .icon { font-size: 38px; line-height: 1; }
  .header h1 { color: #fff; font-size: 40px; font-weight: 900; letter-spacing: 0.5px; }
  .goldline { flex-shrink: 0; height: 6px; background: linear-gradient(90deg, #f0c94a, #d89b1f, #f0c94a); }
  .content { flex: 1; min-height: 0; overflow: hidden; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 26px 55px; gap: 14px; }
  .qrframe { flex-shrink: 0; background: #fff; border: 2px solid #f0c94a; border-radius: 22px; padding: 32px; box-shadow: 0 8px 24px rgba(74, 21, 75, 0.18); }
  .qr { width: 560px; height: 560px; display: block; }
  .urlbox { flex-shrink: 0; width: 100%; max-width: 628px; background: #f3f4f6; border: 1.5px dashed #c9c9d4; border-radius: 10px; padding: 5px 17px; font-size: 16px; color: #4a154b; word-break: break-all; font-family: monospace; text-align: center; font-weight: 700; }
  .howto-title { flex-shrink: 0; font-weight: 900; font-size: 24px; color: #4a154b; display: flex; align-items: center; gap: 10px; }
  .howto-title::before, .howto-title::after { content: ""; height: 3px; width: 32px; background: #f0c94a; border-radius: 2px; }
  .steps { flex-shrink: 0; width: 100%; max-width: 680px; display: flex; flex-direction: row; align-items: stretch; gap: 16px; }
  .step { flex: 1; min-width: 0; display: flex; align-items: center; justify-content: center; gap: 12px; background: #f6eefb; border: 1.5px solid #e7d3ef; border-radius: 14px; padding: 14px 12px; }
  .num { flex-shrink: 0; width: 38px; height: 38px; border-radius: 50%; background: linear-gradient(135deg, #f5d76e, #d89b1f); color: #3d1140; font-weight: 900; display: flex; align-items: center; justify-content: center; font-size: 19px; box-shadow: 0 2px 6px rgba(0,0,0,0.2); }
  .steptext { font-weight: 900; font-size: 23px; line-height: 1.2; color: #3d1140; text-align: center; }
  .footer { flex-shrink: 0; background: #f3d9ea; padding: 8px 10px 12px; text-align: center; position: relative; }
  .footer::before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 4px; background: linear-gradient(90deg, #f0c94a, #d89b1f, #f0c94a); }
  .footer .fname { font-weight: 900; font-size: 23px; color: #3d1140; }
  .footer .sub { font-size: 14px; color: #8a5a86; margin-top: 4px; font-weight: 600; }
</style></head>
<body>
  <div class="card">
    <div class="header"><div class="accent"></div><span class="icon">🏭</span><h1>- หน้าหลักระบบโหลดจ่าย -</h1></div>
    <div class="goldline"></div>
    <div class="content">
      <div class="qrframe"><img class="qr" src="${qrSrc}" alt="QR" /></div>
      <div class="urlbox">${safeUrl}</div>
      <div class="howto-title">วิธีการใช้งาน</div>
      <div class="steps">
        <div class="step"><div class="num">1</div><div class="steptext">สแกน<br>QR CODE</div></div>
        <div class="step"><div class="num">2</div><div class="steptext">แสดง<br>หน้าหลัก</div></div>
        <div class="step"><div class="num">3</div><div class="steptext">เลือก<br>จุดงาน</div></div>
      </div>
    </div>
    <div class="footer">
      <div class="fname">${facilityName}</div>
      <div class="sub">PLP – Process Improvement</div>
    </div>
  </div>
  <script>
    var img = document.querySelector('.qr');
    function go() { setTimeout(function () { window.print(); }, 150); }
    if (img.complete) go(); else img.onload = go;
  </script>
</body></html>`;
  const w = window.open("", "_blank", "width=650,height=920");
  if (!w) return;
  w.document.open();
  w.document.write(html);
  w.document.close();
};

// พิมพ์ QR "คนขับ เช็คอิน" เป็นไฟล์/หน้าเดียว ตามเทมเพลตป้ายติดหน้าโรงงาน (โทนเขียว)
// ป้ายระยะ geofence ตรึงไว้ที่ 200 เมตรเสมอ (ไม่ตามค่าจริงที่ตั้งใน Master Setting)
const printDriverQrTemplate = (url) => {
  const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=${encodeURIComponent(url)}`;
  const facilityName = escapeHtml(settings.facilityName);
  const safeUrl = escapeHtml(url);
  const radiusM = 200;
  const html = `<!doctype html>
<html lang="th"><head><meta charset="utf-8" /><title>QR คนขับ เช็คอิน</title>
<style>
  @page { size: A4; margin: 0; }
  * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Sarabun', sans-serif; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
  html, body { width: 100%; height: 100%; }
  body { display: flex; background: #f2fbf3; }
  .card { flex: 1; display: flex; flex-direction: column; border: 16px solid #cdeecf; box-sizing: border-box; position: relative; overflow: hidden; min-height: 0; }
  .header { flex-shrink: 0; background: linear-gradient(135deg, #16803d 0%, #0f5c2c 100%); padding: 26px 36px; display: flex; align-items: center; gap: 16px; }
  .accent { flex-shrink: 0; width: 6px; height: 46px; border-radius: 4px; background: #f0c94a; }
  .header .icon { font-size: 38px; line-height: 1; }
  .header h1 { color: #fff; font-size: 44px; font-weight: 900; letter-spacing: 0.5px; }
  .header h1 u { text-decoration: none; border-bottom: 4px solid #f0c94a; padding-bottom: 2px; }
  .goldline { flex-shrink: 0; height: 6px; background: linear-gradient(90deg, #f0c94a, #d89b1f, #f0c94a); }
  .content { flex: 1; min-height: 0; overflow: hidden; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 10px 55px; gap: 10px; }
  .infobox { flex-shrink: 0; width: 100%; max-width: 628px; background: #eef2f6; border: 1.5px solid #dbe3ea; border-radius: 12px; padding: 8px 20px; text-align: center; }
  .infobox .line1 { font-weight: 900; font-size: 25px; color: #1f2937; display: flex; align-items: center; justify-content: center; gap: 10px; }
  .infobox .line2 { font-size: 20px; color: #4b5563; margin-top: 4px; font-weight: 600; }
  .qrframe { flex-shrink: 0; background: #fff; border: 2px solid #f0c94a; border-radius: 22px; padding: 30px; box-shadow: 0 8px 24px rgba(15, 92, 44, 0.18); }
  .qr { width: 520px; height: 520px; display: block; }
  .urlbox { flex-shrink: 0; width: 100%; max-width: 628px; background: #f3f4f6; border: 1.5px dashed #c9c9d4; border-radius: 10px; padding: 5px 17px; font-size: 16px; color: #0f5c2c; word-break: break-all; font-family: monospace; text-align: center; font-weight: 700; }
  .howto-title { flex-shrink: 0; font-weight: 900; font-size: 24px; color: #0f5c2c; }
  .steps { flex-shrink: 0; width: 100%; max-width: 720px; display: flex; flex-direction: row; align-items: stretch; gap: 16px; }
  .step { flex: 1; min-width: 0; display: flex; align-items: center; justify-content: center; gap: 10px; background: #eafbee; border: 1.5px solid #cdeecf; border-radius: 14px; padding: 14px 6px; }
  .num { flex-shrink: 0; width: 42px; height: 42px; border-radius: 50%; background: linear-gradient(135deg, #f5d76e, #d89b1f); color: #3d1140; font-weight: 900; display: flex; align-items: center; justify-content: center; font-size: 21px; box-shadow: 0 2px 6px rgba(0,0,0,0.2); }
  .steptext { font-weight: 900; font-size: 25px; line-height: 1.2; color: #0f5c2c; text-align: center; white-space: nowrap; }
  .footer { flex-shrink: 0; background: #cdeecf; padding: 6px 10px 10px; text-align: center; position: relative; }
  .footer::before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 4px; background: linear-gradient(90deg, #f0c94a, #d89b1f, #f0c94a); }
  .footer .fname { font-weight: 900; font-size: 23px; color: #0f5c2c; }
  .footer .sub { font-size: 14px; color: #3f7a52; margin-top: 4px; font-weight: 600; }
</style></head>
<body>
  <div class="card">
    <div class="header"><div class="accent"></div><span class="icon">🚚</span><h1>สแกนเข้าโรงงาน <u>"รับสินค้า"</u></h1></div>
    <div class="goldline"></div>
    <div class="content">
      <div class="infobox">
        <div class="line1">🚛 สำหรับคนขับรถเท่านั้น</div>
        <div class="line2">ใช้งานได้เมื่ออยู่ภายในรัศมี ${radiusM} เมตร จากโรงงาน</div>
      </div>
      <div class="qrframe"><img class="qr" src="${qrSrc}" alt="QR" /></div>
      <div class="urlbox">${safeUrl}</div>
      <div class="howto-title">วิธีการสแกน</div>
      <div class="steps">
        <div class="step"><div class="num">1</div><div class="steptext">สแกน<br>QR-CODE</div></div>
        <div class="step"><div class="num">2</div><div class="steptext">กรอก<br>ทะเบียนรถ</div></div>
        <div class="step"><div class="num">3</div><div class="steptext">ยืนยัน<br>เช็คอิน</div></div>
      </div>
    </div>
    <div class="footer">
      <div class="fname">${facilityName}</div>
      <div class="sub">PLP – Process Improvement</div>
    </div>
  </div>
  <script>
    var img = document.querySelector('.qr');
    function go() { setTimeout(function () { window.print(); }, 150); }
    if (img.complete) go(); else img.onload = go;
  </script>
</body></html>`;
  const w = window.open("", "_blank", "width=650,height=920");
  if (!w) return;
  w.document.open();
  w.document.write(html);
  w.document.close();
};

// พิมพ์ QR "Dashboard ขนส่ง" เป็นไฟล์/หน้าเดียว — เทมเพลตแบบมินิมอล (โทนน้ำเงิน) ไม่มีขั้นตอนใช้งาน
const printDashboardQrTemplate = (url) => {
  const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=${encodeURIComponent(url)}`;
  const facilityName = escapeHtml(settings.facilityName);
  const safeUrl = escapeHtml(url);
  const html = `<!doctype html>
<html lang="th"><head><meta charset="utf-8" /><title>QR Dashboard ขนส่ง</title>
<style>
  @page { size: A4; margin: 0; }
  * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Sarabun', sans-serif; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
  html, body { width: 100%; height: 100%; }
  body { display: flex; background: #eaf2fb; }
  .card { flex: 1; display: flex; flex-direction: column; border: 16px solid #bcd9f0; box-sizing: border-box; position: relative; overflow: hidden; min-height: 0; }
  .header { flex-shrink: 0; background: linear-gradient(135deg, #1e3a8a 0%, #1e293b 100%); padding: 30px 20px; text-align: center; }
  .header h1 { color: #fff; font-size: 50px; font-weight: 900; letter-spacing: 0.5px; }
  .goldline { flex-shrink: 0; height: 6px; background: linear-gradient(90deg, #f0c94a, #d89b1f, #f0c94a); }
  .content { flex: 1; min-height: 0; overflow: hidden; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 18px; }
  .infobox { flex-shrink: 0; width: 100%; max-width: 700px; background: #eef2f6; border: 1.5px solid #cfe0f2; border-radius: 12px; padding: 8px 20px; text-align: center; }
  .infobox .line1 { font-weight: 900; font-size: 33px; color: #1e293b; display: flex; align-items: center; justify-content: center; gap: 10px; }
  .infobox .line2 { font-size: 26px; color: #4b6079; margin-top: 4px; font-weight: 600; }
  .qrframe { flex-shrink: 0; background: #fff; border: 3px solid #1e3a8a; border-radius: 26px; padding: 36px; box-shadow: 0 8px 24px rgba(30, 58, 138, 0.18); }
  .qr { width: 540px; height: 540px; display: block; }
  .urlbox { flex-shrink: 0; width: 100%; max-width: 740px; background: #f3f4f6; border: 1.5px dashed #9fb8d6; border-radius: 10px; padding: 6px 17px; font-size: 18px; color: #1e3a8a; word-break: break-all; font-family: monospace; text-align: center; font-weight: 700; }
  .footer { flex-shrink: 0; background: #bcd9f0; padding: 18px 10px 22px; text-align: center; }
  .footer .fname { font-weight: 900; font-size: 26px; color: #1e293b; }
  .footer .sub { font-size: 16px; color: #4b6079; margin-top: 4px; font-weight: 600; }
</style></head>
<body>
  <div class="card">
    <div class="header"><h1>- Dashboard ขนส่ง -</h1></div>
    <div class="goldline"></div>
    <div class="content">
      <div class="infobox">
        <div class="line1">🚚 สำหรับขนส่ง</div>
        <div class="line2">ใช้ดูคิวรถขนส่งของตัวเองแบบเรียลไทม์</div>
      </div>
      <div class="qrframe"><img class="qr" src="${qrSrc}" alt="QR" /></div>
      <div class="urlbox">${safeUrl}</div>
    </div>
    <div class="goldline"></div>
    <div class="footer">
      <div class="fname">${facilityName}</div>
      <div class="sub">PLP – Process Improvement</div>
    </div>
  </div>
  <script>
    var img = document.querySelector('.qr');
    function go() { setTimeout(function () { window.print(); }, 150); }
    if (img.complete) go(); else img.onload = go;
  </script>
</body></html>`;
  const w = window.open("", "_blank", "width=650,height=920");
  if (!w) return;
  w.document.open();
  w.document.write(html);
  w.document.close();
};

// prefix ป้ายชื่อของ QR การ์ดที่ผูกกับลานโหลด (ตามด้วยชื่อสั้นของลานจาก wh_lanes สด ๆ) —
// แยกจาก QR_ITEMS.label เดิมเพื่อให้แก้ชื่อ/ชื่อสั้นของลานใน Master Setting แล้วสะท้อนที่นี่ทันที
const QR_LANE_LABEL_PREFIX = {
  qc_parts: "ลานโหลด", qc_head: "ลานโหลด", qc_pork: "ลานโหลด",
  loading_parts: "Checker", loading_head: "Checker", loading_pork: "Checker",
  sample_parts: "QC", sample_head: "QC", sample_pork: "QC",
};

// ─── คู่มือการใช้งาน (User Manual) ────────────────────────────────────────────
const MANUAL_SECTIONS = [
  {
    id: "driver", img: "/manual/driver.png", emoji: "🚛", title: "คนขับ – เช็คอินเข้าโรงงาน",
    audience: "คนขับรถบรรทุกที่มาส่ง/รับสินค้า (สแกน QR ที่ป้อมยาม)",
    steps: [
      "เปิดลิงก์/สแกน QR \"คนขับ เช็คอิน\" ระบบจะขอเปิด GPS ก่อนเข้าใช้งาน",
      "กดปุ่ม \"📍 ตรวจสอบตำแหน่ง\" และอนุญาตให้เว็บเข้าถึงตำแหน่ง",
      "ถ้าอยู่นอกระยะที่กำหนด ระบบจะแจ้ง \"อยู่นอกพื้นที่โรงงาน\" และตรวจสอบตำแหน่งซ้ำอัตโนมัติจนกว่าจะเข้าใกล้",
      "เมื่อเข้าเขตแล้ว พิมพ์เลขทะเบียนรถ (ตัวเลขเท่านั้น) แล้วกด \"ค้นหา\"",
      "ตรวจสอบข้อมูลทะเบียน/กลุ่มลูกค้า/Zone ที่ระบบแสดงให้ถูกต้อง แล้วกด \"✅ ยืนยันเช็คอิน\"",
      "ถ้าไม่พบทะเบียนในคิว ระบบจะสร้างรายการ \"Walk-in\" ให้อัตโนมัติ (ยังไม่มีข้อมูล Zone/สินค้า ต้องให้แอดมินเพิ่มทีหลัง)",
    ],
    notes: [
      "ระยะที่อนุญาตให้เช็คอิน (Geofence) ค่าเริ่มต้น 2,000 เมตรจากจุดที่ตั้งไว้ ปรับได้ที่ Master Setting",
      "ระบบจับคู่ทะเบียนด้วยตัวเลขล้วน ไม่สนใจตัวอักษร/ตัวจังหวัด",
      "ถ้าทะเบียนคันนี้เช็คอินครบทุกเที่ยวของวันนี้แล้ว ระบบจะแจ้งเตือนไม่ให้เช็คอินซ้ำ",
    ],
  },
  {
    id: "qc_lane", img: "/manual/qc_lane.png", emoji: "🌡️", title: "ลานโหลด – ตรวจอุณหภูมิรถก่อนเข้าโหลด (3 ลาน)",
    audience: "พนักงานประจำหน้าลานโหลดแต่ละลาน (บทบาท \"ลานโหลด\")",
    steps: [
      "เลือกช่องโหลด (ถ้าระบบเปิดใช้งานเลือกช่องโหลดไว้) ทุกครั้งที่เข้าฟอร์ม",
      "เลือกทะเบียนรถจากรายการ (แสดงเฉพาะรถที่เช็คอิน/กำลังโหลดอยู่)",
      "กรอกอุณหภูมิ (°C) — ต้องกรอกก่อนจึงจะบันทึกได้",
      "แนบรูปประกอบได้ (ไม่บังคับ) สูงสุด 15 รูปต่อครั้ง",
      "กด \"✅ บันทึก QC → {ลาน}\" ระบบจะแจ้งเตือนสีเขียว \"QC ผ่าน\" เมื่อสำเร็จ",
      "วัดซ้ำ/แก้ไขค่าล่าสุดได้เรื่อย ๆ ไม่มีการล็อกหลังบันทึกครั้งแรก",
    ],
    notes: [
      "ถ้าทะเบียนที่เลือกไม่มีคำสั่งโหลดที่ลานนี้ ระบบจะเตือนแต่ยังกดบันทึกต่อได้ (ไม่บล็อก)",
      "รถต้องผ่านขั้นตอนนี้ก่อน จึงจะไปปรากฏในรายการให้ \"Checker\" เลือกโหลดได้",
    ],
  },
  {
    id: "qc_sample", img: "/manual/qc_sample.png", emoji: "📷", title: "QC – สุ่มตรวจสินค้าด้วยรูปภาพ (3 ลาน)",
    audience: "พนักงานตรวจคุณภาพสินค้า (บทบาท \"QC\")",
    steps: [
      "เลือกช่องโหลด (ถ้าเปิดใช้งาน) แล้วเลือกทะเบียนรถ",
      "เขียนโน้ต/หมายเหตุประกอบได้ (ไม่บังคับ)",
      "ถ่ายรูปอย่างน้อย 1 รูป (บังคับ) สูงสุด 15 รูป",
      "กด \"✅ บันทึกตรวจ → {ลาน}\"",
    ],
    notes: [
      "หน้านี้ไม่มีช่องกรอกอุณหภูมิ เป็นการตรวจสอบด้วยภาพถ่ายเท่านั้น (ต่างจาก \"ลานโหลด\" ที่ต้องกรอกอุณหภูมิ)",
      "ตรวจซ้ำได้ไม่จำกัดครั้ง",
    ],
  },
  {
    id: "checker", img: "/manual/checker.png", emoji: "🧺", title: "Checker – โหลดสินค้าจริง + นับตะกร้า + แจ้งรอสินค้า (3 ลาน)",
    audience: "พนักงานหน้าลานที่ทำการโหลดสินค้าขึ้นรถจริง (บทบาท \"Checker\")",
    steps: [
      "เลือกช่องโหลด (ถ้าเปิดใช้งาน)",
      "เลือกทะเบียนรถ — ระบบแสดงเฉพาะรถที่ \"ลานโหลด\" ตรวจอุณหภูมิผ่านแล้ว และยังไม่โหลดเสร็จที่ลานนี้ (มีสัญลักษณ์ ⏳ นำหน้าคันที่กำลังรอสินค้าอยู่)",
      "ถ้าสินค้ายังไม่พร้อม กด \"⏳ รอเติมสินค้า\" แล้วเลือก/พิมพ์เหตุผล (ได้สูงสุด 3 เหตุผลต่อครั้ง)",
      "ถ้าโหลดเสร็จแล้ว กรอกยอดตะกร้า/ตะขอ (ถ้ามี), ถ่ายรูปยืนยัน แล้วกด \"✅ บันทึกโหลดเสร็จ\" — ระบบจะให้ยืนยันอีกครั้งก่อนบันทึก และส่งแจ้งเตือนไป Teams อัตโนมัติ",
    ],
    notes: [
      "แต่ละลาน (ชิ้นส่วน/หัว-เครื่องใน/หมูซีก) มีข้อมูลฟอร์มแยกกัน สลับลานได้โดยข้อมูลที่กรอกไม่หาย",
      "ยอดตะกร้าเป็นข้อมูลเสริม ไม่กรอกก็บันทึกโหลดเสร็จได้",
    ],
  },
  {
    id: "office_wh", img: "/manual/office_wh.png", emoji: "📋", title: "Office คลัง – พิมพ์ใบเบิก/ใบสรุปจ่าย (Picking)",
    audience: "เจ้าหน้าที่คลัง",
    steps: [
      "เปิดหน้า Picking จะเห็นรายการรถวันนี้ เรียงรถที่ต้องทำก่อนไว้บนสุด",
      "เมื่อรถเช็คอินแล้ว (สถานะ \"เข้าโรงงานแล้ว\") กด \"🖨️ เบิก\" เพื่อพิมพ์ใบเบิกสินค้า → สถานะเปลี่ยนเป็น \"กำลังโหลด\"",
      "เมื่อ Checker โหลดเสร็จแล้วอย่างน้อย 1 ลาน กด \"🖨️ สรุป\" เพื่อพิมพ์ใบสรุปจ่าย → สถานะเปลี่ยนเป็น \"โหลดเสร็จ/สรุปแล้ว\"",
      "ใส่ \"สถานะเพิ่มเติม\" (เช่น รอแปรสินค้า, ติดปัญหา IT) เพื่อติดป้ายแจ้งเตือนให้ทุกหน้าเห็น ได้ตามต้องการ",
    ],
    notes: ["ปุ่มแต่ละปุ่มกดได้เฉพาะตอนที่เงื่อนไขก่อนหน้าสำเร็จแล้วเท่านั้น (จะแสดงเป็น \"—\" ถ้ายังกดไม่ได้)"],
  },
  {
    id: "office_plan", img: "/manual/office_plan.png", emoji: "📄", title: "Office วางแผน – ออก Invoice และอัพโหลด PO",
    audience: "เจ้าหน้าที่วางแผน/ออกเอกสาร",
    steps: [
      "หน้า \"ห้องวางแผน\": รถที่พิมพ์ใบสรุปแล้ว (สถานะ \"โหลดเสร็จ/สรุปแล้ว\") จะลอยขึ้นมาบนสุด",
      "กด \"ออก Invoice\" เพื่อออกใบ Invoice → สถานะเปลี่ยนเป็น \"ออก Invoice แล้ว\" (ขั้นตอนสุดท้ายของรถคันนั้น)",
      "หน้า \"อัพโหลด PO\": อัพโหลดไฟล์ PO ของแต่ละช่องทาง (ตลาดสด/Makro/LOTUS หรือช่องทางที่ตั้งเพิ่มไว้) เพื่อจับคู่ทะเบียนรถ ↔ ลานที่ต้องโหลด",
      "ถ้าจับคู่ไม่ได้ (0 รายการ) กด \"🔍 ดู Debug\" เพื่อเทียบโค้ดสินค้าใน Master กับไฟล์ที่อัพโหลด",
    ],
    notes: [
      "ต้องอัพโหลดไฟล์ Master (ที่ Master Setting) ไว้ก่อนแล้ว ไม่เช่นนั้นจะจับคู่ PO ไม่ได้",
      "ไฟล์ PO ที่อัพโหลดจะถูกเก็บแยกตามวันทำงาน และซิงก์ให้ทุกเครื่องเห็นข้อมูลเดียวกันอัตโนมัติ",
    ],
  },
  {
    id: "lg", img: "/manual/lg.png", emoji: "🚚", title: "LG – อัพโหลดคิวรถประจำวัน",
    audience: "ฝ่าย LG/โลจิสติกส์",
    steps: [
      "อัพโหลดไฟล์ Excel/CSV ที่ Export จากระบบ Axons Move (รายงาน Scheduling)",
      "ตรวจสอบตารางตัวอย่างข้อมูลที่ระบบอ่านได้ (ทะเบียน/วันที่/Zone/เวลาเข้า-ออก/กลุ่มลูกค้า)",
      "กด \"✅ ยืนยัน — ตั้งคิวรถ N คัน\" เพื่อบันทึกเป็นคิวของวันนี้",
      "แก้ไข/ลบ/เพิ่มรายการรถแบบ Manual ได้ที่ตารางคิวด้านล่าง",
      "แท็บ \"🗓️ อัพโหลดล่วงหน้า\" ใช้อัพคิวของวันทำงานถัดไปไว้ก่อนได้เลย โดยไม่กระทบคิววันนี้",
    ],
    notes: [
      "การยืนยันไฟล์ใหม่ในแท็บ \"คิวรถวันนี้\" จะ \"แทนที่คิวเดิมทั้งหมด\" ไม่ใช่การเพิ่มต่อ — ถ้ามีรถเช็คอินไปแล้ว ระบบจะจับคู่ทะเบียนกับคิวใหม่ให้อัตโนมัติ",
      "คิวที่อัพในแท็บ \"อัพโหลดล่วงหน้า\" จะถูกดึงมาใช้แทนคิวปัจจุบันเองอัตโนมัติ ทันทีที่ข้ามเวลาตัดรอบวันทำงาน พร้อมปิดยอด/เก็บข้อมูลวันเก่าเข้าประวัติให้เอง ไม่ต้องกด \"ล้างวันใหม่\"",
      "ระบบอ่านข้อมูลตามตำแหน่งคอลัมน์ที่ตายตัว ต้องใช้ไฟล์รูปแบบเดิมเสมอ",
    ],
  },
  {
    id: "dashboards", img: "/manual/dashboards.png", emoji: "📊", title: "Dashboard – ภาพรวมคิวรถวันนี้",
    audience: "ฝ่ายบริหาร/ห้องคุม/จอมอนิเตอร์ และบริษัทขนส่ง (Dashboard ขนส่ง)",
    steps: [
      "Dashboard หลัก แสดงตารางรถทุกคันวันนี้ พร้อมสถานะแต่ละลาน เวลาเข้า-ออก และเอกสารที่พิมพ์แล้ว",
      "ใช้ช่องค้นหาทะเบียนเพื่อกรองรายการ, สลับดูตามลาน (ชิ้นส่วน/หัว-เครื่องใน/หมูซีก) ได้ที่แท็บด้านบน (จอใหญ่)",
      "กด \"⛶\" เพื่อเข้าโหมดเต็มจอ หรือกด \"TV\" เพื่อเข้าโหมดจอทีวี (ตัวอักษรใหญ่ เหมาะกับจอมอนิเตอร์ผนัง)",
      "\"Dashboard ขนส่ง\" คือหน้าย่อ แสดงเฉพาะทะเบียน/กลุ่มลูกค้า/เวลาเข้า-ออก เหมาะเปิดให้บริษัทขนส่ง/คนขับดูรอที่หน้าโรงงาน",
    ],
    notes: [
      "แถบเวลานับถอยหลังก่อนออกจะเปลี่ยนสี เขียว→เหลือง→แดง เมื่อเหลือเวลาน้อยลง",
      "ทะเบียนของคนขับที่เพิ่งเช็คอินจากเครื่องเดียวกันจะถูกไฮไลต์สีฟ้าให้เห็นง่าย",
    ],
  },
  {
    id: "overview_log", img: "/manual/overview_log.png", emoji: "🗒️", title: "Log ภาพรวมการทำงาน",
    audience: "ผู้ดูแล/ฝ่ายข้อมูลการโหลดสินค้า",
    steps: [
      "เลือกวันที่ และกรองตามประเภทเหตุการณ์ (โหลดจ่ายสินค้า / ตรวจอุณหภูมิรถ / ตรวจอุณหภูมิสินค้า) หรือค้นหาทะเบียน/คำค้น",
      "ดูลำดับเหตุการณ์ของแต่ละคันแบบ timeline พร้อมรูปถ่ายประกอบ (คลิกรูปเพื่อดูขนาดใหญ่/บันทึกรูปลงเครื่อง)",
      "กด \"Export Excel\" เพื่อดาวน์โหลดรายการที่กรองไว้",
    ],
    notes: [],
  },
  {
    id: "work_tracking", img: "/manual/work_tracking.png", emoji: "📈", title: "Tracking การทำงาน",
    audience: "ผู้บริหาร/ฝ่ายติดตาม KPI",
    steps: [
      "เลือกวันที่ที่ต้องการดู (ดูข้อมูลวันก่อนหน้าจากประวัติที่บันทึกไว้ได้)",
      "ดูตัวเลขสรุปด้านบน (รถเข้ารวม/ตรงเวลา, เวลาเฉลี่ยอยู่ในโรงงาน, รถออกตรงเวลา/สาย) เทียบกับเมื่อวาน",
      "ดูตารางรายละเอียดทุกขั้นตอนของทุกคันในหน้าเดียว (เข้าโรงงาน → ตรวจอุณหภูมิ/QC/โหลดเสร็จของทั้ง 3 ลาน → ใบสรุป/Invoice → ออกโรงงาน) คลิกหัวคอลัมน์เพื่อจัดเรียง",
      "กด \"Export Excel\" เพื่อดาวน์โหลดตารางทั้งหมด",
    ],
    notes: ["ช่องที่ไม่เกี่ยวกับรถคันนั้น (เช่น รถที่ไม่ได้โหลดลานหมูซีก) จะแสดงเป็นสีเทาอัตโนมัติ"],
  },
  {
    id: "basket_summary", img: "/manual/basket_summary.png", emoji: "🧺", title: "ข้อมูลยอดตะกร้า/ตะขอ",
    audience: "ผู้ดูแลระบบ (บทบาท \"ทั้งหมด\")",
    steps: [
      "เลือกวันที่/ทะเบียนที่ต้องการดู",
      "บันทึกการคืนตะกร้า/ตะขอของแต่ละทะเบียนได้ที่ฟอร์มด้านบน",
      "ดูยอดรวมแยกตามประเภทตะกร้า/ลาน และยอด \"ตะกร้าค้างคืน\" สะสมของแต่ละทะเบียน",
      "กด \"Export Excel\" เพื่อดาวน์โหลดข้อมูลทั้ง 3 ตาราง",
    ],
    notes: ["ยอดค้างคืนคำนวณสะสมข้ามวัน ไม่ถูกล้างเมื่อกด \"ล้างวันใหม่\""],
  },
  {
    id: "waiting_summary", img: "/manual/waiting_summary.png", emoji: "⏳", title: "ข้อมูลการรอสินค้า",
    audience: "ผู้ดูแลระบบ (บทบาท \"ทั้งหมด\")",
    steps: [
      "เลือกวันที่/ทะเบียน เพื่อดูรายการรถที่ Checker แจ้ง \"รอเติมสินค้า\"",
      "รายการที่ยังรออยู่นานเกินเวลาที่กำหนดจะไฮไลต์สีแดงเตือน",
      "กด \"Export Excel\" เพื่อดาวน์โหลดรายการที่กรองไว้",
    ],
    notes: ["ถ้าทะเบียนเดียวกันรอสินค้าซ้ำหลายครั้งในลานเดียวกันวันเดียวกัน ระบบจะเก็บ/แสดงเฉพาะครั้งล่าสุดเท่านั้น"],
  },
  {
    id: "end_of_day", img: "/manual/end_of_day.png", emoji: "🗑️", title: "จบการทำงาน",
    audience: "ผู้ดูแลระบบ (บทบาท \"ทั้งหมด\")",
    steps: [
      "เมื่อจบวันทำงาน กด \"🗑️ ล้างวันใหม่\" ระบบจะเก็บข้อมูลวันนี้เป็นประวัติ แล้วล้างคิว/รถให้พร้อมรับวันใหม่ (ต้องกดยืนยันอีกครั้ง)",
      "ดาวน์โหลดข้อมูลย้อนหลังของวันที่ต้องการเป็นไฟล์ Excel ได้ที่ \"📥 ดาวน์โหลดข้อมูลย้อนหลัง\"",
      "ลบข้อมูลย้อนหลังแบบถาวรได้ที่ \"🗑️ ลบข้อมูลย้อนหลัง\" (กู้คืนไม่ได้)",
    ],
    notes: ["การล้างวันใหม่และลบข้อมูลย้อนหลังเป็นการกระทำที่ย้อนกลับไม่ได้ ควรตรวจสอบให้แน่ใจก่อนกดยืนยัน"],
  },
  {
    id: "admin", img: "/manual/admin.png", emoji: "🛠️", title: "Admin – แก้ไขข้อมูลรถ",
    audience: "ผู้ดูแลระบบ (บทบาท \"ทั้งหมด\")",
    steps: [
      "เลือกทะเบียนรถที่ต้องการแก้ไข",
      "แก้ไขกลุ่มลูกค้า/Zone/สถานะ/อุณหภูมิ QC/สถานะการโหลดของแต่ละลานได้โดยตรง",
      "ถ้ามีทะเบียนซ้ำ (สแกนผิดคัน) ใช้ \"🔀 Merge รถซ้ำ\" เพื่อรวมข้อมูลเป็นคันเดียว",
      "ใช้ \"🔗 ผูก Queue Entry\" เพื่อเชื่อมรถกับรายการคิวที่ถูกต้อง หากจับคู่อัตโนมัติผิด",
      "กด \"💾 บันทึกการแก้ไข\" เพื่อยืนยัน หรือ \"🗑️ ลบรถออกจากระบบ\" หากเป็นรายการที่สแกนผิดทั้งคัน",
    ],
    notes: [],
  },
  {
    id: "qr_page", img: "/manual/qr_page.png", emoji: "📱", title: "QR Code ทั้งหมด",
    audience: "ผู้ดูแลระบบ – ใช้ตอนติดตั้ง/พิมพ์ QR ไปแปะหน้าสถานี",
    steps: [
      "แต่ละกล่องคือ 1 จุดใช้งาน คลิกกล่องเพื่อดู QR ขนาดใหญ่",
      "ใช้ \"📋 คัดลอก\" คัดลอกลิงก์ หรือ \"↗ เปิด\" เพื่อเปิดหน้านั้นทดสอบ หรือ \"💾 บันทึกรูป QR\" เพื่อเซฟรูป QR ไปพิมพ์",
      "กด \"🖨️ พิมพ์ทั้งหมด\" เพื่อพิมพ์ QR ทุกจุดพร้อมกัน",
    ],
    notes: [],
  },
  {
    id: "master_setting", img: "/manual/master_setting.png", emoji: "🔐", title: "Master Setting – ตั้งค่าระบบทั้งหมด",
    audience: "ผู้ดูแลระบบเท่านั้น",
    steps: [
      "เข้าหน้านี้ต้องกรอกรหัส PIN 4 หลักก่อน (ค่าเริ่มต้น 0000 — ควรเปลี่ยนทันทีหลังติดตั้งระบบ)",
      "อัพโหลดไฟล์ Master (รหัสสินค้า ↔ ลานโหลด) — ระบบจะรายงานจำนวนที่จับคู่สำเร็จให้ตรวจสอบทุกครั้ง",
      "ปรับ \"⚙️ ตั้งค่าระบบ\" เช่น เวลาตัดรอบวันทำงาน, ระยะรอสินค้าที่ถือว่าเร่งด่วน, จำนวนรูปสูงสุด, พิกัด/ระยะ Geofence, ชื่อโรงงาน, ราคาต่อหน่วย/VAT",
      "ปรับชื่อ/สี/เปิด-ปิดใช้งานลานที่ \"🏭 ลานโหลด\", จัดการช่องโหลดที่ \"🚪 ช่องโหลด\"",
      "ปรับป้ายชื่อ/อีโมจิของแต่ละตำแหน่งงานที่ \"🧑‍💼 ตำแหน่งงาน\"",
      "จัดการชื่อเรียกลาน, เหตุผลรอสินค้า, ประเภทตะกร้า, ช่องทาง PO ได้ในหมวดย่อยที่เหลือ",
      "เปลี่ยนรหัส PIN ได้ที่ \"🔑 รหัสผ่านเข้าแก้ Master Setting\"",
    ],
    notes: ["การบันทึกค่าตั้งค่ามีผลทันทีเฉพาะเครื่อง/แท็บที่กดบันทึก เครื่องอื่นต้องรีเฟรชหน้าจึงจะเห็นค่าใหม่"],
  },
];

const UserManualPage = ({ onClose }) => {
  const thTd = { border: "1px solid #e5e7eb", padding: "6px 8px", textAlign: "left" };
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 900 }}>📘 คู่มือการใช้งานระบบ</h2>
          <p style={{ margin: "4px 0 0", color: "#6b7280", fontSize: 12 }}>สรุปวิธีใช้งานทุกหน้าในระบบ แยกตามตำแหน่งงาน</p>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onClose}
            style={{ background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            ← กลับหน้า QR Code
          </button>
          <button onClick={() => window.print()}
            style={{ background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            🖨️ Export PDF
          </button>
        </div>
      </div>

      <div style={{ background: "#fff", border: "1px solid #e5e7eb", padding: 16, marginBottom: 14 }}>
        <h3 style={{ margin: "0 0 10px", fontSize: 15, fontWeight: 900 }}>🗺️ ภาพรวมขั้นตอนการทำงาน</h3>
        <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", marginBottom: 12 }}>
          {FLOW_STEPS.map((s, i) => (
            <React.Fragment key={s.key}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, background: "#f3f4f6", padding: "6px 12px" }}>
                <span>{s.emoji}</span><b style={{ fontSize: 12 }}>{s.label}</b>
              </div>
              {i < FLOW_STEPS.length - 1 && <span style={{ color: "#9ca3af" }}>→</span>}
            </React.Fragment>
          ))}
        </div>
        <p style={{ fontSize: 12.5, color: "#374151", lineHeight: 1.9, margin: "0 0 12px" }}>
          1) LG อัพโหลดคิวรถของวัน → 2) คนขับเช็คอิน (สถานะ "เข้าโรงงานแล้ว") → 3) Office คลังพิมพ์ใบเบิก (สถานะ "กำลังโหลด")
          → 4) "ลานโหลด" ตรวจอุณหภูมิรถก่อนเข้าแต่ละลาน (ทำได้ทีละลาน ไม่ต้องครบ 3 ลานพร้อมกัน) → 5) "Checker" โหลดสินค้าจริงที่ลานที่ผ่านการตรวจแล้ว
          → 6) Office คลังพิมพ์ใบสรุปเมื่อโหลดเสร็จอย่างน้อย 1 ลาน (สถานะ "โหลดเสร็จ/สรุปแล้ว") → 7) Office วางแผนออก Invoice (สถานะ "ออก Invoice แล้ว" — ขั้นสุดท้าย)
        </p>

        <h4 style={{ margin: "0 0 8px", fontSize: 13, fontWeight: 900 }}>⚠️ ป้ายชื่อที่มักสับสน</h4>
        <div style={{ overflowX: "auto", marginBottom: 12 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead><tr style={{ background: "#f3f4f6" }}>
              <th style={thTd}>ป้ายชื่อ</th><th style={thTd}>ทำหน้าที่อะไร</th>
            </tr></thead>
            <tbody>
              <tr><td style={thTd}><b>ลานโหลด</b></td><td style={thTd}>ตรวจ "อุณหภูมิ" รถก่อนเข้าโหลด — กรอกองศา + ถ่ายรูป (ไม่ใช่คนโหลดสินค้า)</td></tr>
              <tr><td style={thTd}><b>QC</b></td><td style={thTd}>สุ่มตรวจสินค้าด้วย "รูปภาพ" เท่านั้น ไม่มีช่องกรอกอุณหภูมิ</td></tr>
              <tr><td style={thTd}><b>Checker</b></td><td style={thTd}>ลงมือโหลดสินค้าจริงขึ้นรถ + นับตะกร้า/ตะขอ + แจ้งรอสินค้า</td></tr>
            </tbody>
          </table>
        </div>

        <h4 style={{ margin: "0 0 8px", fontSize: 13, fontWeight: 900 }}>🔧 ค่าเริ่มต้นที่ควรรู้ (ปรับได้ที่ Master Setting)</h4>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <tbody>
              {[
                ["รหัส PIN เข้า Master Setting", "0000"],
                ["ระยะ Geofence สำหรับเช็คอิน", "2,000 เมตร"],
                ["จำนวนรูปสูงสุดต่อครั้ง", "15 รูป"],
                ["เหตุผลรอสินค้าสูงสุด/ครั้ง", "3 เหตุผล"],
                ["เวลาตัดรอบวันทำงาน", "10:00 น."],
                ["ถือว่า \"รอสินค้านาน\" เมื่อเกิน", "20 นาที"],
                ["กรอบเวลานับถอยหลังก่อนออกจากโรงงาน", "240 นาที"],
              ].map(([k, v]) => (
                <tr key={k}><td style={{ ...thTd, fontWeight: 700 }}>{k}</td><td style={thTd}>{v}</td></tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {MANUAL_SECTIONS.map(sec => (
        <div key={sec.id} style={{ background: "#fff", border: "1px solid #e5e7eb", padding: 16, marginBottom: 14 }}>
          <h3 style={{ margin: "0 0 4px", fontSize: 15, fontWeight: 900 }}>{sec.emoji} {sec.title}</h3>
          <p style={{ margin: "0 0 10px", fontSize: 12, color: "#6b7280" }}>👤 {sec.audience}</p>
          {sec.img && (
            <img src={sec.img} alt={sec.title} loading="lazy"
              style={{ display: "block", width: "100%", maxWidth: 480, border: "1px solid #e5e7eb", marginBottom: 12 }} />
          )}
          <ol style={{ margin: "0 0 10px", paddingLeft: 20, fontSize: 12.5, color: "#374151", lineHeight: 1.9 }}>
            {sec.steps.map((s, i) => <li key={i}>{s}</li>)}
          </ol>
          {sec.notes.length > 0 && (
            <div style={{ background: "#fefce8", border: "1px solid #fde68a", padding: "8px 12px", fontSize: 11.5, color: "#78350f", lineHeight: 1.8 }}>
              {sec.notes.map((n, i) => <div key={i}>💡 {n}</div>)}
            </div>
          )}
        </div>
      ))}
    </div>
  );
};

const QRCodePage = () => {
  const base = typeof window !== "undefined" ? `${window.location.origin}${window.location.pathname}` : "";
  const [zoomItem, setZoomItem] = useState(null);
  const [showManual, setShowManual] = useState(false);
  if (showManual) return <UserManualPage onClose={() => setShowManual(false)} />;
  const items = QR_ITEMS.map(i => {
    const laneId = LANE_ENTRY_TAB_TO_LANE_ID[i.mode];
    const lane = laneId && lanes.find(l => l.id === laneId);
    return lane ? { ...i, label: `${QR_LANE_LABEL_PREFIX[i.mode]} ${lane.tinyLabel}` } : i;
  });
  const topRowItems = [QR_HOME_ITEM, ...items.filter(i => i.mode === "driver" || i.mode === "dashboard_transport")];
  const restItems    = items.filter(i => i.mode !== "driver" && i.mode !== "dashboard_transport");

  const urlForMode = mode => mode === "home" ? base : `${base}?mode=${mode}`;

  const renderCard = ({ mode, emoji, label, color, bg }) => {
    const url = urlForMode(mode);
    return (
      <div key={mode} onClick={() => setZoomItem({ mode, emoji, label, color, bg })}
        style={{ background: "#fff", borderRadius: 0, padding: 8, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", border: `2px solid ${color}20`, display: "flex", flexDirection: "column", alignItems: "center", gap: 5, cursor: "pointer" }}>
        <div style={{ width: "100%", background: color, borderRadius: 0, padding: "4px 0", textAlign: "center", color: "#fff", fontWeight: 900, fontSize: 10 }}>
          {emoji} {label}
        </div>
        <QRCodeDisplay url={url} size={80} />
        <div style={{ background: bg, borderRadius: 0, padding: "3px 6px", fontSize: 7, color: "#374151", wordBreak: "break-all", fontFamily: "monospace", width: "100%", boxSizing: "border-box", textAlign: "center" }}>
          {url}
        </div>
        <div style={{ display: "flex", gap: 4, width: "100%" }}>
          <button onClick={e => { e.stopPropagation(); navigator.clipboard?.writeText(url); }}
            style={{ flex: 1, background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "4px 0", fontSize: 9, fontWeight: 700, cursor: "pointer" }}>
            📋 คัดลอก
          </button>
          <a href={url} target="_blank" rel="noreferrer" onClick={e => e.stopPropagation()}
            style={{ flex: 1, background: color, color: "#fff", border: "none", borderRadius: 0, padding: "4px 0", fontSize: 9, fontWeight: 700, cursor: "pointer", textDecoration: "none", textAlign: "center", display: "flex", alignItems: "center", justifyContent: "center" }}>
            ↗ เปิด
          </a>
        </div>
      </div>
    );
  };

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 900 }}>📱 QR Code ทั้งหมด</h2>
          <p style={{ margin: "4px 0 0", color: "#6b7280", fontSize: 12 }}>สแกนเพื่อเข้าหน้าต่าง ๆ โดยตรง · คลิกที่กล่องเพื่อดูขนาดใหญ่</p>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => setShowManual(true)}
            style={{ background: "#0ea5e9", color: "#fff", border: "none", borderRadius: 0, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            📘 คู่มือการใช้งาน
          </button>
          <button onClick={() => window.print()}
            style={{ background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            🖨️ พิมพ์ทั้งหมด
          </button>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 10, marginBottom: 12 }}>
        {topRowItems.map(renderCard)}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 10 }}>
        {restItems.map(renderCard)}
      </div>

      {zoomItem && (() => {
        const url = urlForMode(zoomItem.mode);
        return (
          <div onClick={() => setZoomItem(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, cursor: "zoom-out" }}>
            <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 0, padding: 28, maxWidth: 360, width: "100%", display: "flex", flexDirection: "column", alignItems: "center", gap: 14, cursor: "default" }}>
              <div style={{ width: "100%", background: zoomItem.color, borderRadius: 0, padding: "12px 0", textAlign: "center", color: "#fff", fontWeight: 900, fontSize: 18 }}>
                {zoomItem.emoji} {zoomItem.label}
              </div>
              <QRCodeDisplay url={url} size={280} />
              <div style={{ background: zoomItem.bg, borderRadius: 0, padding: "8px 12px", fontSize: 11, color: "#374151", wordBreak: "break-all", fontFamily: "monospace", width: "100%", boxSizing: "border-box", textAlign: "center" }}>
                {url}
              </div>
              <div style={{ display: "flex", gap: 8, width: "100%" }}>
                <button onClick={() => navigator.clipboard?.writeText(url)}
                  style={{ flex: 1, background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  📋 คัดลอก
                </button>
                <a href={url} target="_blank" rel="noreferrer"
                  style={{ flex: 1, background: zoomItem.color, color: "#fff", border: "none", borderRadius: 0, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer", textDecoration: "none", textAlign: "center", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  ↗ เปิด
                </a>
              </div>
              <button onClick={() => saveQrImage(url, zoomItem.mode)}
                style={{ width: "100%", background: "#16a34a", color: "#fff", border: "none", borderRadius: 0, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                💾 บันทึกรูป QR
              </button>
              {zoomItem.mode === "home" && (
                <button onClick={() => printHomeQrTemplate(url)}
                  style={{ width: "100%", background: "#4a154b", color: "#fff", border: "none", borderRadius: 0, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  🖨️ พิมพ์ป้าย QR Code
                </button>
              )}
              {zoomItem.mode === "driver" && (
                <button onClick={() => printDriverQrTemplate(url)}
                  style={{ width: "100%", background: "#0f5c2c", color: "#fff", border: "none", borderRadius: 0, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  🖨️ พิมพ์ป้าย QR Code
                </button>
              )}
              {zoomItem.mode === "dashboard_transport" && (
                <button onClick={() => printDashboardQrTemplate(url)}
                  style={{ width: "100%", background: "#1e3a8a", color: "#fff", border: "none", borderRadius: 0, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  🖨️ พิมพ์ป้าย QR Code
                </button>
              )}
              <button onClick={() => setZoomItem(null)}
                style={{ width: "100%", background: "transparent", color: "#6b7280", border: "1px solid #e5e7eb", borderRadius: 0, padding: "8px 0", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                ปิด
              </button>
            </div>
          </div>
        );
      })()}
    </div>
  );
};

// ─── ICONS ───────────────────────────────────────────────────────────────────
const Icon = ({ name, size = 20 }) => {
  const icons = {
    truck:     <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M1 3h15v13H1zM16 8h4l3 3v5h-7V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>,
    scan:      <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M3 7V5a2 2 0 012-2h2M17 3h2a2 2 0 012 2v2M21 17v2a2 2 0 01-2 2h-2M7 21H5a2 2 0 01-2-2v-2M7 12h10"/></svg>,
    upload:    <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>,
    download:  <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 3v12"/></svg>,
    clipboard: <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/></svg>,
    camera:    <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M23 19a2 2 0 01-2 2H3a2 2 0 01-2-2V8a2 2 0 012-2h4l2-3h6l2 3h4a2 2 0 012 2z"/><circle cx="12" cy="13" r="4"/></svg>,
    check:     <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2.5" viewBox="0 0 24 24"><path d="M20 6L9 17l-5-5"/></svg>,
    chart:     <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>,
    list:      <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>,
    print:     <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M6 9V2h12v7M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2M6 14h12v8H6v-8z"/></svg>,
    invoice:   <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/></svg>,
    temp:      <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M14 14.76V3.5a2.5 2.5 0 00-5 0v11.26a4.5 4.5 0 105 0z"/></svg>,
    loader:    <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg>,
    clock:     <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>,
    exit:      <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M15 3h4a2 2 0 012 2v14a2 2 0 01-2 2h-4M10 17l5-5-5-5M15 12H3"/></svg>,
    alert:     <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><path d="M12 9v4M12 17h.01"/></svg>,
    x:         <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2.5" viewBox="0 0 24 24"><path d="M18 6L6 18M6 6l12 12"/></svg>,
    bell:      <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9M13.73 21a2 2 0 01-3.46 0"/></svg>,
    plan:      <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>,
    lock:      <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>,
    pig_head:  <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="1.8" viewBox="0 0 24 24" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="13.5" r="6.5"/><path d="M7 8.5 L5 4 L10 7.5"/><path d="M17 8.5 L19 4 L14 7.5"/><ellipse cx="12" cy="17" rx="2.5" ry="1.5"/><circle cx="11" cy="17" r="0.6" fill="currentColor" stroke="none"/><circle cx="13" cy="17" r="0.6" fill="currentColor" stroke="none"/><circle cx="9.5" cy="12" r="0.9" fill="currentColor" stroke="none"/><circle cx="14.5" cy="12" r="0.9" fill="currentColor" stroke="none"/></svg>,
    pig_cuts:  <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="1.8" viewBox="0 0 24 24" strokeLinecap="round" strokeLinejoin="round"><path d="M3 16 Q12 13 21 16 L21 18.5 Q12 21.5 3 18.5 Z"/><path d="M4 11 Q12 8 20 11 L20 13.5 Q12 16.5 4 13.5 Z"/><path d="M5 6 Q12 3 19 6 L19 8.5 Q12 11.5 5 8.5 Z"/></svg>,
    pig_side:  <svg width={size} height={size} fill="none" stroke="currentColor" strokeWidth="1.8" viewBox="0 0 24 24" strokeLinecap="round" strokeLinejoin="round"><path d="M10 2 L10 4.5 C7 5.5 6 9 6 13 C6 17 7.5 20.5 10 22 L14 22 C16.5 20.5 18 17 18 13 C18 9 17 5.5 14 4.5 L14 2"/><line x1="9" y1="11" x2="15" y2="11"/><line x1="8.5" y1="14.5" x2="15.5" y2="14.5"/><line x1="9" y1="18" x2="15" y2="18"/></svg>,
  };
  return icons[name] || null;
};

// ─── STATUS BADGE ─────────────────────────────────────────────────────────────
const StatusBadge = ({ status }) => {
  const s = STATUS_META[status]; if (!s) return null;
  return <span style={{ background: s.bg, color: s.color, padding: "3px 10px", borderRadius: 0, fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>{s.label}</span>;
};

// ─── FLOW PROGRESS ────────────────────────────────────────────────────────────
const FlowProgress = ({ status }) => {
  const cur = getStep(status);
  return (
    <div style={{ display: "flex", alignItems: "center", overflowX: "auto", padding: "2px 0" }}>
      {FLOW_STEPS.map((s, i) => {
        const n = i + 1; const done = n <= cur;
        return (
          <div key={s.key} style={{ display: "flex", alignItems: "center", flexShrink: 0 }}>
            <div style={{ textAlign: "center", width: 42 }}>
              <div style={{ width: 26, height: 26, borderRadius: "50%", margin: "0 auto 2px", background: done ? "#111" : "#e5e7eb", display: "flex", alignItems: "center", justifyContent: "center" }}>
                {done ? <Icon name="check" size={13} /> : <span style={{ color: "#9ca3af", fontSize: 12 }}>{s.emoji}</span>}
              </div>
              <div style={{ fontSize: 9, fontWeight: done ? 700 : 400, color: done ? "#111" : "#9ca3af", lineHeight: 1.2 }}>{s.label}</div>
            </div>
            {i < FLOW_STEPS.length - 1 && <div style={{ width: 12, height: 2, background: n < cur ? "#111" : "#e5e7eb", flexShrink: 0, marginBottom: 12 }} />}
          </div>
        );
      })}
    </div>
  );
};

// ─── MODAL ────────────────────────────────────────────────────────────────────
const Modal = ({ title, onClose, children }) => (
  <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
    <div style={{ background: "#fff", borderRadius: 0, width: "100%", maxWidth: 520, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 25px 60px rgba(0,0,0,0.3)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "20px 24px", borderBottom: "1px solid #e5e7eb" }}>
        <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700 }}>{title}</h3>
        <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: "#6b7280" }}><Icon name="x" size={22} /></button>
      </div>
      <div style={{ padding: 24 }}>{children}</div>
    </div>
  </div>
);

// ─── PRINT MODAL ──────────────────────────────────────────────────────────────
const PrintModal = ({ truck, type, onClose }) => {
  const titles = { pickup: "เบิกพัสดุสินค้า", summary: "ใบสรุปค่าย", invoice: "ใบ Invoice" };
  const title = titles[type];
  return (
    <Modal title={`🖨️ ${title}`} onClose={onClose}>
      <div style={{ border: "2px solid #111", borderRadius: 0, padding: 20, fontFamily: "monospace", fontSize: 13, lineHeight: 1.9 }}>
        <div style={{ textAlign: "center", fontWeight: 900, fontSize: 17, borderBottom: "2px solid #111", paddingBottom: 8, marginBottom: 14 }}>
          🏭 {settings.facilityName}<br/><span style={{ fontSize: 14 }}>{title}</span>
        </div>
        <div>วันที่: {TODAY}</div>
        <div>เลขที่: {type.toUpperCase()}-{truck.id}-{Date.now().toString().slice(-4)}</div>
        <div>ทะเบียนรถ: <b>{truck.plate}</b></div>
        <div>คนขับ: {truck.driver}</div>
        <div style={{ margin: "10px 0", borderTop: "1px dashed #aaa", paddingTop: 10 }}>
          <div>สินค้า: {truck.product}</div>
          <div>จำนวน: {truck.qty} {truck.unit}</div>
          <div>ปลายทาง: {truck.destination}</div>
        </div>
        {type === "summary" && truck.qcLanes && (
          <div style={{ borderTop: "1px dashed #aaa", paddingTop: 10 }}>
            <b>อุณหภูมิ QC:</b>
            {lanes.map(l => <div key={l.id}>{l.shortLabel}: {truck.qcLanes[l.id]?.temp || "–"}°C</div>)}
          </div>
        )}
        {type === "invoice" && (
          <div style={{ borderTop: "1px dashed #aaa", paddingTop: 10 }}>
            <div>ราคาต่อหน่วย: {settings.unitPrice.toFixed(2)} บาท</div>
            <div>รวม: {(truck.qty * settings.unitPrice).toLocaleString()} บาท</div>
            <div>VAT {(settings.vatRate * 100).toFixed(0)}%: {Math.round(truck.qty * settings.unitPrice * settings.vatRate).toLocaleString()} บาท</div>
            <div style={{ fontWeight: 900, fontSize: 15 }}>รวมทั้งสิ้น: {Math.round(truck.qty * settings.unitPrice * (1 + settings.vatRate)).toLocaleString()} บาท</div>
          </div>
        )}
        <div style={{ textAlign: "center", marginTop: 14, borderTop: "1px dashed #aaa", paddingTop: 10, fontSize: 11, color: "#666" }}>
          ผู้รับสินค้า: _______________ / ผู้ส่งสินค้า: _______________
        </div>
      </div>
      <button onClick={() => { window.print(); onClose(); }} style={{ marginTop: 14, width: "100%", background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "13px 0", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>
        🖨️ พิมพ์เอกสาร
      </button>
    </Modal>
  );
};

// ─── PHOTO UPLOADER ───────────────────────────────────────────────────────────
const PhotoUploader = ({ label, value, onChange, onRemove }) => {
  const photos = Array.isArray(value) ? value : (value ? [value] : []);
  return (
    <div style={{ marginBottom: 14 }}>
      {label && <label style={{ display: "block", fontSize: 13, fontWeight: 600, color: "#374151", marginBottom: 5 }}>{label}</label>}
      <div style={{ border: `2px dashed ${photos.length > 0 ? "#6ee7b7" : "#d1d5db"}`, borderRadius: 0, padding: photos.length > 0 ? 10 : 18, background: photos.length > 0 ? "#f0fdf4" : "#fafafa" }}>
        {photos.length > 0
          ? <div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 6 }}>
                {photos.map((src, i) => (
                  <div key={i} style={{ position: "relative" }}>
                    <img src={src} alt="" style={{ width: "100%", aspectRatio: "1", borderRadius: 0, objectFit: "cover", display: "block" }} />
                    {onRemove && (
                      <button onClick={e => { e.stopPropagation(); onRemove(photos.filter((_, j) => j !== i)); }}
                        style={{ position: "absolute", top: 3, right: 3, background: "rgba(0,0,0,0.55)", color: "#fff", border: "none", borderRadius: "50%", width: 20, height: 20, fontSize: 11, fontWeight: 900, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1 }}>
                        ×
                      </button>
                    )}
                  </div>
                ))}
                {photos.length < 15 && (
                  <label style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", aspectRatio: "1", borderRadius: 0, border: "1.5px dashed #d1d5db", cursor: "pointer", background: "#fff", gap: 2 }}>
                    <input type="file" accept="image/*" multiple onChange={onChange} style={{ display: "none" }} />
                    <Icon name="camera" size={18} />
                    <span style={{ color: "#9ca3af", fontSize: 10 }}>เพิ่ม</span>
                  </label>
                )}
              </div>
              <div style={{ textAlign: "center", marginTop: 8, fontSize: 11, color: "#10b981", fontWeight: 700 }}>
                {photos.length} / 15 รูป
              </div>
            </div>
          : <label style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, cursor: "pointer" }}>
              <input type="file" accept="image/*" multiple onChange={onChange} style={{ display: "none" }} />
              <Icon name="camera" size={28} />
              <span style={{ color: "#9ca3af", fontSize: 13 }}>ถ่ายรูป / เลือกจาก Gallery</span>
              <span style={{ color: "#d1d5db", fontSize: 11 }}>สูงสุด 15 รูปต่อครั้ง</span>
            </label>
        }
      </div>
    </div>
  );
};

// ─── TRUCK CARD ───────────────────────────────────────────────────────────────
const TruckCard = ({ t, children, highlight }) => (
  <div style={{ background: "#fff", borderRadius: 0, padding: 18, boxShadow: "0 2px 10px rgba(0,0,0,0.07)", marginBottom: 14, border: highlight ? "2px solid #111" : "1.5px solid #f0f0f0" }}>
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
      <div>
        <div style={{ fontWeight: 900, fontSize: 18, letterSpacing: 1 }}>{t.plate}</div>
        <div style={{ color: "#6b7280", fontSize: 12 }}>{t.driver} · เข้า {t.arrivedAt}</div>
      </div>
      <StatusBadge status={t.status} />
    </div>
    <div style={{ background: "#f9fafb", borderRadius: 0, padding: "8px 12px", fontSize: 13, marginBottom: 10, display: "flex", gap: 14, flexWrap: "wrap" }}>
      <span><b>สินค้า:</b> {t.product}</span>
      <span><b>จำนวน:</b> {t.qty} {t.unit}</span>
      <span><b>ปลายทาง:</b> {t.destination}</span>
    </div>
    <FlowProgress status={t.status} />
    {children && <div style={{ marginTop: 12 }}>{children}</div>}
  </div>
);

// ─────────────────────────────────────────────────────────────────────────────
// VIEWS
// ─────────────────────────────────────────────────────────────────────────────

// ── TIME BAR ──────────────────────────────────────────────────────────────────
const parseExitDatetime = (dateStr, timeStr) => {
  if (!timeStr) return null;
  const timeParts = timeStr.split(":");
  const h = parseInt(timeParts[0], 10);
  const min = parseInt(timeParts[1], 10);
  if (isNaN(h) || isNaN(min)) return null;
  if (dateStr) {
    let [day, month, year] = dateStr.split("/").map(Number);
    // handle M/D/YYYY format (US Excel) — month > 12 means day/month are swapped
    if (month > 12) { [day, month] = [month, day]; }
    // defensively correct a stray Thai Buddhist-era year (e.g. old data saved before
    // toDateStr converted it) — otherwise the exit countdown ends up ~543 years off
    if (year > 2400) year -= 543;
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      const d = new Date(year, month - 1, day, h, min, 0, 0);
      if (h * 60 + min <= settings.workDayCutoffHour * 60) d.setDate(d.getDate() + 1);
      return d;
    }
  }
  // ไม่มี date → fallback วันนี้ + ปรับข้ามคืน
  const d = new Date(); d.setHours(h, min, 0, 0);
  if (h * 60 + min <= settings.workDayCutoffHour * 60) d.setDate(d.getDate() + 1);
  return d;
};

const calcTimeBarInfo = (exitTime, date) => {
  const exitDt = parseExitDatetime(date, exitTime);
  const remaining = exitDt ? Math.round((exitDt - Date.now()) / 60000) : 0;
  const totalWindow = settings.exitTimeWindowMinutes;
  const color = remaining > 60 ? "#22c55e" : remaining > 20 ? "#f59e0b" : "#ef4444";
  const fmtMins = m => { const a = Math.abs(m); return `${Math.floor(a/60)}:${String(a%60).padStart(2,"0")}`; };
  const label = remaining < 0 ? `เกิน ${fmtMins(remaining)} ชม.` : `เหลือ ${fmtMins(remaining)} ชม.`;
  const pct = Math.min(Math.max(remaining / totalWindow, 0), 1) * 100;
  return { remaining, color, label, pct };
};

const TimeBar = ({ exitTime, date, done, invoicedAt, fs, card, hideBar, hideLabel }) => {
  if (!exitTime) return <span style={{ color: "#d1d5db", fontSize: fs ? 15 : 11 }}>—</span>;

  if (done) {
    return (
      <div>
        <div style={{ fontSize: fs ? 15 : card ? 20 : 11, fontWeight: 700, color: "#374151" }}>{exitTime}</div>
        {invoicedAt && <div style={{ fontSize: fs ? 13 : card ? 13 : 10, color: "#6b7280", marginTop: 2 }}>ออกจริง {invoicedAt}</div>}
      </div>
    );
  }

  const { remaining, color, label, pct } = calcTimeBarInfo(exitTime, date);
  const barExtend = card ? null : fs ? { marginLeft: -180, width: "calc(100% + 180px)" } : { marginLeft: -114, width: "calc(100% + 114px)" };
  return (
    <div>
      <div style={{ fontSize: fs ? 16 : card ? 20 : 13, fontWeight: 700, color: remaining <= 0 ? "#ef4444" : "#374151", whiteSpace: "nowrap", lineHeight: "18px" }}>{exitTime}</div>
      {!hideLabel && <div style={{ fontSize: fs ? 13 : card ? 14 : 11, color, marginTop: 2, whiteSpace: "nowrap", fontWeight: 600, lineHeight: "16px" }}>{label}</div>}
      {!hideBar && <div style={{ marginTop: fs ? 14 : 4, height: fs ? 10 : 6, borderRadius: card ? 3 : 0, ...(barExtend ?? {}), background: `linear-gradient(to right, ${color} ${pct}%, #e5e7eb ${pct}%)` }} />}
    </div>
  );
};

const TimeBarTrack = ({ exitTime, date, done }) => {
  if (!exitTime || done) return null;
  const { color, pct, label } = calcTimeBarInfo(exitTime, date);
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ flex: 1, height: 6, borderRadius: 3, background: `linear-gradient(to right, ${color} ${pct}%, #e5e7eb ${pct}%)` }} />
      <div style={{ fontSize: 12, color, fontWeight: 600, whiteSpace: "nowrap" }}>{label}</div>
    </div>
  );
};

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
const exportArchiveExcel = async (dateStr) => {
  const { data, error } = await supabase.from("wh_archive").select("*").eq("archive_date", dateStr).single();
  if (error || !data) { alert("ไม่พบข้อมูลวันที่ " + dateStr); return; }
  const { queue, trucks } = data;
  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const rows = queue.map((q, i) => {
    const truck = trucks.find(t => t.queueId === q.id);
    return {
      "ลำดับ":                          i + 1,
      "วันที่":                         q.date || dateStr,
      "กลุ่มลูกค้า":                   q.customerGroup || "",
      "Zone":                           q.zone || "",
      "ทะเบียนรถ":                      q.plate || "",
      "น้ำหนักจัดรถ":                   "",
      "เวลารถเข้าโรงงาน STD":            q.entryTime || "",
      "เวลารถเข้าโรงงาน ACT":            truck?.arrivedAt || "",
      "เวลาพิมพ์ใบเบิก (Picking)":       truck?.pickingAt || "",
      "สถานะเพิ่มเติม":                  truck?.extraStatus || "",
      "เวลาสถานะเพิ่มเติม":             truck?.extraStatusAt || "",
      "เวลาเข้าโหลดชิ้นส่วน STD":       "",
      "เวลาเข้าโหลดชิ้นส่วน ACT":       truck?.qcLanes?.lane_parts?.doneAt || "",
      "เวลารอสินค้าชิ้นส่วน":           truck?.loadLanes?.lane_parts?.waitingAt || "",
      "เวลาเสร็จสิ้นโหลดชิ้นส่วน":     truck?.loadLanes?.lane_parts?.doneAt || "",
      "เวลาเข้าโหลดหัวเครื่องใน STD":   "",
      "เวลาเข้าโหลดหัวเครื่องใน ACT":   truck?.qcLanes?.lane_head?.doneAt || "",
      "เวลารอสินค้าหัวเครื่องใน":       truck?.loadLanes?.lane_head?.waitingAt || "",
      "เวลาเสร็จสิ้นโหลดหัวเครื่องใน":  truck?.loadLanes?.lane_head?.doneAt || "",
      "เวลาเข้าโหลดหมูซีก STD":         "",
      "เวลาเข้าโหลดหมูซีก ACT":         truck?.qcLanes?.lane_pork?.doneAt || "",
      "เวลารอสินค้าหมูซีก":             truck?.loadLanes?.lane_pork?.waitingAt || "",
      "เวลาเสร็จสิ้นโหลดหมูซีก":       truck?.loadLanes?.lane_pork?.doneAt || "",
      "เวลาทำใบสรุปจ่าย":               truck?.summaryPrintedAt || "",
      "เวลาทำใบ Invoice":                truck?.invoicedAt || "",
      "เวลาออกจากโรงงาน":               q.exitTime || "",
      "WT ลูกค้า":                       "",
      "หมายเหตุ":                        "",
    };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, dateStr);
  XLSX.writeFile(wb, `คิวรถ_${dateStr}.xlsx`);
};

const LANE_LABEL = { lane_parts: "ลานชิ้นส่วน", lane_head: "ลานหัว/เครื่องใน", lane_pork: "ลานหมูซีก" };

const TruckTable = ({ visibleRows, allRows, searchPlate, setSearchPlate, getRemMins, myPlate, simple = false }) => {
  const containerRef = useRef(null);
  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const isMyPlate = (plate) => !!myPlate && plateNum(plate) === plateNum(myPlate) && plateNum(myPlate) !== "";
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [isTvMode, setIsTvMode] = useState(false);
  const isMobile = useIsMobile();
  useEffect(() => {
    const onChange = () => {
      const inFs = !!document.fullscreenElement;
      setIsFullscreen(inFs);
      if (!inFs) setIsTvMode(false);
    };
    document.addEventListener("fullscreenchange", onChange);
    return () => document.removeEventListener("fullscreenchange", onChange);
  }, []);
  const toggleFullscreen = () => {
    if (!document.fullscreenElement) containerRef.current?.requestFullscreen();
    else document.exitFullscreen();
  };
  const toggleTvMode = () => {
    if (!isTvMode) {
      containerRef.current?.requestFullscreen();
      setIsTvMode(true);
    } else {
      document.exitFullscreen();
    }
  };
  const fs = isTvMode;
  const Tick = () => <span style={{ color: "#10b981", fontWeight: 700, fontSize: fs ? 20 : 13 }}>✓</span>;
  const Dash = () => <span style={{ color: "#d1d5db", fontSize: fs ? 18 : 12 }}>—</span>;
  const tdP = fs ? "10px 20px" : "10px 12px";
  return (
    <div ref={containerRef} style={{ display: "flex", flexDirection: "column", height: "100%", background: "#fff" }}>
      <div style={{ padding: fs ? "10px 24px" : "10px 20px", borderBottom: "1px solid #f3f4f6", display: "flex", alignItems: "center", gap: 12, flexShrink: 0 }}>
        <span style={{ fontWeight: 700, fontSize: fs ? 18 : 14, whiteSpace: "nowrap" }}>
          🚛 รถในโรงงานวันนี้ <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: fs ? "2px 8px" : "2px 8px", fontSize: fs ? 14 : 11, marginLeft: 4 }}>{allRows.length}</span>
        </span>
        <div style={{ flex: 1 }} />
        <input
          type="text"
          placeholder="🔍 ค้นหาทะเบียน..."
          value={searchPlate}
          onChange={e => setSearchPlate(e.target.value)}
          style={{ border: "1px solid #e5e7eb", borderRadius: 0, padding: "5px 10px", fontSize: 12, width: 180, outline: "none", display: fs ? "none" : undefined, marginRight: 4 }}
        />
        {!isMobile && <button
          onClick={toggleFullscreen}
          title={isFullscreen && !isTvMode ? "ย่อหน้าต่าง (Esc)" : "ขยายเต็มจอ (Desktop)"}
          style={{ border: "1px solid #e5e7eb", borderRadius: 0, background: "#f9fafb", cursor: "pointer", padding: "4px 8px", fontSize: 15, lineHeight: 1, color: "#374151", flexShrink: 0, display: fs ? "none" : "flex", alignItems: "center", justifyContent: "center" }}
        >
          {isFullscreen && !isTvMode ? "✕" : "⛶"}
        </button>}
        {!isMobile && <button
          onClick={toggleTvMode}
          title={fs ? "ออกจากโหมด TV" : "โหมด Smart TV"}
          style={{ border: "1px solid #e5e7eb", borderRadius: 0, background: fs ? "#111" : "#f9fafb", cursor: "pointer", padding: fs ? "6px 12px" : "4px 8px", fontSize: fs ? 22 : 15, fontWeight: 800, lineHeight: 1, color: fs ? "#fff" : "#374151", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", marginLeft: fs ? "auto" : undefined }}
        >
          {fs ? "✕" : "TV"}
        </button>}
      </div>
      {visibleRows.length === 0
        ? <div style={{ padding: fs ? 40 : 36, textAlign: "center", color: "#9ca3af", fontSize: fs ? 20 : 14 }}>
            {searchPlate.trim() ? `ไม่พบทะเบียน "${searchPlate}"` : "ยังไม่มีรถเข้าโรงงาน"}
          </div>
        : isMobile && !fs
        ? <div style={{ padding: "8px 12px 16px", display: "flex", flexDirection: "column", gap: 10, overflowY: "auto", maxHeight: "calc(100vh - 130px)" }}>
            {visibleRows.map(({ key, date, plate, customerGroup, entryTime, exitTime, truck, assignedLanes }) => {
              const rem = getRemMins({ date, exitTime });
              const urgent = rem < settings.waitingUrgentMinutes && truck?.status !== "invoiced";
              const anyQC = lanes.some(l => truck?.qcLanes?.[l.id]?.done);
              const mine = isMyPlate(plate);
              const cardStyle = { background: urgent ? "#fff5f5" : mine ? "#eff6ff" : "#fff", borderRadius: 0, padding: "14px 16px", boxShadow: "0 1px 6px rgba(0,0,0,0.08)", border: urgent ? "1.5px solid #fca5a5" : mine ? "1.5px solid #bfdbfe" : "1px solid #f3f4f6" };
              if (simple) {
                return (
                  <div key={key} style={cardStyle}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                      <div style={{ fontWeight: 900, fontSize: 22, color: "#111", letterSpacing: 0.5 }}>{plate}</div>
                      <div style={{ fontSize: 11, fontWeight: 700, color: "#374151", background: "#f3f4f6", padding: "4px 10px" }}>{customerGroup || "—"}</div>
                    </div>
                    <div style={{ display: "flex", gap: 16 }}>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 10.5, fontWeight: 700, color: "#9ca3af", marginBottom: 4 }}>เวลาเข้าโรงงาน</div>
                        <div style={{ fontWeight: 700, color: "#3b82f6", fontSize: 16 }}>{entryTime || "—"}</div>
                        <div style={{ fontSize: 10.5, color: "#9ca3af", marginTop: 2 }}>{truck?.arrivedAt ? `เข้าจริง ${truck.arrivedAt}` : "ยังไม่เข้าโรงงาน"}</div>
                      </div>
                      <div style={{ width: 1, alignSelf: "stretch", background: "#e5e7eb" }} />
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 10.5, fontWeight: 700, color: "#9ca3af", marginBottom: 4 }}>เวลาออกจากโรงงาน</div>
                        <TimeBar exitTime={exitTime} date={date} done={truck?.status === "invoiced"} invoicedAt={truck?.invoicedAt} fs={false} card hideBar />
                      </div>
                    </div>
                    {exitTime && truck?.status !== "invoiced" && (
                      <div style={{ marginTop: 12 }}>
                        <TimeBarTrack exitTime={exitTime} date={date} done={truck?.status === "invoiced"} />
                      </div>
                    )}
                  </div>
                );
              }
              return (
                <div key={key} style={cardStyle}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                    <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
                      <div style={{ fontWeight: 900, fontSize: 20, color: "#111", letterSpacing: 0.5 }}>{plate}</div>
                      <div style={{ fontSize: 13, color: "#6b7280" }}>{customerGroup}</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <TimeBar exitTime={exitTime} date={date} done={truck?.status === "invoiced"} invoicedAt={truck?.invoicedAt} fs={false} card hideBar hideLabel />
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 8 }}>
                    <div style={{ fontWeight: 700, color: "#3b82f6", fontSize: 12 }}>{entryTime || "—"}</div>
                    {truck?.arrivedAt && <div style={{ fontSize: 11, color: "#9ca3af" }}>({`เข้าจริง ${truck.arrivedAt}`})</div>}
                  </div>
                  <div style={{ marginBottom: 8 }}>
                    <TimeBarTrack exitTime={exitTime} date={date} done={truck?.status === "invoiced"} />
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
                    {!truck
                      ? <span style={{ fontSize: 12, color: "#9ca3af", background: "#f9fafb", borderRadius: 0, padding: "4px 10px" }}>รอเช็คอิน</span>
                      : !anyQC
                      ? (assignedLanes?.size
                          ? lanes.filter(l => assignedLanes.has(l.id)).map(l => (
                              <span key={l.id} style={{ fontSize: 12, color: "#6b7280", background: "#f3f4f6", borderRadius: 0, padding: "4px 10px", fontWeight: 600 }}>รอเข้าโหลด {l.tinyLabel}</span>
                            ))
                          : <span style={{ fontSize: 12, color: "#6b7280", background: "#f3f4f6", borderRadius: 0, padding: "4px 10px", fontWeight: 600 }}>รอเข้าโหลด</span>)
                      : lanes.map(l => {
                          const loaded = truck.loadLanes?.[l.id]?.done;
                          const qcDone = truck.qcLanes?.[l.id]?.done;
                          const waiting = truck.loadLanes?.[l.id]?.waiting && !loaded;
                          if (loaded) return <span key={l.id} style={{ background: "#10b981", color: "#fff", borderRadius: 0, padding: "4px 12px", fontSize: 12, fontWeight: 700 }}>✓ {l.tinyLabel}</span>;
                          if (waiting) return <span key={l.id} style={{ background: "#fbbf24", color: "#fff", borderRadius: 0, padding: "4px 12px", fontSize: 12, fontWeight: 700 }}>⏳ รอสินค้า {l.tinyLabel}{truck.loadLanes?.[l.id]?.waitingFor ? ` — ${truck.loadLanes[l.id].waitingFor}` : ""}</span>;
                          if (qcDone) return <span key={l.id} style={{ background: "#fff7ed", color: "#f97316", borderRadius: 0, padding: "4px 12px", fontSize: 12, fontWeight: 700, border: "1px solid #fed7aa" }}>กำลังโหลด {l.tinyLabel}</span>;
                          if (assignedLanes?.has(l.id)) return <span key={l.id} style={{ fontSize: 12, color: "#6b7280", background: "#f3f4f6", borderRadius: 0, padding: "4px 10px", fontWeight: 600 }}>รอเข้าโหลด {l.tinyLabel}</span>;
                          return null;
                        })
                    }
                    {truck?.extraStatus && <span style={{ background: "#fee2e2", color: "#991b1b", borderRadius: 0, padding: "4px 12px", fontSize: 12, fontWeight: 700 }}>⚠️ {truck.extraStatus}</span>}
                  </div>
                  <div style={{ display: "flex", gap: 8, marginTop: 10, paddingTop: 10, borderTop: "1px solid #f3f4f6" }}>
                    {[{label:"ใบเบิก", done: truck?.pickupPrinted},{label:"ใบสรุป", done: truck?.summaryPrinted},{label:"Invoice", done: truck?.status === "invoiced"}].map(item => (
                      <span key={item.label} style={{ flex: 1, textAlign: "center", fontSize: 11, color: item.done ? "#10b981" : "#d1d5db", fontWeight: 700 }}>
                        {item.done ? "✓" : "—"} {item.label}
                      </span>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        : <div style={{ overflowX: "auto", overflowY: "auto", flex: 1, maxHeight: fs ? undefined : "calc(100vh - 170px)" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: fs ? 18 : 12, tableLayout: (fs || simple) ? "fixed" : undefined }}>
              <thead style={{ position: "sticky", top: 0, zIndex: 10 }}>
                <tr style={{ background: "#f9fafb" }}>
                  {(simple
                    ? [{l:"ทะเบียน",w:"15%"},{l:"กลุ่มลูกค้า",w:"25%"},{l:"เวลาเข้าโรงงาน",w:"20%"},{l:"เวลาออกจากโรงงาน",w:"40%"}]
                    : fs
                    ? [{l:"ทะเบียน",w:140},{l:"เวลาเข้าโรงงาน",w:160},{l:"เวลาออกจากโรงงาน",w:270},{l:"สถานะ",w:"auto"}]
                    : [{l:"ทะเบียน",w:60},{l:"กลุ่มลูกค้า",w:100},{l:"เวลาเข้าโรงงาน",w:90},{l:"เวลาออกจากโรงงาน",w:200},{l:"สถานะ",w:"auto"},{l:"ใบเบิกสินค้า",w:60},{l:"ใบสรุปจ่าย",w:60},{l:"ใบ Invoice",w:60}]
                  ).map(h => (
                    <th key={h.l} style={{ width: h.w, padding: fs ? "10px 20px" : "9px 12px", textAlign: "left", fontWeight: 700, color: "#374151", whiteSpace: "nowrap", borderBottom: "1px solid #e5e7eb", background: "#f9fafb", fontSize: fs ? 16 : undefined }}>{h.l}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {visibleRows.map(({ key, date, plate, customerGroup, entryTime, exitTime, truck, assignedLanes }) => {
                  const rem = getRemMins({ date, exitTime });
                  const urgent = rem < settings.waitingUrgentMinutes && truck?.status !== "invoiced";
                  const mine = isMyPlate(plate);
                  return (
                    <tr key={key} className={urgent ? "row-urgent" : ""} style={{ borderBottom: "1px solid #f3f4f6", background: mine ? "#eff6ff" : undefined }}>
                      <td style={{ padding: tdP, fontWeight: 800, fontSize: fs ? 22 : undefined }}>{plate}</td>
                      {(!fs || simple) && <td style={{ padding: tdP, color: "#374151", maxWidth: 100, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{customerGroup}</td>}
                      <td style={{ padding: tdP, whiteSpace: "nowrap", verticalAlign: "top" }}>
                        <div style={{ fontWeight: 700, color: "#3b82f6", fontSize: fs ? 16 : 13, lineHeight: "18px" }}>{entryTime || "—"}</div>
                        {truck?.arrivedAt
                          ? <div style={{ fontSize: fs ? 13 : 10, color: "#6b7280", marginTop: 2, lineHeight: "16px" }}>เข้าจริง {truck.arrivedAt}</div>
                          : <div style={{ fontSize: fs ? 13 : 10, color: "#6b7280", marginTop: 2, lineHeight: "16px" }}>(รถยังไม่เข้าโรงงาน)</div>}
                      </td>
                      <td style={{ padding: tdP, verticalAlign: "top" }}><TimeBar exitTime={exitTime} date={date} done={truck?.status === "invoiced"} invoicedAt={truck?.invoicedAt} fs={fs} /></td>
                      {!simple && <td style={{ padding: tdP }}>
                        {!truck
                          ? <span style={{ fontSize: fs ? 16 : 11, color: "#9ca3af" }}>รอเช็คอิน</span>
                          : (() => {
                              const anyQC = lanes.some(l => truck.qcLanes?.[l.id]?.done);
                              if (!anyQC) {
                                if (!assignedLanes?.size) return <span style={{ fontSize: fs ? 16 : 11, color: "#6b7280", fontWeight: 600 }}>รอเข้าโหลด</span>;
                                const pending = lanes.filter(l => assignedLanes.has(l.id));
                                return (
                                  <div style={{ display: "flex", flexWrap: "wrap", gap: fs ? 6 : 4 }}>
                                    {pending.map(l => (
                                      <span key={l.id} style={{ fontSize: fs ? 16 : 11, color: "#6b7280", background: "#f3f4f6", borderRadius: 0, padding: fs ? "5px 14px" : "3px 10px", fontWeight: 700, whiteSpace: "nowrap" }}>
                                        รอเข้าโหลด {l.tinyLabel}
                                      </span>
                                    ))}
                                  </div>
                                );
                              }
                              return (
                                <div style={{ display: "flex", flexWrap: "wrap", gap: fs ? 6 : 4 }}>
                                  {lanes.map(l => {
                                    const loaded = truck.loadLanes?.[l.id]?.done;
                                    const qcDone = truck.qcLanes?.[l.id]?.done;
                                    const waiting = truck.loadLanes?.[l.id]?.waiting && !loaded;
                                    if (loaded) return (
                                      <div key={l.id} style={{ position: "relative", display: "inline-block", background: "#10b981", color: "#fff", borderRadius: 0, padding: fs ? "5px 14px 7px 12px" : "3px 10px 5px 8px", fontSize: fs ? 16 : 11, fontWeight: 700, lineHeight: 1.4 }}>
                                        {l.tinyLabel}
                                        <span style={{ position: "absolute", bottom: -4, right: -4, background: "#059669", border: "2px solid #fff", borderRadius: "50%", width: fs ? 18 : 14, height: fs ? 18 : 14, display: "flex", alignItems: "center", justifyContent: "center", fontSize: fs ? 10 : 8, fontWeight: 900 }}>✓</span>
                                      </div>
                                    );
                                    if (waiting) return (
                                      <div key={l.id} style={{ position: "relative", display: "inline-block", background: "#fbbf24", color: "#fff", borderRadius: 0, padding: fs ? "5px 14px 7px 12px" : "3px 10px 5px 8px", fontSize: fs ? 16 : 11, fontWeight: 700, lineHeight: 1.4, whiteSpace: "nowrap" }}>
                                        รอสินค้า {l.tinyLabel}{truck.loadLanes?.[l.id]?.waitingFor ? ` — ${truck.loadLanes[l.id].waitingFor}` : ""}
                                        <span style={{ position: "absolute", bottom: -4, right: -4, background: "#d97706", border: "2px solid #fff", borderRadius: "50%", width: fs ? 18 : 14, height: fs ? 18 : 14, display: "flex", alignItems: "center", justifyContent: "center", fontSize: fs ? 10 : 8 }}>⏳</span>
                                      </div>
                                    );
                                    if (qcDone) return <span key={l.id} style={{ fontSize: fs ? 16 : 11, color: "#f97316", fontWeight: 700, whiteSpace: "nowrap" }}>กำลังโหลด {l.tinyLabel}</span>;
                                    if (assignedLanes?.has(l.id)) return (
                                      <span key={l.id} style={{ fontSize: fs ? 16 : 11, color: "#6b7280", background: "#f3f4f6", borderRadius: 0, padding: fs ? "5px 14px" : "3px 10px", fontWeight: 700, whiteSpace: "nowrap" }}>
                                        รอเข้าโหลด {l.tinyLabel}
                                      </span>
                                    );
                                    return null;
                                  })}
                                </div>
                              );
                            })()
                        }
                        {truck?.extraStatus && (
                          <div style={{ marginTop: 4 }}>
                            <span style={{ display: "inline-block", background: "#fee2e2", color: "#991b1b", borderRadius: 0, padding: fs ? "3px 10px" : "2px 8px", fontSize: fs ? 14 : 10, fontWeight: 700 }}>⚠️ {truck.extraStatus}</span>
                          </div>
                        )}
                      </td>}
                      {!fs && !simple && <td style={{ padding: tdP }}>{truck?.pickupPrinted ? <Tick/> : <Dash/>}</td>}
                      {!fs && !simple && <td style={{ padding: tdP }}>{truck?.summaryPrinted ? <Tick/> : <Dash/>}</td>}
                      {!fs && !simple && <td style={{ padding: tdP }}>{truck?.status === "invoiced" ? <Tick/> : <Dash/>}</td>}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
      }
    </div>
  );
};

const Dashboard = ({ trucks, queue, onReset, lane, detailMap, title, myPlate, simple = false }) => {
  const [clock, setClock] = useState(() => new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", second: "2-digit" }));
  const [searchPlate, setSearchPlate] = useState("");
  const isMobile = useIsMobile();
  useEffect(() => {
    const id = setInterval(() => setClock(new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", second: "2-digit" })), 1000);
    return () => clearInterval(id);
  }, []);
  const cnt = (s) => trucks.filter(t => t.status === s).length;
  const stats = [
    { label: "คิวรอเข้า",         value: queue.filter(q => !trucks.find(t => t.queueId === q.id)).length, color: "#3b82f6", icon: "list"    },
    { label: "รถเข้าโรงงานแล้ว", value: trucks.length,                                                    color: "#22c55e", icon: "truck"   },
    { label: "กำลังโหลด",         value: cnt("arrived") + cnt("picking"),                                  color: "#f97316", icon: "loader"  },
    { label: "Invoice แล้ว",      value: cnt("invoiced"),                                                  color: "#6b7280", icon: "invoice" },
  ];

  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const toMins = t => { const [h, m] = t.split(":").map(Number); return h * 60 + m; };
  const nowMins = new Date().getHours() * 60 + new Date().getMinutes();
  const getRemMins = (row) => {
    const dt = parseExitDatetime(row.date, row.exitTime);
    return dt ? Math.round((dt - Date.now()) / 60000) : Infinity;
  };
  const usedDash = new Set();
  const matchTruckDash = q => {
    let t = trucks.find(t => t.queueId === q.id && !usedDash.has(t.id));
    if (!t) t = trucks.find(t => !t.queueId && plateNum(t.plate) === plateNum(q.plate) && plateNum(q.plate) !== "" && !usedDash.has(t.id));
    if (t) usedDash.add(t.id);
    return t;
  };
  const dashQueueRows = queue.map(q => {
    const truck = matchTruckDash(q);
    return { key: q.id, date: q.date || "", plate: q.plate, customerGroup: q.customerGroup, entryTime: q.entryTime, exitTime: q.exitTime, truck, assignedLanes: laneMatchForTruck(truck, detailMap) };
  });
  const walkIns = trucks.filter(t => !usedDash.has(t.id));
  const allRows = [
    ...dashQueueRows,
    ...walkIns.map(t => ({ key: t.id, date: t.date || "", plate: t.plate, customerGroup: t.customerGroup || "–", entryTime: t.entryTime || "", exitTime: t.exitTime || "", truck: t, assignedLanes: laneMatchForTruck(t, detailMap) })),
  ].filter(row => !settings.excludedCustomerGroups.includes(row.customerGroup))
  // on a lane-specific tab: hide trucks confirmed (matched against PO/master data) to use a
  // *different* lane; keep unmatched trucks (assignedLanes empty — not known yet) visible on
  // every lane tab until the match becomes clear, and keep hiding ones already done on this lane
  .filter(row => !lane || (
    (!row.assignedLanes?.size || row.assignedLanes.has(lane)) &&
    !row.truck?.loadLanes?.[lane]?.done
  ))
  .sort((a, b) => {
    const rank = row => {
      if (["invoiced", "summary_printed"].includes(row.truck?.status)) return 5;

      if (lane) {
        const qcDone = row.truck?.qcLanes?.[lane]?.done;
        const waiting = row.truck?.loadLanes?.[lane]?.waiting;
        if (qcDone || waiting) return 0;   // กำลังโหลดอยู่ → บนสุด
      } else {
        const anyActive = lanes.some(l =>
          row.truck?.qcLanes?.[l.id]?.done && !row.truck?.loadLanes?.[l.id]?.done
        );
        if (anyActive) return 0;           // กำลังโหลดอยู่ → บนสุด
      }

      if (!row.truck) return 2;            // 2.2 ยังไม่เข้าโรงงาน
      return 1;                            // 2.1 เข้าโรงงานแล้ว แต่ไม่ได้โหลดอยู่
    };
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    return getRemMins(a) - getRemMins(b);
  });
  const visibleRows = searchPlate.trim()
    ? allRows.filter(r => r.plate?.toLowerCase().includes(searchPlate.trim().toLowerCase()))
    : allRows;

  return (
    <div>
      {/* Sticky header */}
      <div style={{ position: "sticky", top: 80, zIndex: 40, background: "#f1f5f9", paddingBottom: isMobile ? 6 : 8, paddingTop: 0 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 4 }}>
          <h2 style={{ margin: 0, fontSize: isMobile ? 16 : 22, fontWeight: 900 }}>{title || (lane ? LANE_LABEL[lane] : "Main Dashboard")}</h2>
        </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 8px rgba(0,0,0,0.07)", overflow: "hidden", marginTop: 8 }}>
        <TruckTable visibleRows={visibleRows} allRows={allRows} searchPlate={searchPlate} setSearchPlate={setSearchPlate} getRemMins={getRemMins} myPlate={myPlate} simple={simple} />
      </div>
    </div>
  );
};

// ── 1. LG UPLOAD (Excel → parse → queue) ─────────────────────────────────────
// ตำแหน่งคอลัมน์ในไฟล์ Excel ล็อคตายตัว (ไม่อ่านจากชื่อ header) — ข้อมูลเริ่มแถวที่ 3
// K=วันที่ D=ทะเบียนรถ AS=กลุ่มลูกค้า L=Zone M=เวลาเข้าโรงงาน N=เวลาออกจากโรงงาน
const LG_UPLOAD_DATA_START_ROW = 2; // แถวที่ 3 (index 0-based)
const LG_UPLOAD_COLS = {
  date:          10, // K
  plate:         3,  // D
  customerGroup: 44, // AS
  zone:          11, // L
  entryTime:     12, // M
  exitTime:      13, // N
};

const parseQueueDateToISO = (dateStr) => {
  if (!dateStr) return DATE_STR();
  const parts = dateStr.split("/");
  if (parts.length !== 3) return DATE_STR();
  const [d, m, y] = parts.map(Number);
  return `${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
};

// Axons Move exports the date as plain text — either a bare "16/08/2569" or a full
// "16/08/2569 17:00:00" — using the Thai Buddhist-era year (2569 = 2026 CE). Every other
// date in the app (SHORT_DATE/DATE_STR/manual entry) is Gregorian, so passing the BE year
// straight through silently threw exit-time countdowns ~543 years into the future.
const toDateStr = (val) => {
  if (val === "" || val == null) return "";
  if (typeof val === "string" && /\d/.test(val)) {
    const m = val.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      const [, d, mo, yRaw] = m;
      let y = Number(yRaw);
      if (y > 2400) y -= 543;
      return `${Number(d)}/${Number(mo)}/${y}`;
    }
    return val.trim();
  }
  const num = typeof val === "number" ? val : parseFloat(val);
  if (!isNaN(num) && num > 1000) {
    const d = new Date(Math.round((num - 25569) * 86400 * 1000));
    return `${d.getUTCDate()}/${d.getUTCMonth() + 1}/${d.getUTCFullYear()}`;
  }
  return String(val).trim();
};

const toHHMM = (val) => {
  if (val === "" || val == null) return "";
  // string containing a time — either bare "17:00[:00]" or a full "16/08/2569 17:00:00" datetime
  if (typeof val === "string") {
    const trimmed = val.trim();
    const m = trimmed.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
    if (m) return `${m[1].padStart(2, "0")}:${m[2]}`;
    if (trimmed.includes("/")) return ""; // date-only string (no time part) — don't misread as a serial number
  }
  // Excel stores time as fraction of a day (0.875 = 21:00)
  const num = typeof val === "number" ? val : parseFloat(val);
  if (!isNaN(num)) {
    const frac = num - Math.floor(num); // strip date part
    const mins = Math.round(frac * 1440);
    const h = Math.floor(mins / 60) % 24;
    const m = mins % 60;
    return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
  }
  return String(val);
};

// เวลาออกจากโรงงาน = เวลาโหลดเสร็จ (ค่าที่อ่านจากไฟล์ Excel) + เวลาทำเอกสาร (settings.docTimeMinutes, ปรับได้ที่หน้าตั้งค่าระบบ)
const addMinutesToHHMM = (hhmm, mins) => {
  if (!hhmm) return "";
  const m = hhmm.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return hhmm;
  const total = ((parseInt(m[1], 10) * 60 + parseInt(m[2], 10) + Number(mins || 0)) % 1440 + 1440) % 1440;
  return `${String(Math.floor(total / 60)).padStart(2, "0")}:${String(total % 60).padStart(2, "0")}`;
};

// เวลาออกจากโรงงานถูกคำนวณและบันทึกทับไว้ใน wh_queue/wh_master(pending_queue) ตอนอัพโหลด/แก้ไข —
// ถ้าแค่เปลี่ยนค่า "เวลาทำเอกสาร" ที่ Master Setting เฉยๆ แถวคิวที่มีอยู่แล้วจะไม่ขยับตาม (ค้างค่าที่คำนวณไว้ครั้งก่อน)
// ฟังก์ชันนี้จึงไล่คำนวณ exitTime ใหม่จาก loadDoneTime + เวลาทำเอกสารตัวล่าสุด แล้วเขียนทับให้ตรงกันทันทีที่ Master Setting ถูกบันทึก
const recomputeQueueExitTimes = async (docMins) => {
  try {
    const { data, error } = await supabase.from("wh_queue").select("*");
    if (!error && data) {
      const updates = data
        .map(row => {
          const q = row.data;
          if (!q?.loadDoneTime) return null;
          const exitTime = addMinutesToHHMM(q.loadDoneTime, docMins);
          return exitTime === q.exitTime ? null : { id: row.id, data: { ...q, exitTime } };
        })
        .filter(Boolean);
      if (updates.length) await supabase.from("wh_queue").upsert(updates);
    }

    const { data: pendingRow } = await supabase.from("wh_master").select("*").eq("id", "pending_queue").single();
    const pending = pendingRow?.data;
    if (Array.isArray(pending?.rows) && pending.rows.length) {
      const newRows = pending.rows.map(q => q.loadDoneTime ? { ...q, exitTime: addMinutesToHHMM(q.loadDoneTime, docMins) } : q);
      await supabase.from("wh_master").upsert({ id: "pending_queue", data: { ...pending, rows: newRows } });
    }
  } catch (e) {
    console.error("recomputeQueueExitTimes ไม่สำเร็จ:", e);
  }
};

// ถ้า exitTime อยู่ก่อนเวลาตัดรอบวันทำงาน (settings.workDayCutoffHour) → วันที่จริงคือ dateStr + 1 (กะข้ามคืน)
const displayDate = (dateStr, exitTime) => {
  if (!dateStr || !exitTime) return dateStr || "";
  const [hStr, minStr] = exitTime.split(":");
  const h = parseInt(hStr, 10); const min = parseInt(minStr, 10);
  if (isNaN(h) || isNaN(min) || h * 60 + min > settings.workDayCutoffHour * 60) return dateStr;
  const parts = dateStr.split("/");
  if (parts.length !== 3) return dateStr;
  let [d, m, y] = parts.map(Number);
  if (m > 12) { [d, m] = [m, d]; }
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + 1);
  return `${dt.getDate()}/${dt.getMonth() + 1}/${dt.getFullYear()}`;
};

const LGUpload = ({ queue, onSetQueue, pendingQueue, onSetPendingQueue, onClearPendingQueue }) => {
  const isMobile = useIsMobile();
  const [mode,     setMode]     = useState("today"); // "today" | "advance"
  const [fileName, setFileName] = useState("");
  const [status,   setStatus]   = useState("idle"); // idle | preview | uploading | done | error
  const [extracted, setExtracted] = useState([]);
  const [errMsg,   setErrMsg]   = useState("");
  const [savedCount, setSavedCount] = useState(0);
  const [clearingPending, setClearingPending] = useState(false);
  const [editId,   setEditId]   = useState(null);
  const [editData, setEditData] = useState({});

  const [addingManual, setAddingManual] = useState(false);
  const [manualData, setManualData] = useState({ date: "", plate: "", customerGroup: "", zone: "", entryTime: "", loadDoneTime: "" });
  const [searchQuery, setSearchQuery] = useState("");

  const [queueSaving, setQueueSaving] = useState(false);
  const startEdit = (q) => { setEditId(q.id); setEditData({ plate: q.plate, customerGroup: q.customerGroup, zone: q.zone || "", entryTime: q.entryTime, loadDoneTime: q.loadDoneTime || "" }); };
  const cancelEdit = () => { setEditId(null); setEditData({}); };
  const saveEdit = async () => {
    setQueueSaving(true);
    try {
      const exitTime = addMinutesToHHMM(editData.loadDoneTime, settings.docTimeMinutes);
      await onSetQueue(queue.map(q => q.id === editId ? { ...q, ...editData, zone: editData.zone, time: editData.entryTime, exitTime } : q), "lg_edit_row");
      setEditId(null); setEditData({});
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setQueueSaving(false);
    }
  };
  const deleteRow = async (id) => {
    if (!window.confirm("ลบรถคันนี้ออกจากคิว?")) return;
    setQueueSaving(true);
    try {
      await onSetQueue(queue.filter(q => q.id !== id), "lg_delete_row");
    } catch (e) {
      alert("ลบไม่สำเร็จ: " + e.message);
    } finally {
      setQueueSaving(false);
    }
  };
  const saveManual = async () => {
    if (!manualData.plate) return;
    setQueueSaving(true);
    try {
      const exitTime = addMinutesToHHMM(manualData.loadDoneTime, settings.docTimeMinutes);
      await onSetQueue([...queue, { id: `M${Date.now()}`, ...manualData, exitTime, date: manualData.date || SHORT_DATE(), time: manualData.entryTime, driver: "", zone: manualData.zone || "", product: "", destination: "", qty: 0, unit: "กก.", loadTime: "" }], "lg_manual_add");
      setManualData({ date: "", plate: "", customerGroup: "", zone: "", entryTime: "", loadDoneTime: "" });
      setAddingManual(false);
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setQueueSaving(false);
    }
  };

  const inputStyle = { border: "1px solid #d1d5db", borderRadius: 0, padding: "4px 8px", fontSize: 12, width: "100%", boxSizing: "border-box" };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setErrMsg("");
    setStatus("idle");

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb   = XLSX.read(ev.target.result, { type: "array", cellDates: false });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });

        if (rows.length <= LG_UPLOAD_DATA_START_ROW) throw new Error("ไม่พบข้อมูลในไฟล์");

        const trucks = [];
        for (let i = LG_UPLOAD_DATA_START_ROW; i < rows.length; i++) {
          const row = rows[i];
          const plate = String(row[LG_UPLOAD_COLS.plate] ?? "").trim();
          if (!plate) continue;
          const loadDoneTime = toHHMM(row[LG_UPLOAD_COLS.exitTime]);
          trucks.push({
            date:          toDateStr(row[LG_UPLOAD_COLS.date]),
            plate,
            customerGroup: String(row[LG_UPLOAD_COLS.customerGroup] ?? "").trim(),
            zone:          String(row[LG_UPLOAD_COLS.zone] ?? "").trim(),
            entryTime:     toHHMM(row[LG_UPLOAD_COLS.entryTime]),
            loadDoneTime,
            exitTime:      addMinutesToHHMM(loadDoneTime, settings.docTimeMinutes),
          });
        }

        if (trucks.length === 0) throw new Error("ไม่พบข้อมูลทะเบียนรถ — ตรวจสอบว่าคอลัมน์ D มีเลขทะเบียนตั้งแต่แถวที่ 3");
        setExtracted(trucks);
        setStatus("preview");
      } catch (err) {
        setErrMsg(err.message);
        setStatus("error");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleConfirm = async () => {
    const newQueue = extracted.map((t, i) => ({
      id:            `Q${Date.now()}-${i}`,
      seq:           i,
      date:          t.date          || "",
      plate:         t.plate        || "",
      driver:        "",
      customerGroup: t.customerGroup || "",
      zone:          t.zone          || "",
      product:       t.customerGroup || "",
      destination:   t.zone          || "",
      qty:           0,
      unit:          "กก.",
      time:          t.entryTime    || "",
      entryTime:     t.entryTime    || "",
      loadTime:      t.loadTime     || "",
      loadDoneTime:  t.loadDoneTime || "",
      exitTime:      t.exitTime     || "",
    }));
    setStatus("uploading");
    setErrMsg("");
    try {
      if (mode === "advance") {
        await onSetPendingQueue(newQueue, fileName);
      } else {
        await onSetQueue(newQueue, "lg_upload_excel");
      }
      setSavedCount(newQueue.length);
      setStatus("done");
      setExtracted([]);
      setFileName("");
    } catch (err) {
      setErrMsg("บันทึกไม่สำเร็จ: " + err.message);
      setStatus("error");
    }
  };

  const clearPending = async () => {
    if (!window.confirm("ยกเลิกคิวรถล่วงหน้าที่อัพโหลดไว้?")) return;
    setClearingPending(true);
    try {
      await onClearPendingQueue();
    } catch (e) {
      alert("ยกเลิกไม่สำเร็จ: " + e.message);
    } finally {
      setClearingPending(false);
    }
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 4px", fontWeight: 900, fontSize: 22 }}>🤝 LG → Upload ตารางคิวรถ</h2>
      <p style={{ margin: "0 0 18px", fontSize: 13, color: "#6b7280" }}>ไฟล์ Export จาก Axons Move — DPI04000 (รายงาน Scheduling)</p>

      {/* Mode tabs: อัพวันนี้ vs อัพล่วงหน้า */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        {[
          { id: "today",   label: "📋 คิวรถวันนี้" },
          { id: "advance", label: "🗓️ อัพโหลดล่วงหน้า" },
        ].map(m => (
          <button key={m.id} onClick={() => { setMode(m.id); setStatus("idle"); setFileName(""); setExtracted([]); setErrMsg(""); }}
            style={{
              flex: 1, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer",
              border: mode === m.id ? "1.5px solid #111" : "1.5px solid #d1d5db",
              background: mode === m.id ? "#111" : "#fff", color: mode === m.id ? "#fff" : "#374151",
              borderRadius: 0,
            }}>
            {m.label}
          </button>
        ))}
      </div>

      {mode === "advance" && (
        <div style={{ padding: "12px 16px", background: "#eff6ff", borderRadius: 0, color: "#1e3a8a", fontSize: 12.5, marginBottom: 14, lineHeight: 1.6 }}>
          ใช้แท็บนี้เพื่ออัพโหลดคิวรถของ<b>วันทำงานถัดไป</b>ล่วงหน้าได้เลย โดยไม่กระทบคิวของวันนี้ —
          ระบบจะดึงคิวนี้มาใช้เป็น "คิวรถวันปัจจุบัน" ให้อัตโนมัติทันทีที่ข้ามเวลาตัดรอบวันทำงาน (ประมาณ {String(settings.workDayCutoffHour).padStart(2, "0")}:00 น.)
          พร้อมปิดยอด/เก็บข้อมูลของวันเก่าเข้าประวัติให้เองโดยไม่ต้องกด "ล้างวันใหม่"
        </div>
      )}

      {/* Pending advance-queue status */}
      {pendingQueue && Array.isArray(pendingQueue.rows) && (
        <div style={{ padding: "14px 16px", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 0, marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontSize: 12.5, color: "#92400e", lineHeight: 1.6 }}>
            ⏳ มีคิวรถล่วงหน้ารออยู่ <b>{pendingQueue.rows.length} คัน</b> ({pendingQueue.fileName || "ไม่ทราบชื่อไฟล์"})
            <br />จะถูกใช้งานอัตโนมัติเป็นคิวของวันที่ <b>{new Date(pendingQueue.forDate).toLocaleDateString("th-TH", { year: "numeric", month: "long", day: "numeric" })}</b>
          </div>
          <button onClick={clearPending} disabled={clearingPending}
            style={{ background: "#fee2e2", color: "#991b1b", border: "none", borderRadius: 0, padding: "7px 14px", fontSize: 12, fontWeight: 700, cursor: clearingPending ? "not-allowed" : "pointer", whiteSpace: "nowrap" }}>
            {clearingPending ? "..." : "ยกเลิกคิวล่วงหน้า"}
          </button>
        </div>
      )}

      {/* Upload zone */}
      <label style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 10, background: fileName ? "#f0fdf4" : "#fafafa", border: `2px dashed ${fileName ? "#6ee7b7" : "#d1d5db"}`, borderRadius: 0, padding: 30, textAlign: "center", cursor: "pointer", marginBottom: 14 }}>
        <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
        <Icon name="upload" size={36} />
        {fileName
          ? <><div style={{ fontWeight: 800, color: "#065f46", fontSize: 14 }}>📄 {fileName}</div><div style={{ fontSize: 12, color: "#6b7280" }}>แตะเพื่อเปลี่ยนไฟล์</div></>
          : <><div style={{ fontWeight: 700, color: "#374151" }}>แตะเพื่ออัปโหลดไฟล์ Excel</div><div style={{ fontSize: 12, color: "#9ca3af" }}>รองรับ .xlsx, .xls, .csv</div></>
        }
      </label>

      {/* Error */}
      {status === "error" && (
        <div style={{ padding: "12px 16px", background: "#fee2e2", borderRadius: 0, color: "#991b1b", fontWeight: 600, fontSize: 13, marginBottom: 14 }}>
          ❌ {errMsg}
          <div style={{ marginTop: 8, fontWeight: 400, fontSize: 12 }}>
            Column ที่รองรับ: <b>ทะเบียนรถ · กลุ่มลูกค้า · Zone · เข้าโรงงาน · เข้าโหลด · ออก</b>
          </div>
        </div>
      )}

      {/* Preview */}
      {status === "preview" && extracted.length > 0 && (
        <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", marginBottom: 14, overflow: "hidden" }}>
          <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontWeight: 800, fontSize: 14 }}>✅ อ่านข้อมูลได้ {extracted.length} คัน</div>
            <span style={{ fontSize: 12, color: "#6b7280" }}>ตรวจสอบแล้วกด "ยืนยัน"</span>
          </div>
          <div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: isMobile ? 11 : 12 }}>
              <thead>
                <tr style={{ background: "#f9fafb" }}>
                  {(isMobile ? ["ทะเบียนรถ","กลุ่มลูกค้า","Zone","เวลา"] : ["วันที่","ทะเบียนรถ","กลุ่มลูกค้า","Zone","เวลาเข้าโรงงาน","เวลาโหลดเสร็จ","เวลาออกจากโรงงาน"]).map(h => (
                    <th key={h} style={{ padding: isMobile ? "6px 6px" : "8px 12px", textAlign: "left", fontWeight: 700, color: "#374151", whiteSpace: "nowrap", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {extracted.map((t, i) => isMobile ? (
                  <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "6px 6px", fontWeight: 800 }}>
                      {t.plate}
                      <div style={{ fontWeight: 400, color: "#9ca3af", fontSize: 10 }}>{displayDate(t.date, t.exitTime)}</div>
                    </td>
                    <td style={{ padding: "6px 6px" }}>{t.customerGroup}</td>
                    <td style={{ padding: "6px 6px", fontWeight: 700, color: "#7c3aed" }}>{t.zone}</td>
                    <td style={{ padding: "6px 6px" }}>
                      <div style={{ fontWeight: 700, color: "#3b82f6" }}>{t.entryTime}</div>
                      <div style={{ fontWeight: 700, color: "#9ca3af" }}>{t.loadDoneTime}</div>
                      <div style={{ fontWeight: 700, color: "#6b7280" }}>{t.exitTime}</div>
                    </td>
                  </tr>
                ) : (
                  <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "8px 12px", color: "#6b7280" }}>{displayDate(t.date, t.exitTime)}</td>
                    <td style={{ padding: "8px 12px", fontWeight: 800 }}>{t.plate}</td>
                    <td style={{ padding: "8px 12px" }}>{t.customerGroup}</td>
                    <td style={{ padding: "8px 12px", fontWeight: 700, color: "#7c3aed" }}>{t.zone}</td>
                    <td style={{ padding: "8px 12px", fontWeight: 700, color: "#3b82f6" }}>{t.entryTime}</td>
                    <td style={{ padding: "8px 12px", fontWeight: 700, color: "#9ca3af" }}>{t.loadDoneTime}</td>
                    <td style={{ padding: "8px 12px", fontWeight: 700, color: "#6b7280" }}>{t.exitTime}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ padding: 16 }}>
            <button onClick={handleConfirm} disabled={status === "uploading"}
              style={{ width: "100%", background: status === "uploading" ? "#6ee7b7" : "#10b981", color: "#fff", border: "none", borderRadius: 0, padding: "13px 0", fontWeight: 700, fontSize: 15, cursor: status === "uploading" ? "not-allowed" : "pointer" }}>
              {status === "uploading" ? "⏳ กำลังบันทึก..." : mode === "advance" ? `✅ ยืนยัน — บันทึกคิวรถล่วงหน้า ${extracted.length} คัน` : `✅ ยืนยัน — ตั้งคิวรถ ${extracted.length} คัน`}
            </button>
          </div>
        </div>
      )}

      {status === "done" && (
        <div style={{ padding: "13px 16px", background: "#d1fae5", borderRadius: 0, color: "#065f46", fontWeight: 700, marginBottom: 14 }}>
          {mode === "advance"
            ? `✅ บันทึกคิวรถล่วงหน้าเรียบร้อย ${savedCount} คัน — จะใช้งานอัตโนมัติเมื่อขึ้นวันทำงานถัดไป`
            : `✅ ตั้งคิวรถเรียบร้อย ${savedCount} คัน — พร้อมให้คนขับ Scan เข้า`}
        </div>
      )}

      {/* Current queue list */}
      {mode === "today" && queue.length > 0 && (() => {
        const filteredQueue = queue.filter(q => q.plate.includes(searchQuery));
        return (
        <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 8px rgba(0,0,0,0.07)", overflow: "hidden" }}>
          <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
            <div style={{ fontWeight: 700, fontSize: 14 }}>
              คิวรถวันนี้ <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{filteredQueue.length}</span>
            </div>
            <input type="text" placeholder="🔍 ค้นหาทะเบียนรถ..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)}
              style={{ border: "1px solid #d1d5db", borderRadius: 0, padding: "6px 12px", fontSize: 12, outline: "none", width: isMobile ? "100%" : 160 }} />
          </div>
          <div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: isMobile ? 11 : 12 }}>
              <thead>
                <tr style={{ background: "#f9fafb" }}>
                  {(isMobile ? ["ทะเบียนรถ","กลุ่มลูกค้า","Zone","เวลา",""] : ["วันที่","ทะเบียนรถ","กลุ่มลูกค้า","Zone","เวลาเข้าโรงงาน","เวลาโหลดเสร็จ","เวลาออกจากโรงงาน",""]).map(h => (
                    <th key={h} style={{ padding: isMobile ? "6px 6px" : "8px 12px", textAlign: "left", fontWeight: 700, color: "#374151", whiteSpace: "nowrap", borderBottom: "1px solid #e5e7eb" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredQueue.map(q => {
                  const isEditing = editId === q.id;
                  const actions = isEditing ? (
                    <div style={{ display: "flex", gap: 4 }}>
                      <button onClick={saveEdit} disabled={queueSaving} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{queueSaving ? "..." : "บันทึก"}</button>
                      <button onClick={cancelEdit} disabled={queueSaving} style={{ background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>ยกเลิก</button>
                    </div>
                  ) : (
                    <div style={{ display: "flex", gap: 4 }}>
                      <button onClick={() => startEdit(q)} disabled={queueSaving} style={{ background: "#eff6ff", color: "#1d4ed8", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>แก้ไข</button>
                      <button onClick={() => deleteRow(q.id)} disabled={queueSaving} style={{ background: "#fee2e2", color: "#991b1b", border: "none", borderRadius: 0, padding: "4px 8px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>ลบ</button>
                    </div>
                  );
                  return isMobile ? (
                    <tr key={q.id} style={{ borderBottom: "1px solid #f3f4f6", background: isEditing ? "#fffbeb" : undefined }}>
                      <td style={{ padding: "6px 6px", fontWeight: 800 }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.plate} onChange={e => setEditData(d => ({ ...d, plate: e.target.value }))} />
                          : <>{q.plate}<div style={{ fontWeight: 400, color: "#9ca3af", fontSize: 10 }}>{displayDate(q.date, q.exitTime)}</div></>}
                      </td>
                      <td style={{ padding: "6px 6px" }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.customerGroup} onChange={e => setEditData(d => ({ ...d, customerGroup: e.target.value }))} />
                          : q.customerGroup}
                      </td>
                      <td style={{ padding: "6px 6px", fontWeight: 700, color: "#7c3aed" }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.zone} placeholder="Zone" onChange={e => setEditData(d => ({ ...d, zone: e.target.value }))} />
                          : q.zone}
                      </td>
                      <td style={{ padding: "6px 6px" }}>
                        {isEditing
                          ? <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                              <input style={inputStyle} value={editData.entryTime} placeholder="เข้า HH:MM" onChange={e => setEditData(d => ({ ...d, entryTime: e.target.value }))} />
                              <input style={inputStyle} value={editData.loadDoneTime} placeholder="โหลดเสร็จ HH:MM" onChange={e => setEditData(d => ({ ...d, loadDoneTime: e.target.value }))} />
                              <div style={{ fontSize: 10, color: "#9ca3af" }}>ออก: {addMinutesToHHMM(editData.loadDoneTime, settings.docTimeMinutes) || "–"}</div>
                            </div>
                          : <>
                              <div style={{ fontWeight: 700, color: "#3b82f6" }}>{q.entryTime}</div>
                              <div style={{ fontWeight: 700, color: "#9ca3af" }}>{q.loadDoneTime}</div>
                              <div style={{ fontWeight: 700, color: "#6b7280" }}>{q.exitTime}</div>
                            </>}
                      </td>
                      <td style={{ padding: "6px 4px", whiteSpace: "nowrap" }}>{actions}</td>
                    </tr>
                  ) : (
                    <tr key={q.id} style={{ borderBottom: "1px solid #f3f4f6", background: isEditing ? "#fffbeb" : undefined }}>
                      <td style={{ padding: "8px 12px", color: "#6b7280" }}>{displayDate(q.date, q.exitTime)}</td>
                      <td style={{ padding: "8px 12px", fontWeight: 800 }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.plate} onChange={e => setEditData(d => ({ ...d, plate: e.target.value }))} />
                          : q.plate}
                      </td>
                      <td style={{ padding: "8px 12px" }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.customerGroup} onChange={e => setEditData(d => ({ ...d, customerGroup: e.target.value }))} />
                          : q.customerGroup}
                      </td>
                      <td style={{ padding: "8px 12px", fontWeight: 700, color: "#7c3aed" }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.zone} placeholder="Zone" onChange={e => setEditData(d => ({ ...d, zone: e.target.value }))} />
                          : q.zone}
                      </td>
                      <td style={{ padding: "8px 12px", fontWeight: 700, color: "#3b82f6" }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.entryTime} placeholder="HH:MM" onChange={e => setEditData(d => ({ ...d, entryTime: e.target.value }))} />
                          : q.entryTime}
                      </td>
                      <td style={{ padding: "8px 12px", fontWeight: 700, color: "#9ca3af" }}>
                        {isEditing
                          ? <input style={inputStyle} value={editData.loadDoneTime} placeholder="HH:MM" onChange={e => setEditData(d => ({ ...d, loadDoneTime: e.target.value }))} />
                          : q.loadDoneTime}
                      </td>
                      <td style={{ padding: "8px 12px", fontWeight: 700, color: "#6b7280" }}>
                        {isEditing
                          ? (addMinutesToHHMM(editData.loadDoneTime, settings.docTimeMinutes) || "–")
                          : q.exitTime}
                      </td>
                      <td style={{ padding: "8px 8px", whiteSpace: "nowrap" }}>{actions}</td>
                    </tr>
                  );
                })}
                {addingManual && (isMobile ? (
                  <tr style={{ borderBottom: "1px solid #f3f4f6", background: "#f0fdf4" }}>
                    <td style={{ padding: "6px 6px" }}>
                      <input style={{ ...inputStyle, marginBottom: 3 }} placeholder="24/4/2026" value={manualData.date} onChange={e => setManualData(d => ({ ...d, date: e.target.value }))} />
                      <input style={inputStyle} placeholder="กข-1234" value={manualData.plate} onChange={e => setManualData(d => ({ ...d, plate: e.target.value }))} />
                    </td>
                    <td style={{ padding: "6px 6px" }}><input style={inputStyle} placeholder="กลุ่มลูกค้า" value={manualData.customerGroup} onChange={e => setManualData(d => ({ ...d, customerGroup: e.target.value }))} /></td>
                    <td style={{ padding: "6px 6px" }}><input style={inputStyle} placeholder="Zone" value={manualData.zone} onChange={e => setManualData(d => ({ ...d, zone: e.target.value }))} /></td>
                    <td style={{ padding: "6px 6px" }}>
                      <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                        <input style={inputStyle} placeholder="เข้า HH:MM" value={manualData.entryTime} onChange={e => setManualData(d => ({ ...d, entryTime: e.target.value }))} />
                        <input style={inputStyle} placeholder="โหลดเสร็จ HH:MM" value={manualData.loadDoneTime} onChange={e => setManualData(d => ({ ...d, loadDoneTime: e.target.value }))} />
                        <div style={{ fontSize: 10, color: "#9ca3af" }}>ออก: {addMinutesToHHMM(manualData.loadDoneTime, settings.docTimeMinutes) || "–"}</div>
                      </div>
                    </td>
                    <td style={{ padding: "6px 4px", whiteSpace: "nowrap" }}>
                      <div style={{ display: "flex", gap: 4 }}>
                        <button onClick={saveManual} disabled={queueSaving} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{queueSaving ? "..." : "บันทึก"}</button>
                        <button onClick={() => setAddingManual(false)} disabled={queueSaving} style={{ background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>ยกเลิก</button>
                      </div>
                    </td>
                  </tr>
                ) : (
                  <tr style={{ borderBottom: "1px solid #f3f4f6", background: "#f0fdf4" }}>
                    <td style={{ padding: "6px 8px" }}><input style={inputStyle} placeholder="24/4/2026" value={manualData.date} onChange={e => setManualData(d => ({ ...d, date: e.target.value }))} /></td>
                    <td style={{ padding: "6px 8px" }}><input style={inputStyle} placeholder="กข-1234" value={manualData.plate} onChange={e => setManualData(d => ({ ...d, plate: e.target.value }))} /></td>
                    <td style={{ padding: "6px 8px" }}><input style={inputStyle} placeholder="กลุ่มลูกค้า" value={manualData.customerGroup} onChange={e => setManualData(d => ({ ...d, customerGroup: e.target.value }))} /></td>
                    <td style={{ padding: "6px 8px" }}><input style={inputStyle} placeholder="Zone" value={manualData.zone} onChange={e => setManualData(d => ({ ...d, zone: e.target.value }))} /></td>
                    <td style={{ padding: "6px 8px" }}><input style={inputStyle} placeholder="HH:MM" value={manualData.entryTime} onChange={e => setManualData(d => ({ ...d, entryTime: e.target.value }))} /></td>
                    <td style={{ padding: "6px 8px" }}><input style={inputStyle} placeholder="HH:MM" value={manualData.loadDoneTime} onChange={e => setManualData(d => ({ ...d, loadDoneTime: e.target.value }))} /></td>
                    <td style={{ padding: "6px 8px", color: "#9ca3af" }}>{addMinutesToHHMM(manualData.loadDoneTime, settings.docTimeMinutes) || "–"}</td>
                    <td style={{ padding: "6px 8px", whiteSpace: "nowrap" }}>
                      <div style={{ display: "flex", gap: 4 }}>
                        <button onClick={saveManual} disabled={queueSaving} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{queueSaving ? "..." : "บันทึก"}</button>
                        <button onClick={() => setAddingManual(false)} disabled={queueSaving} style={{ background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>ยกเลิก</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ padding: "12px 16px", borderTop: "1px solid #f3f4f6" }}>
            <button onClick={() => setAddingManual(true)} style={{ background: "#eff6ff", color: "#1d4ed8", border: "1px dashed #93c5fd", borderRadius: 0, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", width: "100%" }}>
              + เพิ่มทะเบียนรถ Manual
            </button>
          </div>
        </div>
        );
      })()}
    </div>
  );
};

// ── 2. DRIVER SCAN ────────────────────────────────────────────────────────────
const DriverScan = ({ queue, trucks, onScan, skipGeofence }) => {
  const [plate, setPlate] = useState("");
  const [step, setStep] = useState("input"); // "input" | "confirm"
  const [pendingEntry, setPendingEntry] = useState(null);
  const [selectedZone, setSelectedZone] = useState("");
  const [msg, setMsg] = useState(null);
  const geo = useGeofence();

  // ── Geofence Gate ──
  if (!skipGeofence && geo.status !== "inside") {
    return (
      <div style={{ minHeight: "70vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ background: "#fff", borderRadius: 0, padding: "40px 28px", boxShadow: "0 4px 24px rgba(0,0,0,0.10)", textAlign: "center", maxWidth: 380, width: "100%" }}>
          {geo.status === "idle" && (
            <>
              <div style={{ fontSize: 56, marginBottom: 16 }}>📍</div>
              <h2 style={{ margin: "0 0 8px", fontWeight: 900, fontSize: 20 }}>เช็คอินเข้าโรงงาน</h2>
              <p style={{ color: "#6b7280", fontSize: 14, margin: "0 0 24px", lineHeight: 1.6 }}>
                กรุณาอนุญาตการเข้าถึงตำแหน่ง<br />เพื่อยืนยันว่าคุณอยู่ใกล้โรงงาน
              </p>
              <button onClick={geo.start}
                style={{ width: "100%", background: "linear-gradient(135deg, #111 0%, #374151 100%)", color: "#fff", border: "none", borderRadius: 0, padding: "15px 0", fontSize: 16, fontWeight: 700, cursor: "pointer", boxShadow: "0 4px 14px rgba(0,0,0,0.2)" }}>
                📍 ตรวจสอบตำแหน่ง
              </button>
            </>
          )}
          {geo.status === "loading" && (
            <>
              <div style={{ fontSize: 48, marginBottom: 16, animation: "pulse 1.5s infinite" }}>🛰️</div>
              <h2 style={{ margin: "0 0 8px", fontWeight: 900, fontSize: 20 }}>กำลังหาตำแหน่ง...</h2>
              <p style={{ color: "#6b7280", fontSize: 14, margin: 0 }}>รอสักครู่ระบบกำลังตรวจสอบ GPS</p>
              <style>{`@keyframes pulse { 0%,100% { transform: scale(1); } 50% { transform: scale(1.15); } }`}</style>
            </>
          )}
          {geo.status === "outside" && (
            <>
              <div style={{ fontSize: 56, marginBottom: 12 }}>🚫</div>
              <h2 style={{ margin: "0 0 8px", fontWeight: 900, fontSize: 20, color: "#dc2626" }}>อยู่นอกพื้นที่โรงงาน</h2>
              <div style={{ background: "#fef2f2", border: "1.5px solid #fecaca", borderRadius: 0, padding: "16px 20px", marginBottom: 20 }}>
                <div style={{ fontSize: 32, fontWeight: 900, color: "#dc2626", marginBottom: 4 }}>
                  {geo.distance >= 1000 ? `${(geo.distance / 1000).toFixed(1)} กม.` : `${geo.distance} เมตร`}
                </div>
                <div style={{ fontSize: 13, color: "#991b1b", fontWeight: 600 }}>
                  ระยะห่างจากโรงงาน
                </div>
              </div>
              <p style={{ color: "#6b7280", fontSize: 13, margin: "0 0 16px", lineHeight: 1.6 }}>
                กรุณาเดินทางเข้าใกล้โรงงานแล้วลองใหม่<br />ระบบจะตรวจสอบตำแหน่งอัตโนมัติ
              </p>
              <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 0, padding: 12, fontSize: 12, color: "#166534" }}>
                💡 ระบบกำลังติดตามตำแหน่งอยู่ — เมื่อเข้าใกล้โรงงานจะเปิดอัตโนมัติ
              </div>
            </>
          )}
          {geo.status === "error" && (
            <>
              <div style={{ fontSize: 56, marginBottom: 12 }}>⚠️</div>
              <h2 style={{ margin: "0 0 8px", fontWeight: 900, fontSize: 20, color: "#d97706" }}>ไม่สามารถตรวจสอบตำแหน่งได้</h2>
              <div style={{ background: "#fffbeb", border: "1.5px solid #fde68a", borderRadius: 0, padding: "14px 18px", marginBottom: 20, fontSize: 14, color: "#92400e", fontWeight: 600 }}>
                {geo.error}
              </div>
              <button onClick={geo.start}
                style={{ width: "100%", background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "14px 0", fontSize: 15, fontWeight: 700, cursor: "pointer" }}>
                🔄 ลองใหม่
              </button>
            </>
          )}
        </div>
      </div>
    );
  }

  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const matchPlate = (a, b) => plateNum(a) === plateNum(b) && plateNum(a) !== "";

  const handleSearch = () => {
    const p = plate.trim();
    if (!p) return;
    const queueEntries = queue.filter(q => matchPlate(q.plate, p));
    const usedIds = new Set(trucks.map(t => t.queueId).filter(Boolean));
    const nextEntry = queueEntries.find(q => !usedIds.has(q.id));
    if (queueEntries.length > 0 && !nextEntry) {
      setMsg({ t: "warn", text: "⚠️ รถคันนี้เช็คอินครบทุก trip แล้ว" }); return;
    }
    const entry = nextEntry || { id: `WALK-${Date.now()}`, plate: p, driver: "", customerGroup: "", zone: "", product: "", destination: "", qty: 0, unit: "กก.", time: TIME_NOW(), entryTime: TIME_NOW(), loadTime: "", exitTime: "" };
    setPendingEntry(entry);
    setSelectedZone(entry.zone || "");
    setStep("confirm");
    setMsg(null);
  };

  const handleConfirm = () => {
    onScan({ ...pendingEntry, zone: selectedZone, status: "arrived", arrivedAt: TIME_NOW(), queueId: pendingEntry.id, pickupPrinted: false, summaryPrinted: false });
    const isWalkIn = pendingEntry.id.startsWith("WALK-");
    setMsg({ t: isWalkIn ? "walk" : "ok", text: `✅ เช็คอินสำเร็จ! ${pendingEntry.plate}${selectedZone ? ` — ${selectedZone}` : ""}` });
    setPlate(""); setPendingEntry(null); setSelectedZone(""); setStep("input");
  };

  if (step === "confirm" && pendingEntry) return (
    <div>
      <h2 style={{ margin: "0 0 6px", fontWeight: 900, fontSize: 22 }}>🚛 ยืนยันการเช็คอิน</h2>
      <p style={{ margin: "0 0 18px", color: "#6b7280", fontSize: 13 }}>ตรวจสอบข้อมูลแล้วกดยืนยัน</p>
      <div style={{ background: "#fff", borderRadius: 0, padding: 24, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", marginBottom: 16 }}>
        <div style={{ background: "#f9fafb", borderRadius: 0, padding: "16px 20px", marginBottom: 20 }}>
          <div style={{ fontSize: 28, fontWeight: 900, letterSpacing: 2, marginBottom: 4 }}>{pendingEntry.plate}</div>
          {pendingEntry.customerGroup && <div style={{ fontSize: 13, color: "#6b7280", fontWeight: 600 }}>กลุ่มลูกค้า: {pendingEntry.customerGroup}</div>}
        </div>
        {selectedZone && (
          <div style={{ marginBottom: 16 }}>
            <label style={{ display: "block", fontWeight: 700, fontSize: 14, marginBottom: 8 }}>📍 Zone</label>
            <div style={{ width: "100%", border: "2px solid #e5e7eb", borderRadius: 0, padding: "12px 14px", fontSize: 16, fontWeight: 700, boxSizing: "border-box", background: "#f9fafb", color: "#7c3aed" }}>
              {selectedZone}
            </div>
          </div>
        )}
        <button onClick={handleConfirm}
          style={{ width: "100%", background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "14px 0", fontSize: 16, fontWeight: 700, cursor: "pointer", marginBottom: 10 }}>
          ✅ ยืนยันเช็คอิน
        </button>
        <button onClick={() => { setStep("input"); setMsg(null); }}
          style={{ width: "100%", background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "12px 0", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>
          ← กลับ
        </button>
      </div>
    </div>
  );

  return (
    <div>
      <h2 style={{ margin: "0 0 18px", fontWeight: 900, fontSize: 22 }}>🚛 คนขับ → เช็คอินเข้าโรงงาน</h2>
      <div style={{ background: "#fff", borderRadius: 0, padding: 24, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", marginBottom: 16 }}>
        <div style={{ textAlign: "center", marginBottom: 20 }}>
          <div style={{ width: 80, height: 80, background: "#111", borderRadius: 0, margin: "0 auto 10px", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff" }}>
            <Icon name="check" size={40} />
          </div>
          <p style={{ fontWeight: 700, fontSize: 15, margin: 0 }}>เช็คอินเข้าโรงงาน</p>
        </div>
        <input value={plate} onChange={e => setPlate(e.target.value.replace(/\D/g, ''))} onKeyDown={e => e.key === "Enter" && handleSearch()}
          placeholder="กรอกเลขทะเบียนของท่าน เช่น 1234"
          type="tel" inputMode="numeric" pattern="[0-9]*"
          style={{ width: "100%", border: "2px solid #e5e7eb", borderRadius: 0, padding: "14px 16px", fontSize: 18, fontWeight: 800, textAlign: "center", outline: "none", boxSizing: "border-box" }} />
        <button onClick={handleSearch} style={{ marginTop: 10, width: "100%", background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "14px 0", fontSize: 16, fontWeight: 700, cursor: "pointer" }}>
          ค้นหา
        </button>
        {msg && (
          <div style={{ marginTop: 12, padding: "12px 16px", borderRadius: 0, fontWeight: 600, fontSize: 14,
            background: msg.t === "ok" ? "#d1fae5" : msg.t === "walk" ? "#eff6ff" : msg.t === "warn" ? "#fef3c7" : "#fee2e2",
            color:      msg.t === "ok" ? "#065f46" : msg.t === "walk" ? "#1d4ed8" : msg.t === "warn" ? "#92400e" : "#991b1b" }}>
            {msg.text}
          </div>
        )}
      </div>

    </div>
  );
};

// ── 3+6. PICKING ──────────────────────────────────────────────────────────────
const Picking = ({ trucks, queue, onUpdate, detailMapByChannel = {} }) => {
  const isMobile = useIsMobile();

  // รวม queue + walk-in (รถที่เข้าแล้วแต่ยังไม่มีในคิว)
  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const usedPick = new Set();
  const matchTruckPick = q => {
    let t = trucks.find(t => t.queueId === q.id && !usedPick.has(t.id));
    if (!t) t = trucks.find(t => !t.queueId && plateNum(t.plate) === plateNum(q.plate) && plateNum(q.plate) !== "" && !usedPick.has(t.id));
    if (t) usedPick.add(t.id);
    return t;
  };
  const pickQueueRows = queue.map(q => ({ key: q.id, plate: q.plate, customerGroup: q.customerGroup, entryTime: q.entryTime, truck: matchTruckPick(q) }));
  const walkIns = trucks.filter(t => !usedPick.has(t.id));
  const allRows = [
    ...pickQueueRows,
    ...walkIns.map(t => ({ key: t.id, plate: t.plate, customerGroup: t.customerGroup || "–", entryTime: "", truck: t })),
  ].sort((a, b) => {
    const rank = t => {
      if (!t) return 1;                          // รอเช็คอิน
      if (t.summaryPrinted) return 3;            // เสร็จแล้ว → ล่าง
      const can3 = t.status === "arrived";
      const can6 = t.status === "picking" &&
        lanes.some(l => t.loadLanes?.[l.id]?.done) &&
        !lanes.some(l => t.qcLanes?.[l.id]?.done && !t.loadLanes?.[l.id]?.done);
      if (can3 || can6) return 0;                // กดได้เลย → บน
      return 2;                                  // รอขั้นตอนอื่น
    };
    return rank(a.truck) - rank(b.truck);
  }).filter(row => !settings.excludedCustomerGroups.includes(row.customerGroup));

  const [searchQuery, setSearchQuery] = useState("");
  const filteredRows = allRows.filter(r => r.plate.includes(searchQuery));

  const canStep3 = t => t?.status === "arrived";
  const doneStep3 = t => t && ["picking","summary_printed","invoiced"].includes(t.status);
  const canStep6 = t =>
    t?.status === "picking" &&
    lanes.some(l => t.loadLanes?.[l.id]?.done) &&
    !lanes.some(l => t.qcLanes?.[l.id]?.done && !t.loadLanes?.[l.id]?.done);
  const doneStep6 = t => t && ["summary_printed","invoiced"].includes(t.status);

  const ExtraStatusCell = ({ truck }) => {
    const [isEditing, setIsEditing] = useState(false);
    const [val, setVal] = useState("");
    const [saving, setSaving] = useState(false);

    if (!truck) return <span style={{ color: "#d1d5db", fontSize: 12 }}>—</span>;

    if (truck.extraStatus) {
      return (
        <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "#fee2e2", color: "#991b1b", padding: "4px 8px", borderRadius: 0, fontSize: 11, fontWeight: 700 }}>
          <span>⚠️ {truck.extraStatus}</span>
          <button onClick={async () => {
            setSaving(true);
            try { await onUpdate(truck.id, { extraStatus: "" }); }
            catch (e) { alert("บันทึกไม่สำเร็จ: " + e.message); }
            finally { setSaving(false); }
          }} disabled={saving} style={{ background: "transparent", border: "none", color: "#991b1b", cursor: "pointer", padding: 0, fontWeight: 900, fontSize: 12 }}>×</button>
        </div>
      );
    }
    if (isEditing) {
      return (
        <div style={{ display: "flex", gap: 4 }}>
          <input list={`extraStatusOptions-${truck.id}`} autoFocus value={val} onChange={e => setVal(e.target.value)} placeholder="พิมพ์หรือเลือก..." style={{ border: "1px solid #d1d5db", borderRadius: 0, padding: "2px 6px", fontSize: 11, width: 100 }} />
          <datalist id={`extraStatusOptions-${truck.id}`}>
            <option value="รอแปรสินค้า" />
            <option value="ติดปัญหา IT" />
          </datalist>
          <button onClick={async () => {
            if (!val) { setIsEditing(false); return; }
            setSaving(true);
            try {
              await onUpdate(truck.id, { extraStatus: val, extraStatusAt: TIME_NOW() });
              setIsEditing(false);
            } catch (e) {
              alert("บันทึกไม่สำเร็จ: " + e.message);
              setSaving(false);
            }
          }} disabled={saving} style={{ background: "#10b981", color: "#fff", border: "none", borderRadius: 0, padding: "2px 6px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{saving ? "..." : "บันทึก"}</button>
          <button onClick={() => setIsEditing(false)} disabled={saving} style={{ background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "2px 6px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>ยกเลิก</button>
        </div>
      );
    }
    return (
      <button onClick={() => { setIsEditing(true); setVal(""); }} style={{ background: "#f3f4f6", color: "#4b5563", border: "1px dashed #9ca3af", borderRadius: 0, padding: "4px 8px", fontSize: 10, cursor: "pointer", whiteSpace: "nowrap" }}>
        + เพิ่มสถานะ
      </button>
    );
  };

  const StatusCell = ({ truck, compact }) => {
    if (!truck) return <span style={{ fontSize: 11, color: "#9ca3af", fontWeight: 600 }}>รอเช็คอิน</span>;
    const anyQC = lanes.some(l => truck.qcLanes?.[l.id]?.done);
    if (!anyQC) return <span style={{ fontSize: 11, color: "#6b7280", fontWeight: 600 }}>รอเข้าโหลด</span>;
    const fSize = compact ? 10 : 11;
    return (
      <div style={{ display: "flex", flexWrap: "wrap", gap: compact ? 3 : 5, alignItems: "center" }}>
        {lanes.map(l => {
          const loaded = truck.loadLanes?.[l.id]?.done;
          const qcDone = truck.qcLanes?.[l.id]?.done;
          const waiting = truck.loadLanes?.[l.id]?.waiting && !loaded;
          if (loaded) return (
            <div key={l.id} style={{ position: "relative", display: "inline-block", background: "#10b981", color: "#fff", borderRadius: 0, padding: compact ? "2px 7px 4px 6px" : "3px 10px 5px 8px", fontSize: fSize, fontWeight: 700, lineHeight: 1.4 }}>
              {l.tinyLabel}
              <span style={{ position: "absolute", bottom: -4, right: -4, background: "#059669", border: "2px solid #fff", borderRadius: "50%", width: 14, height: 14, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8, fontWeight: 900 }}>✓</span>
            </div>
          );
          if (waiting) return (
            <div key={l.id} style={{ position: "relative", display: "inline-block", background: "#fbbf24", color: "#fff", borderRadius: 0, padding: compact ? "2px 7px 4px 6px" : "3px 10px 5px 8px", fontSize: fSize, fontWeight: 700, lineHeight: 1.4, whiteSpace: "nowrap" }}>
              รอสินค้า {l.tinyLabel}
              <span style={{ position: "absolute", bottom: -4, right: -4, background: "#d97706", border: "2px solid #fff", borderRadius: "50%", width: 14, height: 14, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8 }}>⏳</span>
            </div>
          );
          if (qcDone) return (
            <span key={l.id} style={{ fontSize: fSize, color: "#f97316", fontWeight: 700, whiteSpace: "nowrap" }}>กำลังโหลด {l.tinyLabel}</span>
          );
          return null;
        })}
      </div>
    );
  };

  const Step3Cell = ({ truck }) => {
    const [saving, setSaving] = useState(false);
    if (!truck) return <span style={{ color: "#d1d5db", fontSize: 12 }}>—</span>;
    if (doneStep3(truck)) return <span style={{ color: "#10b981", fontWeight: 700, fontSize: 13 }}>✓</span>;
    if (canStep3(truck)) return (
      <button onClick={async () => {
        setSaving(true);
        try { await onUpdate(truck.id, { pickupPrinted: true, status: "picking", pickingAt: TIME_NOW() }); }
        catch (e) { alert("บันทึกไม่สำเร็จ: " + e.message); setSaving(false); }
      }} disabled={saving}
        style={{ background: "#c2410c", color: "#fff", border: "none", borderRadius: 0, padding: "5px 8px", fontWeight: 700, fontSize: 11, cursor: "pointer", whiteSpace: "nowrap" }}>
        {saving ? "⏳" : "🖨️ เบิก"}
      </button>
    );
    return <span style={{ color: "#d1d5db", fontSize: 12 }}>—</span>;
  };

  const Step6Cell = ({ truck }) => {
    const [saving, setSaving] = useState(false);
    if (!truck) return <span style={{ color: "#d1d5db", fontSize: 12 }}>—</span>;
    if (doneStep6(truck)) return <span style={{ color: "#10b981", fontWeight: 700, fontSize: 13 }}>✓</span>;
    if (canStep6(truck)) return (
      <button onClick={async () => {
        setSaving(true);
        try { await onUpdate(truck.id, { summaryPrinted: true, summaryPrintedAt: TIME_NOW(), status: "summary_printed" }); }
        catch (e) { alert("บันทึกไม่สำเร็จ: " + e.message); setSaving(false); }
      }} disabled={saving}
        style={{ background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 0, padding: "5px 8px", fontWeight: 700, fontSize: 11, cursor: "pointer", whiteSpace: "nowrap" }}>
        {saving ? "⏳" : "🖨️ สรุป"}
      </button>
    );
    return <span style={{ color: "#d1d5db", fontSize: 12 }}>—</span>;
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 18px", fontWeight: 900, fontSize: 22 }}>📦 ห้อง Picking</h2>

      <div style={{ background: "#fff", borderRadius: 0, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.07)" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>
            📋 คิวรถวันนี้ <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{filteredRows.length}</span>
          </div>
          <input type="text" placeholder="🔍 ค้นหาทะเบียนรถ..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)}
            style={{ border: "1px solid #d1d5db", borderRadius: 0, padding: "6px 12px", fontSize: 12, outline: "none", width: isMobile ? "100%" : 160 }} />
        </div>
        {filteredRows.length === 0
          ? <div style={{ padding: 36, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีคิวรถ</div>
          : (
          <div style={{ overflowX: "auto", overflowY: "auto", maxHeight: "calc(100vh - 190px)" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: isMobile ? 11 : 12 }}>
              <thead style={{ position: "sticky", top: 0, zIndex: 10 }}>
                <tr style={{ background: "#f9fafb" }}>
                  {(isMobile
                    ? ["ทะเบียน","ลาน","เวลา / สถานะ","Action"]
                    : ["ทะเบียน","กลุ่มลูกค้า", ...lanes.map(l => l.tinyLabel), "เวลาเข้าโรงงาน","สถานะ","สถานะเพิ่มเติม","③ พิมพ์ใบเบิกสินค้า","⑥ พิมพ์ใบสรุปจ่าย"]
                  ).map(h => (
                    <th key={h} style={{ padding: isMobile ? "7px 6px" : "9px 12px", textAlign: "left", fontWeight: 700, color: "#374151", whiteSpace: "nowrap", borderBottom: "1px solid #e5e7eb", background: "#f9fafb" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredRows.map(({ key, plate, customerGroup, entryTime, truck }) => {
                  const matchedLanes = laneMatchForTruck({ plate, customerGroup }, detailMapByChannel);
                  return isMobile ? (
                    <tr key={key} style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td style={{ padding: "8px 6px", fontWeight: 800 }}>
                        {plate}
                        <div style={{ fontWeight: 400, color: "#6b7280", fontSize: 10 }}>{customerGroup}</div>
                      </td>
                      <td style={{ padding: "8px 6px" }}>
                        <div style={{ display: "flex", gap: 6 }}>
                          {lanes.map(l => (
                            <span key={l.id} style={{ fontSize: 10, fontWeight: 800, color: matchedLanes.has(l.id) ? l.color : "#d1d5db" }}>
                              {l.tinyLabel}{matchedLanes.has(l.id) ? "✓" : ""}
                            </span>
                          ))}
                        </div>
                      </td>
                      <td style={{ padding: "8px 6px" }}>
                        <div style={{ fontWeight: 700, color: "#3b82f6" }}>{entryTime || "—"}</div>
                        <div style={{ marginTop: 4 }}><StatusCell truck={truck} compact /></div>
                        <div style={{ marginTop: 4 }}><ExtraStatusCell truck={truck} /></div>
                      </td>
                      <td style={{ padding: "8px 6px", whiteSpace: "nowrap" }}>
                        <div style={{ display: "flex", flexDirection: "column", gap: 4, alignItems: "flex-start" }}>
                          <Step3Cell truck={truck} />
                          <Step6Cell truck={truck} />
                        </div>
                      </td>
                    </tr>
                  ) : (
                    <tr key={key} style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td style={{ padding: "10px 12px", fontWeight: 800 }}>{plate}</td>
                      <td style={{ padding: "10px 12px", color: "#374151" }}>{customerGroup}</td>
                      {lanes.map(l => (
                        <td key={l.id} style={{ padding: "10px 12px", textAlign: "center" }}>
                          {matchedLanes.has(l.id)
                            ? <span style={{ color: l.color, fontWeight: 900, fontSize: 16 }}>✓</span>
                            : <span style={{ color: "#e5e7eb" }}>—</span>}
                        </td>
                      ))}
                      <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                        <div style={{ fontWeight: 700, color: "#3b82f6" }}>{entryTime || "—"}</div>
                        {truck?.arrivedAt
                          ? <div style={{ fontSize: 10, color: "#6b7280", marginTop: 2 }}>เข้าจริง {truck.arrivedAt}</div>
                          : <div style={{ fontSize: 10, color: "#6b7280", marginTop: 2 }}>(รถยังไม่เข้าโรงงาน)</div>}
                      </td>
                      <td style={{ padding: "10px 12px" }}><StatusCell truck={truck} /></td>
                      <td style={{ padding: "10px 12px" }}><ExtraStatusCell truck={truck} /></td>
                      <td style={{ padding: "10px 12px" }}><Step3Cell truck={truck} /></td>
                      <td style={{ padding: "10px 12px" }}><Step6Cell truck={truck} /></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

    </div>
  );
};

// ── 3.5 BAY SELECT (เลือกช่องโหลดก่อนเข้าฟอร์ม QC/QC สุ่ม/Checker ทุกครั้ง — ไม่จำค่าไว้) ──
// เปิด/ปิดได้ทั้งระบบที่ settings.enableBaySelection (Master Setting → 🚪 ช่องโหลด)
const BaySelect = ({ laneId, actLane, title, onSelect, onBack }) => {
  const laneBays = bays.filter(b => b.laneId === laneId).sort((a, b) => a.sortOrder - b.sortOrder);
  return (
    <div style={{ minHeight: "70vh", display: "flex", alignItems: "center", justifyContent: "center", padding: "16px 20px", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
      <div style={{ maxWidth: 420, width: "100%" }}>
        <h2 style={{ textAlign: "center", fontWeight: 900, fontSize: 20, margin: "0 0 4px" }}>{title}</h2>
        <p style={{ textAlign: "center", color: "#6b7280", fontSize: 12, margin: "0 0 28px" }}>เลือกช่องโหลด</p>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {laneBays.map(b => (
            <button key={b.id} onClick={() => onSelect(b.id)}
              style={{ background: "#fff", border: "1.5px solid #e5e7eb", borderRadius: 0, padding: "16px 14px", fontSize: 15, fontWeight: 700, cursor: "pointer", textAlign: "center", color: actLane.color }}>
              {b.label}
            </button>
          ))}
        </div>
        {onBack && (
          <button onClick={onBack} style={{ marginTop: 24, width: "100%", background: "transparent", border: "none", color: "#6b7280", fontSize: 13, fontWeight: 600, cursor: "pointer", textAlign: "center" }}>
            ← กลับ
          </button>
        )}
      </div>
    </div>
  );
};

// ── 4. QC (per-lane) ──────────────────────────────────────────────────────────
const QC = ({ trucks, onUpdate, laneId, detailMapByChannel = {}, onBack }) => {
  const [selId,     setSelId]     = useState("");
  const lane = laneId;
  const [temp,      setTemp]      = useState("");
  const [photo,     setPhoto]     = useState(null);
  const [flashLane, setFlashLane] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [bayId,     setBayId]     = useState(null);

  // รับทุกรถที่สถานะ "picking" (พิมพ์เบิกแล้ว)
  const eligible = trucks.filter(t => ["arrived", "picking"].includes(t.status) && !settings.excludedCustomerGroups.includes(t.customerGroup));
  const sel      = trucks.find(t => t.id === selId) || null;
  const actLane  = lanes.find(l => l.id === lane);
  const thisLaneQCd = sel?.qcLanes?.[lane]?.done;
  const bay      = settings.enableBaySelection ? bays.find(b => b.id === bayId) : null;

  if (settings.enableBaySelection && !bayId) {
    return <BaySelect laneId={lane} actLane={actLane} title={`ลานโหลด → ${actLane.label}`} onSelect={setBayId} onBack={onBack} />;
  }

  const handlePhoto = e => {
    const files = Array.from(e.target.files).slice(0, settings.maxPhotoUploads); if (!files.length) return;
    Promise.all(files.map(compressImage)).then(newPhotos => {
      setPhoto(prev => {
        const p = Array.isArray(prev) ? prev : (prev ? [prev] : []);
        return [...p, ...newPhotos].slice(0, settings.maxPhotoUploads);
      });
    }).catch(err => alert(err.message));
  };

  const handleSubmit = async () => {
    if (!sel || !temp || uploading) return;
    setUploading(true);
    try {
      const photoUrls = await uploadPhotos(`qc`, sel.plate, Array.isArray(photo) ? photo : (photo ? [photo] : []));
      const qcLanes = { ...(sel.qcLanes || {}), [lane]: { done: true, temp, photos: photoUrls, doneAt: TIME_NOW(), bayId: bay?.id || null } };
      await onUpdate(sel.id, { qcLanes });
      setFlashLane(lane); setTemp(""); setPhoto(null);
      setTimeout(() => setFlashLane(null), 2500);
    } catch (e) {
      alert("อัพโหลดรูปไม่สำเร็จ: " + e.message);
    } finally {
      setUploading(false);
    }
  };

  const hasAnyQC = t => t.qcLanes && Object.values(t.qcLanes).some(l => l.done);

  return (
    <div>
      <h2 style={{ margin: "0 0 18px", fontWeight: 900, fontSize: 22, display: "flex", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <span>ลานโหลด → {actLane.label}{bay ? ` · ${bay.label}` : ""}</span>
        {bay && (
          <button onClick={() => setBayId(null)}
            style={{ background: "none", border: "none", color: "#2563eb", fontSize: 13, fontWeight: 700, cursor: "pointer", textDecoration: "underline" }}>
            เปลี่ยนช่องโหลด
          </button>
        )}
      </h2>

      {flashLane && (
        <div style={{ padding: "13px 16px", background: "#d1fae5", borderRadius: 0, color: "#065f46", fontWeight: 700, marginBottom: 14, display: "flex", gap: 8, alignItems: "center" }}>
          <Icon name="check" size={18} /> QC ผ่าน → พร้อมเข้า {lanes.find(l => l.id === flashLane)?.label}
        </div>
      )}
      {/* เลือกรถ */}
      <div style={{ background: "#fff", borderRadius: 0, padding: 18, boxShadow: "0 2px 10px rgba(0,0,0,0.07)", marginBottom: 14 }}>
        <label style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#374151", marginBottom: 8 }}>เลือกทะเบียนรถ</label>
        <select value={selId} onChange={e => {
            const id = e.target.value;
            setSelId(id); setTemp(""); setPhoto(null);
            const t = trucks.find(tt => tt.id === id);
            if (t) {
              const lanes = laneMatchForTruck(t, detailMapByChannel);
              if (lanes.size > 0 && !lanes.has(lane)) {
                alert(`ทะเบียนนี้ไม่มีโหลดสินค้า ${actLane.label}`);
              }
            }
          }}
          style={{ width: "100%", border: "1.5px solid #e5e7eb", borderRadius: 0, padding: "11px 12px", fontSize: 15, outline: "none", boxSizing: "border-box" }}>
          <option value="">-- เลือกทะเบียนรถที่รอเข้าโหลด --</option>
          {eligible.map(t => <option key={t.id} value={t.id}>{t.loadLanes?.[lane]?.waiting ? "⏳ " : ""}{t.plate} · {t.customerGroup || t.product}</option>)}
        </select>
        {sel && (
          <div style={{ marginTop: 10 }}>
            <div style={{ background: "#f9fafb", borderRadius: 0, padding: "8px 12px", fontSize: 13, marginBottom: 8 }}>
              <b>{sel.product}</b>{sel.destination ? ` → ${sel.destination}` : ""}
            </div>
            {/* สรุป QC รายลาน */}
            <div style={{ display: "flex", gap: 6 }}>
              {lanes.map(l => {
                const qc = sel.qcLanes?.[l.id];
                return (
                  <div key={l.id} style={{ flex: 1, background: qc?.done ? l.bg : "#f3f4f6", border: `1px solid ${qc?.done ? l.border : "#e5e7eb"}`, borderRadius: 0, padding: "6px 4px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: qc?.done ? l.color : "#9ca3af", lineHeight: 1.4 }}>{l.shortLabel}</div>
                    <div style={{ fontSize: 10, fontWeight: 800, color: qc?.done ? l.color : "#9ca3af" }}>
                      {qc?.done ? `${qc.temp}°C ✓` : "รอ QC"}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* ฟอร์มลาน */}
      <div style={{ background: actLane.bg, border: `2px solid ${actLane.border}`, borderRadius: 0, padding: 18, marginBottom: 12 }}>
        {thisLaneQCd && (
          <div style={{ padding: "9px 12px", background: "#d1fae5", borderRadius: 0, color: "#065f46", fontWeight: 700, marginBottom: 12, fontSize: 13 }}>
            ✅ {actLane.label} QC แล้ว: {sel.qcLanes[lane].temp}°C — สามารถวัดซ้ำได้
          </div>
        )}
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 800, fontSize: 15, color: actLane.color }}>{actLane.label}</div>
          <div style={{ fontSize: 12, color: "#6b7280" }}>วัดอุณหภูมิก่อนรถเข้าลานนี้</div>
        </div>
        <label style={{ display: "block", fontSize: 13, fontWeight: 600, color: "#374151", marginBottom: 5 }}>อุณหภูมิ (°C)</label>
        <input value={temp} onChange={e => setTemp(e.target.value)} type="number" placeholder="-4"
          style={{ width: "100%", border: `2px solid ${actLane.border}`, borderRadius: 0, padding: "12px 14px", fontSize: 26, fontWeight: 900, outline: "none", boxSizing: "border-box", color: actLane.color, background: "#fff", textAlign: "center", marginBottom: 12 }} />
        <PhotoUploader label="📷 ถ่ายรูปอุณหภูมิ" value={photo} onChange={handlePhoto} onRemove={setPhoto} />
      </div>

      <button onClick={handleSubmit} disabled={!sel || !temp || uploading}
        style={{ width: "100%", background: sel && temp && !uploading ? actLane.color : "#e5e7eb", color: sel && temp && !uploading ? "#fff" : "#9ca3af", border: "none", borderRadius: 0, padding: "14px 0", fontWeight: 700, fontSize: 15, cursor: sel && temp && !uploading ? "pointer" : "default", marginBottom: 20 }}>
        {uploading ? "⏳ กำลังอัพโหลดรูป..." : !sel ? "เลือกทะเบียนรถก่อน" : !temp ? "กรอกอุณหภูมิก่อน" : `✅ บันทึก QC → ${actLane.label}`}
      </button>

    </div>
  );
};

// ── 4b. RANDOM SAMPLE CHECK (per-lane, photo only, no temp) ───────────────────
const RandomSampleCheck = ({ trucks, onUpdate, laneId, detailMapByChannel = {}, onBack }) => {
  const [selId,     setSelId]     = useState("");
  const lane = laneId;
  const [photo,     setPhoto]     = useState(null);
  const [note,      setNote]      = useState("");
  const [flashLane, setFlashLane] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [bayId,     setBayId]     = useState(null);

  const eligible = trucks.filter(t => ["arrived", "picking"].includes(t.status) && !settings.excludedCustomerGroups.includes(t.customerGroup));
  const sel      = trucks.find(t => t.id === selId) || null;
  const actLane  = lanes.find(l => l.id === lane);
  const photos   = Array.isArray(photo) ? photo : (photo ? [photo] : []);
  const thisLaneChecked = sel?.sampleLanes?.[lane]?.done;
  const bay      = settings.enableBaySelection ? bays.find(b => b.id === bayId) : null;

  if (settings.enableBaySelection && !bayId) {
    return <BaySelect laneId={lane} actLane={actLane} title={`ตรวจอุณหภูมิ → ${actLane.label}`} onSelect={setBayId} onBack={onBack} />;
  }

  const handlePhoto = e => {
    const files = Array.from(e.target.files).slice(0, settings.maxPhotoUploads); if (!files.length) return;
    Promise.all(files.map(compressImage)).then(newPhotos => {
      setPhoto(prev => {
        const p = Array.isArray(prev) ? prev : (prev ? [prev] : []);
        return [...p, ...newPhotos].slice(0, settings.maxPhotoUploads);
      });
    }).catch(err => alert(err.message));
  };

  const handleSubmit = async () => {
    if (!sel || !photos.length || uploading) return;
    setUploading(true);
    try {
      const photoUrls = await uploadPhotos(`sample`, sel.plate, photos);
      const sampleLanes = { ...(sel.sampleLanes || {}), [lane]: { done: true, photos: photoUrls, note, doneAt: TIME_NOW(), bayId: bay?.id || null } };
      await onUpdate(sel.id, { sampleLanes });
      setFlashLane(lane); setPhoto(null); setNote("");
      setTimeout(() => setFlashLane(null), 2500);
    } catch (e) {
      alert("อัพโหลดรูปไม่สำเร็จ: " + e.message);
    } finally {
      setUploading(false);
    }
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 18px", fontWeight: 900, fontSize: 22, display: "flex", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <span>ตรวจอุณหภูมิ → {actLane.label}{bay ? ` · ${bay.label}` : ""}</span>
        {bay && (
          <button onClick={() => setBayId(null)}
            style={{ background: "none", border: "none", color: "#2563eb", fontSize: 13, fontWeight: 700, cursor: "pointer", textDecoration: "underline" }}>
            เปลี่ยนช่องโหลด
          </button>
        )}
      </h2>

      {flashLane && (
        <div style={{ padding: "13px 16px", background: "#d1fae5", borderRadius: 0, color: "#065f46", fontWeight: 700, marginBottom: 14, display: "flex", gap: 8, alignItems: "center" }}>
          <Icon name="check" size={18} /> บันทึกตรวจแล้ว → {actLane.label}
        </div>
      )}
      {/* เลือกรถ */}
      <div style={{ background: "#fff", borderRadius: 0, padding: 18, boxShadow: "0 2px 10px rgba(0,0,0,0.07)", marginBottom: 14 }}>
        <label style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#374151", marginBottom: 8 }}>เลือกทะเบียนรถ</label>
        <select value={selId} onChange={e => {
            const id = e.target.value;
            setSelId(id); setPhoto(null); setNote("");
            const t = trucks.find(tt => tt.id === id);
            if (t) {
              const lanes = laneMatchForTruck(t, detailMapByChannel);
              if (lanes.size > 0 && !lanes.has(lane)) {
                alert(`ทะเบียนนี้ไม่มีโหลดสินค้า ${actLane.label}`);
              }
            }
          }}
          style={{ width: "100%", border: "1.5px solid #e5e7eb", borderRadius: 0, padding: "11px 12px", fontSize: 15, outline: "none", boxSizing: "border-box" }}>
          <option value="">-- เลือกทะเบียนรถ --</option>
          {eligible.map(t => <option key={t.id} value={t.id}>{t.plate} · {t.customerGroup || t.product}</option>)}
        </select>
        {sel && (
          <div style={{ marginTop: 10 }}>
            <div style={{ background: "#f9fafb", borderRadius: 0, padding: "8px 12px", fontSize: 13, marginBottom: 8 }}>
              <b>{sel.product}</b>{sel.destination ? ` → ${sel.destination}` : ""}
            </div>
            {/* สรุปตรวจรายลาน */}
            <div style={{ display: "flex", gap: 6 }}>
              {lanes.map(l => {
                const sc = sel.sampleLanes?.[l.id];
                return (
                  <div key={l.id} style={{ flex: 1, background: sc?.done ? l.bg : "#f3f4f6", border: `1px solid ${sc?.done ? l.border : "#e5e7eb"}`, borderRadius: 0, padding: "6px 4px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: sc?.done ? l.color : "#9ca3af", lineHeight: 1.4 }}>{l.shortLabel}</div>
                    <div style={{ fontSize: 10, fontWeight: 800, color: sc?.done ? l.color : "#9ca3af" }}>
                      {sc?.done ? "📷 ตรวจแล้ว" : "รอตรวจ"}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* ฟอร์มลาน */}
      <div style={{ background: actLane.bg, border: `2px solid ${actLane.border}`, borderRadius: 0, padding: 18, marginBottom: 12 }}>
        {thisLaneChecked && (
          <div style={{ padding: "9px 12px", background: "#d1fae5", borderRadius: 0, color: "#065f46", fontWeight: 700, marginBottom: 12, fontSize: 13 }}>
            ✅ {actLane.label} ตรวจแล้ว — สามารถตรวจซ้ำได้
          </div>
        )}
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 800, fontSize: 15, color: actLane.color }}>{actLane.label}</div>
          <div style={{ fontSize: 12, color: "#6b7280" }}>ตรวจอุณหภูมิสินค้าก่อนโหลดขึ้นรถ — ถ่ายรูปยืนยัน</div>
        </div>
        <textarea
          placeholder="Note (ถ้ามี)"
          value={note}
          onChange={e => setNote(e.target.value)}
          rows={2}
          style={{ width: "100%", border: `1.5px solid ${actLane.border}`, borderRadius: 0, padding: "10px 12px", fontSize: 13, outline: "none", boxSizing: "border-box", marginBottom: 12, resize: "vertical", fontFamily: "inherit" }}
        />
        <PhotoUploader label="📷 ถ่ายรูปอุณหภูมิ" value={photo} onChange={handlePhoto} onRemove={setPhoto} />
      </div>

      <button onClick={handleSubmit} disabled={!sel || !photos.length || uploading}
        style={{ width: "100%", background: sel && photos.length && !uploading ? actLane.color : "#e5e7eb", color: sel && photos.length && !uploading ? "#fff" : "#9ca3af", border: "none", borderRadius: 0, padding: "14px 0", fontWeight: 700, fontSize: 15, cursor: sel && photos.length && !uploading ? "pointer" : "default", marginBottom: 20 }}>
        {uploading ? "⏳ กำลังอัพโหลดรูป..." : !sel ? "เลือกทะเบียนรถก่อน" : !photos.length ? "แนบรูปก่อน" : `✅ บันทึกตรวจ → ${actLane.label}`}
      </button>

    </div>
  );
};

// ── 5. LOADING YARD (per-lane gate) ───────────────────────────────────────────
const LoadingYard = ({ trucks, onUpdate, laneId, masterLane = [], onBack }) => {
  const [activeLane, setActiveLane] = useState(laneId ?? "lane_parts");
  const [bayId, setBayId] = useState(null);
  const emptyBaskets = () => ({ ...Object.fromEntries(basketTypes.map(b => [b.key, ""])), payer: "" });
  const [forms, setForms] = useState({
    lane_parts: { selId: "", photo: null, note: "", flash: false, uploading: false, baskets: emptyBaskets() },
    lane_head:  { selId: "", photo: null, note: "", flash: false, uploading: false, baskets: emptyBaskets() },
    lane_pork:  { selId: "", photo: null, note: "", flash: false, uploading: false, baskets: emptyBaskets() },
  });
  const setF = (lId, upd) => setForms(p => ({ ...p, [lId]: { ...p[lId], ...upd } }));
  const curLane = lanes.find(l => l.id === activeLane);
  const form    = forms[activeLane];
  const setBasket = (key, val) => setF(activeLane, { baskets: { ...form.baskets, [key]: val } });
  const basketTotal = basketTypes.filter(b => b.countsInTotal).reduce((sum, b) => sum + (Number(form.baskets?.[b.key]) || 0), 0);
  const [basketsOpen, setBasketsOpen] = useState(false);
  const [waitingModal, setWaitingModal] = useState(false);
  const [waitingReasons, setWaitingReasons] = useState([""]);
  const MAX_WAITING_REASONS = settings.maxWaitingReasons;
  const setReasonAt = (idx, val) => setWaitingReasons(rs => rs.map((r, i) => i === idx ? val : r));
  const addReasonField = () => setWaitingReasons(rs => rs.length >= MAX_WAITING_REASONS ? rs : [...rs, ""]);
  const removeReasonField = (idx) => setWaitingReasons(rs => rs.length <= 1 ? rs : rs.filter((_, i) => i !== idx));
  const bay = settings.enableBaySelection ? bays.find(b => b.id === bayId) : null;

  if (settings.enableBaySelection && !bayId) {
    return <BaySelect laneId={activeLane} actLane={curLane} title={`Checker ${curLane.label}`} onSelect={setBayId} onBack={onBack} />;
  }

  // ตัวเลือก dropdown ของ popup "รอสินค้าอะไร" — ชื่อสินค้าจาก Master ลานโหลด กรองเฉพาะลานที่เปิดอยู่
  const waitingOptions = (() => {
    const names = new Set();
    for (const m of masterLane) {
      if (m.laneKey === activeLane && m.productNameTha) names.add(m.productNameTha);
    }
    if (names.size) return [...names].sort((a, b) => a.localeCompare(b, "th"));
    if (waitingReasonPresets.length) return waitingReasonPresets.map(r => r.label);
    // default เดิมในโค้ด — ใช้ก่อนจะมีข้อมูลใน wh_waiting_reasons เลย
    return ["รอเบิกสินค้าจากคลัง", "รอแพ็ค/ชั่งน้ำหนักสินค้า", "รอสินค้าจากไลน์ผลิต", "รอรถขนย้ายภายใน"];
  })();

  // รถที่ QC ลานนี้ผ่านแล้ว และยังไม่ได้โหลดลานนี้
  const eligibleForLane = (laneId) => trucks.filter(t =>
    ["arrived", "picking"].includes(t.status) &&
    !settings.excludedCustomerGroups.includes(t.customerGroup) &&
    t.qcLanes?.[laneId]?.done &&
    !t.loadLanes?.[laneId]?.done
  );
  const eligible = eligibleForLane(activeLane);
  const sel = trucks.find(t => t.id === form.selId) || null;

  const handlePhoto = lId => e => {
    const files = Array.from(e.target.files).slice(0, settings.maxPhotoUploads); if (!files.length) return;
    Promise.all(files.map(compressImage)).then(newPhotos => {
      setForms(prev => {
        const f = prev[lId];
        const curPhotos = Array.isArray(f.photo) ? f.photo : (f.photo ? [f.photo] : []);
        return { ...prev, [lId]: { ...f, photo: [...curPhotos, ...newPhotos].slice(0, settings.maxPhotoUploads) } };
      });
    }).catch(err => alert(err.message));
  };

  const openWaitingModal = () => {
    if (!sel || form.uploading) return;
    setWaitingReasons([""]);
    setWaitingModal(true);
  };

  const confirmWaiting = async () => {
    const combinedReason = waitingReasons.map(r => r.trim()).filter(Boolean).join(", ");
    if (!sel || form.uploading || !combinedReason) return;
    setWaitingModal(false);
    setF(activeLane, { uploading: true });
    try {
      const loadLanes = { ...(sel.loadLanes || {}), [activeLane]: { ...(sel.loadLanes?.[activeLane] || {}), waiting: true, waitingAt: TIME_NOW(), waitingFor: combinedReason, note: form.note, bayId: bay?.id || null } };
      await onUpdate(sel.id, { loadLanes });
      setF(activeLane, { selId: "", photo: null, note: "", uploading: false, baskets: emptyBaskets() });
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
      setF(activeLane, { uploading: false });
    }
  };

  const handleLoad = async () => {
    if (!sel || form.uploading) return;
    if (!window.confirm(`ยืนยัน: บันทึกโหลดเสร็จ ${sel.plate}?`)) return;
    setF(activeLane, { uploading: true });
    try {
      const photos = Array.isArray(form.photo) ? form.photo : (form.photo ? [form.photo] : []);
      const photoUrls = await uploadPhotos(`loading/${activeLane}`, sel.plate, photos);
      const existing = sel.loadLanes?.[activeLane] || {};
      const baskets = Object.fromEntries(basketTypes.map(b => [b.key, Number(form.baskets?.[b.key]) || 0]));
      const basketPayer = (form.baskets?.payer || "").trim();
      const loadLanes = { ...(sel.loadLanes || {}), [activeLane]: { ...existing, done: true, photos: photoUrls, note: form.note, doneAt: TIME_NOW(), baskets, basketPayer, bayId: bay?.id || null } };
      await onUpdate(sel.id, { loadLanes });
      setF(activeLane, { selId: "", photo: null, note: "", flash: true, uploading: false, baskets: emptyBaskets() });
      setTimeout(() => setF(activeLane, { flash: false }), 2500);
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
      setF(activeLane, { uploading: false });
    }
  };

  const doneTrucks = trucks.filter(t => ["summary_printed","invoiced"].includes(t.status));

  const LaneSummary = ({ t }) => (
    <div style={{ display: "flex", gap: 6, margin: "8px 0" }}>
      {lanes.map(l => {
        const qc  = t.qcLanes?.[l.id];
        const ld  = t.loadLanes?.[l.id];
        const bg  = ld?.done ? l.bg : qc?.done ? "#fef9c3" : "#f9fafb";
        const bdr = ld?.done ? l.border : qc?.done ? "#fde047" : "#e5e7eb";
        return (
          <div key={l.id} style={{ flex: 1, background: bg, border: `1px solid ${bdr}`, borderRadius: 0, padding: "5px 4px", textAlign: "center" }}>
            <div style={{ fontSize: 9, fontWeight: 700, color: ld?.done ? l.color : qc?.done ? "#713f12" : "#9ca3af" }}>{l.tinyLabel}</div>
            <div style={{ fontSize: 9, fontWeight: 800, color: ld?.done ? l.color : qc?.done ? "#713f12" : "#9ca3af", lineHeight: 1.3 }}>
              {ld?.done ? `✓ ${ld.doneAt}` : qc?.done ? `QC ${qc.temp}°C` : "–"}
            </div>
            {ld?.photo && <img src={ld.photo} alt="" style={{ width: "100%", borderRadius: 0, marginTop: 2, height: 28, objectFit: "cover" }} />}
          </div>
        );
      })}
    </div>
  );

  return (
    <div>
      <h2 style={{ margin: "0 0 18px", fontWeight: 900, fontSize: 22, display: "flex", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <span>Checker {curLane.label}{bay ? ` · ${bay.label}` : ""}</span>
        {bay && (
          <button onClick={() => setBayId(null)}
            style={{ background: "none", border: "none", color: "#2563eb", fontSize: 13, fontWeight: 700, cursor: "pointer", textDecoration: "underline" }}>
            เปลี่ยนช่องโหลด
          </button>
        )}
      </h2>

      {/* ฟอร์มลาน */}
      <div style={{ background: curLane.bg, border: `2px solid ${curLane.border}`, borderRadius: 0, padding: 20, marginBottom: 16 }}>
        {form.flash && (
          <div style={{ padding: "11px 14px", background: "#d1fae5", borderRadius: 0, color: "#065f46", fontWeight: 700, marginBottom: 12, display: "flex", gap: 8, alignItems: "center" }}>
            <Icon name="check" size={16} /> บันทึกโหลดเสร็จ → Checker {curLane.label}
          </div>
        )}
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontWeight: 900, fontSize: 16, color: curLane.color }}>Checker {curLane.label}</div>
        </div>
        <label style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#374151", marginBottom: 6 }}>เลือกทะเบียนรถ</label>
        <select value={form.selId} onChange={e => setF(activeLane, { selId: e.target.value })}
          style={{ width: "100%", border: `1.5px solid ${curLane.border}`, borderRadius: 0, padding: "11px 12px", fontSize: 15, outline: "none", boxSizing: "border-box", marginBottom: 12, background: "#fff" }}>
          <option value="">-- เลือกทะเบียนรถ --</option>
          {eligible.map(t => <option key={t.id} value={t.id}>{t.loadLanes?.[activeLane]?.waiting ? "⏳ " : ""}{t.plate} · {t.customerGroup || t.product}</option>)}
        </select>
        {sel && (
          <div style={{ background: "#fff", borderRadius: 0, padding: "9px 12px", fontSize: 13, marginBottom: 12, border: `1px solid ${curLane.border}`, display: "flex", flexDirection: "column", gap: 4 }}>
            <span><b>กลุ่มลูกค้า:</b> {sel.customerGroup || sel.product}</span>
            {sel.destination && <span><b>ปลายทาง:</b> {sel.destination}</span>}
          </div>
        )}
        <textarea
          placeholder="Note (ถ้ามี)"
          value={form.note}
          onChange={e => setF(activeLane, { note: e.target.value })}
          rows={2}
          style={{ width: "100%", border: `1.5px solid ${curLane.border}`, borderRadius: 0, padding: "10px 12px", fontSize: 13, outline: "none", boxSizing: "border-box", marginBottom: 12, resize: "vertical", fontFamily: "inherit" }}
        />
        <PhotoUploader label="📷 ถ่ายรูปหลังโหลดเสร็จ" value={form.photo} onChange={handlePhoto(activeLane)} onRemove={photos => setF(activeLane, { photo: photos })} />

        <div style={{ marginTop: 4, marginBottom: 12 }}>
          <div onClick={() => setBasketsOpen(o => !o)}
            style={{ display: "flex", alignItems: "center", justifyContent: "space-between", fontWeight: 800, fontSize: 13, color: "#374151", marginBottom: basketsOpen ? 8 : 0, cursor: "pointer" }}>
            <span>🧺 ยอดตะกร้า / ตะขอ <span style={{ fontWeight: 400, color: "#9ca3af" }}>(Optional)</span></span>
            <span style={{ fontSize: 11, color: "#9ca3af" }}>{basketsOpen ? "▲" : "▼"}</span>
          </div>
          {basketsOpen && (
            <>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(90px, 1fr))", gap: 8, marginBottom: 8 }}>
                {basketTypes.map(b => (
                  <div key={b.key}>
                    <label style={{ display: "block", fontSize: 11, color: "#6b7280", marginBottom: 4 }}>{b.label}</label>
                    <input type="number" min="0" inputMode="numeric" value={form.baskets?.[b.key] ?? ""} onChange={e => setBasket(b.key, e.target.value)}
                      style={{ width: "100%", border: `1.5px solid ${curLane.border}`, borderRadius: 0, padding: "8px 10px", fontSize: 14, outline: "none", boxSizing: "border-box", background: "#fff" }} />
                  </div>
                ))}
                <div>
                  <label style={{ display: "block", fontSize: 11, color: "#6b7280", marginBottom: 4 }}>รวมตะกร้า</label>
                  <div style={{ padding: "8px 10px", fontSize: 14, fontWeight: 700, color: curLane.color, background: "#fff", border: `1.5px solid ${curLane.border}`, boxSizing: "border-box" }}>{basketTotal}</div>
                </div>
              </div>
              <input
                placeholder="ชื่อผู้จ่ายตะกร้า"
                value={form.baskets?.payer ?? ""}
                onChange={e => setBasket("payer", e.target.value)}
                style={{ width: "100%", border: `1.5px solid ${curLane.border}`, borderRadius: 0, padding: "9px 10px", fontSize: 13, outline: "none", boxSizing: "border-box" }}
              />
            </>
          )}
        </div>

        <button onClick={openWaitingModal} disabled={!sel || form.uploading}
          style={{ width: "100%", background: sel && !form.uploading ? "#f59e0b" : "#e5e7eb", color: sel && !form.uploading ? "#fff" : "#9ca3af", border: "none", borderRadius: 0, padding: "13px 0", fontWeight: 700, fontSize: 15, cursor: sel && !form.uploading ? "pointer" : "default", marginBottom: 8 }}>
          ⏳ รอเติมสินค้า
        </button>
        <button onClick={handleLoad} disabled={!sel || form.uploading}
          style={{ width: "100%", background: sel && !form.uploading ? curLane.color : "#e5e7eb", color: sel && !form.uploading ? "#fff" : "#9ca3af", border: "none", borderRadius: 0, padding: "13px 0", fontWeight: 700, fontSize: 15, cursor: sel && !form.uploading ? "pointer" : "default" }}>
          {form.uploading ? "⏳ กำลังอัพโหลดรูป..." : "✅ บันทึกโหลดเสร็จ"}
        </button>
      </div>

      {waitingModal && (
        <div
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}
          onClick={() => setWaitingModal(false)}
        >
          <div
            style={{ background: "#fff", borderRadius: 0, padding: 22, width: "min(420px, 100%)", boxShadow: "0 10px 30px rgba(0,0,0,0.25)" }}
            onClick={e => e.stopPropagation()}
          >
            <div style={{ fontWeight: 900, fontSize: 17, marginBottom: 2 }}>⏳ {sel?.plate} — รอสินค้าอะไร?</div>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 14 }}>พิมพ์เอง หรือเลือกจากตัวเลือกด้านล่าง — รอได้หลายอย่าง กด "+ เพิ่มรายการ"</div>
            {waitingReasons.map((reason, idx) => (
              <div key={idx} style={{ display: "flex", gap: 6, marginBottom: 10 }}>
                <input
                  list="waitingReasonOptions"
                  autoFocus={idx === 0}
                  value={reason}
                  onChange={e => setReasonAt(idx, e.target.value)}
                  placeholder="เช่น รอเบิกสินค้าจากคลัง..."
                  style={{ flex: 1, border: "1.5px solid #d1d5db", borderRadius: 0, padding: "11px 12px", fontSize: 15, outline: "none", boxSizing: "border-box" }}
                />
                {waitingReasons.length > 1 && (
                  <button onClick={() => removeReasonField(idx)}
                    style={{ background: "#f3f4f6", color: "#6b7280", border: "none", borderRadius: 0, padding: "0 14px", fontSize: 16, fontWeight: 700, cursor: "pointer" }}>
                    ×
                  </button>
                )}
              </div>
            ))}
            <datalist id="waitingReasonOptions">
              {waitingOptions.map(opt => <option key={opt} value={opt} />)}
            </datalist>
            {waitingReasons.length < MAX_WAITING_REASONS && (
              <button onClick={addReasonField}
                style={{ background: "transparent", color: "#f59e0b", border: "1.5px dashed #f59e0b", borderRadius: 0, padding: "8px 0", width: "100%", fontWeight: 700, fontSize: 13, cursor: "pointer", marginBottom: 16 }}>
                + เพิ่มรายการ
              </button>
            )}
            <div style={{ display: "flex", gap: 8, marginTop: waitingReasons.length < MAX_WAITING_REASONS ? 0 : 16 }}>
              <button onClick={() => setWaitingModal(false)}
                style={{ flex: 1, background: "#f3f4f6", color: "#374151", border: "none", borderRadius: 0, padding: "12px 0", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>
                ยกเลิก
              </button>
              <button onClick={confirmWaiting} disabled={!waitingReasons.some(r => r.trim())}
                style={{ flex: 1, background: waitingReasons.some(r => r.trim()) ? "#f59e0b" : "#e5e7eb", color: waitingReasons.some(r => r.trim()) ? "#fff" : "#9ca3af", border: "none", borderRadius: 0, padding: "12px 0", fontWeight: 700, fontSize: 14, cursor: waitingReasons.some(r => r.trim()) ? "pointer" : "default" }}>
                ยืนยัน
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};

// ── 6.5 LOADING LOG (chat-style feed of completed loads) ─────────────────────
// the work-day rolls over at settings.workDayCutoffHour, not midnight — must stay
// in sync everywhere a "which work-day does this belong to" decision is made
// (cycleDateStr, handleReset's archive-date calc, DATE_STR, SHORT_DATE,
// parseExitDatetime, displayDate), otherwise a time/date can get filed under the
// wrong work-day in one place but not another. All of those now read the same
// settings.workDayCutoffHour value instead of each hardcoding their own hour.
const cycleDateStr = () => {
  const d = new Date();
  if (d.getHours() < settings.workDayCutoffHour) d.setDate(d.getDate() - 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

// เลื่อนวันที่ (รูปแบบ cycleDateStr เดียวกัน) ไปกี่วันก็ได้ — ใช้คำนวณ "วันทำงานถัดไป/ก่อนหน้า"
const addDaysToDateStr = (dateStr, days) => {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + days);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

// วันทำงานที่จะเริ่มขึ้นเมื่อข้ามเวลาตัดรอบครั้งถัดไป — ใช้ตอนอัพโหลดคิวรถ "ล่วงหน้า" จาก LG
// เพื่อรู้ว่าคิวที่อัพนี้ควรถูกดึงมาใช้เป็นคิวปัจจุบันตอนไหน (ดู promotePendingQueueIfDue ใน App())
const nextCycleDateStr = () => addDaysToDateStr(cycleDateStr(), 1);

// doneAt is just "HH:mm"; events before the cutoff belong to the work-day's
// *next* calendar date (truck loaded after midnight, e.g. 00:14, still files under
// the previous work-day) — show that real date next to the time so it's not confusing.
const formatLogTime = (doneAt, workDate) => {
  const hour = parseInt(doneAt.split(":")[0], 10);
  if (hour >= settings.workDayCutoffHour) return doneAt;
  const [y, m, d] = workDate.split("-").map(Number);
  const realDate = new Date(y, m - 1, d + 1);
  const dd = String(realDate.getDate()).padStart(2, "0");
  const mm = String(realDate.getMonth() + 1).padStart(2, "0");
  return `${doneAt} (${dd}/${mm})`;
};

// converts "HH:mm" into a sortable minute value that respects the work-day
// cutoff: times before the cutoff belong to the *next* calendar day, so they must
// sort after evening times of the same work-day (e.g. 18:00 < 00:14 next day).
const workTimeValue = (doneAt) => {
  const [h, m] = doneAt.split(":").map(Number);
  return (h < settings.workDayCutoffHour ? h + 24 : h) * 60 + m;
};

const buildLaneEvents = (list, sources) => {
  const events = [];
  for (const t of list || []) {
    for (const src of sources) {
      for (const lane of lanes) {
        const ld = t[src.field]?.[lane.id];
        if (ld?.done && ld?.doneAt) {
          events.push({
            key: `${t.id}_${src.field}_${lane.id}`,
            truckId: t.id,
            field: src.field,
            laneId: lane.id,
            plate: t.plate,
            customerGroup: t.customerGroup || "",
            zone: t.zone || "",
            laneLabel: lane.tinyLabel,
            laneColor: lane.color,
            laneBg: lane.bg,
            bayLabel: ld.bayId ? (bays.find(b => b.id === ld.bayId)?.label || ld.bayId) : "",
            doneAt: ld.doneAt,
            photos: ld.photos || [],
            doneLabel: src.doneLabel,
            note: src.field === "qcLanes" ? (ld.temp != null ? `${ld.temp}°C` : "") : ld.note,
          });
        }
      }
    }
  }
  return events.sort((a, b) => workTimeValue(a.doneAt) - workTimeValue(b.doneAt));
};

const EventLog = ({ trucks, title, emptyMsg }) => {
  const isMobile = useIsMobile();
  const today = cycleDateStr();
  const [date, setDate] = useState(today);
  const [sourceFilter, setSourceFilter] = useState("");
  const [plateFilter, setPlateFilter] = useState("");
  const [extraFilter, setExtraFilter] = useState("");
  const [archiveTrucks, setArchiveTrucks] = useState(null);
  const [loadingArchive, setLoadingArchive] = useState(false);
  const [zoomUrl, setZoomUrl] = useState(null);

  // always check the archive for the selected date — even when it equals "today" by
  // the wall-clock cutoff — because pressing "ล้างวันใหม่" archives that work-day and
  // empties the live trucks table before the work-day cutoff passes; without this check
  // the Log would keep showing the now-empty live data instead of the archived snapshot
  useEffect(() => {
    setArchiveTrucks(null);
    setLoadingArchive(date !== today);
    supabase.from("wh_archive").select("trucks").eq("archive_date", date).single()
      .then(({ data }) => setArchiveTrucks(data?.trucks ?? null))
      .finally(() => setLoadingArchive(false));
  }, [date, today]);

  const allSources = Object.values(LOG_SOURCES);
  const activeSources = sourceFilter ? allSources.filter(s => s.field === sourceFilter) : allSources;
  const sourceTrucks = archiveTrucks ?? (date === today ? trucks : []);
  const events = buildLaneEvents(sourceTrucks, activeSources).filter(ev => {
    const matchesPlate = !plateFilter.trim() || ev.plate?.toLowerCase().includes(plateFilter.trim().toLowerCase());
    const extraQuery = extraFilter.trim().toLowerCase();
    const matchesExtra = !extraQuery
      || ev.zone?.toLowerCase().includes(extraQuery)
      || ev.customerGroup?.toLowerCase().includes(extraQuery)
      || ev.note?.toLowerCase().includes(extraQuery);
    return matchesPlate && matchesExtra;
  });

  const truckById = new Map((sourceTrucks || []).map(t => [t.id, t]));
  const latestLoadDoneAt = truckId => {
    const t = truckById.get(truckId);
    let latest = null;
    for (const lane of lanes) {
      const ld = t?.loadLanes?.[lane.id];
      if (ld?.done && ld?.doneAt && (latest == null || workTimeValue(ld.doneAt) > workTimeValue(latest))) latest = ld.doneAt;
    }
    return latest;
  };

  const groupsByTruck = [];
  const groupIndex = new Map();
  for (const ev of events) {
    let group = groupIndex.get(ev.truckId);
    if (!group) {
      group = { truckId: ev.truckId, plate: ev.plate, customerGroup: ev.customerGroup, zone: ev.zone, doneAt: ev.doneAt, loadDoneAt: latestLoadDoneAt(ev.truckId), lanes: [] };
      groupIndex.set(ev.truckId, group);
      groupsByTruck.push(group);
    }
    group.lanes.push(ev);
    if (workTimeValue(ev.doneAt) > workTimeValue(group.doneAt)) group.doneAt = ev.doneAt;
  }
  // trucks with a "โหลดเสร็จ" action (any lane) always rank above trucks without one,
  // sorted by recency within each bucket
  groupsByTruck.sort((a, b) => {
    if (!!a.loadDoneAt !== !!b.loadDoneAt) return a.loadDoneAt ? -1 : 1;
    const aTime = a.loadDoneAt ? workTimeValue(a.loadDoneAt) : workTimeValue(a.doneAt);
    const bTime = b.loadDoneAt ? workTimeValue(b.loadDoneAt) : workTimeValue(b.doneAt);
    return bTime - aTime;
  });
  // within each truck, group entries by physical lane (ลานชิ้นส่วน/หัวเครื่องใน/หมูซีก) first —
  // a truck using multiple lanes shows each lane's full sequence together instead of interleaved —
  // lane groups are ordered by their most recent activity (the lane that finished first sinks
  // to the bottom), then by work step (ลานโหลด → QC → Checker), then by time
  const FIELD_ORDER = { qcLanes: 0, sampleLanes: 1, loadLanes: 2 };
  for (const group of groupsByTruck) {
    const laneLatest = {};
    for (const ev of group.lanes) {
      const t = workTimeValue(ev.doneAt);
      if (laneLatest[ev.laneId] == null || t > laneLatest[ev.laneId]) laneLatest[ev.laneId] = t;
    }
    group.lanes.sort((a, b) =>
      laneLatest[b.laneId] - laneLatest[a.laneId]
      || FIELD_ORDER[a.field] - FIELD_ORDER[b.field]
      || workTimeValue(a.doneAt) - workTimeValue(b.doneAt)
    );
  }

  const inp = { border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none", width: isMobile ? "100%" : undefined };

  const exportExcel = () => {
    const rows = [];
    for (const group of groupsByTruck) {
      for (const ev of group.lanes) {
        rows.push({
          "ทะเบียน": group.plate,
          "กลุ่มลูกค้า": group.customerGroup || "",
          "Zone": group.zone || "",
          "ประเภทงาน": ev.laneLabel,
          "จุดจอด": ev.bayLabel || "",
          "รายการ": ev.doneLabel,
          "หมายเหตุ": ev.note || "",
          "เวลา": formatLogTime(ev.doneAt, date),
        });
      }
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), date);
    XLSX.writeFile(wb, `Log_ภาพรวม_${date}.xlsx`);
  };

  return (
    <div style={{ maxWidth: 560, width: "100%", margin: "0 auto", boxSizing: "border-box" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <h2 style={{ margin: 0, fontWeight: 900, fontSize: isMobile ? 18 : 22 }}>{title}</h2>
        <button onClick={exportExcel} disabled={groupsByTruck.length === 0}
          style={{ background: groupsByTruck.length === 0 ? "#e5e7eb" : "#16a34a", color: groupsByTruck.length === 0 ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "7px 16px", fontSize: 13, fontWeight: 700, cursor: groupsByTruck.length === 0 ? "default" : "pointer", whiteSpace: "nowrap" }}>
          Export Excel
        </button>
      </div>
      <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", gap: 8, marginBottom: 8 }}>
        <input type="date" value={date} onChange={e => setDate(e.target.value)} style={inp} />
        <select value={sourceFilter} onChange={e => setSourceFilter(e.target.value)} style={{ ...inp, flex: isMobile ? undefined : 1, minWidth: isMobile ? undefined : 160 }}>
          <option value="">ทั้งหมด</option>
          {allSources.map(s => <option key={s.field} value={s.field}>{s.filterLabel}</option>)}
        </select>
      </div>
      <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", gap: 8, marginBottom: 16 }}>
        <input type="text" placeholder="ค้นหาทะเบียนรถ" value={plateFilter} onChange={e => setPlateFilter(e.target.value)} style={{ ...inp, flex: isMobile ? undefined : 1, minWidth: isMobile ? undefined : 140 }} />
        <input type="text" placeholder="ค้นหา Zone / กลุ่มลูกค้า / Note" value={extraFilter} onChange={e => setExtraFilter(e.target.value)} style={{ ...inp, flex: isMobile ? undefined : 1, minWidth: isMobile ? undefined : 160 }} />
      </div>

      {loadingArchive && <div style={{ textAlign: "center", color: "#9ca3af", padding: 30 }}>กำลังโหลด...</div>}

      {!loadingArchive && groupsByTruck.length === 0 && (
        <div style={{ textAlign: "center", color: "#9ca3af", padding: 30 }}>{emptyMsg}</div>
      )}

      {!loadingArchive && groupsByTruck.map(group => (
        <div key={group.truckId} style={{ background: "#fff", borderRadius: 0, padding: isMobile ? 12 : 14, marginBottom: 12, boxShadow: "0 2px 8px rgba(0,0,0,0.06)", border: "1px solid #e5e7eb" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4, flexWrap: "wrap", gap: 6 }}>
            <b style={{ fontSize: 15 }}>{group.plate}</b>
          </div>
          {(group.customerGroup || group.zone) && (
            <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 10 }}>
              {group.customerGroup && <span>กลุ่มลูกค้า: {group.customerGroup}</span>}
              {group.customerGroup && group.zone && <span> · </span>}
              {group.zone && <span>Zone: {group.zone}</span>}
            </div>
          )}
          {group.lanes.map((ev, i) => (
            <div key={ev.key} style={{ borderTop: i === 0 ? "none" : "1px dashed #e5e7eb", paddingTop: i === 0 ? 0 : 10, marginTop: i === 0 ? 0 : 10 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: ev.photos.length || ev.note ? 8 : 0, flexWrap: "wrap", gap: 6 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ background: ev.laneBg, color: ev.laneColor, borderRadius: 0, padding: "3px 8px", fontSize: 12, fontWeight: 700 }}>
                    {ev.laneLabel}
                  </span>
                  {ev.bayLabel && (
                    <span style={{ background: "#f3f4f6", color: "#374151", borderRadius: 0, padding: "3px 8px", fontSize: 12, fontWeight: 700 }}>
                      {ev.bayLabel}
                    </span>
                  )}
                  <span style={{ color: "#16a34a", fontWeight: 700, fontSize: 13 }}>{ev.doneLabel}</span>
                </div>
                <span style={{ fontSize: 12, color: "#9ca3af" }}>{formatLogTime(ev.doneAt, date)}</span>
              </div>
              {ev.note && <div style={{ fontSize: 13, color: "#374151", marginBottom: 8 }}>{ev.note}</div>}
              {ev.photos.length > 0 && (
                <div style={{ display: "grid", gridTemplateColumns: `repeat(auto-fill, minmax(${isMobile ? 64 : 80}px, 1fr))`, gap: 6 }}>
                  {ev.photos.map((url, idx) => (
                    <img key={idx} src={url} alt="" onClick={() => setZoomUrl(url)}
                      style={{ width: "100%", aspectRatio: "1", objectFit: "cover", borderRadius: 0, cursor: "pointer" }} />
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      ))}

      {zoomUrl && (
        <div onClick={() => setZoomUrl(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 1000, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 14, padding: 20, cursor: "zoom-out" }}>
          <img src={zoomUrl} alt="" onClick={e => e.stopPropagation()} style={{ maxWidth: "100%", maxHeight: "75vh", borderRadius: 0, cursor: "default" }} />
          <button onClick={e => { e.stopPropagation(); saveImageToDevice(zoomUrl, `photo-${Date.now()}.jpg`); }}
            style={{ background: "#16a34a", color: "#fff", border: "none", borderRadius: 0, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
            💾 บันทึกรูปลงเครื่อง
          </button>
        </div>
      )}
    </div>
  );
};

const LOG_SOURCES = {
  loadLanes:   { field: "loadLanes",   filterLabel: "การโหลดจ่ายสินค้า",      doneLabel: "✅ Checker นับตะกร้าและโหลดเสร็จแล้ว" },
  qcLanes:     { field: "qcLanes",     filterLabel: "การตรวจอุณหภูมิรถขนส่ง", doneLabel: "🌡️ ลานโหลด ตรวจสอบอุณหภูมิรถแล้ว" },
  sampleLanes: { field: "sampleLanes", filterLabel: "การตรวจอุณหภูมิสินค้า",   doneLabel: "📷 QC สุ่มตรวจอุณหภูมิสินค้าแล้ว" },
};

const OverviewLog = ({ trucks }) => (
  <EventLog trucks={trucks} title="💬 Log ภาพรวมการทำงาน" emptyMsg="ยังไม่มีรายการทำงาน" />
);

// ── 6.6 BASKET / HOOK SUMMARY (ยอดตะกร้า,ตะขอรายวัน) ─────────────────────────
const BasketSummary = ({ trucks }) => {
  const isMobile = useIsMobile();
  const today = cycleDateStr();
  const [date, setDate] = useState(today);
  const [plateFilter, setPlateFilter] = useState("");
  const [archiveTrucks, setArchiveTrucks] = useState(null);
  const [loadingArchive, setLoadingArchive] = useState(false);

  useEffect(() => {
    setArchiveTrucks(null);
    setLoadingArchive(date !== today);
    supabase.from("wh_archive").select("trucks").eq("archive_date", date).single()
      .then(({ data }) => setArchiveTrucks(data?.trucks ?? null))
      .finally(() => setLoadingArchive(false));
  }, [date, today]);

  const sourceTrucks = archiveTrucks ?? (date === today ? trucks : []);

  // ── การคืนตะกร้า (สะสมข้ามวัน ไม่ล้างตอนปิดงาน) ──
  const zeroBaskets = () => Object.fromEntries(basketTypes.map(b => [b.key, 0]));
  const emptyReturnForm = () => ({ plate: "", ...Object.fromEntries(basketTypes.map(b => [b.key, ""])) });
  const [allArchiveTrucks, setAllArchiveTrucks] = useState([]);
  const [returns, setReturns] = useState([]);
  const [returnForm, setReturnForm] = useState(emptyReturnForm());
  const [savingReturn, setSavingReturn] = useState(false);

  const fetchReturns = () => {
    supabase.from("wh_basket_returns").select("data").then(({ data }) => setReturns((data || []).map(r => r.data)));
  };

  useEffect(() => {
    supabase.from("wh_archive").select("trucks").then(({ data }) => setAllArchiveTrucks((data || []).map(r => r.trucks || [])));
    fetchReturns();
  }, []);

  // ตัดตัวอักษร/ขีด/เว้นวรรคออก เหลือแค่ตัวเลข — กันแยกยอดผิดคันเวลาพิมพ์ทะเบียนต่างรูปแบบ
  // (เช่น "1กข-1234" vs "1กข 1234") ใช้ key เดียวกับที่หน้าอื่นในแอปนี้ใช้จับคู่ทะเบียนกันอยู่แล้ว
  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";

  // ยอดออกสะสมทุกวัน (archive ทั้งหมด + คิววันนี้ที่ยังไม่ปิดงาน) แยกตามทะเบียนรถ
  const issuedByPlate = {};
  for (const t of [...allArchiveTrucks.flat(), ...trucks]) {
    if (!t?.plate) continue;
    const key = plateNum(t.plate);
    if (!key) continue;
    for (const l of lanes) {
      const ld = t.loadLanes?.[l.id];
      if (!ld?.baskets) continue;
      const cur = issuedByPlate[key] || { plate: t.plate, ...zeroBaskets() };
      for (const b of basketTypes) cur[b.key] += ld.baskets[b.key] || 0;
      issuedByPlate[key] = cur;
    }
  }

  // ยอดคืนสะสม แยกตามทะเบียนรถ
  const returnedByPlate = {};
  for (const r of returns) {
    if (!r?.plate) continue;
    const key = plateNum(r.plate);
    if (!key) continue;
    const cur = returnedByPlate[key] || zeroBaskets();
    for (const b of basketTypes) cur[b.key] += r[b.key] || 0;
    returnedByPlate[key] = cur;
  }

  const outstandingRows = Object.keys(issuedByPlate).map(key => {
    const issued = issuedByPlate[key];
    const plate = issued.plate;
    const returned = returnedByPlate[key] || zeroBaskets();
    const out = Object.fromEntries(basketTypes.map(b => [b.key, issued[b.key] - returned[b.key]]));
    const total = basketTypes.reduce((sum, b) => sum + out[b.key], 0);
    return { plate, out, total };
  }).filter(r => r.total > 0).sort((a, b) => b.total - a.total);

  const outstandingGrandTotal = outstandingRows.reduce((sum, r) => sum + r.total, 0);

  const submitReturn = async () => {
    const plate = returnForm.plate.trim();
    if (!plate) { alert("กรุณากรอกทะเบียนรถ"); return; }
    const counts = Object.fromEntries(basketTypes.map(b => [b.key, Number(returnForm[b.key]) || 0]));
    if (!Object.values(counts).some(v => v > 0)) {
      alert("กรุณากรอกจำนวนตะกร้า/ตะขอที่คืนอย่างน้อย 1 ช่อง");
      return;
    }
    setSavingReturn(true);
    try {
      const id = `return_${plate}_${Date.now()}`;
      const { error } = await supabase.from("wh_basket_returns").insert({ id, data: { plate, ...counts, returnedAt: new Date().toISOString() } });
      if (error) throw error;
      setReturnForm(emptyReturnForm());
      fetchReturns();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setSavingReturn(false);
    }
  };

  const th = { padding: "10px 12px", textAlign: "center", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb", background: "#f9fafb", whiteSpace: "nowrap" };
  const td = { padding: "10px 12px", textAlign: "center", borderBottom: "1px solid #f3f4f6" };
  const inp = { border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none", width: isMobile ? "100%" : undefined };

  const totals = {};
  const payersByLane = {};
  for (const l of lanes) { totals[l.id] = zeroBaskets(); payersByLane[l.id] = new Set(); }

  const detailRows = [];
  for (const t of sourceTrucks) {
    for (const l of lanes) {
      const ld = t.loadLanes?.[l.id];
      if (!ld?.baskets) continue;
      for (const b of basketTypes) totals[l.id][b.key] += ld.baskets[b.key] || 0;
      if (ld.basketPayer) payersByLane[l.id].add(ld.basketPayer);
      detailRows.push({ key: `${t.id}_${l.id}`, plate: t.plate, lane: l, baskets: ld.baskets, payer: ld.basketPayer, doneAt: ld.doneAt });
    }
  }
  detailRows.sort((a, b) => (a.doneAt || "").localeCompare(b.doneAt || ""));
  const filteredDetailRows = plateFilter.trim()
    ? detailRows.filter(r => r.plate?.toLowerCase().includes(plateFilter.trim().toLowerCase()))
    : detailRows;

  const laneTotal = (l) => basketTypes.filter(b => b.countsInTotal).reduce((sum, b) => sum + totals[l.id][b.key], 0);

  const exportExcel = () => {
    const summaryRows = [
      ...basketTypes.map(b => ({
        "สีตะกร้า": b.label,
        ...Object.fromEntries(lanes.map(l => [l.tinyLabel, totals[l.id][b.key] || 0])),
      })),
      {
        "สีตะกร้า": "รวมตะกร้า",
        ...Object.fromEntries(lanes.map(l => [l.tinyLabel, laneTotal(l)])),
      },
      {
        "สีตะกร้า": "ผู้จ่าย",
        ...Object.fromEntries(lanes.map(l => [l.tinyLabel, [...payersByLane[l.id]].join(", ") || ""])),
      },
    ];
    const detailRowsForExport = filteredDetailRows.map(r => ({
      "ทะเบียน": r.plate,
      "ลาน": r.lane.tinyLabel,
      ...Object.fromEntries(basketTypes.map(b => [b.label, r.baskets[b.key] || 0])),
      "ผู้จ่าย": r.payer || "",
      "เวลา": r.doneAt || "",
    }));
    const outstandingForExport = outstandingRows.map(r => ({
      "ทะเบียน": r.plate,
      ...Object.fromEntries(basketTypes.map(b => [b.label, r.out[b.key]])),
      "รวมค้างคืน": r.total,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "สรุป");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detailRowsForExport), "รายละเอียด");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(outstandingForExport), "ค้างคืน");
    XLSX.writeFile(wb, `ตะกร้า_${date}.xlsx`);
  };

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <h2 style={{ margin: 0, fontWeight: 900, fontSize: 22 }}>🧺 ข้อมูลยอดตะกร้า/ตะขอ</h2>
        <button onClick={exportExcel}
          style={{ background: "#16a34a", color: "#fff", border: "none", borderRadius: 0, padding: "7px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>
          Export Excel
        </button>
      </div>

      <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", gap: 8, marginBottom: 20 }}>
        <input type="date" value={date} onChange={e => setDate(e.target.value)} style={inp} />
        <input type="text" placeholder="ค้นหาทะเบียนรถ" value={plateFilter} onChange={e => setPlateFilter(e.target.value)} style={{ ...inp, flex: isMobile ? undefined : 1, minWidth: isMobile ? undefined : 160 }} />
      </div>

      {loadingArchive && <div style={{ textAlign: "center", color: "#9ca3af", padding: 30 }}>กำลังโหลด...</div>}

      {!loadingArchive && <>
      <div style={{ background: "#fff", borderRadius: 0, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.07)", marginBottom: 20 }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 700, fontSize: 14 }}>
          🔄 บันทึกการคืนตะกร้า/ตะขอ
        </div>
        <div style={{ padding: 16 }}>
          <input
            list="basketReturnPlates"
            placeholder="ทะเบียนรถ"
            value={returnForm.plate}
            onChange={e => setReturnForm(f => ({ ...f, plate: e.target.value }))}
            style={{ ...inp, width: "100%", marginBottom: 8 }}
          />
          <datalist id="basketReturnPlates">
            {Object.values(issuedByPlate).map(v => <option key={v.plate} value={v.plate} />)}
          </datalist>
          <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "1fr 1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
            {basketTypes.map(b => (
              <div key={b.key}>
                <label style={{ display: "block", fontSize: 11, color: "#6b7280", marginBottom: 4 }}>{b.label}</label>
                <input type="number" min="0" inputMode="numeric" value={returnForm[b.key]}
                  onChange={e => setReturnForm(fm => ({ ...fm, [b.key]: e.target.value }))}
                  style={{ ...inp, width: "100%" }} />
              </div>
            ))}
          </div>
          <button onClick={submitReturn} disabled={savingReturn}
            style={{ width: "100%", background: savingReturn ? "#e5e7eb" : "#16a34a", color: savingReturn ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "11px 0", fontWeight: 700, fontSize: 14, cursor: savingReturn ? "default" : "pointer" }}>
            {savingReturn ? "⏳ กำลังบันทึก..." : "✅ บันทึกคืนตะกร้า"}
          </button>
        </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 0, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.07)", marginBottom: 20 }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr>
                <th style={{ ...th, textAlign: "left" }}>สีตะกร้า</th>
                {lanes.map(l => <th key={l.id} style={{ ...th, color: l.color }}>{l.emoji} {l.tinyLabel}</th>)}
              </tr>
            </thead>
            <tbody>
              {basketTypes.map(b => (
                <tr key={b.key}>
                  <td style={{ ...td, textAlign: "left", fontWeight: 700 }}>{b.label}</td>
                  {lanes.map(l => <td key={l.id} style={td}>{totals[l.id][b.key] || 0}</td>)}
                </tr>
              ))}
              <tr>
                <td style={{ ...td, textAlign: "left", fontWeight: 800, background: "#f9fafb" }}>รวมตะกร้า</td>
                {lanes.map(l => <td key={l.id} style={{ ...td, fontWeight: 800, background: "#f9fafb" }}>{laneTotal(l)}</td>)}
              </tr>
              <tr>
                <td style={{ ...td, textAlign: "left", fontWeight: 700, color: "#6b7280" }}>ผู้จ่าย</td>
                {lanes.map(l => <td key={l.id} style={{ ...td, fontSize: 12, color: "#6b7280" }}>{[...payersByLane[l.id]].join(", ") || "—"}</td>)}
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 0, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.07)" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 700, fontSize: 14 }}>
          📋 รายการต่อคันรถ <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{filteredDetailRows.length}</span>
        </div>
        {filteredDetailRows.length === 0
          ? <div style={{ padding: 36, textAlign: "center", color: "#9ca3af" }}>{plateFilter.trim() ? `ไม่พบทะเบียน "${plateFilter}"` : "ยังไม่มีข้อมูลตะกร้าวันนี้"}</div>
          : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr>
                    <th style={th}>ทะเบียน</th>
                    <th style={th}>ลาน</th>
                    {basketTypes.map(b => <th key={b.key} style={th}>{b.label}</th>)}
                    <th style={th}>ผู้จ่าย</th>
                    <th style={th}>เวลา</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredDetailRows.map(r => (
                    <tr key={r.key}>
                      <td style={{ ...td, fontWeight: 700 }}>{r.plate}</td>
                      <td style={{ ...td, color: r.lane.color }}>{r.lane.tinyLabel}</td>
                      {basketTypes.map(b => <td key={b.key} style={td}>{r.baskets[b.key] || 0}</td>)}
                      <td style={td}>{r.payer || "—"}</td>
                      <td style={td}>{r.doneAt || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )
        }
      </div>

      <div style={{ background: "#fff", borderRadius: 0, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.07)", marginTop: 20 }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 700, fontSize: 14 }}>
          📦 ตะกร้าค้างคืน (สะสมทุกวัน) <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{outstandingRows.length}</span>
        </div>
        {outstandingRows.length === 0
          ? <div style={{ padding: 36, textAlign: "center", color: "#9ca3af" }}>ไม่มีตะกร้า/ตะขอค้างคืน</div>
          : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr>
                    <th style={{ ...th, textAlign: "left" }}>ทะเบียน</th>
                    {basketTypes.map(b => <th key={b.key} style={th}>{b.label}</th>)}
                    <th style={th}>รวมค้างคืน</th>
                  </tr>
                </thead>
                <tbody>
                  {outstandingRows.map(r => (
                    <tr key={r.plate}>
                      <td style={{ ...td, textAlign: "left", fontWeight: 700 }}>{r.plate}</td>
                      {basketTypes.map(b => <td key={b.key} style={td}>{r.out[b.key]}</td>)}
                      <td style={{ ...td, fontWeight: 800, color: "#dc2626" }}>{r.total}</td>
                    </tr>
                  ))}
                  <tr>
                    <td style={{ ...td, textAlign: "left", fontWeight: 800, background: "#f9fafb" }}>รวมทั้งหมด</td>
                    <td style={{ ...td, fontWeight: 800, background: "#f9fafb" }} colSpan={basketTypes.length}></td>
                    <td style={{ ...td, fontWeight: 800, background: "#f9fafb", color: "#dc2626" }}>{outstandingGrandTotal}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          )
        }
      </div>
      </>}
    </div>
  );
};

// ── 6.7 WAITING-FOR-PRODUCT SUMMARY (ข้อมูลการรอสินค้า) ────────────────────────
const WaitingSummary = ({ trucks }) => {
  const isMobile = useIsMobile();
  const today = cycleDateStr();
  const [date, setDate] = useState(today);
  const [plateFilter, setPlateFilter] = useState("");
  const [archiveTrucks, setArchiveTrucks] = useState(null);
  const [loadingArchive, setLoadingArchive] = useState(false);
  const [nowMin, setNowMin] = useState(() => workTimeValue(TIME_NOW()));

  useEffect(() => {
    const id = setInterval(() => setNowMin(workTimeValue(TIME_NOW())), 30000);
    return () => clearInterval(id);
  }, []);

  useEffect(() => {
    setArchiveTrucks(null);
    setLoadingArchive(date !== today);
    supabase.from("wh_archive").select("trucks").eq("archive_date", date).single()
      .then(({ data }) => setArchiveTrucks(data?.trucks ?? null))
      .finally(() => setLoadingArchive(false));
  }, [date, today]);

  const sourceTrucks = archiveTrucks ?? (date === today ? trucks : []);

  const th = { padding: "10px 12px", textAlign: "center", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb", background: "#f9fafb", whiteSpace: "nowrap" };
  const td = { padding: "10px 12px", textAlign: "center", borderBottom: "1px solid #f3f4f6" };
  const inp = { border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none", width: isMobile ? "100%" : undefined };

  // log of every wait period (ongoing or resolved) — loadLanes[lane] only keeps the
  // latest waitingAt/waitingFor, so a truck that waited more than once per lane in a
  // day only shows its most recent wait period here, not every one
  const rows = [];
  for (const t of sourceTrucks) {
    for (const l of lanes) {
      const ld = t.loadLanes?.[l.id];
      if (!ld?.waiting || !ld.waitingAt) continue;
      const ongoing = !ld.done;
      const endAt = ongoing ? null : ld.doneAt;
      const durationMin = ongoing
        ? Math.max(0, nowMin - workTimeValue(ld.waitingAt))
        : (endAt ? Math.max(0, workTimeValue(endAt) - workTimeValue(ld.waitingAt)) : null);
      rows.push({ key: `${t.id}_${l.id}`, plate: t.plate, customerGroup: t.customerGroup, lane: l, waitingFor: ld.waitingFor, startAt: ld.waitingAt, endAt, ongoing, durationMin });
    }
  }
  rows.sort((a, b) => {
    if (a.ongoing !== b.ongoing) return a.ongoing ? -1 : 1;
    if (a.ongoing) return (b.durationMin ?? 0) - (a.durationMin ?? 0);
    return workTimeValue(b.endAt || "00:00") - workTimeValue(a.endAt || "00:00");
  });

  const filtered = plateFilter.trim()
    ? rows.filter(r => r.plate?.toLowerCase().includes(plateFilter.trim().toLowerCase()))
    : rows;

  const exportExcel = () => {
    const rowsForExport = filtered.map(r => ({
      "ทะเบียน": r.plate,
      "กลุ่มลูกค้า": r.customerGroup || "",
      "ลาน": r.lane.tinyLabel,
      "สินค้าที่รอ": r.waitingFor || "",
      "ตั้งแต่": r.startAt || "",
      "ถึง": r.ongoing ? "" : (r.endAt || ""),
      "ระยะเวลา (นาที)": r.durationMin ?? "",
      "สถานะ": r.ongoing ? "กำลังรอ" : "เสร็จแล้ว",
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rowsForExport), "รอสินค้า");
    XLSX.writeFile(wb, `รอสินค้า_${date}.xlsx`);
  };

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <h2 style={{ margin: 0, fontWeight: 900, fontSize: 22 }}>⏳ ข้อมูลการรอสินค้า</h2>
        <button onClick={exportExcel}
          style={{ background: "#16a34a", color: "#fff", border: "none", borderRadius: 0, padding: "7px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>
          Export Excel
        </button>
      </div>

      <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", gap: 8, marginBottom: 16 }}>
        <input type="date" value={date} onChange={e => setDate(e.target.value)} style={inp} />
        <input type="text" placeholder="ค้นหาทะเบียนรถ" value={plateFilter} onChange={e => setPlateFilter(e.target.value)}
          style={{ ...inp, flex: isMobile ? undefined : 1, minWidth: isMobile ? undefined : 160 }} />
      </div>

      {loadingArchive && <div style={{ textAlign: "center", color: "#9ca3af", padding: 30 }}>กำลังโหลด...</div>}

      {!loadingArchive && (
      <div style={{ background: "#fff", borderRadius: 0, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.07)" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 700, fontSize: 14 }}>
          🕓 ประวัติการรอสินค้า <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{filtered.length}</span>
        </div>
        {filtered.length === 0
          ? <div style={{ padding: 36, textAlign: "center", color: "#9ca3af" }}>{plateFilter.trim() ? `ไม่พบทะเบียน "${plateFilter}"` : "ยังไม่มีการรอสินค้าวันนี้"}</div>
          : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr>
                    <th style={{ ...th, textAlign: "left" }}>ทะเบียน</th>
                    <th style={{ ...th, textAlign: "left" }}>กลุ่มลูกค้า</th>
                    <th style={th}>ลาน</th>
                    <th style={{ ...th, textAlign: "left" }}>สินค้าที่รอ</th>
                    <th style={th}>ตั้งแต่</th>
                    <th style={th}>ถึง</th>
                    <th style={th}>ระยะเวลา</th>
                    <th style={th}>สถานะ</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map(r => {
                    const urgent = r.ongoing && (r.durationMin ?? 0) >= settings.waitingUrgentMinutes;
                    return (
                      <tr key={r.key} style={{ background: urgent ? "#fff5f5" : undefined }}>
                        <td style={{ ...td, textAlign: "left", fontWeight: 700 }}>{r.plate}</td>
                        <td style={{ ...td, textAlign: "left", color: "#6b7280" }}>{r.customerGroup || "—"}</td>
                        <td style={{ ...td, color: r.lane.color, fontWeight: 700 }}>{r.lane.tinyLabel}</td>
                        <td style={{ ...td, textAlign: "left" }}>{r.waitingFor || "—"}</td>
                        <td style={td}>{r.startAt || "—"}</td>
                        <td style={td}>{r.ongoing ? "—" : (r.endAt || "—")}</td>
                        <td style={{ ...td, fontWeight: 700, color: urgent ? "#dc2626" : "#374151" }}>{r.durationMin != null ? formatMinsDelta(r.durationMin) : "—"}</td>
                        <td style={td}>
                          {r.ongoing
                            ? <span style={{ background: "#fef3c7", color: "#92400e", borderRadius: 0, padding: "3px 10px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>⏳ กำลังรอ</span>
                            : <span style={{ background: "#d1fae5", color: "#065f46", borderRadius: 0, padding: "3px 10px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>✓ เสร็จแล้ว</span>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        }
      </div>
      )}
    </div>
  );
};

// ── 7. PLANNING ───────────────────────────────────────────────────────────────
const Planning = ({ trucks, queue, onUpdate }) => {
  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const usedPlan = new Set();
  const matchTruckPlan = q => {
    let t = trucks.find(t => t.queueId === q.id && !usedPlan.has(t.id));
    if (!t) t = trucks.find(t => !t.queueId && plateNum(t.plate) === plateNum(q.plate) && plateNum(q.plate) !== "" && !usedPlan.has(t.id));
    if (t) usedPlan.add(t.id);
    return t;
  };
  const planQueueRows = queue.map(q => ({ key: q.id, plate: q.plate, customerGroup: q.customerGroup, truck: matchTruckPlan(q) }));
  const walkIns = trucks.filter(t => !usedPlan.has(t.id));
  const allRows = [
    ...planQueueRows,
    ...walkIns.map(t => ({ key: t.id, plate: t.plate, customerGroup: t.customerGroup || "–", truck: t })),
  ].sort((a, b) => {
    const rank = t => {
      if (!t) return 1;
      if (t.status === "invoiced") return 3;     // ออกแล้ว → ล่าง
      if (t.status === "summary_printed") return 0; // ออกได้เลย → บน
      return 2;
    };
    return rank(a.truck) - rank(b.truck);
  }).filter(row => !settings.excludedCustomerGroups.includes(row.customerGroup));

  const Tick = () => <span style={{ color: "#10b981", fontWeight: 700, fontSize: 13 }}>✓</span>;
  const Dash = () => <span style={{ color: "#d1d5db", fontSize: 12 }}>—</span>;

  const InvoiceCell = ({ truck }) => {
    const [saving, setSaving] = useState(false);
    if (!truck || !["summary_printed", "invoiced"].includes(truck.status)) return <Dash />;
    if (truck.status === "invoiced") return <Tick />;
    return (
      <button onClick={async () => {
        setSaving(true);
        try { await onUpdate(truck.id, { invoiceDone: true, status: "invoiced", invoicedAt: TIME_NOW() }); }
        catch (e) { alert("บันทึกไม่สำเร็จ: " + e.message); setSaving(false); }
      }} disabled={saving}
        style={{ background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "5px 8px", fontWeight: 700, fontSize: 11, cursor: "pointer", whiteSpace: "nowrap" }}>
        {saving ? "⏳" : "ออก Invoice"}
      </button>
    );
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 18px", fontWeight: 900, fontSize: 22 }}>📄 ห้องวางแผน</h2>

      <div style={{ background: "#fff", borderRadius: 0, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.07)" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 700, fontSize: 14 }}>
          📋 คิวรถวันนี้ <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{allRows.length}</span>
        </div>
        {allRows.length === 0
          ? <div style={{ padding: 36, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีคิวรถ</div>
          : (
          <div style={{ overflowX: "auto", overflowY: "auto", maxHeight: "calc(100vh - 190px)" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead style={{ position: "sticky", top: 0, zIndex: 10 }}>
                <tr style={{ background: "#f9fafb" }}>
                  {["ทะเบียน","กลุ่มลูกค้า","เวลาเข้าโรงงาน","สถานะ","③ ใบเบิกสินค้า","⑥ ใบสรุปจ่าย","⑦ ใบ Invoice"].map(h => (
                    <th key={h} style={{ padding: "9px 12px", textAlign: "left", fontWeight: 700, color: "#374151", whiteSpace: "nowrap", borderBottom: "1px solid #e5e7eb", background: "#f9fafb" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {allRows.map(({ key, plate, customerGroup, entryTime, truck }) => (
                  <tr key={key} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "10px 12px", fontWeight: 800 }}>{plate}</td>
                    <td style={{ padding: "10px 12px", color: "#374151" }}>{customerGroup}</td>
                    <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                      <div style={{ fontWeight: 700, color: "#3b82f6" }}>{entryTime || "—"}</div>
                      {truck?.arrivedAt
                        ? <div style={{ fontSize: 10, color: "#6b7280", marginTop: 2 }}>เข้าจริง {truck.arrivedAt}</div>
                        : <div style={{ fontSize: 10, color: "#6b7280", marginTop: 2 }}>(รถยังไม่เข้าโรงงาน)</div>}
                    </td>
                    <td style={{ padding: "10px 12px" }}>
                      {!truck
                        ? <span style={{ fontSize: 11, color: "#9ca3af", fontWeight: 600 }}>รอเช็คอิน</span>
                        : (() => {
                            const anyQC = lanes.some(l => truck.qcLanes?.[l.id]?.done);
                            return (
                              <div>
                                {!anyQC
                                  ? <span style={{ fontSize: 11, color: "#6b7280", fontWeight: 600 }}>รอเข้าโหลด</span>
                                  : <div style={{ display: "flex", flexWrap: "wrap", gap: 5, alignItems: "center" }}>
                                      {lanes.map(l => {
                                        const loaded = truck.loadLanes?.[l.id]?.done;
                                        const qcDone = truck.qcLanes?.[l.id]?.done;
                                        const waiting = truck.loadLanes?.[l.id]?.waiting && !loaded;
                                        if (loaded) return (
                                          <div key={l.id} style={{ position: "relative", display: "inline-block", background: "#10b981", color: "#fff", borderRadius: 0, padding: "3px 10px 5px 8px", fontSize: 11, fontWeight: 700, lineHeight: 1.4 }}>
                                            {l.tinyLabel}
                                            <span style={{ position: "absolute", bottom: -4, right: -4, background: "#059669", border: "2px solid #fff", borderRadius: "50%", width: 14, height: 14, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8, fontWeight: 900 }}>✓</span>
                                          </div>
                                        );
                                        if (waiting) return (
                                          <div key={l.id} style={{ position: "relative", display: "inline-block", background: "#fbbf24", color: "#fff", borderRadius: 0, padding: "3px 10px 5px 8px", fontSize: 11, fontWeight: 700, lineHeight: 1.4, whiteSpace: "nowrap" }}>
                                            รอสินค้า {l.tinyLabel}
                                            <span style={{ position: "absolute", bottom: -4, right: -4, background: "#d97706", border: "2px solid #fff", borderRadius: "50%", width: 14, height: 14, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 8 }}>⏳</span>
                                          </div>
                                        );
                                        if (qcDone) return (
                                          <span key={l.id} style={{ fontSize: 11, color: "#f97316", fontWeight: 700, whiteSpace: "nowrap" }}>กำลังโหลด {l.tinyLabel}</span>
                                        );
                                        return null;
                                      })}
                                    </div>
                                }
                              </div>
                            );
                          })()
                      }
                    </td>
                    <td style={{ padding: "10px 12px" }}>{truck?.pickupPrinted ? <Tick/> : <Dash/>}</td>
                    <td style={{ padding: "10px 12px" }}>{truck?.summaryPrinted ? <Tick/> : <Dash/>}</td>
                    <td style={{ padding: "10px 12px" }}><InvoiceCell truck={truck} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};

// ── DOWNLOAD ─────────────────────────────────────────────────────────────────
const Download = ({ onReset }) => {
  const [exportDate, setExportDate] = useState("");
  const [loading, setLoading] = useState(false);
  const [archives, setArchives] = useState([]);
  const [deleteDate, setDeleteDate] = useState("");
  const [deleting, setDeleting] = useState(false);

  const loadArchives = () => {
    supabase.from("wh_archive").select("archive_date").order("archive_date", { ascending: false })
      .then(({ data }) => setArchives((data || []).map(r => r.archive_date)));
  };

  useEffect(() => { loadArchives(); }, []);

  const handleDownload = async () => {
    if (!exportDate) return;
    setLoading(true);
    await exportArchiveExcel(exportDate);
    setLoading(false);
  };

  const handleDeleteArchive = async () => {
    if (!deleteDate) return;
    if (!window.confirm(`ลบข้อมูล Archive วันที่ ${deleteDate} ถาวร?`)) return;
    setDeleting(true);
    await supabase.from("wh_archive").delete().eq("archive_date", deleteDate);
    setDeleteDate("");
    loadArchives();
    setDeleting(false);
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 20px", fontWeight: 900, fontSize: 22 }}>จบการทำงาน</h2>

      <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", padding: 24, maxWidth: 480, marginBottom: 20 }}>
        <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 6 }}>🗑️ ล้างวันใหม่</div>
        <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 14 }}>ล้างข้อมูลรถและคิวทั้งหมด แล้วเริ่มต้นวันใหม่ ข้อมูลจะถูก archive ไว้ก่อน</div>
        <button onClick={onReset}
          style={{ background: "#fee2e2", color: "#991b1b", border: "1.5px solid #fca5a5", borderRadius: 0, padding: "12px 0", fontWeight: 700, fontSize: 14, cursor: "pointer", width: "100%" }}>
          🗑️ ล้างวันใหม่
        </button>
      </div>

      <h3 style={{ margin: "0 0 12px", fontWeight: 800, fontSize: 16 }}>📥 ดาวน์โหลดข้อมูลย้อนหลัง</h3>
      <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", padding: 24, maxWidth: 480 }}>
        <label style={{ display: "block", fontWeight: 700, fontSize: 13, marginBottom: 8 }}>เลือกวันที่</label>
        <input
          type="date"
          value={exportDate}
          onChange={e => setExportDate(e.target.value)}
          style={{ width: "100%", border: "1px solid #d1d5db", borderRadius: 0, padding: "10px 12px", fontSize: 14, boxSizing: "border-box", marginBottom: 12 }}
        />
        <button
          onClick={handleDownload}
          disabled={!exportDate || loading}
          style={{ width: "100%", background: exportDate ? "#111" : "#e5e7eb", color: exportDate ? "#fff" : "#9ca3af", border: "none", borderRadius: 0, padding: "13px 0", fontSize: 15, fontWeight: 700, cursor: exportDate ? "pointer" : "default" }}
        >
          {loading ? "กำลังดาวน์โหลด..." : "⬇️ ดาวน์โหลด Excel"}
        </button>
        {archives.length > 0 && (
          <div style={{ marginTop: 20 }}>
            <div style={{ fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 8 }}>ข้อมูลที่มีใน Archive</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {archives.map(d => (
                <button key={d} onClick={() => setExportDate(d)}
                  style={{ background: exportDate === d ? "#111" : "#f3f4f6", color: exportDate === d ? "#fff" : "#374151", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                  {d}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      <h3 style={{ margin: "24px 0 12px", fontWeight: 800, fontSize: 16 }}>🗑️ ลบข้อมูลย้อนหลัง</h3>
      <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", padding: 24, maxWidth: 480 }}>
        <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 14 }}>เลือกวันที่แล้วลบข้อมูล Archive — การลบจะไม่สามารถกู้คืนได้</div>
        {archives.length > 0 && (
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 8 }}>เลือกจาก Archive ที่มี</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {archives.map(d => (
                <button key={d} onClick={() => setDeleteDate(d)}
                  style={{ background: deleteDate === d ? "#991b1b" : "#f3f4f6", color: deleteDate === d ? "#fff" : "#374151", border: "none", borderRadius: 0, padding: "4px 10px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                  {d}
                </button>
              ))}
            </div>
          </div>
        )}
        <input
          type="date"
          value={deleteDate}
          onChange={e => setDeleteDate(e.target.value)}
          style={{ width: "100%", border: "1px solid #d1d5db", borderRadius: 0, padding: "10px 12px", fontSize: 14, boxSizing: "border-box", marginBottom: 12 }}
        />
        <button
          onClick={handleDeleteArchive}
          disabled={!deleteDate || deleting}
          style={{ width: "100%", background: deleteDate ? "#fee2e2" : "#e5e7eb", color: deleteDate ? "#991b1b" : "#9ca3af", border: deleteDate ? "1.5px solid #fca5a5" : "none", borderRadius: 0, padding: "13px 0", fontSize: 15, fontWeight: 700, cursor: deleteDate ? "pointer" : "default" }}
        >
          {deleting ? "กำลังลบ..." : `🗑️ ลบ Archive${deleteDate ? ` วันที่ ${deleteDate}` : ""}`}
        </button>
      </div>
    </div>
  );
};

// ─── COLLAPSIBLE CARD (พับ/กางกล่องตั้งค่า — คลิกหัวข้อเพื่อเข้าไปแก้) ─────────────
const Collapsible = ({ title, defaultOpen = false, children }) => {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 1px 6px rgba(0,0,0,0.07)" }}>
      <button onClick={() => setOpen(o => !o)}
        style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, background: "none", border: "none", padding: "16px 20px", cursor: "pointer", textAlign: "left", fontFamily: "inherit" }}>
        <span style={{ fontWeight: 800, fontSize: 14, color: "#111" }}>{title}</span>
        <span style={{ fontSize: 12, color: "#9ca3af", transform: open ? "rotate(180deg)" : "none", transition: "transform 0.15s", flexShrink: 0 }}>▾</span>
      </button>
      {open && <div style={{ padding: "0 20px 16px" }}>{children}</div>}
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// SYSTEM SETTINGS (wh_settings) — ค่าที่เคย hardcode ในโค้ด แก้ตรงนี้แทน
// ─────────────────────────────────────────────────────────────────────────────
const SystemSettings = () => {
  const [form, setForm] = useState({
    workDayCutoffHour:    settings.workDayCutoffHour,
    waitingUrgentMinutes: settings.waitingUrgentMinutes,
    maxPhotoUploads:      settings.maxPhotoUploads,
    maxWaitingReasons:    settings.maxWaitingReasons,
    geofenceLat:          settings.geofence.lat,
    geofenceLng:          settings.geofence.lng,
    geofenceRadiusM:      settings.geofence.radiusM,
    facilityName:            settings.facilityName,
    siteTitle:               settings.siteTitle,
    unitPrice:               settings.unitPrice,
    vatRate:                 settings.vatRate,
    exitTimeWindowMinutes:   settings.exitTimeWindowMinutes,
    docTimeMinutes:          settings.docTimeMinutes,
    excludedCustomerGroups:  settings.excludedCustomerGroups.join(", "),
  });
  const [saving, setSaving] = useState(false);
  const [msg,    setMsg]    = useState("");

  const set = (key) => (e) => setForm(f => ({ ...f, [key]: e.target.value }));

  // ตรวจช่วงค่าให้ตรงกับที่ SETTERS ใน lib/settings.js ยอมรับจริง — ถ้าไม่เช็คตรงนี้ก่อน
  // กรอกค่านอกช่วง (เช่น ปล่อยว่าง lat/lng, ใส่ 24 ในเวลาตัดรอบ) จะเห็น "บันทึกสำเร็จ" ทั้งที่
  // saveSetting เขียนลง DB ไปแล้วแต่ SETTERS ปฏิเสธเงียบๆ ไม่ update ค่าใน memory เลย
  const validate = () => {
    const n = v => Number(v);
    if (!Number.isFinite(n(form.workDayCutoffHour)) || n(form.workDayCutoffHour) < 0 || n(form.workDayCutoffHour) >= 24)
      return "เวลาตัดรอบวันทำงานต้องเป็นตัวเลข 0-23";
    if (!Number.isFinite(n(form.waitingUrgentMinutes)) || n(form.waitingUrgentMinutes) <= 0)
      return "นาทีที่ถึงจะขึ้น urgent ต้องมากกว่า 0";
    if (!Number.isFinite(n(form.maxPhotoUploads)) || n(form.maxPhotoUploads) < 1 || n(form.maxPhotoUploads) > 15)
      return "จำนวนรูปสูงสุดต้องอยู่ระหว่าง 1-15 รูป";
    if (!Number.isFinite(n(form.maxWaitingReasons)) || n(form.maxWaitingReasons) <= 0)
      return "จำนวนเหตุผลรอสินค้าสูงสุดต้องมากกว่า 0";
    if (!Number.isFinite(n(form.geofenceLat)) || !Number.isFinite(n(form.geofenceLng)) || String(form.geofenceLat).trim() === "" || String(form.geofenceLng).trim() === "")
      return "Geofence lat/lng ต้องเป็นตัวเลข ห้ามเว้นว่าง";
    if (!Number.isFinite(n(form.geofenceRadiusM)) || n(form.geofenceRadiusM) <= 0)
      return "รัศมี geofence ต้องมากกว่า 0";
    if (!form.facilityName.trim())
      return "กรุณากรอกชื่อโรงงาน";
    if (!form.siteTitle.trim())
      return "กรุณากรอกชื่อเว็บ";
    if (!Number.isFinite(n(form.unitPrice)) || n(form.unitPrice) < 0)
      return "ราคาต่อหน่วยต้องไม่ติดลบ";
    if (!Number.isFinite(n(form.vatRate)) || n(form.vatRate) < 0)
      return "อัตรา VAT ต้องไม่ติดลบ";
    if (!Number.isFinite(n(form.exitTimeWindowMinutes)) || n(form.exitTimeWindowMinutes) <= 0)
      return "ช่วงเวลานับถอยหลังก่อนออกโรงงานต้องมากกว่า 0";
    if (!Number.isFinite(n(form.docTimeMinutes)) || n(form.docTimeMinutes) < 0)
      return "เวลาทำเอกสารต้องไม่ติดลบ";
    return "";
  };

  const save = async () => {
    const err = validate();
    if (err) { setMsg(""); alert("บันทึกไม่ได้: " + err); return; }
    setSaving(true);
    setMsg("");
    try {
      await Promise.all([
        saveSetting("work_day_cutoff_hour",   Number(form.workDayCutoffHour)),
        saveSetting("waiting_urgent_minutes", Number(form.waitingUrgentMinutes)),
        saveSetting("max_photo_uploads",      Number(form.maxPhotoUploads)),
        saveSetting("max_waiting_reasons",    Number(form.maxWaitingReasons)),
        saveSetting("geofence", { lat: Number(form.geofenceLat), lng: Number(form.geofenceLng), radiusM: Number(form.geofenceRadiusM) }),
        saveSetting("facility_name",            form.facilityName),
        saveSetting("site_title",               form.siteTitle),
        saveSetting("unit_price",               Number(form.unitPrice)),
        saveSetting("vat_rate",                 Number(form.vatRate)),
        saveSetting("exit_time_window_minutes", Number(form.exitTimeWindowMinutes)),
        saveSetting("doc_time_minutes",         Number(form.docTimeMinutes)),
        saveSetting("excluded_customer_groups", form.excludedCustomerGroups.split(",").map(s => s.trim()).filter(Boolean)),
      ]);
      await recomputeQueueExitTimes(Number(form.docTimeMinutes));
      setMsg("✅ บันทึกการตั้งค่าสำเร็จ — มีผลทันทีในเซสชันนี้ เครื่องอื่นต้องรีเฟรชหน้าถึงจะเห็นค่าใหม่");
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setSaving(false);
    }
  };

  const lbl  = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp  = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="⚙️ ตั้งค่าระบบ">
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
        <div>
          <label style={lbl}>เวลาตัดรอบวันทำงาน (0-23)</label>
          <input type="number" min="0" max="23" value={form.workDayCutoffHour} onChange={set("workDayCutoffHour")} style={inp} />
        </div>
        <div>
          <label style={lbl}>รอเกินกี่นาทีถึงขึ้น "urgent"</label>
          <input type="number" min="1" value={form.waitingUrgentMinutes} onChange={set("waitingUrgentMinutes")} style={inp} />
        </div>
        <div>
          <label style={lbl}>จำนวนรูปสูงสุดต่อครั้ง (1-15)</label>
          <input type="number" min="1" max="15" value={form.maxPhotoUploads} onChange={set("maxPhotoUploads")} style={inp} />
        </div>
        <div>
          <label style={lbl}>จำนวนเหตุผลรอสินค้าสูงสุด</label>
          <input type="number" min="1" value={form.maxWaitingReasons} onChange={set("maxWaitingReasons")} style={inp} />
        </div>
      </div>
      <div style={{ fontWeight: 700, fontSize: 12, color: "#6b7280", margin: "10px 0 6px" }}>📍 Geofence เช็คอินคนขับ</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 12 }}>
        <div>
          <label style={lbl}>Latitude</label>
          <input type="number" step="any" value={form.geofenceLat} onChange={set("geofenceLat")} style={inp} />
        </div>
        <div>
          <label style={lbl}>Longitude</label>
          <input type="number" step="any" value={form.geofenceLng} onChange={set("geofenceLng")} style={inp} />
        </div>
        <div>
          <label style={lbl}>รัศมี (เมตร)</label>
          <input type="number" min="1" value={form.geofenceRadiusM} onChange={set("geofenceRadiusM")} style={inp} />
        </div>
      </div>
      <div style={{ fontWeight: 700, fontSize: 12, color: "#6b7280", margin: "10px 0 6px" }}>🏭 ข้อมูลโรงงาน / เอกสาร</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
        <div>
          <label style={lbl}>ชื่อโรงงาน (แสดงบนหัวจอและใบพิมพ์)</label>
          <input value={form.facilityName} onChange={set("facilityName")} style={inp} placeholder="เช่น โรงงานพระพุทธบาท" />
        </div>
        <div>
          <label style={lbl}>ชื่อเว็บ (แสดงบนแท็บ browser)</label>
          <div style={{ display: "flex" }}>
            <input value={form.siteTitle} onChange={set("siteTitle")} style={{ ...inp, borderRight: "none" }} placeholder="เช่น KK" />
            <div style={{ ...inp, borderLeft: "none", background: "#e5e7eb", color: "#9ca3af", flexShrink: 0, width: "auto", display: "flex", alignItems: "center" }}>Loading</div>
          </div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
        <div>
          <label style={lbl}>ราคาต่อหน่วย (บาท) — ใบ Invoice</label>
          <input type="number" min="0" step="any" value={form.unitPrice} onChange={set("unitPrice")} style={inp} />
        </div>
        <div>
          <label style={lbl}>อัตรา VAT (0-1, เช่น 0.07 = 7%)</label>
          <input type="number" min="0" step="0.01" value={form.vatRate} onChange={set("vatRate")} style={inp} />
        </div>
        <div>
          <label style={lbl}>ช่วงเวลานับถอยหลังก่อนออกโรงงาน (นาที)</label>
          <input type="number" min="1" value={form.exitTimeWindowMinutes} onChange={set("exitTimeWindowMinutes")} style={inp} />
        </div>
        <div>
          <label style={lbl}>เวลาทำเอกสารหลังโหลดเสร็จ (นาที)</label>
          <input type="number" min="0" value={form.docTimeMinutes} onChange={set("docTimeMinutes")} style={inp} />
        </div>
        <div>
          <label style={lbl}>กลุ่มลูกค้าที่ยกเว้นจาก Dashboard/คิว (คั่นด้วย ,)</label>
          <input value={form.excludedCustomerGroups} onChange={set("excludedCustomerGroups")} style={inp} placeholder="เช่น CPFTH" />
        </div>
      </div>
      {msg && <div style={{ padding: "8px 10px", background: "#d1fae5", color: "#065f46", fontSize: 12, fontWeight: 700, marginBottom: 10 }}>{msg}</div>}
      <button onClick={save} disabled={saving}
        style={{ width: "100%", background: saving ? "#e5e7eb" : "#111", color: saving ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "11px 0", fontWeight: 700, fontSize: 14, cursor: saving ? "default" : "pointer" }}>
        {saving ? "⏳ กำลังบันทึก..." : "💾 บันทึกการตั้งค่า"}
      </button>
    </Collapsible>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// LANE ALIASES (wh_lane_aliases) — เพิ่มชื่อเรียกลานแบบอื่นๆ ที่ Master ใช้
// ─────────────────────────────────────────────────────────────────────────────
const LaneAliasSettings = () => {
  const [rows, setRows] = useState(() => Object.entries(laneAliases).map(([alias, laneKey]) => ({ alias, laneKey })));
  const [newAlias, setNewAlias] = useState("");
  const [newLane, setNewLane] = useState(lanes[0].id);
  const [busy, setBusy] = useState(false);

  const refresh = () => setRows(Object.entries(laneAliases).map(([alias, laneKey]) => ({ alias, laneKey })));

  const add = async () => {
    setBusy(true);
    try {
      await saveLaneAlias(newAlias, newLane);
      setNewAlias("");
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(false);
    }
  };

  const remove = async (alias) => {
    if (!window.confirm(`ลบชื่อเรียกลาน "${alias}"?`)) return;
    try {
      await deleteLaneAlias(alias);
      refresh();
    } catch (e) {
      alert("ลบไม่สำเร็จ: " + e.message);
    }
  };

  const lbl  = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp  = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="🏷️ ชื่อเรียกลานอื่นๆ (Lane Aliases)">
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 12 }}>ใช้ตอนอัปโหลด Master แล้วคอลัมน์ "ลานโหลด" สะกดไม่ตรงกับที่ระบบรู้จัก — ไม่ต้องแก้โค้ด เพิ่มที่นี่ได้เลย</div>
      {rows.length > 0 && (
        <div style={{ marginBottom: 12 }}>
          {rows.map(r => (
            <div key={r.alias} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #f3f4f6", fontSize: 13 }}>
              <span>"{r.alias}" → {lanes.find(l => l.id === r.laneKey)?.tinyLabel || r.laneKey}</span>
              <button onClick={() => remove(r.alias)} style={{ background: "none", border: "none", color: "#dc2626", cursor: "pointer", fontSize: 13 }}>🗑️ ลบ</button>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr auto", gap: 8, alignItems: "end" }}>
        <div>
          <label style={lbl}>ชื่อเรียกลาน (ตามที่เจอในไฟล์)</label>
          <input value={newAlias} onChange={e => setNewAlias(e.target.value)} style={inp} placeholder="เช่น หัวเครื่องในหมู" />
        </div>
        <div>
          <label style={lbl}>คือลาน</label>
          <select value={newLane} onChange={e => setNewLane(e.target.value)} style={inp}>
            {lanes.map(l => <option key={l.id} value={l.id}>{l.tinyLabel}</option>)}
          </select>
        </div>
        <button onClick={add} disabled={busy}
          style={{ background: busy ? "#e5e7eb" : "#111", color: busy ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: busy ? "default" : "pointer", whiteSpace: "nowrap" }}>
          + เพิ่ม
        </button>
      </div>
    </Collapsible>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// WAITING REASONS (wh_waiting_reasons) — รายการเหตุผลรอสินค้า (fallback)
// ─────────────────────────────────────────────────────────────────────────────
const WaitingReasonSettings = () => {
  const [rows, setRows] = useState(() => [...waitingReasonPresets]);
  const [newLabel, setNewLabel] = useState("");
  const [busy, setBusy] = useState(false);

  const refresh = () => setRows([...waitingReasonPresets]);

  const add = async () => {
    setBusy(true);
    try {
      await addWaitingReason(newLabel);
      setNewLabel("");
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(false);
    }
  };

  const remove = async (row) => {
    if (!window.confirm(`ลบเหตุผล "${row.label}"?`)) return;
    try {
      await deleteWaitingReason(row.id);
      refresh();
    } catch (e) {
      alert("ลบไม่สำเร็จ: " + e.message);
    }
  };

  const lbl  = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp  = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="💬 รายการเหตุผลรอสินค้า (fallback)">
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 12 }}>ใช้เฉพาะตอนไฟล์ Master ไม่มีชื่อสินค้า match กับลานนั้นเลย — ปกติ dropdown จะดึงชื่อสินค้าจาก Master ก่อนเสมอ</div>
      {rows.length > 0 && (
        <div style={{ marginBottom: 12 }}>
          {rows.map(r => (
            <div key={r.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #f3f4f6", fontSize: 13 }}>
              <span>{r.label}</span>
              <button onClick={() => remove(r)} style={{ background: "none", border: "none", color: "#dc2626", cursor: "pointer", fontSize: 13 }}>🗑️ ลบ</button>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 8, alignItems: "end" }}>
        <div>
          <label style={lbl}>เหตุผลใหม่</label>
          <input value={newLabel} onChange={e => setNewLabel(e.target.value)} style={inp} placeholder="เช่น รอเบิกสินค้าจากคลัง" />
        </div>
        <button onClick={add} disabled={busy}
          style={{ background: busy ? "#e5e7eb" : "#111", color: busy ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: busy ? "default" : "pointer", whiteSpace: "nowrap" }}>
          + เพิ่ม
        </button>
      </div>
    </Collapsible>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// BASKET TYPES (wh_basket_types) — ประเภทตะกร้า/ตะขอ
// ─────────────────────────────────────────────────────────────────────────────
const BasketTypeSettings = () => {
  const [rows, setRows] = useState(() => [...basketTypes]);
  const [newKey, setNewKey] = useState("");
  const [newLabel, setNewLabel] = useState("");
  const [newCounts, setNewCounts] = useState(true);
  const [busy, setBusy] = useState(false);

  const refresh = () => setRows([...basketTypes]);

  const add = async () => {
    setBusy(true);
    try {
      await saveBasketType(newKey, newLabel, newCounts);
      setNewKey("");
      setNewLabel("");
      setNewCounts(true);
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(false);
    }
  };

  const remove = async (b) => {
    if (!window.confirm(`ลบประเภท "${b.label}"? ถ้าเคยมีรถบันทึกข้อมูลด้วยประเภทนี้ไว้ ข้อมูลจะยังอยู่ในฐานข้อมูลแต่จะไม่แสดงในหน้าเว็บอีก`)) return;
    try {
      await deleteBasketType(b.key);
      refresh();
    } catch (e) {
      alert("ลบไม่สำเร็จ: " + e.message);
    }
  };

  const lbl  = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp  = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="🧺 ประเภทตะกร้า/ตะขอ">
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 12 }}>4 ตัวแรก (เหลืองใหญ่/เล็ก, เทา, ตะขอ) เป็น default ในระบบ ลบไม่ได้ — เพิ่มประเภทใหม่ได้ที่นี่ แต่ห้ามเปลี่ยน/ลบ "รหัส (key)" ที่มีข้อมูลบันทึกไว้แล้ว</div>
      {rows.length > 0 && (
        <div style={{ marginBottom: 12 }}>
          {rows.map(b => (
            <div key={b.key} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #f3f4f6", fontSize: 13 }}>
              <span>{b.label} <span style={{ color: "#9ca3af", fontSize: 11 }}>({b.key}{b.countsInTotal ? ", นับรวมในรวมตะกร้า" : ""})</span></span>
              {!["yellowBig", "yellowSmall", "gray", "hooks"].includes(b.key) && (
                <button onClick={() => remove(b)} style={{ background: "none", border: "none", color: "#dc2626", cursor: "pointer", fontSize: 13 }}>🗑️ ลบ</button>
              )}
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1.5fr auto auto", gap: 8, alignItems: "end" }}>
        <div>
          <label style={lbl}>รหัส (key, ภาษาอังกฤษ ไม่มีเว้นวรรค)</label>
          <input value={newKey} onChange={e => setNewKey(e.target.value)} style={inp} placeholder="เช่น blueBig" />
        </div>
        <div>
          <label style={lbl}>ชื่อที่แสดง</label>
          <input value={newLabel} onChange={e => setNewLabel(e.target.value)} style={inp} placeholder="เช่น น้ำเงิน (ใหญ่)" />
        </div>
        <div>
          <label style={lbl}>นับรวม?</label>
          <select value={newCounts ? "1" : "0"} onChange={e => setNewCounts(e.target.value === "1")} style={inp}>
            <option value="1">นับรวม</option>
            <option value="0">ไม่นับรวม</option>
          </select>
        </div>
        <button onClick={add} disabled={busy}
          style={{ background: busy ? "#e5e7eb" : "#111", color: busy ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: busy ? "default" : "pointer", whiteSpace: "nowrap" }}>
          + เพิ่ม
        </button>
      </div>
    </Collapsible>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// DETAIL SOURCES (wh_detail_sources) — ช่องทาง PO + คอลัมน์ Excel ต่อช่องทาง
// ─────────────────────────────────────────────────────────────────────────────
const DetailSourceSettings = () => {
  const [rows, setRows] = useState(() => [...detailSources]);
  const dCols = defaultDetailCols();
  const emptyForm = { id: "", label: "", emoji: "📦", plateCol: String(dCols.plateCol), productCodeCol: String(dCols.productCodeCol), groupFlagCol: String(dCols.groupFlagCol), matchKeywords: "" };
  const [form, setForm] = useState(emptyForm);
  const [busy, setBusy] = useState(false);
  const DEFAULT_IDS = ["wet_market", "modern_trade", "others"];

  const refresh = () => setRows([...detailSources]);
  const set = (key) => (e) => setForm(f => ({ ...f, [key]: e.target.value }));

  const add = async () => {
    setBusy(true);
    try {
      await saveDetailSource(form.id, {
        label: form.label,
        emoji: form.emoji,
        color: "#6b7280",
        bg: "#f3f4f6",
        plateCol: Number(form.plateCol),
        productCodeCol: Number(form.productCodeCol),
        groupFlagCol: Number(form.groupFlagCol),
        matchKeywords: form.matchKeywords.split(",").map(k => k.trim()).filter(Boolean),
      });
      setForm(emptyForm);
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(false);
    }
  };

  const remove = async (s) => {
    if (!window.confirm(`ลบช่องทาง "${s.label}"? ไฟล์ที่เคยอัปโหลดไว้จะยังอยู่ในฐานข้อมูลแต่จะไม่แสดงในหน้าเว็บอีก`)) return;
    try {
      await deleteDetailSource(s.id);
      refresh();
    } catch (e) {
      alert("ลบไม่สำเร็จ: " + e.message);
    }
  };

  const lbl  = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp  = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="📦 ช่องทาง PO (Detail Sources)">
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 12 }}>
        3 ช่องทางแรก (ตลาดสด/Makro/LOTUS) เป็น default ในระบบ ลบไม่ได้ — เพิ่มช่องทางใหม่ได้ที่นี่ ถ้า retailer วางคอลัมน์ plate/รหัสสินค้า/กลุ่มลูกค้าในตำแหน่งไม่ตรงกับ default (65/20/11) ตั้งเลขคอลัมน์ (นับจาก 0) ให้ตรงได้เลย
      </div>
      {rows.length > 0 && (
        <div style={{ marginBottom: 12 }}>
          {rows.map(s => (
            <div key={s.id} style={{ padding: "6px 0", borderBottom: "1px solid #f3f4f6", fontSize: 13 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span>{s.emoji} {s.label} <span style={{ color: "#9ca3af", fontSize: 11 }}>({s.id})</span></span>
                {!DEFAULT_IDS.includes(s.id) && (
                  <button onClick={() => remove(s)} style={{ background: "none", border: "none", color: "#dc2626", cursor: "pointer", fontSize: 13 }}>🗑️ ลบ</button>
                )}
              </div>
              <div style={{ fontSize: 11, color: "#9ca3af" }}>
                คอลัมน์: plate={s.plateCol}, รหัสสินค้า={s.productCodeCol}, กลุ่มลูกค้า={s.groupFlagCol} · คำจับคู่: {(s.matchKeywords || []).join(", ") || "—"}
              </div>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
        <div>
          <label style={lbl}>รหัสช่องทาง (id, ภาษาอังกฤษ)</label>
          <input value={form.id} onChange={set("id")} style={inp} placeholder="เช่น makro" />
        </div>
        <div>
          <label style={lbl}>ชื่อที่แสดง</label>
          <input value={form.label} onChange={set("label")} style={inp} placeholder="เช่น Makro" />
        </div>
        <div>
          <label style={lbl}>Emoji</label>
          <input value={form.emoji} onChange={set("emoji")} style={inp} />
        </div>
        <div>
          <label style={lbl}>คำจับคู่กลุ่มลูกค้า (คั่นด้วย ,)</label>
          <input value={form.matchKeywords} onChange={set("matchKeywords")} style={inp} placeholder="เช่น makro" />
        </div>
        <div>
          <label style={lbl}>คอลัมน์ทะเบียนรถ (0-based)</label>
          <input type="number" value={form.plateCol} onChange={set("plateCol")} style={inp} />
        </div>
        <div>
          <label style={lbl}>คอลัมน์รหัสสินค้า (0-based)</label>
          <input type="number" value={form.productCodeCol} onChange={set("productCodeCol")} style={inp} />
        </div>
        <div>
          <label style={lbl}>คอลัมน์กลุ่มลูกค้า (0-based)</label>
          <input type="number" value={form.groupFlagCol} onChange={set("groupFlagCol")} style={inp} />
        </div>
      </div>
      <button onClick={add} disabled={busy}
        style={{ width: "100%", background: busy ? "#e5e7eb" : "#111", color: busy ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "9px 0", fontWeight: 700, fontSize: 13, cursor: busy ? "default" : "pointer" }}>
        + เพิ่มช่องทาง
      </button>
    </Collapsible>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// LANES (wh_lanes) — ป้ายชื่อ/สี/emoji ของลานโหลด
// ─────────────────────────────────────────────────────────────────────────────
const LaneSettings = () => {
  const [rows, setRows] = useState(() => [...lanes]);
  const [busy, setBusy] = useState(null);

  const refresh = () => setRows([...lanes]);
  const editField = (id, key) => (e) => setRows(rs => rs.map(r => r.id === id ? { ...r, [key]: e.target.value } : r));

  const save = async (row) => {
    setBusy(row.id);
    try {
      await saveLane(row.id, { label: row.label, shortLabel: row.shortLabel, tinyLabel: row.tinyLabel, emoji: row.emoji, color: row.color, bg: row.bg, border: row.border, sortOrder: row.sortOrder, enabled: row.enabled });
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(null);
    }
  };

  const toggleEnabled = async (row, v) => {
    setBusy(row.id);
    try {
      await saveLane(row.id, { label: row.label, shortLabel: row.shortLabel, tinyLabel: row.tinyLabel, emoji: row.emoji, color: row.color, bg: row.bg, border: row.border, sortOrder: row.sortOrder, enabled: v });
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(null);
    }
  };

  const lbl  = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp  = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="🏭 ลานโหลด (Lanes)">
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 12 }}>
        แก้ป้ายชื่อ/สี/emoji ของ 3 ลานได้ที่นี่ — รหัสลาน (id) เปลี่ยนไม่ได้เพราะผูกกับ QR/URL ที่พิมพ์ใช้งานอยู่แล้ว การเพิ่ม/ลดจำนวนลานต้องแก้โค้ดส่วน routing เพิ่มเติม ปิดใช้งานลานจะซ่อนลานนั้นจากเมนูเลือกลานของ QC/ลานโหลด/Checker (URL/QR เดิมยังเปิดเข้ามาได้แต่จะเจอข้อความแจ้งว่าลานปิดใช้งานอยู่)
      </div>
      {rows.map(r => (
        <div key={r.id} style={{ display: "grid", gridTemplateColumns: "auto 1fr 1fr 1fr 56px auto auto", gap: 8, alignItems: "end", padding: "10px 0", borderBottom: "1px solid #f3f4f6" }}>
          <div>
            <label style={lbl}>id</label>
            <div style={{ fontSize: 12, color: "#9ca3af", padding: "9px 0" }}>{r.id}</div>
          </div>
          <div>
            <label style={lbl}>ชื่อเต็ม</label>
            <input value={r.label} onChange={editField(r.id, "label")} style={inp} />
          </div>
          <div>
            <label style={lbl}>ชื่อสั้น</label>
            <input value={r.tinyLabel} onChange={editField(r.id, "tinyLabel")} style={inp} />
          </div>
          <div>
            <label style={lbl}>Emoji</label>
            <input value={r.emoji} onChange={editField(r.id, "emoji")} style={inp} />
          </div>
          <div>
            <label style={lbl}>สี</label>
            <input type="color" value={r.color} onChange={editField(r.id, "color")} style={{ ...inp, padding: 2, height: 38 }} />
          </div>
          <div>
            <label style={lbl}>ใช้งาน</label>
            <div style={{ padding: "9px 0" }}>
              <Switch checked={r.enabled !== false} onChange={(v) => toggleEnabled(r, v)} disabled={busy === r.id} />
            </div>
          </div>
          <button onClick={() => save(r)} disabled={busy === r.id}
            style={{ background: busy === r.id ? "#e5e7eb" : "#111", color: busy === r.id ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "9px 14px", fontWeight: 700, fontSize: 13, cursor: busy === r.id ? "default" : "pointer", whiteSpace: "nowrap" }}>
            💾 บันทึก
          </button>
        </div>
      ))}
    </Collapsible>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ROLES (wh_roles) — ป้ายชื่อ/emoji ของตำแหน่งงานในหน้าเลือกตำแหน่งงาน
// ─────────────────────────────────────────────────────────────────────────────
const RoleSettings = () => {
  const [rows, setRows] = useState(() => [...roles]);
  const [busy, setBusy] = useState(null);

  const refresh = () => setRows([...roles]);
  const editField = (id, key) => (e) => setRows(rs => rs.map(r => r.id === id ? { ...r, [key]: e.target.value } : r));

  const save = async (row) => {
    setBusy(row.id);
    try {
      await saveRole(row.id, { label: row.label, emoji: row.emoji, img: row.img });
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(null);
    }
  };

  const lbl  = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp  = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="🧑‍💼 ตำแหน่งงาน (Roles)">
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 12 }}>
        แก้ป้ายชื่อ/emoji ของตำแหน่งงานได้ที่นี่ — รหัสตำแหน่ง (id) เปลี่ยนไม่ได้และเพิ่ม/ลบไม่ได้ เพราะผูกกับสิทธิ์เข้าถึงเมนูของแต่ละตำแหน่งในโค้ดโดยตรง
      </div>
      {/* single shared grid (not one grid per row) so the "id"/label/emoji columns line up across every row */}
      <div style={{ display: "grid", gridTemplateColumns: "minmax(70px,auto) 1fr 90px auto", columnGap: 10, rowGap: 6, alignItems: "center" }}>
        <div style={lbl}>id</div>
        <div style={lbl}>ชื่อที่แสดง</div>
        <div style={lbl}>Emoji</div>
        <div />
        {rows.map(r => (
          <React.Fragment key={r.id}>
            <div style={{ fontSize: 12, color: "#9ca3af", borderTop: "1px solid #f3f4f6", padding: "10px 0" }}>{r.id}</div>
            <div style={{ borderTop: "1px solid #f3f4f6", padding: "6px 0" }}>
              <input value={r.label} onChange={editField(r.id, "label")} style={inp} />
            </div>
            <div style={{ borderTop: "1px solid #f3f4f6", padding: "6px 0" }}>
              <input value={r.emoji || ""} onChange={editField(r.id, "emoji")} style={inp} />
            </div>
            <div style={{ borderTop: "1px solid #f3f4f6", padding: "6px 0" }}>
              <button onClick={() => save(r)} disabled={busy === r.id}
                style={{ background: busy === r.id ? "#e5e7eb" : "#111", color: busy === r.id ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "9px 14px", fontWeight: 700, fontSize: 13, cursor: busy === r.id ? "default" : "pointer", whiteSpace: "nowrap" }}>
                💾 บันทึก
              </button>
            </div>
          </React.Fragment>
        ))}
      </div>
    </Collapsible>
  );
};

const Switch = ({ checked, onChange, disabled }) => (
  <button
    onClick={() => !disabled && onChange(!checked)}
    disabled={disabled}
    style={{
      width: 44, height: 24, borderRadius: 999, border: "none", padding: 2, flexShrink: 0,
      background: checked ? "#16a34a" : "#d1d5db", cursor: disabled ? "default" : "pointer",
      display: "flex", alignItems: "center", justifyContent: checked ? "flex-end" : "flex-start",
      transition: "background 0.15s",
    }}>
    <span style={{ width: 20, height: 20, borderRadius: "50%", background: "#fff", boxShadow: "0 1px 2px rgba(0,0,0,0.3)" }} />
  </button>
);

// ─────────────────────────────────────────────────────────────────────────────
// BAYS (wh_bays) — ช่องโหลดย่อยในแต่ละลาน + สวิตช์เปิด/ปิดการบังคับเลือกช่องโหลดทั้งระบบ
// ─────────────────────────────────────────────────────────────────────────────
const BaySettings = () => {
  const [rows, setRows] = useState(() => [...bays]);
  const [newLabel, setNewLabel] = useState(() => Object.fromEntries(lanes.map(l => [l.id, ""])));
  const [busy, setBusy] = useState(null);
  const [enabled, setEnabled] = useState(settings.enableBaySelection);
  const [toggling, setToggling] = useState(false);

  const refresh = () => setRows([...bays]);
  const editLabel = (id) => (e) => setRows(rs => rs.map(r => r.id === id ? { ...r, label: e.target.value } : r));

  const save = async (row) => {
    setBusy(row.id);
    try {
      await saveBay(row.id, row.label);
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      refresh(); // ทั้งสำเร็จและพัง ก็ sync กลับไปตามค่าจริงใน `bays` เสมอ กันช่องค้างข้อความที่ไม่ได้บันทึกจริง
      setBusy(null);
    }
  };

  const add = async (laneId) => {
    const label = (newLabel[laneId] || "").trim();
    setBusy(`add_${laneId}`);
    try {
      await addBay(laneId, label);
      setNewLabel(nl => ({ ...nl, [laneId]: "" }));
      refresh();
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(null);
    }
  };

  const remove = async (row) => {
    if (!window.confirm(`ลบ "${row.label}"?`)) return;
    setBusy(row.id);
    try {
      await deleteBay(row.id);
      refresh();
    } catch (e) {
      alert("ลบไม่สำเร็จ: " + e.message);
    } finally {
      setBusy(null);
    }
  };

  const toggleEnabled = async (v) => {
    setToggling(true);
    try {
      await saveSetting("enable_bay_selection", v);
      setEnabled(v);
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setToggling(false);
    }
  };

  const inp = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <Collapsible title="🚪 ช่องโหลด">
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 0 16px", borderBottom: "1px solid #f3f4f6", marginBottom: 14 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#111" }}>เปิดใช้งานการเลือกช่องโหลด</div>
          <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>ปิดแล้วหน้า ลานโหลด/QC/Checker จะข้ามหน้าเลือกช่องโหลดไปเข้าฟอร์มเลย ไม่บันทึกช่องโหลดลง DB</div>
        </div>
        <Switch checked={enabled} onChange={toggleEnabled} disabled={toggling} />
      </div>
      {lanes.map(lane => {
        const laneBays = rows.filter(b => b.laneId === lane.id).sort((a, b) => a.sortOrder - b.sortOrder);
        return (
          <div key={lane.id} style={{ marginBottom: 18 }}>
            <div style={{ fontWeight: 700, fontSize: 12, color: lane.color, marginBottom: 8 }}>{lane.tinyLabel} ({laneBays.length} ช่องโหลด)</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>
              {laneBays.map(r => (
                <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 4, border: "1.5px solid #e5e7eb", borderRadius: 0, padding: "6px 6px 6px 10px", background: "#fff" }}>
                  <input value={r.label} onChange={editLabel(r.id)} onBlur={() => { if (r.label !== bays.find(b => b.id === r.id)?.label) save(r); }} disabled={busy === r.id}
                    style={{ border: "none", outline: "none", fontSize: 13, fontWeight: 700, width: 140, background: "transparent" }} />
                  <button onClick={() => remove(r)} disabled={busy === r.id}
                    style={{ background: "none", border: "none", color: busy === r.id ? "#e5e7eb" : "#dc2626", cursor: busy === r.id ? "default" : "pointer", fontSize: 16, lineHeight: 1, padding: "2px 4px" }}>
                    ×
                  </button>
                </div>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 8 }}>
              <input value={newLabel[lane.id] || ""} onChange={e => setNewLabel(nl => ({ ...nl, [lane.id]: e.target.value }))} style={inp} placeholder={`เช่น ช่องโหลด ${laneBays.length + 1}`} />
              <button onClick={() => add(lane.id)} disabled={busy === `add_${lane.id}`}
                style={{ background: busy === `add_${lane.id}` ? "#e5e7eb" : "#111", color: busy === `add_${lane.id}` ? "#9ca3af" : "#fff", border: "none", borderRadius: 0, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: busy === `add_${lane.id}` ? "default" : "pointer", whiteSpace: "nowrap" }}>
                + เพิ่มช่อง
              </button>
            </div>
          </div>
        );
      })}
    </Collapsible>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// ADMIN
// ─────────────────────────────────────────────────────────────────────────────
const Admin = ({ trucks, queue, onUpdate, onDeleteTruck }) => {
  const [selId,   setSelId]   = useState("");
  const [form,    setForm]    = useState(null);
  const [msg,     setMsg]     = useState("");
  const [mergeId, setMergeId] = useState("");

  const truck  = trucks.find(t => t.id === selId);

  // unique values from queue entries matching this truck's plate
  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const matchedQueue   = truck ? queue.filter(q => plateNum(q.plate) === plateNum(truck.plate)) : [];
  const duplicateTrucks = truck ? trucks.filter(t => t.id !== selId && plateNum(t.plate) === plateNum(truck.plate)) : [];
  const queueZones  = [...new Set(matchedQueue.map(q => q.zone).filter(Boolean))];
  const queueGroups = [...new Set(matchedQueue.map(q => q.customerGroup).filter(Boolean))];

  useEffect(() => {
    if (!truck) { setForm(null); setMsg(""); return; }
    setForm({
      queueId:       truck.queueId       || "",
      customerGroup: truck.customerGroup || "",
      zone:          truck.zone          || "",
      entryTime:     truck.entryTime     || "",
      exitTime:      truck.exitTime      || "",
      status:        truck.status        || "arrived",
      qcLanes:       JSON.parse(JSON.stringify(truck.qcLanes   || {})),
      loadLanes:     JSON.parse(JSON.stringify(truck.loadLanes || {})),
    });
    setMsg("");
  }, [selId]);

  const linkQueue = (qid) => {
    const q = matchedQueue.find(q => q.id === qid);
    if (!q) return;
    setForm(f => ({ ...f, queueId: q.id, zone: q.zone || "", customerGroup: q.customerGroup || "", entryTime: q.entryTime || "", exitTime: q.exitTime || "" }));
  };

  const [saving, setSaving] = useState(false);
  const save = async () => {
    setSaving(true);
    try {
      await onUpdate(selId, form);
      setMsg("✅ บันทึกแล้ว");
      setTimeout(() => setMsg(""), 2500);
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setSaving(false);
    }
  };

  const [merging, setMerging] = useState(false);
  const handleMerge = async () => {
    const src = trucks.find(t => t.id === mergeId);
    if (!src || !truck) return;
    if (!window.confirm(`Merge ข้อมูลจาก ${src.plate} (zone: ${src.zone || "–"}) เข้า ${truck.plate} (zone: ${truck.zone || "–"}) แล้วลบรถต้นทางออก?`)) return;
    setMerging(true);
    try {
      // merge lane-by-lane: T-1 (target) มีสิทธิ์ก่อน ถ้า T-1 ไม่มีค่อยเอาของ T-2
      // (เดิมลืม sampleLanes — ข้อมูลสุ่มตรวจของรถต้นทางเลยหายไปเงียบๆ ตอน merge ทั้งที่ยัง
      // ไม่ได้ถูกคัดลอกไปไหน ทั้งที่ตัวรถต้นทางถูกลบทิ้ง (ตอนนี้ soft-delete แล้วก็จริง แต่ merge
      // เองก็ควรพาข้อมูลไปครบตั้งแต่ต้น ไม่ใช่พึ่งการกู้คืนทีหลัง)
      const mergedQC     = { ...(src.qcLanes     || {}) };
      const mergedLoad   = { ...(src.loadLanes   || {}) };
      const mergedSample = { ...(src.sampleLanes || {}) };
      for (const l of lanes) {
        if (truck.qcLanes?.[l.id]?.done)     mergedQC[l.id]     = truck.qcLanes[l.id];
        if (truck.loadLanes?.[l.id]?.done)   mergedLoad[l.id]   = truck.loadLanes[l.id];
        if (truck.sampleLanes?.[l.id]?.done) mergedSample[l.id] = truck.sampleLanes[l.id];
      }
      await onUpdate(selId, { qcLanes: mergedQC, loadLanes: mergedLoad, sampleLanes: mergedSample });
      await onDeleteTruck(mergeId);
      setMergeId("");
      setMsg("✅ Merge สำเร็จ — ลบรถซ้ำแล้ว");
      setTimeout(() => setMsg(""), 3000);
    } catch (e) {
      alert("Merge ไม่สำเร็จ: " + e.message);
    } finally {
      setMerging(false);
    }
  };

  const setQC   = (lid, key, val) => setForm(f => ({ ...f, qcLanes:   { ...f.qcLanes,   [lid]: { ...(f.qcLanes[lid]   || {}), [key]: val } } }));
  const setLoad = (lid, key, val) => setForm(f => ({ ...f, loadLanes: { ...f.loadLanes, [lid]: { ...(f.loadLanes[lid] || {}), [key]: val } } }));
  const resetQC   = (lid) => setForm(f => ({ ...f, qcLanes:   { ...f.qcLanes,   [lid]: {} } }));
  const resetLoad = (lid) => setForm(f => ({ ...f, loadLanes: { ...f.loadLanes, [lid]: {} } }));

  const card  = { background: "#fff", borderRadius: 0, padding: "16px 20px", marginBottom: 14, boxShadow: "0 1px 6px rgba(0,0,0,0.07)" };
  const lbl   = { display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 };
  const inp   = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none" };

  return (
    <div style={{ maxWidth: 640, margin: "0 auto" }}>
      <h2 style={{ fontWeight: 900, fontSize: 22, marginBottom: 16 }}>⚙️ Admin — แก้ไขข้อมูล</h2>

      {/* Truck selector */}
      <div style={card}>
        <label style={lbl}>เลือกทะเบียนรถ</label>
        <select value={selId} onChange={e => setSelId(e.target.value)} style={{ ...inp, fontSize: 15 }}>
          <option value="">— เลือกรถ —</option>
          {trucks.map(t => (
            <option key={t.id} value={t.id}>{t.plate}</option>
          ))}
        </select>
      </div>

      {/* Merge — แสดงเฉพาะเมื่อมีรถ plate เดียวกันในระบบ */}
      {truck && duplicateTrucks.length > 0 && (
        <div style={{ ...card, border: "1.5px solid #fde047", background: "#fefce8" }}>
          <div style={{ fontWeight: 800, fontSize: 14, marginBottom: 10 }}>🔀 Merge รถซ้ำ</div>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 10 }}>
            พบทะเบียน <b>{truck.plate}</b> ในระบบ {duplicateTrucks.length + 1} คัน — เลือกรถที่จะดูดข้อมูล (source) เข้ามารวมกับตัวนี้ แล้วรถ source จะถูกลบ
          </div>
          <select value={mergeId} onChange={e => setMergeId(e.target.value)} style={{ ...inp, marginBottom: 10 }}>
            <option value="">— เลือกรถ source —</option>
            {duplicateTrucks.map(t => (
              <option key={t.id} value={t.id}>
                {t.plate} · zone: {t.zone || "–"} · {STATUS_META[t.status]?.label || t.status}
                {t.qcLanes ? ` · QC: ${lanes.filter(l => t.qcLanes[l.id]?.done).map(l => l.tinyLabel).join(", ") || "ยังไม่มี"}` : ""}
              </option>
            ))}
          </select>
          <button onClick={handleMerge} disabled={!mergeId || merging}
            style={{ width: "100%", background: mergeId ? "#854d0e" : "#e5e7eb", color: mergeId ? "#fff" : "#9ca3af", border: "none", borderRadius: 0, padding: "11px 0", fontSize: 14, fontWeight: 700, cursor: mergeId ? "pointer" : "default" }}>
            {merging ? "⏳ กำลัง Merge..." : "🔀 Merge และลบรถ source"}
          </button>
        </div>
      )}

      {truck && form && (
        <>
          {/* ผูก Queue Entry */}
          <div style={card}>
            <label style={lbl}>🔗 ผูก Queue Entry (เปลี่ยนข้อมูลจาก LG)</label>
            <select value={form.queueId} onChange={e => linkQueue(e.target.value)} style={{ ...inp, fontSize: 13 }}>
              <option value="">— เลือก Queue Entry —</option>
              {matchedQueue.map(q => (
                <option key={q.id} value={q.id}>
                  {q.zone || "–"} · {q.customerGroup || "–"} · เข้า {q.entryTime || "–"} · ออก {q.exitTime || "–"}
                </option>
              ))}
            </select>
            {form.queueId && (
              <div style={{ fontSize: 11, color: "#6b7280", marginTop: 6 }}>
                ผูกอยู่กับ: Zone <b>{form.zone || "–"}</b> · กลุ่ม <b>{form.customerGroup || "–"}</b> · ออก <b>{form.exitTime || "–"}</b>
              </div>
            )}
          </div>

          {/* กลุ่มลูกค้า + Zone */}
          <div style={card}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <div>
                <label style={lbl}>👥 กลุ่มลูกค้า</label>
                <select value={form.customerGroup} onChange={e => setForm(f => ({ ...f, customerGroup: e.target.value }))} style={inp}>
                  <option value="">— ไม่ระบุ —</option>
                  {queueGroups.map(g => <option key={g} value={g}>{g}</option>)}
                </select>
              </div>
              <div>
                <label style={lbl}>📍 Zone</label>
                <select value={form.zone} onChange={e => setForm(f => ({ ...f, zone: e.target.value }))} style={inp}>
                  <option value="">— ไม่ระบุ —</option>
                  {queueZones.map(z => <option key={z} value={z}>{z}</option>)}
                </select>
              </div>
            </div>
          </div>

          {/* Status */}
          <div style={card}>
            <label style={lbl}>🚦 สถานะ</label>
            <select value={form.status} onChange={e => setForm(f => ({ ...f, status: e.target.value }))} style={inp}>
              {Object.entries(STATUS_META).map(([k, v]) => (
                <option key={k} value={k}>{v.label}</option>
              ))}
            </select>
          </div>

          {/* QC Lanes */}
          <div style={card}>
            <div style={{ fontWeight: 800, fontSize: 14, marginBottom: 12 }}>🌡️ QC ลานโหลด</div>
            {lanes.map(l => {
              const qc = form.qcLanes[l.id] || {};
              return (
                <div key={l.id} style={{ borderRadius: 0, border: `1.5px solid ${l.border}`, background: l.bg, padding: "12px 14px", marginBottom: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                    <div style={{ fontWeight: 700, fontSize: 13, color: l.color }}>{l.tinyLabel}</div>
                    <button onClick={() => resetQC(l.id)}
                      style={{ background: "#fee2e2", color: "#991b1b", border: "none", borderRadius: 0, padding: "3px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                      ↩ Reset QC
                    </button>
                  </div>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                    <div style={{ flex: "0 0 auto" }}>
                      <label style={{ ...lbl, marginBottom: 4 }}>อุณหภูมิ (°C)</label>
                      <input value={qc.temp || ""} onChange={e => setQC(l.id, "temp", e.target.value)}
                        style={{ ...inp, width: 110 }} placeholder="เช่น -18" />
                    </div>
                    <label style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, fontSize: 13, cursor: "pointer", marginTop: 18 }}>
                      <input type="checkbox" checked={!!qc.done} onChange={e => setQC(l.id, "done", e.target.checked)} style={{ width: 16, height: 16 }} />
                      QC แล้ว
                    </label>
                  </div>
                  {qc.doneAt && <div style={{ fontSize: 11, color: "#6b7280", marginTop: 6 }}>บันทึกเมื่อ: {qc.doneAt}</div>}
                </div>
              );
            })}
          </div>

          {/* Load Lanes */}
          <div style={card}>
            <div style={{ fontWeight: 800, fontSize: 14, marginBottom: 12 }}>🏗️ สถานะลานโหลด</div>
            {lanes.map(l => {
              const ld = form.loadLanes[l.id] || {};
              return (
                <div key={l.id} style={{ borderRadius: 0, border: `1.5px solid ${l.border}`, background: l.bg, padding: "12px 14px", marginBottom: 10 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                    <div style={{ fontWeight: 700, fontSize: 13, color: l.color }}>{l.tinyLabel}</div>
                    <button onClick={() => resetLoad(l.id)}
                      style={{ background: "#fee2e2", color: "#991b1b", border: "none", borderRadius: 0, padding: "3px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                      ↩ Reset
                    </button>
                  </div>
                  <div style={{ display: "flex", gap: 20 }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={!!ld.waiting} onChange={e => setLoad(l.id, "waiting", e.target.checked)} style={{ width: 16, height: 16 }} />
                      รอสินค้า
                    </label>
                    <label style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={!!ld.done} onChange={e => setLoad(l.id, "done", e.target.checked)} style={{ width: 16, height: 16 }} />
                      โหลดเสร็จ
                    </label>
                  </div>
                  {(ld.waitingAt || ld.doneAt) && (
                    <div style={{ fontSize: 11, color: "#6b7280", marginTop: 6 }}>
                      {ld.waitingAt && <span>รอ: {ld.waitingAt} </span>}
                      {ld.doneAt   && <span>เสร็จ: {ld.doneAt}</span>}
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          {msg && <div style={{ textAlign: "center", color: "#10b981", fontWeight: 700, marginBottom: 12, fontSize: 15 }}>{msg}</div>}
          <button onClick={save} disabled={saving}
            style={{ width: "100%", background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "14px 0", fontSize: 16, fontWeight: 700, cursor: "pointer", marginBottom: 10 }}>
            {saving ? "⏳ กำลังบันทึก..." : "💾 บันทึกการแก้ไข"}
          </button>
          <button onClick={() => { if (window.confirm(`ลบรถ ${truck.plate} ออกจากระบบ?`)) onDeleteTruck(truck.id); }}
            style={{ width: "100%", background: "#fee2e2", color: "#991b1b", border: "1.5px solid #fca5a5", borderRadius: 0, padding: "12px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
            🗑️ ลบรถออกจากระบบ (กรณีสแกนทะเบียนผิด)
          </button>
        </>
      )}
    </div>
  );
};

// ─────────────────────────────────────────────────────────────────────────────
// DETAIL LOADING (channels/columns มาจาก wh_detail_sources — ดู src/lib/masterData.js)
// ─────────────────────────────────────────────────────────────────────────────

// Map Thai lane names from Master file → lane IDs used in system — default/fallback
// ที่ยังใช้งานได้แม้ยังไม่มีข้อมูลใน wh_lane_aliases เลย ส่วน alias ใหม่ๆ ที่เจอทีหลัง
// (สะกดต่างออกไป) ให้เพิ่มผ่านหน้า Admin แทน ไม่ต้องแก้โค้ดนี้
const LANE_NAME_MAP = {
  "ชิ้นส่วน":        "lane_parts",
  "หัว/เครื่องใน":   "lane_head",
  "หัวเครื่องใน":    "lane_head",
  "หมูซีก":          "lane_pork",
  // fallback: if value already is a lane ID, pass through
  "lane_parts":       "lane_parts",
  "lane_head":        "lane_head",
  "lane_pork":        "lane_pork",
};
const normalizeLaneKey = (raw) => {
  const t = String(raw || "").trim();
  if (!t) return null;
  if (laneAliases[t]) return laneAliases[t];
  if (LANE_NAME_MAP[t]) return LANE_NAME_MAP[t];
  // partial match fallback
  if (t.includes("ชิ้นส่วน")) return "lane_parts";
  if (t.includes("หัว") || t.includes("เครื่องใน")) return "lane_head";
  if (t.includes("หมูซีก") || t.includes("ซีก") || t.includes("หมู")) return "lane_pork";
  return null;
};
const normalizeProductCode = (val) => String(val || "").replace(/\.0+$/, "").trim().replace(/^0+(\d)/, "$1");

// joins order rows (plate, productCode) against the master lane list → plate→Set(laneKey)
const buildDetailMap = (masterLane, allDetailRows) => {
  const map = {};
  for (const row of allDetailRows) {
    const rowCode = normalizeProductCode(row.productCode);
    const match = masterLane.find(m => normalizeProductCode(m.productCode) === rowCode);
    if (!match) continue;
    const k = String(row.plate).replace(/\s/g, "").toUpperCase();
    if (!map[k]) map[k] = new Set();
    map[k].add(match.laneKey);
  }
  return map;
};

// one detailMap per PO channel, so a plate's lanes never leak across channels
const buildDetailMapByChannel = (masterLane, rowsByChannel) => {
  const result = {};
  for (const src of detailSources) result[src.id] = buildDetailMap(masterLane, rowsByChannel[src.id] || []);
  return result;
};

// LG's "กลุ่มลูกค้า" is a comma-separated list of sub-brand tags, e.g.
// "MT-Lotus-CPFM,MT-LotusB2C,Wetmarket" — split & classify each token to a PO channel.
// Unrecognized tags (CPFTH, FARM, ...) are simply ignored, not one of the known channels.
// matchKeywords ต่อช่องทางมาจาก wh_detail_sources (ดู src/lib/masterData.js) — เพิ่ม
// retailer ใหม่หรือคำที่ใช้จับคู่เพิ่มได้ผ่านหน้า Admin ไม่ต้องแก้โค้ดนี้
const normalizeChannels = (customerGroup) => {
  const channels = new Set();
  for (const raw of String(customerGroup || "").split(",")) {
    const t = raw.trim().toLowerCase();
    if (!t) continue;
    const src = detailSources.find(s => (s.matchKeywords || []).some(kw => t.includes(kw.toLowerCase())));
    if (src) channels.add(src.id);
  }
  return channels;
};

// lanes a specific truck round needs, scoped to only the PO channel(s) its LG
// "กลุ่มลูกค้า" says it's running — a plate that does 2 rounds for 2 different
// customers in one day must not have both rounds' lanes merged together.
// when กลุ่มลูกค้า is blank/unrecognized (no channel to scope to), fall back to
// checking every channel — otherwise plates with no customerGroup set never match at all
const laneMatchForTruck = (truck, detailMapByChannel) => {
  const num = (String(truck?.plate).match(/\d+/g) || []).pop() || "";
  if (!num) return new Set();
  const channels = normalizeChannels(truck?.customerGroup);
  const searchChannels = channels.size ? [...channels] : Object.keys(detailMapByChannel);
  const lanes = new Set();
  for (const ch of searchChannels) {
    const map = detailMapByChannel[ch] || {};
    const matched = Object.entries(map).find(([k]) => (String(k).match(/\d+/g) || []).pop() === num);
    if (matched) matched[1].forEach(l => lanes.add(l));
  }
  return lanes;
};

const DETAIL_LS_VER = "v2"; // bump when encoding/format changes — forces re-upload

const DetailLoading = ({ masterLane, onDetailChange }) => {
  // cache แคชไว้ต่อ "วันทำงาน" (คัดตาม cycleDateStr, ตัดรอบ 10:00 น. เดียวกับที่อื่น) —
  // ไม่งั้นถ้ายังไม่มีใครอัปโหลดไฟล์ PO ของวันใหม่ หน้านี้จะดึง localStorage ของเมื่อวานมาโชว์
  // ปนกับของวันนี้แทนที่จะว่างรอไฟล์ใหม่
  const markDetailCacheFresh = () => { try { localStorage.setItem("wh_detail_date", cycleDateStr()); } catch {} };
  const initSrc = () => {
    try {
      const stale = localStorage.getItem("wh_detail_ver") !== DETAIL_LS_VER || localStorage.getItem("wh_detail_date") !== cycleDateStr();
      if (stale) {
        ["wh_detail_src", "wh_detail_names"].forEach(k => localStorage.removeItem(k));
        localStorage.setItem("wh_detail_ver", DETAIL_LS_VER);
        markDetailCacheFresh();
        return {};
      }
      return JSON.parse(localStorage.getItem("wh_detail_src") || "{}");
    } catch { return {}; }
  };
  const initNames = () => {
    try { return JSON.parse(localStorage.getItem("wh_detail_names") || "{}"); } catch { return {}; }
  };

  const [srcData,     setSrcData]     = useState(() => ({ ...Object.fromEntries(detailSources.map(s => [s.id, []])), ...initSrc() }));
  const [fileNames,   setFileNames]   = useState(initNames);
  const [showDebug,   setShowDebug]   = useState(false);

  // On mount: fetch from Supabase (shared), fallback to localStorage
  useEffect(() => {
    fetchDetailSrc().then(remote => {
      if (remote) {
        const merged = {};
        const names = { ...initNames() };
        for (const [srcId, { rows, fileName }] of Object.entries(remote)) {
          merged[srcId] = rows;
          if (fileName) names[srcId] = fileName;
        }
        markDetailCacheFresh();
        setSrcData(prev => {
          const next = { ...prev, ...merged };
          localStorage.setItem("wh_detail_src", JSON.stringify(next));
          return next;
        });
        setFileNames(prev => {
          const next = { ...prev, ...names };
          localStorage.setItem("wh_detail_names", JSON.stringify(next));
          return next;
        });
        Object.entries(merged).forEach(([srcId, rows]) => {
          if (rows?.length) onDetailChange(srcId, rows);
        });
      } else {
        const stored = initSrc();
        Object.entries(stored).forEach(([srcId, rows]) => {
          if (rows?.length) onDetailChange(srcId, rows);
        });
      }
    });

    const channel = supabase.channel("detail-src-sync")
      .on("postgres_changes", { event: "*", schema: "public", table: "wh_master" }, async (payload) => {
        const row = payload.new;
        if (!row?.id) return;
        const today = cycleDateStr();
        const src = detailSources.find(s => `detail_${s.id}_${today}` === row.id);
        if (!src) return;
        const srcId = src.id;
        // Re-fetch แทนการอ่านจาก payload เพราะ JSONB ขนาดใหญ่อาจถูกตัดใน realtime
        const { data: fresh } = await supabase.from("wh_master").select("data").eq("id", row.id).single();
        const payload2 = fresh?.data || row.data || {};
        const rows = Array.isArray(payload2) ? payload2 : (payload2.rows || []);
        const fileName = payload2.file_name || "";
        markDetailCacheFresh();
        setSrcData(prev => {
          const next = { ...prev, [srcId]: rows };
          localStorage.setItem("wh_detail_src", JSON.stringify(next));
          return next;
        });
        if (fileName) {
          setFileNames(prev => {
            const next = { ...prev, [srcId]: fileName };
            localStorage.setItem("wh_detail_names", JSON.stringify(next));
            return next;
          });
        }
        onDetailChange(srcId, rows);
      })
      .subscribe();

    return () => supabase.removeChannel(channel);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Parse source file — ตำแหน่งคอลัมน์ (plateCol/productCodeCol/groupFlagCol) มาจาก
  // wh_detail_sources ต่อช่องทาง (แต่ละ retailer วางคอลัมน์ไม่ตรงกันได้ ปรับได้ที่ Admin)
  const parseSourceFile = (file, srcId) => {
    const src = detailSources.find(s => s.id === srcId) || {};
    const dCols = defaultDetailCols();
    const plateCol = src.plateCol ?? dCols.plateCol, productCodeCol = src.productCodeCol ?? dCols.productCodeCol, groupFlagCol = src.groupFlagCol ?? dCols.groupFlagCol;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      let parsed;
      try {
        let wb;
        if (/\.csv$/i.test(file.name)) {
          // CSV with Thai Windows-874 encoding: decode bytes manually
          const text = new TextDecoder("windows-874").decode(new Uint8Array(ev.target.result));
          wb = XLSX.read(text, { type: "string" });
        } else {
          wb = XLSX.read(ev.target.result, { type: "array", raw: true, codepage: 874 });
        }
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
        const dataRows = rows.slice(1).filter(r => r[plateCol] && String(r[plateCol]).trim() !== "");
        parsed = dataRows.map(r => ({
          plate:        String(r[plateCol] || "").trim(),
          productCode:  normalizeProductCode(r[productCodeCol]),
          groupFlag:    String(r[groupFlagCol] || "").trim(),
        })).filter(r => r.plate && r.productCode);

        markDetailCacheFresh();
        setSrcData(prev => {
          const next = { ...prev, [srcId]: parsed };
          localStorage.setItem("wh_detail_src", JSON.stringify(next));
          return next;
        });
        setFileNames(prev => {
          const next = { ...prev, [srcId]: file.name };
          localStorage.setItem("wh_detail_names", JSON.stringify(next));
          return next;
        });
        onDetailChange(srcId, parsed);
      } catch(e) {
        alert("อ่านไฟล์ไม่สำเร็จ: " + e.message);
        return;
      }

      try {
        const { error } = await supabase.from("wh_master").upsert({ id: `detail_${srcId}_${cycleDateStr()}`, data: { rows: parsed, file_name: file.name } });
        if (error) throw error;
      } catch (e) {
        alert("บันทึกขึ้นเซิร์ฟเวอร์ไม่สำเร็จ — เครื่องอื่นจะไม่เห็นข้อมูลนี้: " + e.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Compute plate → lanes mapping from all sources + master
  // memoized: this is an O(rows × masterLane) nested scan (thousands × ~350), and
  // without useMemo it re-ran on every render — including the app's own 15s clock
  // tick re-rendering whichever tab is active — which is what made this page feel
  // stuck compared to others
  const allDetail = useMemo(
    () => Object.values(srcData).flat(),
    [srcData]
  );
  const plateLaneMap = useMemo(() => {
    const map = {};
    for (const row of allDetail) {
      const rowCode = normalizeProductCode(row.productCode);
      const match = (masterLane || []).find(m => normalizeProductCode(m.productCode) === rowCode);
      if (!match) continue;
      const plateKey = String(row.plate).replace(/\s/g, "").toUpperCase();
      if (!map[plateKey]) map[plateKey] = new Set();
      map[plateKey].add(match.laneKey);
    }
    return map;
  }, [allDetail, masterLane]);

  return (
    <div style={{ padding: 20 }}>
      <h2 style={{ margin: "0 0 20px", fontSize: 20, fontWeight: 900 }}>📋 Detail Loading</h2>

      {/* Source upload buttons (1-3) */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16, marginBottom: 24 }}>
        {detailSources.map(src => (
          <div key={src.id} style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", padding: 20 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
              <span style={{ fontSize: 28 }}>{src.emoji}</span>
              <div>
                <div style={{ fontWeight: 800, fontSize: 15 }}>{src.label}</div>
                {fileNames[src.id] && (
                  <div style={{ fontSize: 11, color: src.color, fontWeight: 700, marginTop: 2 }}>
                    ✅ {fileNames[src.id]}
                  </div>
                )}
              </div>
            </div>
            <label style={{ display: "block", background: src.bg, color: src.color, border: `1.5px dashed ${src.color}`, borderRadius: 0, padding: "12px 0", textAlign: "center", fontSize: 13, fontWeight: 700, cursor: "pointer", transition: "opacity 0.2s" }}
              onMouseOver={e => e.currentTarget.style.opacity = "0.8"}
              onMouseOut={e => e.currentTarget.style.opacity = "1"}>
              {fileNames[src.id] ? "🔄 เปลี่ยนไฟล์" : "⬆️ อัปโหลดไฟล์"} {src.label}
              <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
                onChange={e => { if (e.target.files[0]) parseSourceFile(e.target.files[0], src.id); e.target.value = ""; }} />
            </label>
          </div>
        ))}
      </div>

      {/* Result table: plate → lanes */}
      {Object.keys(plateLaneMap).length > 0 && (
        <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", overflow: "hidden" }}>
          <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 700, fontSize: 14 }}>
            📊 สรุป ทะเบียนรถ → ลานโหลด
            <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 8 }}>{Object.keys(plateLaneMap).length} คัน</span>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#f9fafb" }}>
                  {["ทะเบียนรถ", ...lanes.map(l => l.tinyLabel)].map(h => (
                    <th key={h} style={{ padding: "10px 16px", textAlign: "left", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.entries(plateLaneMap).map(([plate, plateLanes]) => (
                  <tr key={plate} style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "8px 16px", fontWeight: 800, fontFamily: "monospace" }}>{plate}</td>
                    {lanes.map(l => (
                      <td key={l.id} style={{ padding: "8px 16px", textAlign: "center" }}>
                        {plateLanes.has(l.id)
                          ? <span style={{ color: l.color, fontWeight: 800, fontSize: 16 }}>✓</span>
                          : <span style={{ color: "#e5e7eb", fontSize: 14 }}>—</span>
                        }
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {Object.keys(plateLaneMap).length === 0 && (masterLane || []).length > 0 && allDetail.length > 0 && (
        <div style={{ background: "#fffbeb", border: "1.5px solid #fde68a", borderRadius: 0, padding: 20, color: "#92400e", fontWeight: 600, fontSize: 14 }}>
          ⚠️ ไม่พบรหัสสินค้าที่ Match กับ Master — ตรวจสอบว่า Product Code ในไฟล์ตรงกับ Master หรือไม่
          <button onClick={() => setShowDebug(v => !v)} style={{ marginLeft: 12, background: "#92400e", color: "#fff", border: "none", borderRadius: 0, padding: "4px 12px", fontSize: 12, cursor: "pointer", fontWeight: 700 }}>
            {showDebug ? "ซ่อน" : "🔍 ดู Debug"}
          </button>
        </div>
      )}

      {showDebug && (masterLane || []).length > 0 && allDetail.length > 0 && (
        <div style={{ background: "#1e1e2e", borderRadius: 0, padding: 16, marginTop: 12, fontFamily: "monospace", fontSize: 12 }}>
          <div style={{ color: "#cba6f7", fontWeight: 700, marginBottom: 8 }}>🔍 Debug: เปรียบเทียบรหัสสินค้า</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div>
              <div style={{ color: "#a6e3a1", fontWeight: 700, marginBottom: 4 }}>Master (5 แรก)</div>
              {(masterLane || []).slice(0, 5).map((m, i) => (
                <div key={i} style={{ color: "#cdd6f4", padding: "2px 0" }}>
                  <span style={{ color: "#f9e2af" }}>[{typeof m.productCode}]</span> "{m.productCode}" → {m.laneKey || <span style={{ color: "#f38ba8" }}>null (lane ไม่ match)</span>}
                </div>
              ))}
            </div>
            <div>
              <div style={{ color: "#89b4fa", fontWeight: 700, marginBottom: 4 }}>Source (5 แรก)</div>
              {allDetail.slice(0, 5).map((r, i) => (
                <div key={i} style={{ color: "#cdd6f4", padding: "2px 0" }}>
                  <span style={{ color: "#f9e2af" }}>[{typeof r.productCode}]</span> "{r.productCode}" → plate: {r.plate}
                </div>
              ))}
            </div>
          </div>
          <div style={{ color: "#6c7086", marginTop: 8, fontSize: 11 }}>
            ถ้า type ต่างกัน (number vs string) หรือ leading zeros หาย → นั่นคือสาเหตุ
          </div>
        </div>
      )}

      {(masterLane || []).length === 0 && (
        <div style={{ background: "#f9fafb", border: "1.5px dashed #d1d5db", borderRadius: 0, padding: 20, color: "#9ca3af", fontWeight: 600, fontSize: 14, textAlign: "center" }}>
          🗂️ กรุณาอัปโหลดไฟล์ Master ลานโหลดก่อน (เมนู "Master Setting")
        </div>
      )}
    </div>
  );
};

// ─── PIN SETTINGS (รหัสผ่านเข้าแก้ Master Setting) ───────────────────────────
const PinSettings = () => {
  const [pin,        setPin]        = useState("");
  const [confirmPin, setConfirmPin] = useState("");
  const [saving,     setSaving]     = useState(false);
  const [msg,        setMsg]        = useState("");

  const inp = { width: "100%", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "9px 12px", fontSize: 14, fontWeight: 600, boxSizing: "border-box", outline: "none", textAlign: "center", letterSpacing: 4 };

  const save = async () => {
    if (!/^\d{4}$/.test(pin)) { setMsg(""); alert("รหัสผ่านต้องเป็นตัวเลข 4 หลัก"); return; }
    if (pin !== confirmPin) { setMsg(""); alert("รหัสผ่านทั้งสองช่องไม่ตรงกัน"); return; }
    setSaving(true);
    setMsg("");
    try {
      await saveSetting("master_setting_pin", pin);
      setPin(""); setConfirmPin("");
      setMsg("✅ เปลี่ยนรหัสผ่านสำเร็จ — มีผลทันทีในเซสชันนี้ เครื่องอื่นต้องรีเฟรชหน้าถึงจะเห็นค่าใหม่");
    } catch (e) {
      alert("บันทึกไม่สำเร็จ: " + e.message);
    } finally {
      setSaving(false);
    }
  };

  return (
    <Collapsible title="🔑 รหัสผ่านเข้าแก้ Master Setting">
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
        <div>
          <label style={{ display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 }}>รหัสผ่านใหม่ (4 หลัก)</label>
          <input type="password" inputMode="numeric" maxLength={4} placeholder="••••" value={pin}
            onChange={e => setPin(e.target.value.replace(/\D/g, "").slice(0, 4))} style={inp} />
        </div>
        <div>
          <label style={{ display: "block", fontWeight: 700, fontSize: 12, color: "#6b7280", marginBottom: 6 }}>ยืนยันรหัสผ่านใหม่</label>
          <input type="password" inputMode="numeric" maxLength={4} placeholder="••••" value={confirmPin}
            onChange={e => setConfirmPin(e.target.value.replace(/\D/g, "").slice(0, 4))} style={inp} />
        </div>
      </div>
      <button onClick={save} disabled={saving}
        style={{ background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "9px 20px", fontSize: 13, fontWeight: 700, cursor: saving ? "default" : "pointer" }}>
        {saving ? "กำลังบันทึก..." : "บันทึกรหัสผ่าน"}
      </button>
      {msg && <div style={{ marginTop: 10, fontSize: 12, color: "#065f46" }}>{msg}</div>}
    </Collapsible>
  );
};

// ─── MASTER UPLOAD (แยกออกมาจาก Detail Loading ให้มีหน้าของตัวเอง) ──────────
const MasterUpload = ({ masterLane, onMasterChange }) => {
  const [fileName, setFileName] = useState(() => {
    try { return JSON.parse(localStorage.getItem("wh_detail_names") || "{}").master || ""; } catch { return ""; }
  });
  const [masterDebug, setMasterDebug] = useState(null); // { total, matched, sampleCol3 }
  const [uploadLog, setUploadLog] = useState([]);
  const [unlocked,   setUnlocked]   = useState(false);
  const [pinInput,   setPinInput]   = useState("");
  const [pinError,   setPinError]   = useState(false);

  const checkPin = () => {
    if (pinInput === settings.masterSettingPin) { setUnlocked(true); setPinError(false); }
    else { setPinError(true); setPinInput(""); }
  };

  useEffect(() => {
    supabase.from("wh_master_upload_log").select("id, data").then(({ data }) => {
      const fetched = (data || []).map(r => ({ id: r.id, ...r.data }));
      // merge instead of overwrite — an optimistic entry from an upload that happens
      // before this fetch resolves would otherwise get silently wiped off the list
      setUploadLog(prev => {
        const merged = [...fetched];
        for (const p of prev) {
          if (!merged.some(m => m.fileName === p.fileName && m.uploadedAt === p.uploadedAt)) merged.push(p);
        }
        return merged.sort((a, b) => (b.uploadedAt || "").localeCompare(a.uploadedAt || ""));
      });
    });
  }, []);

  // Parse master file — columns found by header name (ProductBKey=productCode, ลานโหลด=laneId, ProductNameTha=name)
  // falls back to the old fixed positions (col B / col E) if a header isn't found, for older master files
  const parseMasterFile = (file) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array", raw: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
        const headerRow = rows[0] || [];
        const findCol = (name) => headerRow.findIndex(h => String(h || "").trim().toLowerCase() === name.toLowerCase());
        const nameColIdx = findCol("ProductNameTha");
        const codeColIdx0 = findCol("ProductBKey");
        const laneColIdx0 = findCol("ลานโหลด");
        const codeColIdx = codeColIdx0 >= 0 ? codeColIdx0 : settings.masterFileFallbackCols.productCode;
        const laneColIdx = laneColIdx0 >= 0 ? laneColIdx0 : settings.masterFileFallbackCols.lane;
        const dataRows = rows.slice(1).filter(r => r[codeColIdx] && String(r[codeColIdx]).trim() !== "" && String(r[codeColIdx]).trim() !== "SAP");
        const mapped = dataRows.map(r => ({
          productCode: normalizeProductCode(r[codeColIdx]),
          laneKey:     normalizeLaneKey(r[laneColIdx]),
          rawCol3:     String(r[laneColIdx] || "").trim(),
          productNameTha: nameColIdx >= 0 ? String(r[nameColIdx] || "").trim() : "",
        }));
        const parsed = mapped.filter(r => r.productCode && r.laneKey);
        setMasterDebug({
          total:    dataRows.length,
          matched:  parsed.length,
          sampleCol3: [...new Set(mapped.map(r => r.rawCol3))].slice(0, 6),
        });
        setFileName(file.name);
        try {
          const names = JSON.parse(localStorage.getItem("wh_detail_names") || "{}");
          localStorage.setItem("wh_detail_names", JSON.stringify({ ...names, master: file.name }));
        } catch {}
        onMasterChange(parsed);

        const uploadId = `upload_${Date.now()}`;
        const logEntry = { id: uploadId, fileName: file.name, uploadedAt: new Date().toISOString(), matched: parsed.length, total: dataRows.length, rows: parsed };
        setUploadLog(prev => [logEntry, ...prev]);
        supabase.from("wh_master_upload_log").insert({ id: uploadId, data: logEntry })
          .then(({ error }) => {
            if (error) {
              console.error("บันทึกประวัติการอัพโหลดไม่สำเร็จ:", error);
              // roll back the optimistic row — otherwise it looks saved until the next reload silently drops it
              setUploadLog(prev => prev.filter(h => h.id !== uploadId));
              alert("อัปโหลด Master สำเร็จ แต่บันทึกประวัติการอัพโหลดไม่สำเร็จ: " + error.message);
            }
          });
      } catch(e) {
        alert("อ่านไฟล์ Master ไม่สำเร็จ: " + e.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const downloadRows = (rows, fileSuffix) => {
    const sheetRows = (rows || []).map(m => ({
      "ProductBKey":    m.productCode,
      "ลานโหลด":        m.rawCol3 || m.laneKey,
      "ProductNameTha": m.productNameTha || "",
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheetRows), "Master");
    XLSX.writeFile(wb, `master_${fileSuffix}.xlsx`);
  };

  const deleteUploadLogEntry = async (entry, index) => {
    if (!window.confirm(`ลบประวัติ "${entry.fileName}" ออกจากรายการ?`)) return;
    setUploadLog(prev => prev.filter((_, i) => i !== index));
    if (entry.id) {
      const { error } = await supabase.from("wh_master_upload_log").delete().eq("id", entry.id);
      if (error) alert("ลบไม่สำเร็จ: " + error.message);
    }
  };

  if (!unlocked) {
    return (
      <div style={{ minHeight: "60vh", display: "flex", alignItems: "center", justifyContent: "center", padding: "16px 20px", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <div style={{ maxWidth: 320, width: "100%", textAlign: "center" }}>
          <div style={{ fontSize: 40, marginBottom: 10 }}>🔒</div>
          <h2 style={{ fontWeight: 900, fontSize: 18, margin: "0 0 4px" }}>Master Setting</h2>
          <p style={{ color: "#6b7280", fontSize: 12, margin: "0 0 20px" }}>กรุณาใส่รหัสผ่านก่อนเข้าแก้ไข</p>
          <input
            type="password"
            inputMode="numeric"
            maxLength={4}
            value={pinInput}
            onChange={e => { setPinInput(e.target.value.replace(/\D/g, "").slice(0, 4)); setPinError(false); }}
            onKeyDown={e => { if (e.key === "Enter") checkPin(); }}
            placeholder="••••"
            autoFocus
            style={{ width: "100%", textAlign: "center", fontSize: 26, letterSpacing: 8, border: `2px solid ${pinError ? "#dc2626" : "#e5e7eb"}`, borderRadius: 0, padding: "12px 0", outline: "none", boxSizing: "border-box", marginBottom: 12 }}
          />
          {pinError && <div style={{ color: "#dc2626", fontSize: 12, fontWeight: 700, marginBottom: 12 }}>รหัสไม่ถูกต้อง</div>}
          <button onClick={checkPin} style={{ width: "100%", background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "12px 0", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>
            ยืนยัน
          </button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ padding: 20 }}>
      <h2 style={{ margin: "0 0 20px", fontSize: 20, fontWeight: 900 }}>🗂️ Master Setting</h2>

      <div style={{ display: "flex", flexWrap: "wrap", alignItems: "stretch", gap: 20 }}>
      <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", padding: 20, border: "2px solid #111", flex: 1, minWidth: 320, height: 220, overflowY: "auto", boxSizing: "border-box" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
          <span style={{ fontSize: 28 }}>🗂️</span>
          <div>
            <div style={{ fontWeight: 800, fontSize: 15 }}>Master ลานโหลด</div>
            {fileName && (
              <div style={{ fontSize: 11, color: "#111", fontWeight: 700, marginTop: 2 }}>
                ✅ {fileName}
              </div>
            )}
          </div>
        </div>
        <label style={{ display: "block", background: "#111", color: "#fff", border: "none", borderRadius: 0, padding: "12px 0", textAlign: "center", fontSize: 13, fontWeight: 700, cursor: "pointer", transition: "opacity 0.2s" }}
          onMouseOver={e => e.currentTarget.style.opacity = "0.8"}
          onMouseOut={e => e.currentTarget.style.opacity = "1"}>
          {fileName ? "🔄 เปลี่ยน Master" : "⬆️ อัปโหลด Master"}
          <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={e => { if (e.target.files[0]) parseMasterFile(e.target.files[0]); e.target.value = ""; }} />
        </label>
        {masterDebug && (
          <div style={{ marginTop: 10, fontSize: 11, padding: "8px 10px", borderRadius: 0, background: masterDebug.matched === 0 ? "#fee2e2" : "#d1fae5", color: masterDebug.matched === 0 ? "#991b1b" : "#065f46" }}>
            {masterDebug.matched === 0
              ? <>❌ Match 0/{masterDebug.total} — ค่าใน col D: {masterDebug.sampleCol3.map(v => `"${v}"`).join(", ")}</>
              : <>✅ Match {masterDebug.matched}/{masterDebug.total} รหัส</>}
          </div>
        )}
        {(masterLane || []).length > 0 && (
          <div style={{ marginTop: 10, fontSize: 12, color: "#6b7280" }}>
            มีข้อมูล Master อยู่แล้ว {masterLane.length.toLocaleString()} รายการ
          </div>
        )}
      </div>

      <div style={{ background: "#fff", borderRadius: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.08)", flex: 1, minWidth: 320, height: 220, display: "flex", flexDirection: "column" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 700, fontSize: 14, flexShrink: 0 }}>
          🕓 ประวัติการอัพโหลด <span style={{ background: "#111", color: "#fff", borderRadius: 0, padding: "2px 8px", fontSize: 11, marginLeft: 4 }}>{uploadLog.length}</span>
        </div>
        {uploadLog.length === 0
          ? <div style={{ padding: 30, textAlign: "center", color: "#9ca3af", fontSize: 13 }}>ยังไม่มีประวัติการอัพโหลด</div>
          : (
            <div style={{ overflowX: "auto", overflowY: "auto", flex: 1 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: "#f9fafb", position: "sticky", top: 0, zIndex: 1 }}>
                    <th style={{ padding: "8px 12px", textAlign: "left", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb" }}>ไฟล์</th>
                    <th style={{ padding: "8px 12px", textAlign: "left", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>วันที่/เวลา</th>
                    <th style={{ padding: "8px 12px", textAlign: "center", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb" }}></th>
                    <th style={{ padding: "8px 12px", textAlign: "center", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>Match</th>
                    <th style={{ padding: "8px 12px", textAlign: "center", fontWeight: 700, color: "#374151", borderBottom: "1px solid #e5e7eb" }}></th>
                  </tr>
                </thead>
                <tbody>
                  {uploadLog.map((h, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td style={{ padding: "8px 12px", fontSize: 11 }}>{h.fileName}</td>
                      <td style={{ padding: "8px 12px", whiteSpace: "nowrap", color: "#6b7280" }}>
                        {h.uploadedAt ? new Date(h.uploadedAt).toLocaleString("th-TH", { dateStyle: "short", timeStyle: "short" }) : "—"}
                      </td>
                      <td style={{ padding: "8px 12px", textAlign: "center" }}>
                        {h.rows && h.rows.length > 0 ? (
                          <button onClick={() => downloadRows(h.rows, h.uploadedAt ? h.uploadedAt.slice(0, 10) : String(i))}
                            title="ดาวน์โหลดไฟล์นี้"
                            style={{ background: "none", border: "none", padding: 0, cursor: "pointer", color: "#6b7280", display: "inline-flex" }}>
                            <Icon name="download" size={16} />
                          </button>
                        ) : (
                          <span style={{ color: "#e5e7eb" }} title="ไม่มีข้อมูลให้ดาวน์โหลด">
                            <Icon name="download" size={16} />
                          </span>
                        )}
                      </td>
                      <td style={{ padding: "8px 12px", textAlign: "center", color: h.matched === 0 ? "#dc2626" : "#065f46", fontWeight: 700, fontSize: 11 }}>
                        {h.matched}/{h.total}
                      </td>
                      <td style={{ padding: "8px 12px", textAlign: "center" }}>
                        <button onClick={() => deleteUploadLogEntry(h, i)}
                          title="ลบประวัตินี้"
                          style={{ background: "none", border: "none", padding: 0, cursor: "pointer", fontSize: 14 }}>
                          🗑️
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )
        }
      </div>
      </div>

      <div style={{ marginTop: 20, display: "grid", gridTemplateColumns: "1fr", gap: 14 }}>
        <SystemSettings />
        <LaneSettings />
        <BaySettings />
        <RoleSettings />
        <LaneAliasSettings />
        <WaitingReasonSettings />
        <BasketTypeSettings />
        <DetailSourceSettings />
        <PinSettings />
      </div>
    </div>
  );
};

// ─── WORK TRACKING (Power BI–style matrix table) ─────────────────────────────
const WT_GROUPS = [
  { id: "info",  label: "ข้อมูลรถ/สินค้า",         span: 6, dark: "#0f172a", mid: "#1e293b" },
  { id: "entry", label: "เข้าโรงงาน / เบิกสินค้า", span: 5, dark: "#1d4ed8", mid: "#2563eb" },
  { id: "parts", label: "🥩 ลานชิ้นส่วน",         span: 15, dark: "#c2410c", mid: "#ea580c" },
  { id: "head",  label: "🐷 ลานหัว/เครื่องใน",     span: 15, dark: "#6d28d9", mid: "#7c3aed" },
  { id: "pork",  label: "🐖 ลานหมูซีก",            span: 15, dark: "#9f1239", mid: "#be123c" },
  { id: "docs",  label: "เอกสาร",                 span: 2, dark: "#0d9488", mid: "#0f766e" },
  { id: "exit",  label: "ออกโรงงาน",              span: 2, dark: "#475569", mid: "#64748b" },
];
const WT_GROUP_MAP = Object.fromEntries(WT_GROUPS.map(g => [g.id, g]));
const WT_CELL_BG  = { info: "#f8fafc", entry: "#eff6ff", parts: "#fff7ed", head: "#f5f3ff", pork: "#fff1f2", docs: "#f0fdfa", exit: "#f8fafc" };

const GRP_LANE = { parts: "lane_parts", head: "lane_head", pork: "lane_pork" };

// ─── STAT TILE (Tracking summary cards) ──────────────────────────────────────
// de-emphasis bars in a light tint of the tile's own accent hue, current (last) bar in full accent
const Sparkline = ({ data, accent, height = 30, barWidth = 5, gap = 3 }) => {
  if (!data || !data.length) return null;
  const max = Math.max(1, ...data);
  return (
    <div style={{ display: "flex", alignItems: "flex-end", gap, height, flexShrink: 0 }}>
      {data.map((v, i) => {
        const h = Math.max(2, Math.round((v / max) * height));
        const isLast = i === data.length - 1;
        return <div key={i} style={{ width: barWidth, height: h, borderRadius: "3px 3px 0 0", background: isLast ? accent : `${accent}2E` }} />;
      })}
    </div>
  );
};

// e.g. 68 → "1 ชม. 8 น.", 45 → "45 น."
const formatMinsDelta = (mins) => (mins < 60 ? `${mins} น.` : `${Math.floor(mins / 60)} ชม. ${mins % 60} น.`);

// plain "yesterday's value" note — no delta math, no color, just the number
const PrevNote = ({ label, value }) => (
  <div style={{ display: "flex", alignItems: "center", gap: 4, minWidth: 0 }}>
    <span style={{ fontSize: 10.5, color: "#898781", fontWeight: 600, whiteSpace: "nowrap" }}>{label}</span>
    <span style={{ fontSize: 11, color: "#52514e", fontWeight: 800, whiteSpace: "nowrap" }}>{value}</span>
  </div>
);

const StatTile = ({ icon, accent, title, value, unit, prevValue, sparkline }) => (
  <div style={{ background: "#fff", borderRadius: 16, padding: "16px 16px 14px", boxShadow: "0 1px 2px rgba(11,11,11,0.06), 0 8px 20px rgba(11,11,11,0.07)", display: "flex", flexDirection: "column", gap: 12, minWidth: 0 }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ width: 30, height: 30, borderRadius: "50%", background: `${accent}1A`, color: accent, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <Icon name={icon} size={15} />
      </div>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#52514e", lineHeight: 1.25 }}>{title}</div>
    </div>

    <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", gap: 8 }}>
      <div style={{ fontSize: 26, fontWeight: 900, color: "#0b0b0b", lineHeight: 1, whiteSpace: "nowrap" }}>
        {value}{unit ? <span style={{ fontSize: 13, fontWeight: 700, color: "#898781", marginLeft: 3 }}>{unit}</span> : null}
      </div>
      <Sparkline data={sparkline} accent={accent} />
    </div>

    <PrevNote label="เมื่อวาน" value={prevValue} />
  </div>
);

// combined tile for exit on-time % + late count — two figures, one "รถออก" story
const ExitStatTile = ({ icon, accent, title, onTimePct, onTimeCount, prevOnTimePct, lateCount, latePct, prevLatePct, sparkline }) => (
  <div style={{ background: "#fff", borderRadius: 16, padding: "16px 16px 14px", boxShadow: "0 1px 2px rgba(11,11,11,0.06), 0 8px 20px rgba(11,11,11,0.07)", display: "flex", flexDirection: "column", gap: 12, minWidth: 0 }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ width: 30, height: 30, borderRadius: "50%", background: `${accent}1A`, color: accent, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <Icon name={icon} size={15} />
      </div>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#52514e", lineHeight: 1.25 }}>{title}</div>
    </div>

    <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", gap: 8 }}>
      <div style={{ display: "flex", alignItems: "flex-end", gap: 12 }}>
        <div>
          <div style={{ fontSize: 10.5, color: "#898781", fontWeight: 600, marginBottom: 2, whiteSpace: "nowrap" }}>ตรงเวลา{onTimeCount != null ? ` (${onTimeCount} คัน)` : ""}</div>
          <div style={{ fontSize: 26, fontWeight: 900, color: "#0b0b0b", lineHeight: 1, whiteSpace: "nowrap" }}>
            {onTimePct != null ? onTimePct : "—"}{onTimePct != null ? <span style={{ fontSize: 13, fontWeight: 700, color: "#898781", marginLeft: 2 }}>%</span> : null}
          </div>
        </div>
        <div style={{ width: 1, alignSelf: "stretch", background: "#e5e7eb" }} />
        <div>
          <div style={{ fontSize: 10.5, color: "#898781", fontWeight: 600, marginBottom: 2, whiteSpace: "nowrap" }}>สาย{lateCount != null ? ` (${lateCount} คัน)` : ""}</div>
          <div style={{ fontSize: 26, fontWeight: 900, color: "#e34948", lineHeight: 1, whiteSpace: "nowrap" }}>
            {latePct != null ? latePct : "—"}{latePct != null ? <span style={{ fontSize: 13, fontWeight: 700, color: "#898781", marginLeft: 2 }}>%</span> : null}
          </div>
        </div>
      </div>
      <Sparkline data={sparkline} accent={accent} />
    </div>

    <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
      <PrevNote label="เมื่อวานตรงเวลา" value={prevOnTimePct != null ? `${prevOnTimePct}%` : "—"} />
      <PrevNote label="เมื่อวานสาย" value={prevLatePct != null ? `${prevLatePct}%` : "—"} />
    </div>
  </div>
);

// combined tile for total entered count + on-time % — one "รถเข้า" story
const EntryStatTile = ({ icon, accent, title, count, prevCount, onTimePct, onTimeCount, prevOnTimePct, sparkline }) => (
  <div style={{ background: "#fff", borderRadius: 16, padding: "16px 16px 14px", boxShadow: "0 1px 2px rgba(11,11,11,0.06), 0 8px 20px rgba(11,11,11,0.07)", display: "flex", flexDirection: "column", gap: 12, minWidth: 0 }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{ width: 30, height: 30, borderRadius: "50%", background: `${accent}1A`, color: accent, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <Icon name={icon} size={15} />
      </div>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#52514e", lineHeight: 1.25 }}>{title}</div>
    </div>

    <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", gap: 8 }}>
      <div style={{ display: "flex", alignItems: "flex-end", gap: 12 }}>
        <div>
          <div style={{ fontSize: 10.5, color: "#898781", fontWeight: 600, marginBottom: 2, whiteSpace: "nowrap" }}>เข้าโรงงาน</div>
          <div style={{ fontSize: 26, fontWeight: 900, color: "#0b0b0b", lineHeight: 1, whiteSpace: "nowrap" }}>
            {count}<span style={{ fontSize: 13, fontWeight: 700, color: "#898781", marginLeft: 2 }}>คัน</span>
          </div>
        </div>
        <div style={{ width: 1, alignSelf: "stretch", background: "#e5e7eb" }} />
        <div>
          <div style={{ fontSize: 10.5, color: "#898781", fontWeight: 600, marginBottom: 2, whiteSpace: "nowrap" }}>ตรงเวลา{onTimeCount != null ? ` (${onTimeCount} คัน)` : ""}</div>
          <div style={{ fontSize: 26, fontWeight: 900, color: "#1baf7a", lineHeight: 1, whiteSpace: "nowrap" }}>
            {onTimePct != null ? onTimePct : "—"}{onTimePct != null ? <span style={{ fontSize: 13, fontWeight: 700, color: "#898781", marginLeft: 2 }}>%</span> : null}
          </div>
        </div>
      </div>
      <Sparkline data={sparkline} accent={accent} />
    </div>

    <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
      <PrevNote label="เมื่อวานเข้า" value={`${prevCount ?? "—"} คัน`} />
      <PrevNote label="เมื่อวานตรงเวลา" value={prevOnTimePct != null ? `${prevOnTimePct}%` : "—"} />
    </div>
  </div>
);

// ── สถานะของ log ปิดวันทำงาน (wh_archive_audit.status) → สี/ป้ายที่แสดงในหน้า Tracking ──
const WORK_LOG_STATUS_STYLE = {
  SUCCESS:  { bg: "#dcfce7", fg: "#166534", label: "SUCCESS" },
  WARNING:  { bg: "#fef3c7", fg: "#92400e", label: "WARNING" },
  ERROR:    { bg: "#fee2e2", fg: "#991b1b", label: "ERROR" },
  PARTIAL:  { bg: "#dbeafe", fg: "#1e40af", label: "PARTIAL" },
  NO_DATA:  { bg: "#f3f4f6", fg: "#6b7280", label: "NO DATA" },
};
const WorkLogStatusBadge = ({ status }) => {
  const s = WORK_LOG_STATUS_STYLE[status] || { bg: "#f3f4f6", fg: "#6b7280", label: status || "ไม่มี log" };
  return <span style={{ background: s.bg, color: s.fg, fontWeight: 800, fontSize: 11, padding: "3px 8px", borderRadius: 0, whiteSpace: "nowrap" }}>{s.label}</span>;
};

// ── ตารางสุขภาพข้อมูลรายวัน (โหมดช่วงวันที่ / ทั้งหมด ในหน้า Tracking) ──
// อ่านจาก view wh_daily_health (ดู supabase-add-work-log-and-health-check.sql) — เทียบจำนวนที่
// archive ไว้จริงกับ log การปิดวันทำงานล่าสุดของวันนั้น ใช้ตอบคำถาม "วันไหนข้อมูลผิดปกติ/หายไป"
const DailyHealthTable = ({ rows, loading, onViewDate }) => {
  const th = { padding: "9px 12px", textAlign: "left", fontWeight: 700, color: "#fff", background: "#1e293b", whiteSpace: "nowrap" };
  const td = { padding: "9px 12px", borderBottom: "1px solid #f1f5f9", whiteSpace: "nowrap" };
  return (
    <div style={{ background: "#fff", boxShadow: "0 4px 16px rgba(0,0,0,0.1)", marginBottom: 16, overflowX: "auto" }}>
      <div style={{ padding: "12px 16px", fontWeight: 800, fontSize: 14, borderBottom: "1px solid #f1f5f9" }}>
        🩺 สุขภาพข้อมูลรายวัน — ตรวจว่าวันไหนข้อมูลหาย/ผิดปกติจากประวัติการปิดวันทำงานจริง
      </div>
      {loading && <div style={{ textAlign: "center", color: "#9ca3af", padding: 30 }}>กำลังโหลด...</div>}
      {!loading && rows.length === 0 && <div style={{ textAlign: "center", color: "#9ca3af", padding: 30 }}>ไม่พบข้อมูล archive ในช่วงที่เลือก</div>}
      {!loading && rows.length > 0 && (
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12.5, minWidth: 720 }}>
          <thead><tr>
            <th style={th}>วันที่</th><th style={th}>รถ</th><th style={th}>คิว</th>
            <th style={th}>สถานะล่าสุด</th><th style={th}>ปิดวันโดย</th><th style={th}>เวลา</th>
            <th style={th}>รายละเอียด</th><th style={th}>จำนวนครั้งที่ปิด</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {rows.map(r => {
              const suspicious = r.trucks_count === 0 && r.queue_count === 0;
              return (
                <tr key={r.archive_date} style={{ background: suspicious ? "#fef2f2" : r.is_reconstructed ? "#fffbeb" : "#fff" }}>
                  <td style={{ ...td, fontWeight: 700 }}>{r.archive_date}{r.is_reconstructed ? " 🛠️" : ""}</td>
                  <td style={td}>{r.trucks_count}</td>
                  <td style={td}>{r.queue_count}</td>
                  <td style={td}><WorkLogStatusBadge status={r.last_status} /></td>
                  <td style={td}>{r.last_source || "—"}</td>
                  <td style={td}>{r.last_run_at ? new Date(r.last_run_at).toLocaleString("th-TH") : "—"}</td>
                  <td style={{ ...td, maxWidth: 260, overflow: "hidden", textOverflow: "ellipsis", color: "#991b1b" }} title={r.last_detail || ""}>{r.last_detail || "—"}</td>
                  <td style={td}>{r.run_count}{r.run_count > 1 ? " ⚠️ ปิดซ้ำ" : ""}</td>
                  <td style={td}>
                    <button onClick={() => onViewDate(r.archive_date)}
                      style={{ background: "#111827", color: "#fff", border: "none", padding: "4px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer" }}>
                      ดูรายวัน
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
    </div>
  );
};

// ── Data Health Check ของวันเดียว: Input/Processed/Output/Missing/Error/Status ──
// (ตามข้อ 8 ที่ขอ) + log การปิดวันทำงานดิบของวันนั้นแบบขยายดูได้ ตอบคำถาม "ข้อมูลหายไปตรงไหน เมื่อไหร่"
const DataHealthPanel = ({ date, today, activeTrucks, activeQueue, audit }) => {
  const [expanded, setExpanded] = useState(false);
  if (date === today) {
    return (
      <div style={{ background: "#eff6ff", color: "#1e3a8a", padding: "8px 14px", fontSize: 12.5, fontWeight: 700, marginBottom: 12 }}>
        🩺 วันนี้ยังไม่ปิดวันทำงาน (ข้อมูลจะถูกตรวจสุขภาพหลังปิดวันแล้ว) — ตอนนี้มีรถ {activeTrucks.length} คัน, คิว {activeQueue.length} รายการ
      </div>
    );
  }
  const latest = audit[0]; // audit เรียง ts desc มาแล้ว
  const inputCount  = latest ? latest.trucks_before + latest.queue_before : null;
  const outputCount = activeTrucks.length + activeQueue.length;
  const plateCounts = {};
  for (const t of activeTrucks) { if (t.plate) plateCounts[t.plate] = (plateCounts[t.plate] || 0) + 1; }
  const duplicateCount = Object.values(plateCounts).filter(c => c > 1).length;
  const errorRuns = audit.filter(a => a.action === "error" || a.status === "ERROR").length;
  const warningRuns = audit.filter(a => a.status === "WARNING").length;
  const status = errorRuns > 0 ? "ERROR" : warningRuns > 0 ? "WARNING" : outputCount === 0 ? "NO_DATA" : (latest?.status || (audit.length ? latest?.status : null));
  const missing = inputCount != null ? Math.max(0, inputCount - outputCount) : null;

  const cell = { padding: "10px 14px", textAlign: "center", borderRight: "1px solid #f1f5f9" };
  return (
    <div style={{ background: "#fff", boxShadow: "0 2px 10px rgba(0,0,0,0.08)", marginBottom: 16 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 14px", cursor: "pointer" }}
        onClick={() => setExpanded(v => !v)}>
        <span style={{ fontWeight: 800, fontSize: 14 }}>🩺 Data Health Check — {date}</span>
        <span style={{ fontSize: 12, color: "#6b7280" }}>{expanded ? "▲ ซ่อน log" : "▼ ดู log การปิดวันทำงาน"}</span>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", borderTop: "1px solid #f1f5f9" }}>
        <div style={cell}><div style={{ fontSize: 11, color: "#9ca3af" }}>Input</div><div style={{ fontWeight: 800 }}>{inputCount ?? "ไม่มี log"}</div></div>
        <div style={cell}><div style={{ fontSize: 11, color: "#9ca3af" }}>Output (ที่แสดงจริง)</div><div style={{ fontWeight: 800 }}>{outputCount}</div></div>
        <div style={cell}><div style={{ fontSize: 11, color: "#9ca3af" }}>Missing</div><div style={{ fontWeight: 800, color: missing > 0 ? "#dc2626" : "#111827" }}>{missing ?? "—"}</div></div>
        <div style={cell}><div style={{ fontSize: 11, color: "#9ca3af" }}>Duplicate (ทะเบียนซ้ำ)</div><div style={{ fontWeight: 800, color: duplicateCount > 0 ? "#dc2626" : "#111827" }}>{duplicateCount}</div></div>
        <div style={cell}><div style={{ fontSize: 11, color: "#9ca3af" }}>Error runs</div><div style={{ fontWeight: 800, color: errorRuns > 0 ? "#dc2626" : "#111827" }}>{errorRuns}</div></div>
        <div style={{ ...cell, borderRight: "none" }}><div style={{ fontSize: 11, color: "#9ca3af" }}>สถานะ</div><div style={{ marginTop: 2 }}><WorkLogStatusBadge status={status} /></div></div>
      </div>
      {expanded && (
        <div style={{ borderTop: "1px solid #f1f5f9", overflowX: "auto" }}>
          {audit.length === 0
            ? <div style={{ padding: 16, color: "#9ca3af", fontSize: 12.5 }}>ไม่มี log การปิดวันทำงานของวันนี้ (อาจเป็นข้อมูลเก่าก่อนติดตั้งระบบ audit log)</div>
            : <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12, minWidth: 640 }}>
                <thead><tr style={{ background: "#f9fafb" }}>
                  {["เวลา", "Source", "Action", "ก่อน (รถ/คิว)", "หลัง (รถ/คิว)", "สถานะ", "Ref ID", "รายละเอียด"].map(h => (
                    <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 700, color: "#6b7280", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {audit.map(a => (
                    <tr key={a.id}>
                      <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{new Date(a.ts).toLocaleString("th-TH")}</td>
                      <td style={{ padding: "8px 10px" }}>{a.source}</td>
                      <td style={{ padding: "8px 10px" }}>{a.action}</td>
                      <td style={{ padding: "8px 10px" }}>{a.trucks_before} / {a.queue_before}</td>
                      <td style={{ padding: "8px 10px" }}>{a.trucks_written ?? "—"} / {a.queue_written ?? "—"}</td>
                      <td style={{ padding: "8px 10px" }}><WorkLogStatusBadge status={a.status} /></td>
                      <td style={{ padding: "8px 10px", fontFamily: "monospace", fontSize: 11 }}>{a.ref_id || "—"}</td>
                      <td style={{ padding: "8px 10px", maxWidth: 320 }}>{a.detail || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>}
        </div>
      )}
    </div>
  );
};

const WorkTracking = ({ trucks, queue, detailMapByChannel = {}, masterLane = [] }) => {
  const today = cycleDateStr();
  const [date, setDate]       = useState(today);
  const [archiveData, setArchiveData] = useState(null);
  const [loadingArchive, setLoadingArchive] = useState(false);
  const [prevArchiveData, setPrevArchiveData] = useState(null);
  const [histDetailMapByChannel, setHistDetailMapByChannel] = useState(null); // for a past date being viewed
  const [sortCol, setSortCol] = useState("arrivedAt");
  const [sortDir, setSortDir] = useState(1);

  // โหมดดูข้อมูล — วันเดียว (ตารางละเอียด + สรุปสถิติ) หรือช่วงวันที่/ทั้งหมด (ตารางสุขภาพข้อมูลรายวัน)
  const [dateMode, setDateMode] = useState("single"); // 'single' | 'range' | 'all'
  const [rangeStart, setRangeStart] = useState(addDaysToDateStr(today, -6));
  const [rangeEnd, setRangeEnd] = useState(today);
  const [rangeHealth, setRangeHealth] = useState([]);
  const [loadingRange, setLoadingRange] = useState(false);

  // ── Data Health Check + audit log ของวันที่กำลังดู (โหมดวันเดียว) ──
  const [dayAudit, setDayAudit] = useState([]);
  useEffect(() => {
    if (dateMode !== "single") return;
    supabase.from("wh_archive_audit").select("*").eq("archive_date", date).order("ts", { ascending: false })
      .then(({ data }) => setDayAudit(data || []));
  }, [date, dateMode]);

  // ── ตารางสุขภาพข้อมูลรายวัน (โหมดช่วงวันที่/ทั้งหมด) ──
  useEffect(() => {
    if (dateMode === "single") return;
    setLoadingRange(true);
    let q = supabase.from("wh_daily_health").select("*").order("archive_date", { ascending: false });
    if (dateMode === "range") q = q.gte("archive_date", rangeStart).lte("archive_date", rangeEnd);
    q.then(({ data }) => setRangeHealth(data || [])).finally(() => setLoadingRange(false));
  }, [dateMode, rangeStart, rangeEnd]);

  // วันปัจจุบัน → ใช้ detailMap สด, วันย้อนหลัง → ดึงไฟล์ PO ของวันนั้นมาคำนวณใหม่
  // (Master ลานโหลดไม่ได้เก็บย้อนหลัง จึงใช้ตัวปัจจุบันร่วมกับไฟล์ PO ของวันที่ดู)
  const effectiveDetailMapByChannel = date === today ? detailMapByChannel : (histDetailMapByChannel || {});

  // lane groups the truck's PO data says it does NOT need — cell renders as light gray (N/A)
  const irrelevantGrps = t => {
    const lanes = laneMatchForTruck(t, effectiveDetailMapByChannel);
    if (!lanes.size) return new Set();
    return new Set(Object.keys(GRP_LANE).filter(g => !lanes.has(GRP_LANE[g])));
  };

  useEffect(() => {
    setArchiveData(null);
    setLoadingArchive(date !== today);
    // is_reconstructed/reconstructed_note มาจาก supabase-add-archive-reconstructed-flag.sql —
    // ถ้ายังไม่ได้รัน migration นั้น fallback ไป select แบบเดิมแทนเพื่อไม่ให้หน้า Tracking พังทั้งหน้า
    supabase.from("wh_archive").select("trucks, queue, is_reconstructed, reconstructed_note").eq("archive_date", date).single()
      .then(res => res.error
        ? supabase.from("wh_archive").select("trucks, queue").eq("archive_date", date).single()
        : res)
      .then(({ data }) => setArchiveData(data ?? null))
      .finally(() => setLoadingArchive(false));
  }, [date, today]);

  // ก่อนหน้า 1 วัน — ใช้เทียบ delta ใน summary card เท่านั้น
  useEffect(() => {
    setPrevArchiveData(null);
    const d = new Date(date);
    d.setDate(d.getDate() - 1);
    const prevDate = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    let cancelled = false;
    supabase.from("wh_archive").select("trucks, queue").eq("archive_date", prevDate).single()
      .then(({ data }) => { if (!cancelled) setPrevArchiveData(data ?? null); });
    return () => { cancelled = true; };
  }, [date]);

  useEffect(() => {
    if (date === today) { setHistDetailMapByChannel(null); return; }
    let cancelled = false;
    fetchDetailSrc(date).then(remote => {
      if (cancelled) return;
      const rowsByChannel = {};
      for (const src of detailSources) rowsByChannel[src.id] = remote?.[src.id]?.rows || [];
      setHistDetailMapByChannel(buildDetailMapByChannel(masterLane, rowsByChannel));
    });
    return () => { cancelled = true; };
  }, [date, today, masterLane]);

  const activeTrucks = archiveData?.trucks ?? (date === today ? trucks : []);
  const activeQueue  = archiveData?.queue  ?? (date === today ? queue  : []);
  const pNum = s => (String(s).match(/\d+/g) || []).pop() || "";
  const getQ = t => activeQueue.find(q => q.id === t.queueId) || activeQueue.find(q => pNum(q.plate) === pNum(t.plate) && pNum(q.plate) !== "");

  // ── ตรวจสุขภาพข้อมูลของวันที่กำลังดู — เตือนแทนที่จะแสดงตารางว่างเงียบๆ ──
  // (กันซ้ำกรณีแบบ archive วันที่ 6-8 ก.ย. 2026: แถวถูกเขียนทับด้วยคิวของ "วันถัดไป" หลัง
  // race condition ใน promoteIfDue — ดู supabase-fix-archive-race-condition.sql)
  const thaiDateToISO = s => {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(s || ""));
    if (!m) return null;
    return `${m[3]}-${String(m[2]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;
  };
  const integrityWarning = (() => {
    if (date === today) return null;
    if (loadingArchive) return null;
    if (!archiveData) return `ไม่พบข้อมูลของวันที่ ${date} ในระบบ archive เลย (อาจยังไม่ถึงเวลาปิดวันทำงาน หรือข้อมูลสูญหาย — ตรวจสอบ wh_archive_audit)`;
    const mismatched = activeQueue.find(q => {
      const iso = thaiDateToISO(q.date);
      return iso && iso !== date;
    });
    if (mismatched) return `ข้อมูลคิวบางรายการมีวันที่ไม่ตรงกับ ${date} (พบ "${mismatched.date}") — แถว archive นี้อาจถูกเขียนทับผิดวันจากบั๊ก race condition เดิม ไม่ควรเชื่อถือข้อมูลชุดนี้`;
    if (activeTrucks.length === 0 && activeQueue.length === 0) return `วันที่ ${date} มี archive แต่ไม่มีข้อมูลรถ/คิวเลย — ถ้าเป็นวันทำงานปกติ ให้ตรวจสอบว่าข้อมูลสูญหายหรือไม่`;
    return null;
  })();

  // helper รวม note/รูปภาพ/bay/ตะกร้า/สถานะรอสินค้า ของแต่ละลาน ให้แสดงเป็น tooltip บนเซลล์เวลา
  // แทนการเปิดคอลัมน์แยกรายฟิลด์ (จะทำให้ตารางกว้างเกินไป) — ข้อมูลยังดูได้ครบ แค่ hover เอา
  // แต่ละลานสร้างคอลัมน์ชุดเดียวกัน 15 คอลัมน์ (QC 4 + สุ่มตรวจ 4 + โหลด 7) — ทำเป็นฟังก์ชันเดียว
  // ใช้ซ้ำ 3 ลาน กันพิมพ์ผิด/ลืมอัพเดตไม่ตรงกันระหว่างลาน (เดิมเป็น tooltip เดียวรวมทุกอย่าง
  // ย้ายมาเป็นคอลัมน์จริงให้เห็นครบโดยไม่ต้อง hover ตามที่ขอ)
  const basketsSummary = load => {
    if (!load?.baskets) return "—";
    const s = Object.entries(load.baskets).filter(([, v]) => v).map(([k, v]) => `${k}=${v}`).join(", ");
    return s || "—";
  };
  const waitingSummary = load => {
    if (!load?.waiting) return "—";
    return `รอตั้งแต่ ${load.waitingAt || "?"}${load.waitingFor ? ` (${load.waitingFor})` : ""}`;
  };
  const laneCols = (laneId, grp, laneLabel) => [
    { id: `qc_${grp}`,          grp, label: "ตรวจอุณหภูมิรถ", align: "center", get: t => t.qcLanes?.[laneId]?.doneAt },
    { id: `qc_${grp}_temp`,     grp, label: "อุณหภูมิ",        align: "center", type: "text", get: t => t.qcLanes?.[laneId]?.temp ? `${t.qcLanes[laneId].temp}°C` : "—" },
    { id: `qc_${grp}_bay`,      grp, label: "ช่อง QC",         align: "center", type: "text", get: t => t.qcLanes?.[laneId]?.bayId || "—" },
    { id: `qc_${grp}_photos`,   grp, label: "รูป QC",          align: "center", type: "text", get: t => t.qcLanes?.[laneId]?.photos?.length || "—" },
    { id: `sample_${grp}`,        grp, label: "QC สุ่มตรวจ",     align: "center", get: t => t.sampleLanes?.[laneId]?.doneAt },
    { id: `sample_${grp}_note`,   grp, label: "หมายเหตุสุ่มตรวจ", align: "left",   type: "text", get: t => t.sampleLanes?.[laneId]?.note || "—" },
    { id: `sample_${grp}_bay`,    grp, label: "ช่องสุ่มตรวจ",    align: "center", type: "text", get: t => t.sampleLanes?.[laneId]?.bayId || "—" },
    { id: `sample_${grp}_photos`, grp, label: "รูปสุ่มตรวจ",     align: "center", type: "text", get: t => t.sampleLanes?.[laneId]?.photos?.length || "—" },
    { id: `load_${grp}`,          grp, label: "โหลดเสร็จ",       align: "center", get: t => t.loadLanes?.[laneId]?.doneAt },
    { id: `load_${grp}_waiting`,  grp, label: "รอสินค้า",        align: "left",   type: "text", get: t => waitingSummary(t.loadLanes?.[laneId]) },
    { id: `load_${grp}_bay`,      grp, label: "ช่องโหลด",        align: "center", type: "text", get: t => t.loadLanes?.[laneId]?.bayId || "—" },
    { id: `load_${grp}_baskets`,  grp, label: "ตะกร้า/ตะขอ",     align: "left",   type: "text", get: t => basketsSummary(t.loadLanes?.[laneId]) },
    { id: `load_${grp}_payer`,    grp, label: "ผู้จ่ายตะกร้า",   align: "left",   type: "text", get: t => t.loadLanes?.[laneId]?.basketPayer || "—" },
    { id: `load_${grp}_note`,     grp, label: "หมายเหตุโหลด",    align: "left",   type: "text", get: t => t.loadLanes?.[laneId]?.note || "—" },
    { id: `load_${grp}_photos`,   grp, label: "รูปโหลดเสร็จ",    align: "center", type: "text", get: t => t.loadLanes?.[laneId]?.photos?.length || "—" },
  ];

  const COLS = [
    { id: "plate",            grp: "info",   label: "ทะเบียน",    align: "left",   get: t => t.plate },
    { id: "customerGroup",    grp: "info",   label: "กลุ่มลูกค้า", align: "left",   type: "text", get: t => t.customerGroup || "—" },
    { id: "driver",           grp: "info",   label: "คนขับ",      align: "left",   type: "text", get: t => t.driver || getQ(t)?.driver || "—" },
    { id: "product",          grp: "info",   label: "สินค้า",      align: "left",   type: "text", get: t => t.product || getQ(t)?.product || "—" },
    { id: "destination",      grp: "info",   label: "ปลายทาง",    align: "left",   type: "text", get: t => t.destination || t.zone || getQ(t)?.destination || getQ(t)?.zone || "—" },
    { id: "qty",              grp: "info",   label: "จำนวน",      align: "left",   type: "text", get: t => (t.qty || getQ(t)?.qty) ? `${t.qty ?? getQ(t)?.qty} ${t.unit || getQ(t)?.unit || ""}`.trim() : "—" },
    { id: "entrySTD",         grp: "entry",  label: "STD เข้า",   align: "center", get: t => getQ(t)?.entryTime },
    { id: "arrivedAt",        grp: "entry",  label: "ACT เข้า",   align: "center", get: t => t.arrivedAt },
    { id: "entryDelta",       grp: "entry",  label: "เข้าเร็ว/ช้า",   align: "center", get: t => {
      const q = getQ(t);
      if (!q?.entryTime || !t.arrivedAt) return null;
      return workTimeValue(t.arrivedAt) - workTimeValue(q.entryTime);
    } },
    { id: "pickingAt",        grp: "entry",  label: "พิมพ์ใบเบิก", align: "center", get: t => t.pickingAt },
    { id: "extraStatus",      grp: "entry",  label: "สถานะเพิ่มเติม", align: "left", type: "text", get: t => t.extraStatus || "—", title: t => t.extraStatusAt ? `ตั้งเมื่อ ${t.extraStatusAt}` : undefined },
    ...laneCols("lane_parts", "parts"),
    ...laneCols("lane_head",  "head"),
    ...laneCols("lane_pork",  "pork"),
    { id: "summaryPrintedAt", grp: "docs",   label: "ใบสรุป",         align: "center", get: t => t.summaryPrintedAt },
    { id: "invoicedAt",       grp: "docs",   label: "Invoice",         align: "center", get: t => t.invoicedAt },
    { id: "exitSTD",          grp: "exit",   label: "STD ออก",         align: "center", get: t => getQ(t)?.exitTime },
    { id: "exitDelta",        grp: "exit",   label: "ออกเร็ว/ช้า",    align: "center", get: t => {
      const q = getQ(t);
      if (!q?.exitTime || !t.invoicedAt) return null;
      return workTimeValue(t.invoicedAt) - workTimeValue(q.exitTime);
    } },
  ];

  const TEXT_COL_IDS = COLS.filter(c => c.type === "text").map(c => c.id);
  const isTimeCol = id => !["plate", ...TEXT_COL_IDS].includes(id);

  const handleSort = (id) => {
    setSortCol(c => { if (c === id) { setSortDir(d => d === 1 ? -1 : 1); return c; } setSortDir(1); return id; });
  };

  const sortVal = (t, colId) => {
    const col = COLS.find(c => c.id === colId);
    const v = col?.get(t);
    if (v == null || v === "" || v === "—") return sortDir === 1 ? "zz" : "";
    if (colId === "entryDelta" || colId === "exitDelta") {
      const diff = Number(v);
      return String(diff + 10000).padStart(5, "0");
    }
    if (isTimeCol(colId)) {
      const n = workTimeValue(v);
      return String(n).padStart(5, "0");
    }
    return String(v);
  };

  const sorted = [...activeTrucks].sort((a, b) => {
    const va = sortVal(a, sortCol), vb = sortVal(b, sortCol);
    return va < vb ? -sortDir : va > vb ? sortDir : 0;
  });

  const thBase = { padding: "7px 10px", fontWeight: 700, whiteSpace: "nowrap", userSelect: "none", borderRight: "1px solid rgba(255,255,255,0.18)" };

  const deltaLabel = diff => {
    if (diff == null) return "";
    const late = diff > 0;
    return `${late ? "-" : "+"}${formatMinsDelta(Math.abs(diff))}`;
  };

  // สร้างจาก COLS โดยตรง (single source of truth เดียวกับตารางบนจอ) — คอลัมน์ใหม่ที่เพิ่มใน COLS
  // ต่อไปจะโผล่ใน Export Excel ให้อัตโนมัติเสมอ ไม่ต้องมาแก้ export แยกอีกจุดแล้วลืมอัพเดตให้ตรงกัน
  // คอลัมน์ของ 3 ลาน (parts/head/pork) ใช้ label ซ้ำกัน (แยกกันด้วยกลุ่มบนจอ) — ต้องเติมชื่อลาน
  // นำหน้าตอน export ไม่งั้น key ใน object ชนกัน คอลัมน์ของลานหลังจะเขียนทับลานก่อนหน้า
  const GRP_EXPORT_PREFIX = { parts: "ชิ้นส่วน", head: "หัว/เครื่องใน", pork: "หมูซีก" };
  const exportExcel = () => {
    const rows = sorted.map(t => {
      const row = {};
      for (const col of COLS) {
        const key = GRP_EXPORT_PREFIX[col.grp] ? `${GRP_EXPORT_PREFIX[col.grp]} - ${col.label}` : col.label;
        const v = col.get(t);
        if (col.id === "entryDelta" || col.id === "exitDelta") row[key] = deltaLabel(v);
        else row[key] = (v == null || v === "—") ? "" : v;
      }
      row["สถานะ"] = t.status || "";
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, date);
    XLSX.writeFile(wb, `Tracking_${date}.xlsx`);
  };

  // ── สรุปสถิติ: รถเข้า/ออกตรงเวลา, เวลาเฉลี่ยอยู่ในโรงงาน ──
  // ตัวช่วยทั่วไป — ใช้คำนวณได้ทั้งวันที่กำลังดู และวันก่อนหน้า (สำหรับ delta)
  const computeStats = (trucksArr, queueArr) => {
    const gq = t => queueArr.find(q => q.id === t.queueId) || queueArr.find(q => pNum(q.plate) === pNum(t.plate) && pNum(q.plate) !== "");
    const entered = trucksArr.filter(t => t.arrivedAt);
    const entryDiffs = entered
      .map(t => { const q = gq(t); return q?.entryTime ? workTimeValue(t.arrivedAt) - workTimeValue(q.entryTime) : null; })
      .filter(d => d != null);
    const enteredOnTime = entryDiffs.filter(d => d <= 0).length;

    const exited = trucksArr.filter(t => t.invoicedAt);
    const exitDiffs = exited
      .map(t => { const q = gq(t); return q?.exitTime ? workTimeValue(t.invoicedAt) - workTimeValue(q.exitTime) : null; })
      .filter(d => d != null);
    const exitedOnTime = exitDiffs.filter(d => d <= 0).length;
    const exitedLate = exitDiffs.filter(d => d > 0).length;

    const stays = trucksArr
      .filter(t => t.arrivedAt && t.invoicedAt)
      .map(t => workTimeValue(t.invoicedAt) - workTimeValue(t.arrivedAt))
      .filter(d => d >= 0);
    const avgStay = stays.length ? Math.round(stays.reduce((a, b) => a + b, 0) / stays.length) : null;

    return {
      enteredCount: entered.length,
      enteredOnTimeCount: enteredOnTime,
      enteredOnTimePct: entryDiffs.length ? Math.round(enteredOnTime / entryDiffs.length * 100) : null,
      avgStay,
      exitedOnTimeCount: exitedOnTime,
      exitedOnTimePct: exitDiffs.length ? Math.round(exitedOnTime / exitDiffs.length * 100) : null,
      exitedLateCount: exitedLate,
      exitedLatePct: exitDiffs.length ? Math.round(exitedLate / exitDiffs.length * 100) : null,
    };
  };

  // แบ่งวันทำงาน (24 ชม. นับจากเวลาตัดรอบ) เป็น 8 ช่วง ๆ ละ 3 ชม. สำหรับ mini bar chart
  const bucketize = (items, timeFn, valueFn = null, buckets = 8) => {
    const startMin = settings.workDayCutoffHour * 60;
    const bucketMin = (24 * 60) / buckets;
    const sums = Array(buckets).fill(0);
    const counts = Array(buckets).fill(0);
    for (const it of items) {
      const time = timeFn(it);
      if (!time) continue;
      const idx = Math.min(buckets - 1, Math.max(0, Math.floor((workTimeValue(time) - startMin) / bucketMin)));
      counts[idx]++;
      sums[idx] += valueFn ? valueFn(it) : 1;
    }
    return valueFn ? sums.map((s, i) => (counts[i] ? Math.round(s / counts[i]) : 0)) : sums;
  };

  const cur = computeStats(activeTrucks, activeQueue);
  const prevTrucks = prevArchiveData?.trucks ?? [];
  const prevQueue  = prevArchiveData?.queue  ?? [];
  const prev = computeStats(prevTrucks, prevQueue);

  const avgStayLabel = cur.avgStay != null ? `${Math.floor(cur.avgStay / 60)} ชม ${cur.avgStay % 60} นาที` : "—";
  const prevAvgStayLabel = prev.avgStay != null ? formatMinsDelta(prev.avgStay) : "—";

  const staysWithEntry = activeTrucks.filter(t => t.arrivedAt && t.invoicedAt);

  const statCards = [
    {
      type: "entry", icon: "truck", accent: "#2a78d6", title: "รถเข้าโรงงาน (รวม / ตรงเวลา)",
      count: cur.enteredCount, prevCount: prev.enteredCount,
      onTimePct: cur.enteredOnTimePct, onTimeCount: cur.enteredOnTimeCount, prevOnTimePct: prev.enteredOnTimePct,
      sparkline: bucketize(activeTrucks.filter(t => t.arrivedAt), t => t.arrivedAt),
    },
    {
      icon: "clock", accent: "#eda100", title: "เวลาเฉลี่ยอยู่ในโรงงาน",
      value: avgStayLabel, unit: "", prevValue: prevAvgStayLabel,
      sparkline: bucketize(staysWithEntry, t => t.arrivedAt, t => workTimeValue(t.invoicedAt) - workTimeValue(t.arrivedAt)),
    },
    {
      type: "exit", icon: "exit", accent: "#4a3aa7", title: "รถออก (ตรงเวลา / สาย)",
      onTimePct: cur.exitedOnTimePct, onTimeCount: cur.exitedOnTimeCount, prevOnTimePct: prev.exitedOnTimePct,
      lateCount: cur.exitedLateCount, latePct: cur.exitedLatePct, prevLateCount: prev.exitedLateCount, prevLatePct: prev.exitedLatePct,
      sparkline: bucketize(activeTrucks.filter(t => t.invoicedAt), t => t.invoicedAt),
    },
  ];

  const modeBtn = (m, label) => (
    <button onClick={() => setDateMode(m)}
      style={{ background: dateMode === m ? "#111827" : "#fff", color: dateMode === m ? "#fff" : "#111827", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "6px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
      {label}
    </button>
  );

  return (
    <div>
      {/* Top bar */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <h2 style={{ margin: 0, fontWeight: 900, fontSize: 20 }}>Tracking การทำงาน</h2>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {modeBtn("single", "รายวัน")}
          {modeBtn("range", "ช่วงวันที่")}
          {modeBtn("all", "ทั้งหมด")}
          {dateMode === "single" && <>
            <button onClick={() => setDate(today)}
              style={{ background: date === today ? "#2563eb" : "#fff", color: date === today ? "#fff" : "#111827", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "6px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
              วันนี้
            </button>
            <button onClick={() => setDate(addDaysToDateStr(today, -1))}
              style={{ background: date === addDaysToDateStr(today, -1) ? "#2563eb" : "#fff", color: date === addDaysToDateStr(today, -1) ? "#fff" : "#111827", border: "1.5px solid #d1d5db", borderRadius: 0, padding: "6px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
              เมื่อวาน
            </button>
            <input type="date" value={date} onChange={e => setDate(e.target.value)}
              style={{ border: "1.5px solid #d1d5db", borderRadius: 0, padding: "7px 11px", fontSize: 13, fontWeight: 600, outline: "none" }} />
            <button onClick={exportExcel}
              style={{ background: "#16a34a", color: "#fff", border: "none", borderRadius: 0, padding: "7px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>
              Export Excel
            </button>
          </>}
          {dateMode === "range" && <>
            <input type="date" value={rangeStart} onChange={e => setRangeStart(e.target.value)}
              style={{ border: "1.5px solid #d1d5db", borderRadius: 0, padding: "7px 11px", fontSize: 13, fontWeight: 600, outline: "none" }} />
            <span style={{ color: "#9ca3af" }}>ถึง</span>
            <input type="date" value={rangeEnd} onChange={e => setRangeEnd(e.target.value)}
              style={{ border: "1.5px solid #d1d5db", borderRadius: 0, padding: "7px 11px", fontSize: 13, fontWeight: 600, outline: "none" }} />
          </>}
        </div>
      </div>

      {dateMode !== "single" && (
        <DailyHealthTable rows={rangeHealth} loading={loadingRange} onViewDate={d => { setDate(d); setDateMode("single"); }} />
      )}

      {dateMode === "single" && <>
      {archiveData?.is_reconstructed && (
        <div style={{ background: "#fffbeb", color: "#92400e", border: "1.5px solid #fde68a", borderRadius: 0, padding: "10px 14px", fontSize: 13, fontWeight: 700, marginBottom: 12 }}>
          🛠️ ข้อมูลวันนี้เป็นข้อมูลกู้คืนบางส่วน (ไม่ใช่ต้นฉบับ) — {archiveData.reconstructed_note || "กู้จาก R2 photo timestamp หลังข้อมูลเดิมถูกเขียนทับ ไม่มี STD/ใบเบิก/Invoice/กลุ่มลูกค้า"}
        </div>
      )}
      {integrityWarning && (
        <div style={{ background: "#fef2f2", color: "#991b1b", border: "1.5px solid #fecaca", borderRadius: 0, padding: "10px 14px", fontSize: 13, fontWeight: 700, marginBottom: 12 }}>
          ⚠️ {integrityWarning}
        </div>
      )}

      <DataHealthPanel date={date} today={today} activeTrucks={activeTrucks} activeQueue={activeQueue} audit={dayAudit} />

      {/* Summary stat cards */}
      {!loadingArchive && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 14, marginBottom: 16 }}>
          {statCards.map(c => {
            if (c.type === "entry") return (
              <EntryStatTile key={c.title} icon={c.icon} accent={c.accent} title={c.title}
                count={c.count} prevCount={c.prevCount} onTimePct={c.onTimePct} onTimeCount={c.onTimeCount} prevOnTimePct={c.prevOnTimePct}
                sparkline={c.sparkline} />
            );
            if (c.type === "exit") return (
              <ExitStatTile key={c.title} icon={c.icon} accent={c.accent} title={c.title}
                onTimePct={c.onTimePct} onTimeCount={c.onTimeCount} prevOnTimePct={c.prevOnTimePct} lateCount={c.lateCount} latePct={c.latePct} prevLatePct={c.prevLatePct}
                sparkline={c.sparkline} />
            );
            return (
              <StatTile key={c.title} icon={c.icon} accent={c.accent} title={c.title} value={c.value} unit={c.unit}
                prevValue={c.prevValue} sparkline={c.sparkline} />
            );
          })}
        </div>
      )}

      {loadingArchive && <div style={{ textAlign: "center", color: "#9ca3af", padding: 40 }}>กำลังโหลด...</div>}

      {!loadingArchive && (
        <div style={{ overflowX: "auto", boxShadow: "0 4px 16px rgba(0,0,0,0.12)", borderRadius: 0 }}>
          <table style={{ borderCollapse: "collapse", minWidth: 980, width: "100%", fontSize: 12 }}>
            <thead>
              {/* Group header row */}
              <tr>
                {WT_GROUPS.map(g => (
                  <th key={g.id} colSpan={g.span}
                    style={{ ...thBase, background: g.dark, color: "#fff", fontSize: 11, textAlign: "center", padding: "9px 10px", borderBottom: "1px solid rgba(255,255,255,0.12)" }}>
                    {g.label}
                  </th>
                ))}
              </tr>
              {/* Column header row */}
              <tr>
                {COLS.map(col => {
                  const g = WT_GROUP_MAP[col.grp];
                  const sorted_ = sortCol === col.id;
                  return (
                    <th key={col.id} onClick={() => handleSort(col.id)}
                      style={{ ...thBase, background: g.mid, color: "#fff", fontSize: 11, textAlign: col.align, cursor: "pointer", padding: "7px 10px" }}>
                      {col.label}{sorted_ ? (sortDir === 1 ? " ▲" : " ▼") : ""}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {sorted.length === 0 && (
                <tr><td colSpan={COLS.length} style={{ textAlign: "center", color: "#9ca3af", padding: 40, background: "#fff" }}>{date === today ? "ยังไม่มีข้อมูลในวันนี้" : `ไม่มีข้อมูลของวันที่ ${date}`}</td></tr>
              )}
              {sorted.map((t, i) => {
                const q = getQ(t);
                const stdDiff = q?.entryTime && t.arrivedAt ? workTimeValue(t.arrivedAt) - workTimeValue(q.entryTime) : null;
                return (
                  <tr key={t.id} style={{ borderBottom: "1px solid #e5e7eb" }}
                    onMouseEnter={e => e.currentTarget.style.filter = "brightness(0.96)"}
                    onMouseLeave={e => e.currentTarget.style.filter = ""}>
                    {COLS.map(col => {
                      const cellBg = i % 2 === 0 ? WT_CELL_BG[col.grp] : "#fff";
                      const val = col.get(t);

                      if (col.id === "plate") return (
                        <td key={col.id} style={{ padding: "9px 12px", background: cellBg, borderRight: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>
                          <span style={{ fontWeight: 900, fontSize: 13, letterSpacing: 0.5 }}>{val}</span>
                        </td>
                      );

                      if (col.type === "text") return (
                        <td key={col.id} title={col.title ? col.title(t) : (val && val !== "—" ? val : undefined)}
                          style={{ padding: "9px 10px", background: cellBg, borderRight: "1px solid #e5e7eb", maxWidth: 110, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", textAlign: col.align || "left" }}>
                          <span style={{ fontSize: 12, color: "#374151" }}>{val || "—"}</span>
                        </td>
                      );

                      if (col.id === "arrivedAt") {
                        const late = stdDiff != null && stdDiff > 0;
                        const early = stdDiff != null && stdDiff < 0;
                        return (
                          <td key={col.id} style={{ padding: "7px 10px", background: cellBg, borderRight: "1px solid #e5e7eb", textAlign: "center", whiteSpace: "nowrap" }}>
                            {val
                              ? <span style={{ fontWeight: 700, color: "#2563eb", fontSize: 13 }}>{val}</span>
                              : <span style={{ color: "#d1d5db" }}>—</span>}
                          </td>
                        );
                      }

                      if (col.id === "entryDelta") {
                        const late = stdDiff != null && stdDiff > 0;
                        const early = stdDiff != null && stdDiff < 0;
                        return (
                          <td key={col.id} style={{ padding: "7px 10px", background: cellBg, borderRight: "1px solid #e5e7eb", textAlign: "center", whiteSpace: "nowrap" }}>
                            {stdDiff != null
                              ? <span>
                                  <span style={{ fontWeight: 700, color: late ? "#dc2626" : early ? "#16a34a" : "#1d4ed8", fontSize: 13 }}>
                                    {late ? "-" : "+"}
                                  </span>
                                  <span style={{ fontWeight: 700, color: late ? "#dc2626" : early ? "#16a34a" : "#1d4ed8", fontSize: 13 }}>
                                    {formatMinsDelta(Math.abs(stdDiff))}
                                  </span>
                                </span>
                              : <span style={{ color: "#e5e7eb" }}>—</span>}
                          </td>
                        );
                      }

                      if (col.id === "exitDelta") {
                        const q = getQ(t);
                        const exitDiff = q?.exitTime && t.invoicedAt ? workTimeValue(t.invoicedAt) - workTimeValue(q.exitTime) : null;
                        const late = exitDiff != null && exitDiff > 0;
                        const early = exitDiff != null && exitDiff < 0;
                        return (
                          <td key={col.id} style={{ padding: "7px 10px", background: cellBg, borderRight: "1px solid #e5e7eb", textAlign: "center", whiteSpace: "nowrap" }}>
                            {exitDiff != null
                              ? <span>
                                  <span style={{ fontWeight: 700, color: late ? "#dc2626" : early ? "#16a34a" : "#1d4ed8", fontSize: 13 }}>
                                    {late ? "-" : "+"}
                                  </span>
                                  <span style={{ fontWeight: 700, color: late ? "#dc2626" : early ? "#16a34a" : "#1d4ed8", fontSize: 13 }}>
                                    {formatMinsDelta(Math.abs(exitDiff))}
                                  </span>
                                </span>
                              : <span style={{ color: "#e5e7eb" }}>—</span>}
                          </td>
                        );
                      }

                      // ลานที่ไม่เกี่ยวข้องกับทะเบียนนี้ (ตามข้อมูล PO ที่อัพโหลด) — เทาอ่อน
                      if (GRP_LANE[col.grp] && irrelevantGrps(t).has(col.grp)) {
                        return <td key={col.id} style={{ padding: "7px 10px", background: "#e5e7eb", borderRight: "1px solid #e5e7eb" }} />;
                      }

                      // Generic time cell
                      const g = WT_GROUP_MAP[col.grp];
                      const cellTitle = col.title ? col.title(t) : undefined;
                      return (
                        <td key={col.id} title={cellTitle} style={{ padding: "7px 10px", background: cellBg, borderRight: "1px solid #e5e7eb", textAlign: "center", whiteSpace: "nowrap" }}>
                          {val
                            ? <span style={{ fontWeight: 700, color: g.mid, fontSize: 13 }}>{val}{cellTitle ? " 🛈" : ""}</span>
                            : <span style={{ color: "#e5e7eb" }}>—</span>}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      </>}
    </div>
  );
};

// ─── ROLE SELECT (landing page: เลือกตำแหน่งงานก่อนเข้าระบบ) ──────────────────
// label/emoji/img ของแต่ละตำแหน่งงานมาจาก wh_roles (ดู src/lib/masterData.js) — id
// คงที่เสมอเพราะ ROLE_TABS/LANE_SELECT_ROLES ผูก logic กับ id เหล่านี้ตรงๆ

const LOADING_TABS = ["loading_parts", "loading_head", "loading_pork"];

// roles that land on a lane-picker card screen instead of auto-opening a tab
const LANE_SELECT_ROLES = ["qc", "loading", "checker", "office_wh", "office_plan", "lg"];

const ROLE_TABS = {
  qc:           ["qc_parts", "qc_head", "qc_pork"],
  loading:      [...LOADING_TABS],
  checker:      ["sample_parts", "sample_head", "sample_pork"],
  office_wh:    ["picking"],
  office_plan:  ["planning", "detail_loading"],
  lg:           ["lg"],
  dashboard_only: ["dashboard", "dashboard_transport"],
  loading_data: ["overview_log"],
  tracking:     ["work_tracking"],
  all:          null,
};

// QC/ลานโหลด/Checker tabs each work a single lane — used to hide a tab's menu card
// (LaneSelect) and to block its kiosk URL (?mode=...) when that lane is ปิดใช้งาน
// (wh_lanes.enabled === false) ใน Master Setting > ลานโหลด (Lanes)
const LANE_ENTRY_TAB_TO_LANE_ID = {
  qc_parts: "lane_parts", qc_head: "lane_head", qc_pork: "lane_pork",
  loading_parts: "lane_parts", loading_head: "lane_head", loading_pork: "lane_pork",
  sample_parts: "lane_parts", sample_head: "lane_head", sample_pork: "lane_pork",
};

const RoleSelect = ({ onSelect }) => {
  const squareOptions = roles.slice(0, 6);
  const wideOptions   = roles.slice(6);
  return (
    <div style={{ height: "100dvh", overflow: "hidden", background: "#f1f5f9", display: "flex", alignItems: "center", justifyContent: "center", padding: "16px 20px", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
      <div style={{ maxWidth: 420, width: "100%" }}>
        <h2 style={{ textAlign: "center", fontWeight: 900, fontSize: 20, margin: "0 0 4px" }}>เลือกตำแหน่งงาน</h2>
        <p style={{ textAlign: "center", color: "#6b7280", fontSize: 12, margin: "0 0 28px" }}>ระบบจะแสดงเฉพาะเมนูที่เกี่ยวข้องกับตำแหน่งงานของคุณ</p>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          {squareOptions.map(r => (
            <button key={r.id} onClick={() => onSelect(r.id)}
              style={{ background: "#fff", border: "1.5px solid #e5e7eb", borderRadius: 0, aspectRatio: "1 / 1", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 6, padding: 8, fontSize: 15, fontWeight: 700, cursor: "pointer", textAlign: "center" }}>
              {r.img ? <img src={r.img} alt="" style={{ width: 36, height: "auto" }} /> : r.icon ? <Icon name={r.icon} size={26} /> : <span style={{ fontSize: 26 }}>{r.emoji}</span>}
              {r.label}
            </button>
          ))}
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 24 }}>
          {wideOptions.map(r => (
            <button key={r.id} onClick={() => onSelect(r.id)}
              style={{ background: "#fff", border: "1.5px solid #e5e7eb", borderRadius: 0, padding: "14px 8px", fontSize: 15, fontWeight: 700, cursor: "pointer", textAlign: "center" }}>
              {r.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};

// ─── LANE SELECT (เลือกเมนูย่อยหลังเลือกตำแหน่งงาน) ───────────────────────────
const LaneSelect = ({ tabs, roleLabel, onSelect, onBack }) => (
  <div style={{ height: "100dvh", overflow: "hidden", background: "#f1f5f9", display: "flex", alignItems: "center", justifyContent: "center", padding: "16px 20px", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
    <div style={{ maxWidth: 420, width: "100%" }}>
      <h2 style={{ textAlign: "center", fontWeight: 900, fontSize: 20, margin: "0 0 4px" }}>{roleLabel}</h2>
      <p style={{ textAlign: "center", color: "#6b7280", fontSize: 12, margin: "0 0 28px" }}>เลือกเมนูที่ต้องการเข้าใช้งาน</p>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {tabs.map(t => (
          <button key={t.id} onClick={() => onSelect(t.id)}
            style={{ background: "#fff", border: "1.5px solid #e5e7eb", borderRadius: 0, padding: "16px 14px", fontSize: 15, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 12 }}>
            <Icon name={t.icon} size={22} />
            {t.label}
          </button>
        ))}
      </div>
      <button onClick={onBack} style={{ marginTop: 24, width: "100%", background: "transparent", border: "none", color: "#6b7280", fontSize: 13, fontWeight: 600, cursor: "pointer", textAlign: "center" }}>
        ← กลับหน้าเลือกตำแหน่งงาน
      </button>
    </div>
  </div>
);

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────
const fetchQueue  = async () => { const { data } = await supabase.from("wh_queue").select("*");  return (data || []).map(r => r.data).sort((a, b) => (a.seq ?? Infinity) - (b.seq ?? Infinity)); };
// กรอง is_deleted ออก (soft-delete, ดู supabase-add-soft-delete-trucks.sql) — ถ้ายังไม่ได้รัน
// migration นั้น คอลัมน์นี้จะยังไม่มี ให้ fallback ไป select แบบเดิมแทนเพื่อไม่ให้แอปพังทั้งระบบ
const fetchTrucks = async () => {
  let res = await supabase.from("wh_trucks").select("*").eq("is_deleted", false);
  if (res.error) res = await supabase.from("wh_trucks").select("*");
  if (res.error) throw res.error;
  return (res.data || []).map(r => r.data);
};
const fetchMaster = async () => { const { data } = await supabase.from("wh_master").select("*").eq("id", "master"); return data && data[0] ? (data[0].data || []) : []; };
const fetchPendingQueue = async () => { const { data } = await supabase.from("wh_master").select("*").eq("id", "pending_queue"); return data && data[0] ? (data[0].data || null) : null; };
const fetchDetailSrc = async (date = cycleDateStr()) => {
  const ids = detailSources.map(s => `detail_${s.id}_${date}`);
  const { data } = await supabase.from("wh_master").select("*").in("id", ids);
  if (!data || data.length === 0) return null;
  const result = {};
  for (const row of data) {
    const src = detailSources.find(s => `detail_${s.id}_${date}` === row.id);
    if (!src) continue;
    const payload = row.data || {};
    result[src.id] = {
      rows:     Array.isArray(payload) ? payload : (payload.rows || []),
      fileName: payload.file_name || "",
    };
  }
  return result;
};

export default function App() {
  const isMobile = useIsMobile();
  const isNarrow = useIsNarrow();
  const defaultTabForRole = (r) => {
    if (LANE_SELECT_ROLES.includes(r)) {
      const allowed = ROLE_TABS[r];
      return allowed?.length === 1 ? allowed[0] : "";
    }
    const allowed = ROLE_TABS[r];
    return (allowed && !allowed.includes("dashboard")) ? allowed[0] : "dashboard";
  };
  const [role,       setRole]       = useState("");
  const [prevTab,    setPrevTab]    = useState(null); // tab to return to when backing out of ช่องโหลด selection
  const handleSelectRole = (r) => { setRole(r); setTab(defaultTabForRole(r)); setPrevTab(null); };
  const handleChangeRole = () => { setRole(""); setTab("dashboard"); setPrevTab(null); };
  const [queue,      setQueue]      = useState([]);
  const [trucks,     setTrucks]     = useState([]);
  const [pendingQueue, setPendingQueue] = useState(null); // คิวรถที่ LG อัพโหลดล่วงหน้า รอถูกดึงมาใช้เป็นคิวปัจจุบัน
  const [dataWarning, setDataWarning] = useState(null); // แจ้งเตือนจาก close_work_day_tx เมื่อพบข้อมูลลดลงผิดปกติตอนปิดวันทำงาน
  const [masterLane, setMasterLane] = useState(() => {
    try { return JSON.parse(localStorage.getItem("wh_master_cache") || "[]"); } catch { return []; }
  });
  const [myPlate, setMyPlate] = useState(() => localStorage.getItem("wh_my_plate") || "");
  const [detailMapByChannel, setDetailMapByChannel] = useState({}); // { channelId: plate→Set(lanes) }
  const [srcVersion, setSrcVersion] = useState(0);  // bumped when source files change
  const [tab,        setTab]        = useState("dashboard");
  const [dashLane,   setDashLane]   = useState("main");
  const [time,       setTime]       = useState(TIME_NOW());
  const [loading,    setLoading]    = useState(true);
  const [headerClock, setHeaderClock] = useState(() => new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", second: "2-digit" }));

  // Kiosk modes via URL parameter ?mode=<role>
  const _urlMode = typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("mode") : null;
  const isDriverMode      = _urlMode === "driver";
  const isQcPartsMode      = _urlMode === "qc_parts";
  const isQcHeadMode       = _urlMode === "qc_head";
  const isQcPorkMode       = _urlMode === "qc_pork";
  const isLoadingPartsMode = _urlMode === "loading_parts";
  const isLoadingHeadMode  = _urlMode === "loading_head";
  const isLoadingPorkMode  = _urlMode === "loading_pork";
  const isSamplePartsMode  = _urlMode === "sample_parts";
  const isSampleHeadMode   = _urlMode === "sample_head";
  const isSamplePorkMode   = _urlMode === "sample_pork";
  const isDashboardTransportMode = _urlMode === "dashboard_transport";

  useEffect(() => { const id = setInterval(() => setTime(TIME_NOW()), 15000); return () => clearInterval(id); }, []);
  useEffect(() => { const id = setInterval(() => setHeaderClock(new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", second: "2-digit" })), 1000); return () => clearInterval(id); }, []);

  useEffect(() => {
    fetchQueue().then(setQueue);
    fetchTrucks().then(setTrucks).catch(err => console.error("fetchTrucks failed:", err));
    fetchPendingQueue().then(setPendingQueue);
    fetchMaster().then(rows => {
      if (rows.length > 0) {
        setMasterLane(rows);
        localStorage.setItem("wh_master_cache", JSON.stringify(rows));
      }
    });

    const channel = supabase.channel("app-sync")
      .on("postgres_changes", { event: "*", schema: "public", table: "wh_queue" },  () => fetchQueue().then(setQueue))
      .on("postgres_changes", { event: "*", schema: "public", table: "wh_trucks" }, () => fetchTrucks().then(setTrucks).catch(err => console.error("fetchTrucks failed:", err)))
      .on("postgres_changes", { event: "*", schema: "public", table: "wh_master", filter: "id=eq.pending_queue" }, () => fetchPendingQueue().then(setPendingQueue))
      .subscribe();

    return () => supabase.removeChannel(channel);
  }, []);

  // คิวรถล่วงหน้า (อัพจาก LG) → ดึงมาใช้เป็นคิวปัจจุบันอัตโนมัติทันทีที่ข้ามเวลาตัดรอบวันทำงาน
  // เช็คทุก 60 วิ (ไม่มี cron ฝั่ง server จึงอาศัยเครื่องที่เปิดหน้านี้ค้างไว้เป็นตัวจับเวลาแทน) —
  // ใช้ >= แทน === เทียบวันที่ กันกรณีไม่มีเครื่องไหนเปิดหน้าทิ้งไว้ตอนข้ามรอบพอดี (เช่น วันหยุดยาว)
  // แล้วพลาดจังหวะไปเลย
  useEffect(() => {
    const promoteIfDue = async () => {
      try {
        const pending = await fetchPendingQueue();
        if (!pending || !Array.isArray(pending.rows) || !pending.forDate) return;
        if (pending.forDate > cycleDateStr()) return;

        // วันที่ปิดจริงคือ "วันก่อนหน้าที่คิวใหม่จะมีผล" เสมอ ไม่ใช่ cycleDateStr() ปัจจุบัน —
        // เพราะถ้าไม่มีเครื่องไหนเปิดค้างไว้พอดีตอนข้ามรอบ กว่าจะมีเครื่องมาเช็คซ้ำอีกที cycleDateStr()
        // อาจเลยวันที่ตั้งใจไว้ไปแล้วหลายวัน คำนวณจาก forDate จึงได้วันที่ถูกต้องเสมอ
        const archiveDate = addDaysToDateStr(pending.forDate, -1);

        // ปิดวันทำงานผ่าน RPC เดียว (close_work_day_tx, ดู supabase-fix-archive-race-condition.sql) —
        // ทำ archive+wipe+ใส่คิวใหม่ทั้งหมดในทรานแซกชันเดียวฝั่ง DB พร้อม advisory lock กันชนกับ
        // close_work_day() (cron) หรือ promoteIfDue() ของแท็บ/เครื่องอื่นที่เปิดค้างไว้พร้อมกัน —
        // เดิมโค้ดนี้ทำ read (client) แล้วค่อย upsert/delete/insert แยก 4 ก้าว ซึ่งมีช่วงเวลาที่อีก
        // แท็บหนึ่งมาแทรกกลางได้ ทำให้ archive ของวันที่ปิดไปแล้วถูกเขียนทับด้วยคิววันใหม่ที่เพิ่ง
        // promote เข้าไป (สาเหตุที่ข้อมูลวันที่ 6-8 ก.ย. หายไปจริง)
        const { data: result, error: rpcErr } = await supabase.rpc("close_work_day_tx", {
          p_archive_date: archiveDate,
          p_new_queue: pending.rows,
          p_source: "client_promote",
        });
        // RPC ยังไม่มี (ยังไม่ได้รัน SQL migration) หรือ error อื่น — ห้ามลบข้อมูลจริงทิ้งโดยไม่มี RPC
        // คุ้มครอง ปล่อยผ่านรอบนี้ไป รอบถัดไป (60 วิ) จะลองใหม่ เพราะยังไม่ได้แตะข้อมูลอะไรเลย
        if (rpcErr) { console.error("promoteIfDue: close_work_day_tx ไม่สำเร็จ, ข้ามรอบนี้", rpcErr); return; }
        if (result?.warning) setDataWarning(result.warning);

        await supabase.from("wh_master").delete().eq("id", "pending_queue");

        fetchQueue().then(setQueue);
        fetchTrucks().then(setTrucks).catch(() => {});
        setPendingQueue(null);
      } catch (e) {
        console.error("promoteIfDue (pending queue) failed:", e);
      }
    };
    promoteIfDue();
    const id = setInterval(promoteIfDue, 60000);
    return () => clearInterval(id);
  }, []);

  // Recompute detailMapByChannel whenever masterLane or source files change
  useEffect(() => {
    if (!masterLane.length) return;
    try {
      const stored = JSON.parse(localStorage.getItem("wh_detail_src") || "{}");
      setDetailMapByChannel(buildDetailMapByChannel(masterLane, stored));
    } catch {}
  }, [masterLane, srcVersion]);

  const handleMasterChange = async (rows) => {
    setMasterLane(rows);
    localStorage.setItem("wh_master_cache", JSON.stringify(rows));
    await supabase.from("wh_master").upsert({ id: "master", data: rows });
  };

  const handleDetailChange = () => {
    setSrcVersion(v => v + 1);
  };

  const plateNum = s => (String(s).match(/\d+/g) || []).pop() || "";

  const calcTimeDiffStr = (std, actual) => {
    if (!std || !actual) return "";
    const [h1, m1] = std.split(":").map(Number);
    const [h2, m2] = actual.split(":").map(Number);
    if (isNaN(h1) || isNaN(h2)) return "";
    const diffMin = (h2 * 60 + m2) - (h1 * 60 + m1);
    if (diffMin === 0) return "(ตรงเวลา)";
    const absDiff = Math.abs(diffMin);
    const hrs = Math.floor(absDiff / 60);
    const mins = absDiff % 60;
    const str = `${hrs}:${mins.toString().padStart(2, '0')} ชม.`;
    return diffMin < 0 ? `(ก่อน ${str})` : `(สาย ${str})`;
  };

  const handleScan = async (t) => {
    if (t.id.startsWith("WALK-")) {
      const qData = {
        id: t.id,
        seq: queue.length + 1,
        date: SHORT_DATE(),
        plate: t.plate,
        customerGroup: t.customerGroup || "",
        zone: t.zone || "",
        entryTime: t.entryTime || TIME_NOW(),
        exitTime: t.exitTime || "",
        time: t.entryTime || TIME_NOW()
      };
      await supabase.from("wh_queue").upsert({ id: t.id, data: qData });
      setQueue(prev => [...prev.filter(q => q.id !== t.id), qData]);
    }
    await supabase.from("wh_trucks").upsert({ id: t.id, data: t });
    setTrucks(prev => [...prev.filter(tr => tr.id !== t.id), t]);

    localStorage.setItem("wh_my_plate", t.plate);
    setMyPlate(t.plate);

    const actualTime = TIME_NOW();
    const diffStr = calcTimeDiffStr(t.entryTime, actualTime);
  };

  const handleUpdate = async (id, upd) => {
    const truck = trucks.find(t => t.id === id);
    if (!truck) throw new Error("ไม่พบข้อมูลรถคันนี้ในเครื่อง กรุณารีเฟรชหน้าจอแล้วลองใหม่");

    if (upd.loadLanes) {
       for (const lane of Object.keys(upd.loadLanes)) {
         if (upd.loadLanes[lane].done && (!truck.loadLanes || !truck.loadLanes[lane] || !truck.loadLanes[lane].done)) {
           const lName = lanes.find(l => l.id === lane)?.tinyLabel || lane;
           const imgs = upd.loadLanes[lane].photos || [];
           sendTeamsNotification(`✅ โหลดเสร็จ — รถ ${truck.plate}`, { "ลานโหลด": lName, "เวลาโหลดเสร็จ": upd.loadLanes[lane].doneAt || TIME_NOW() }, imgs);
         }
       }
    }

    const updated = { ...truck, ...upd };
    const { error } = await supabase.from("wh_trucks").upsert({ id, data: updated });
    if (error) throw error;
    setTrucks(prev => prev.map(t => t.id === id ? updated : t));

    if (upd.status === "invoiced" && myPlate && plateNum(updated.plate) === plateNum(myPlate)) {
      localStorage.removeItem("wh_my_plate");
      setMyPlate("");
    }
  };

  const handleDeleteTruck = async (id) => {
    // soft delete (ดู supabase-add-soft-delete-trucks.sql) — ไม่ DELETE จริง กันข้อมูลรถหายถาวร
    // ถ้ากดผิด/merge พลาด ยังกู้คืนได้ก่อนวันทำงานนั้นจะปิด (close_work_day_tx ไม่นับแถวนี้เข้า archive)
    const { error } = await supabase.from("wh_trucks")
      .update({ is_deleted: true, deleted_at: new Date().toISOString(), deleted_by: role || "unknown" })
      .eq("id", id);
    if (error) {
      window.alert("ลบไม่สำเร็จ — อาจยังไม่ได้รัน supabase-add-soft-delete-trucks.sql: " + error.message);
      throw error;
    }
    setTrucks(prev => prev.filter(t => t.id !== id));
  };

  const handleReset = async () => {
    if (!window.confirm("ล้างข้อมูลทั้งหมดสำหรับวันใหม่?")) return;
    // archive date = วันรอบงานที่เพิ่งปิด (same work-day cutoff as cycleDateStr, kept in sync via that helper)
    const archiveDate = cycleDateStr();
    // ผ่าน RPC เดียวกับ promoteIfDue/cron (close_work_day_tx) แทนการ upsert+delete ตรงๆ —
    // กันไม่ให้กดปุ่มนี้ชนกับ cron หรือ promoteIfDue ที่อาจกำลังปิดวันเดียวกันอยู่พอดี
    const { data: result, error } = await supabase.rpc("close_work_day_tx", {
      p_archive_date: archiveDate,
      p_new_queue: null,
      p_source: "manual_reset",
    });
    if (error) { window.alert("ล้างวันใหม่ไม่สำเร็จ: " + error.message); return; }
    if (result?.warning) setDataWarning(result.warning);
  };

  const handleSetQueue = async (newQueue, source = "lg_edit") => {
    // ผ่าน RPC เดียว (set_queue_tx, ดู supabase-fix-queue-overwrite-race.sql) แทนการ
    // delete()+upsert() แยก 2 คำสั่งจาก client เดิม — ทำใน transaction เดียวฝั่ง DB (ถ้า insert
    // fail จะ rollback การ delete ให้อัตโนมัติ) และมี guard ปฏิเสธการเขียนทับที่ดูผิดปกติ
    // (คิวหายไปเกินครึ่ง/หายหมด) ป้องกันเคสไฟล์ผิด/เน็ตหลุดระหว่างอัพโหลดทำให้คิวทั้งวันหายถาวร
    // (สาเหตุที่คาดว่าทำให้ข้อมูลวันที่ 11 ก.ย. หายไปจริง)
    const callSetQueue = (force) => supabase.rpc("set_queue_tx", {
      p_new_queue: newQueue,
      p_source: source,
      p_force: force,
    });
    let { data: result, error } = await callSetQueue(false);
    if (error) throw new Error(error.message);
    if (result?.blocked) {
      const confirmed = window.confirm(
        `${result.reason}\n\nถ้าตั้งใจบันทึกข้อมูลนี้จริง กด ตกลง เพื่อยืนยันอีกครั้ง (ถ้าไม่แน่ใจ กด ยกเลิก แล้วตรวจไฟล์/ข้อมูลก่อน)`
      );
      if (!confirmed) throw new Error("ยกเลิกการบันทึก — ระบบตรวจพบว่าข้อมูลคิวอาจไม่ครบ");
      ({ data: result, error } = await callSetQueue(true));
      if (error) throw new Error(error.message);
    }
    setQueue(newQueue);
    // merge walk-in trucks กับ queue entries แบบ one-to-one เรียงตาม seq
    const sortedQueue = [...newQueue].sort((a, b) => (a.seq ?? 0) - (b.seq ?? 0));
    const usedQueueIds = new Set();
    for (const truck of trucks) {
      const match = sortedQueue.find(q => plateNum(q.plate) === plateNum(truck.plate) && plateNum(q.plate) !== "" && !usedQueueIds.has(q.id));
      if (!match) continue;
      usedQueueIds.add(match.id);
      await supabase.from("wh_trucks").upsert({ id: truck.id, data: { ...truck, plate: match.plate, customerGroup: match.customerGroup, zone: match.zone, queueId: match.id, entryTime: match.entryTime, exitTime: match.exitTime } });
    }
  };

  // คิวรถ "ล่วงหน้า" ที่ LG อัพไว้ล่วงหน้าสำหรับวันทำงานถัดไป — เก็บแยกจากคิวปัจจุบันใน wh_master
  // แล้วให้ promoteIfDue (ใน useEffect ด้านบน) ดึงมาใช้แทนคิวปัจจุบันเองตอนข้ามเวลาตัดรอบ
  const handleSetPendingQueue = async (rows, fileName) => {
    const payload = { rows, fileName, forDate: nextCycleDateStr(), uploadedAt: new Date().toISOString() };
    const { error } = await supabase.from("wh_master").upsert({ id: "pending_queue", data: payload });
    if (error) throw new Error(error.message);
    setPendingQueue(payload);
  };

  const handleClearPendingQueue = async () => {
    const { error } = await supabase.from("wh_master").delete().eq("id", "pending_queue");
    if (error) throw new Error(error.message);
    setPendingQueue(null);
  };

  const badge = {
    driver:        queue.filter(q => !trucks.find(t => t.queueId === q.id)).length,
    picking:       trucks.filter(t => t.status === "arrived").length,
    qc_parts:      trucks.filter(t => ["arrived","picking"].includes(t.status) && !t.qcLanes?.lane_parts?.done).length,
    qc_head:       trucks.filter(t => ["arrived","picking"].includes(t.status) && !t.qcLanes?.lane_head?.done).length,
    qc_pork:       trucks.filter(t => ["arrived","picking"].includes(t.status) && !t.qcLanes?.lane_pork?.done).length,
    loading_parts: trucks.filter(t => t.status === "picking" && t.qcLanes?.lane_parts?.done && !t.loadLanes?.lane_parts?.done).length,
    loading_head:  trucks.filter(t => t.status === "picking" && t.qcLanes?.lane_head?.done  && !t.loadLanes?.lane_head?.done).length,
    loading_pork:  trucks.filter(t => t.status === "picking" && t.qcLanes?.lane_pork?.done  && !t.loadLanes?.lane_pork?.done).length,
    sample_parts:  trucks.filter(t => ["arrived","picking"].includes(t.status) && !t.sampleLanes?.lane_parts?.done).length,
    sample_head:   trucks.filter(t => ["arrived","picking"].includes(t.status) && !t.sampleLanes?.lane_head?.done).length,
    sample_pork:   trucks.filter(t => ["arrived","picking"].includes(t.status) && !t.sampleLanes?.lane_pork?.done).length,
    planning:      trucks.filter(t => t.status === "summary_printed").length,
  };

  const laneById = (id) => lanes.find(l => l.id === id) || {};
  const tabs = [
    { id: "dashboard", label: "Dashboard", icon: "chart"     },
    { id: "dashboard_transport", label: "Dashboard ขนส่ง", icon: "chart" },
    { id: "lg",        label: "LG",      icon: "upload"    },
    { id: "driver",    label: "คนขับ",   icon: "scan"      },
    { id: "picking",   label: "Picking", icon: "clipboard" },
    { id: "qc_parts",      label: `ลานโหลด ${laneById("lane_parts").tinyLabel}`,      icon: "temp"      },
    { id: "qc_head",       label: `ลานโหลด ${laneById("lane_head").tinyLabel}`, icon: "temp"      },
    { id: "qc_pork",       label: `ลานโหลด ${laneById("lane_pork").tinyLabel}`,       icon: "temp"      },
    { id: "loading_parts", label: `Checker ${laneById("lane_parts").label}`, icon: "pig_cuts"  },
    { id: "loading_head",  label: `Checker ${laneById("lane_head").label}`,  icon: "pig_head"  },
    { id: "loading_pork",  label: `Checker ${laneById("lane_pork").label}`,  icon: "pig_side"  },
    { id: "sample_parts",  label: `QC ${laneById("lane_parts").tinyLabel}`,          icon: "camera"    },
    { id: "sample_head",   label: `QC ${laneById("lane_head").tinyLabel}`,     icon: "camera"    },
    { id: "sample_pork",   label: `QC ${laneById("lane_pork").tinyLabel}`,           icon: "camera"    },
    { id: "basket_summary", label: "ข้อมูลยอดตะกร้า/ตะขอ", icon: "clipboard" },
    { id: "waiting_summary", label: "ข้อมูลการรอสินค้า", icon: "clock" },
    { id: "overview_log",   label: "Log ภาพรวมการทำงาน",  icon: "list"  },
    { id: "work_tracking",  label: "Tracking การทำงาน",  icon: "chart" },
    { id: "planning",      label: "Invoice",       icon: "plan"      },
    { id: "detail_loading", label: "อัพโหลด PO", icon: "clipboard" },
    { id: "download",       label: "จบการทำงาน",       icon: "invoice"   },
    { id: "admin",          label: "Admin",          icon: "plan"      },
    { id: "master_upload", label: "Master Setting", icon: "clipboard" },
    { id: "qr",             label: "QR Code",       icon: "scan"      },
  ];

  // ── Kiosk header helper ──
  const KioskHeader = ({ emoji, title, color = "#111" }) => (
    <div style={{ background: color, color: "#fff", padding: "0 14px", position: "sticky", top: 0, zIndex: 100, height: 56, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
      <span style={{ fontSize: 22 }}>{emoji}</span>
      <div style={{ fontWeight: 800, fontSize: 16 }}>{title}</div>
    </div>
  );

  const isKioskMode = isDriverMode || isQcPartsMode || isQcHeadMode || isQcPorkMode
    || isLoadingPartsMode || isLoadingHeadMode || isLoadingPorkMode
    || isSamplePartsMode || isSampleHeadMode || isSamplePorkMode
    || isDashboardTransportMode;
  const isLaneEnabled = (tabId) => {
    const laneId = LANE_ENTRY_TAB_TO_LANE_ID[tabId];
    return !laneId || lanes.find(l => l.id === laneId)?.enabled !== false;
  };
  const allowedTabIds = ROLE_TABS[role] || null;
  const visibleTabs = (allowedTabIds ? tabs.filter(t => allowedTabIds.includes(t.id)) : tabs).filter(t => isLaneEnabled(t.id));

  // ── Kiosk URL (?mode=qc_parts เป็นต้น) ของลานที่ถูกปิดใช้งานไว้ — QR/URL เดิมยังเปิด
  // เข้ามาได้ แต่แจ้งเตือนแทนที่จะเปิดฟอร์มกรอกข้อมูลของลานนั้น
  if (LANE_ENTRY_TAB_TO_LANE_ID[_urlMode] && !isLaneEnabled(_urlMode)) {
    const modeLabel = tabs.find(t => t.id === _urlMode)?.label || _urlMode;
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
        <div style={{ textAlign: "center", maxWidth: 360 }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>🚫</div>
          <div style={{ fontWeight: 800, fontSize: 18, marginBottom: 8 }}>ลานนี้ปิดใช้งานอยู่</div>
          <div style={{ color: "#6b7280", fontSize: 14 }}>{modeLabel} ถูกปิดใช้งานชั่วคราวจาก Master Setting กรุณาติดต่อผู้ดูแลระบบ</div>
        </div>
      </div>
    );
  }

  // ── Role select (เลือกตำแหน่งงานก่อนเข้าระบบ) ──
  if (!isKioskMode && !role) {
    return <RoleSelect onSelect={handleSelectRole} />;
  }

  // ── Lane select (เลือกเมนูย่อยหลังเลือกตำแหน่งงาน) ──
  if (!isKioskMode && role && !tab) {
    const roleLabel = roles.find(r => r.id === role)?.label || "";
    return <LaneSelect tabs={visibleTabs} roleLabel={roleLabel} onSelect={setTab} onBack={handleChangeRole} />;
  }

  // ── Driver-only mode ──
  if (isDriverMode) {
    return (
      <div style={{ minHeight: "100vh", background: "linear-gradient(180deg, #f8fafc 0%, #e2e8f0 100%)", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="🚛" title="เช็คอินคนขับ" />
        <div style={{ maxWidth: 480, margin: "0 auto", padding: "20px 14px 60px" }}>
          <DriverScan queue={queue} trucks={trucks} onScan={handleScan} />
        </div>
      </div>
    );
  }

  // ── Dashboard ขนส่ง kiosk mode ──
  if (isDashboardTransportMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="📊" title="Dashboard ขนส่ง" color="#0ea5e9" />
        <div style={{ padding: isMobile ? "8px 10px 80px" : "8px 14px 14px" }}>
          <Dashboard trucks={trucks} queue={queue} onReset={handleReset} lane={null} detailMap={detailMapByChannel} title="Dashboard ขนส่ง" myPlate={myPlate} simple />
        </div>
      </div>
    );
  }

  // ── QC kiosk modes ──
  if (isQcPartsMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="🌡️" title={tabs.find(t => t.id === "qc_parts").label} color="#0369a1" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <QC trucks={trucks} onUpdate={handleUpdate} laneId="lane_parts" detailMapByChannel={detailMapByChannel} />
        </div>
      </div>
    );
  }

  if (isQcHeadMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="🌡️" title={tabs.find(t => t.id === "qc_head").label} color="#0369a1" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <QC trucks={trucks} onUpdate={handleUpdate} laneId="lane_head" detailMapByChannel={detailMapByChannel} />
        </div>
      </div>
    );
  }

  if (isQcPorkMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="🌡️" title={tabs.find(t => t.id === "qc_pork").label} color="#0369a1" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <QC trucks={trucks} onUpdate={handleUpdate} laneId="lane_pork" detailMapByChannel={detailMapByChannel} />
        </div>
      </div>
    );
  }

  // ── Loading lane kiosk modes ──
  if (isLoadingPartsMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="🥩" title={tabs.find(t => t.id === "loading_parts").label} color="#c2410c" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <LoadingYard trucks={trucks} onUpdate={handleUpdate} laneId="lane_parts" masterLane={masterLane} />
        </div>
      </div>
    );
  }

  if (isLoadingHeadMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="🐷" title={tabs.find(t => t.id === "loading_head").label} color="#7c3aed" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <LoadingYard trucks={trucks} onUpdate={handleUpdate} laneId="lane_head" masterLane={masterLane} />
        </div>
      </div>
    );
  }

  if (isLoadingPorkMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="🐖" title={tabs.find(t => t.id === "loading_pork").label} color="#be123c" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <LoadingYard trucks={trucks} onUpdate={handleUpdate} laneId="lane_pork" masterLane={masterLane} />
        </div>
      </div>
    );
  }

  // ── Random sample temp-check kiosk modes ──
  if (isSamplePartsMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="📷" title={tabs.find(t => t.id === "sample_parts").label} color="#0d9488" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <RandomSampleCheck trucks={trucks} onUpdate={handleUpdate} laneId="lane_parts" detailMapByChannel={detailMapByChannel} />
        </div>
      </div>
    );
  }

  if (isSampleHeadMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="📷" title={tabs.find(t => t.id === "sample_head").label} color="#0d9488" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <RandomSampleCheck trucks={trucks} onUpdate={handleUpdate} laneId="lane_head" detailMapByChannel={detailMapByChannel} />
        </div>
      </div>
    );
  }

  if (isSamplePorkMode) {
    return (
      <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
        <KioskHeader emoji="📷" title={tabs.find(t => t.id === "sample_pork").label} color="#0d9488" />
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 14px 60px" }}>
          <RandomSampleCheck trucks={trucks} onUpdate={handleUpdate} laneId="lane_pork" detailMapByChannel={detailMapByChannel} />
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Sarabun','Noto Sans Thai',sans-serif" }}>
      <div style={{ background: "#111", color: "#fff", padding: "0 14px", position: "sticky", top: 0, zIndex: 100, height: 80, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flex: isMobile ? 1 : "none" }}>
          <div style={{ flexShrink: 0 }}>
            <div style={{ fontWeight: 800, fontSize: isMobile ? 12 : 14, lineHeight: 1.2 }}>ระบบโหลดสินค้า{settings.facilityName}</div>
            {isNarrow && <div style={{ fontSize: 10, color: "#9ca3af", marginTop: 2 }}>{TODAY} {headerClock}</div>}
          </div>
          <select value={tab} onChange={e => {
              const v = e.target.value;
              if (v === "__change_role__") { handleChangeRole(); return; }
              setPrevTab(tab); setTab(v); setDashLane("main");
            }}
            style={{ flex: "none", background: "#1f2937", color: "#f9fafb", border: "1px solid #374151", borderRadius: 0, padding: isMobile ? "6px 10px" : "6px 12px", fontSize: 13, fontWeight: 700, cursor: "pointer", outline: "none", marginLeft: isMobile ? "auto" : 6 }}>
            {visibleTabs.map(t => {
              const n = badge[t.id] || 0;
              return <option key={t.id} value={t.id}>{t.label}{n > 0 ? ` · ${n}` : ""}</option>;
            })}
            <option value="__change_role__">(กลับหน้าหลัก)</option>
          </select>
          {!isMobile && tab === "dashboard" && (
            <div style={{ display: "flex", gap: 2 }}>
              {[
                { id: "main",       label: "Main"           },
                { id: "lane_parts", label: "ชิ้นส่วน"       },
                { id: "lane_head",  label: "หัว/เครื่องใน"  },
                { id: "lane_pork",  label: "หมูซีก"         },
              ].map(l => {
                const active = dashLane === l.id;
                return (
                  <button key={l.id} onClick={() => setDashLane(l.id)}
                    style={{ background: "transparent", color: active ? "#fff" : "#9ca3af", border: "none", borderBottom: active ? "2px solid #fff" : "2px solid transparent", padding: "4px 10px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
                    {l.label}
                  </button>
                );
              })}
            </div>
          )}
        </div>
        {!isNarrow && <div style={{ color: "#f9fafb", fontSize: 15, fontWeight: 700, fontVariantNumeric: "tabular-nums", flexShrink: 0 }}>{TODAY} {headerClock}</div>}
      </div>
      {dataWarning && (
        <div style={{ background: "#fef3c7", color: "#92400e", padding: "10px 16px", fontSize: 13, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, borderBottom: "1.5px solid #f59e0b" }}>
          <span>⚠️ {dataWarning} — ดูรายละเอียดใน "Log ภาพรวมการทำงาน" (wh_archive_audit)</span>
          <button onClick={() => setDataWarning(null)} style={{ background: "transparent", border: "none", color: "#92400e", fontWeight: 900, fontSize: 15, cursor: "pointer", flexShrink: 0 }}>✕</button>
        </div>
      )}
      <div style={{ maxWidth: tab === "dashboard" || tab === "dashboard_transport" || tab === "work_tracking" ? "none" : tab === "picking" ? 1400 : 960, margin: "0 auto", padding: tab === "dashboard" || tab === "dashboard_transport" ? (isMobile ? "8px 10px 80px" : "8px 14px 14px") : (isMobile ? "16px 12px 80px" : "20px 14px 100px") }}>
        {tab === "dashboard" && <Dashboard trucks={trucks} queue={queue} onReset={handleReset} lane={dashLane === "main" ? null : dashLane} detailMap={detailMapByChannel} myPlate={myPlate} />}
        {tab === "dashboard_transport" && <Dashboard trucks={trucks} queue={queue} onReset={handleReset} lane={null} detailMap={detailMapByChannel} title="Dashboard ขนส่ง" myPlate={myPlate} simple />}
        {tab === "basket_summary" && <BasketSummary trucks={trucks} />}
        {tab === "waiting_summary" && <WaitingSummary trucks={trucks} />}
        {tab === "qr"        && <QRCodePage />}
        {tab === "lg"        && <LGUpload queue={queue} onSetQueue={handleSetQueue} pendingQueue={pendingQueue} onSetPendingQueue={handleSetPendingQueue} onClearPendingQueue={handleClearPendingQueue} />}
        {tab === "driver"    && <DriverScan queue={queue} trucks={trucks} onScan={handleScan} skipGeofence />}
        {tab === "picking"   && <Picking trucks={trucks} queue={queue} onUpdate={handleUpdate} detailMapByChannel={detailMapByChannel} />}
        {tab === "qc_parts"  && <QC trucks={trucks} onUpdate={handleUpdate} laneId="lane_parts" detailMapByChannel={detailMapByChannel} onBack={() => setTab(prevTab)} />}
        {tab === "qc_head"   && <QC trucks={trucks} onUpdate={handleUpdate} laneId="lane_head" detailMapByChannel={detailMapByChannel} onBack={() => setTab(prevTab)} />}
        {tab === "qc_pork"   && <QC trucks={trucks} onUpdate={handleUpdate} laneId="lane_pork" detailMapByChannel={detailMapByChannel} onBack={() => setTab(prevTab)} />}
        {tab === "loading_parts" && <LoadingYard trucks={trucks} onUpdate={handleUpdate} laneId="lane_parts" masterLane={masterLane} onBack={() => setTab(prevTab)} />}
        {tab === "loading_head"  && <LoadingYard trucks={trucks} onUpdate={handleUpdate} laneId="lane_head" masterLane={masterLane} onBack={() => setTab(prevTab)} />}
        {tab === "loading_pork"  && <LoadingYard trucks={trucks} onUpdate={handleUpdate} laneId="lane_pork" masterLane={masterLane} onBack={() => setTab(prevTab)} />}
        {tab === "sample_parts"  && <RandomSampleCheck trucks={trucks} onUpdate={handleUpdate} laneId="lane_parts" detailMapByChannel={detailMapByChannel} onBack={() => setTab(prevTab)} />}
        {tab === "sample_head"   && <RandomSampleCheck trucks={trucks} onUpdate={handleUpdate} laneId="lane_head" detailMapByChannel={detailMapByChannel} onBack={() => setTab(prevTab)} />}
        {tab === "sample_pork"   && <RandomSampleCheck trucks={trucks} onUpdate={handleUpdate} laneId="lane_pork" detailMapByChannel={detailMapByChannel} onBack={() => setTab(prevTab)} />}
        {tab === "overview_log"  && <OverviewLog trucks={trucks} />}
        {tab === "work_tracking" && <WorkTracking trucks={trucks} queue={queue} detailMapByChannel={detailMapByChannel} masterLane={masterLane} />}
        {tab === "planning"      && <Planning trucks={trucks} queue={queue} onUpdate={handleUpdate} />}
        {tab === "master_upload" && <MasterUpload masterLane={masterLane} onMasterChange={handleMasterChange} />}
        {tab === "detail_loading" && <DetailLoading masterLane={masterLane} onDetailChange={handleDetailChange} />}
        {tab === "download"       && <Download onReset={handleReset} />}
        {tab === "admin"          && <Admin trucks={trucks} queue={queue} onUpdate={handleUpdate} onDeleteTruck={handleDeleteTruck} />}
      </div>

      {/* QR Code Modal */}
    </div>
  );
}
