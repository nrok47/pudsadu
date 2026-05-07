// ============================================================
// Google Apps Script – ระบบทะเบียนครุภัณฑ์ + คลังสื่อ
// Spreadsheet ID: 12t737wkzyAW9UVUlN4XJS9UpGUdn3b5GdO2tMiJKA0A
// ============================================================

const SS_ID      = "12t737wkzyAW9UVUlN4XJS9UpGUdn3b5GdO2tMiJKA0A";
const EQUIP_TAB  = "ท.คุมนอกระบบGF+AMS";
const MEDIA_TAB  = "คลังสื่อ";
const IN_TAB     = "รับเข้า";
const OUT_TAB    = "จ่ายออก";

const EQUIP_HEADERS = [
  "ลำดับ", "รหัสครุภัณฑ์", "รหัส GF", "ปีงบประมาณ",
  "ประเภท", "หมวด/ชนิด", "ชื่อครุภัณฑ์", "สถานที่ตั้ง",
  "กลุ่มงานรับผิดชอบ", "ยี่ห้อ", "รุ่น/แบบ",
  "ราคา/หน่วย", "หน่วย", "วันที่รับ", "สถานะ"
];

const MEDIA_HEADERS = [
  "รหัสสื่อ", "กลุ่มงาน", "ชื่อสื่อ", "หน่วย", "ลิงก์"
];

const IN_HEADERS = [
  "วันที่", "เลขที่เอกสาร", "หมายเหตุ",
  "รหัสสื่อ", "จำนวน", "สถานะ"
];

const OUT_HEADERS = [
  "วันที่", "เลขที่เอกสาร", "ผู้รับ", "หมายเหตุ",
  "รหัสสื่อ", "จำนวน", "สถานะ"
];

// ── entry point ────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ระบบทะเบียนครุภัณฑ์ + คลังสื่อ")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── helpers ────────────────────────────────────────────────
function ss_() {
  return SpreadsheetApp.openById(SS_ID);
}

function ensureSheet_(name, headers) {
  const wb = ss_();
  let sh = wb.getSheetByName(name);
  if (!sh) {
    sh = wb.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function sheetData_(sh) {
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];
  return vals.slice(1).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j] ?? ""; });
    return obj;
  });
}

function dateStr_(v) {
  if (!v || v === "") return "";
  if (v instanceof Date) {
    const y = v.getFullYear();
    const be = (y >= 1960 && y <= 1975) ? y + 600 : y;
    return `${v.getDate()}/${v.getMonth() + 1}/${be}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(1\d{3})-(\d{2})-(\d{2})/);
  if (m) {
    const y = parseInt(m[1]);
    const be = (y >= 1960 && y <= 1975) ? y + 600 : y;
    return `${parseInt(m[3])}/${parseInt(m[2])}/${be}`;
  }
  return s;
}

// ── EQUIPMENT ──────────────────────────────────────────────
function getEquipment() {
  const sh = ss_().getSheetByName(EQUIP_TAB);
  if (!sh) return [];
  return sheetData_(sh).map(r => ({
    _row:         r._row,
    fsn:          r["ลำดับ"],
    equipment_id: r["รหัสครุภัณฑ์"],
    gf_id:        r["รหัส GF"],
    fiscal_year:  r["ปีงบประมาณ"],
    category:     r["ประเภท"],
    tags:         r["หมวด/ชนิด"],
    name:         r["ชื่อครุภัณฑ์"],
    location:     r["สถานที่ตั้ง"],
    responsible:  r["กลุ่มงานรับผิดชอบ"],
    brand:        r["ยี่ห้อ"],
    model:        r["รุ่น/แบบ"],
    unit_price:   r["ราคา/หน่วย"],
    unit:         r["หน่วย"],
    received_date: dateStr_(r["วันที่รับ"]),
    status:       r["สถานะ"],
  }));
}

function saveEquipment(data, rowIndex) {
  const sh = ss_().getSheetByName(EQUIP_TAB);
  if (!sh) throw new Error("ไม่พบชีต " + EQUIP_TAB);
  const row = [
    data.fsn, data.equipment_id, data.gf_id, data.fiscal_year,
    data.category, data.tags, data.name, data.location,
    data.responsible, data.brand, data.model, data.unit_price,
    data.unit, data.received_date, data.status
  ];
  if (rowIndex) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return true;
}

function deleteEquipment(rowIndex) {
  const sh = ss_().getSheetByName(EQUIP_TAB);
  if (!sh) throw new Error("ไม่พบชีต " + EQUIP_TAB);
  sh.deleteRow(rowIndex);
  return true;
}

// ── MEDIA ITEMS ────────────────────────────────────────────
function getMediaItems() {
  const mediaSh = ensureSheet_(MEDIA_TAB, MEDIA_HEADERS);
  const inSh    = ensureSheet_(IN_TAB,    IN_HEADERS);
  const outSh   = ensureSheet_(OUT_TAB,   OUT_HEADERS);

  const items = sheetData_(mediaSh).map(r => ({
    _row:       r._row,
    media_id:   r["รหัสสื่อ"],
    group_raw:  r["กลุ่มงาน"],
    name:       r["ชื่อสื่อ"],
    unit:       r["หน่วย"],
    link:       r["ลิงก์"],
  }));

  // aggregate stock-in
  const inMap  = {};
  sheetData_(inSh).forEach(r => {
    const code = r["รหัสสื่อ"], qty = Number(r["จำนวน"]) || 0;
    const st   = String(r["สถานะ"] || "");
    if (code && st !== "ยกเลิก") inMap[code] = (inMap[code] || 0) + qty;
  });

  // aggregate stock-out
  const outMap = {};
  sheetData_(outSh).forEach(r => {
    const code = r["รหัสสื่อ"], qty = Number(r["จำนวน"]) || 0;
    const st   = String(r["สถานะ"] || "");
    if (code && st !== "ยกเลิก") outMap[code] = (outMap[code] || 0) + qty;
  });

  return items.map(item => {
    const mid  = item.media_id;
    const recv = inMap[mid]  || 0;
    const issu = outMap[mid] || 0;
    const bal  = recv - issu;
    let status = "มีเหลือ";
    if (bal <= 0) status = "หมด";
    else if (recv > 0 && bal / recv < 0.2) status = "ใกล้หมด";
    return { ...item, received: recv, issued: issu, balance: bal, status };
  });
}

function saveMediaItem(data, rowIndex) {
  const sh = ensureSheet_(MEDIA_TAB, MEDIA_HEADERS);
  const row = [data.media_id, data.group_raw, data.name, data.unit, data.link];
  if (rowIndex) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return true;
}

function deleteMediaItem(rowIndex) {
  const sh = ensureSheet_(MEDIA_TAB, MEDIA_HEADERS);
  sh.deleteRow(rowIndex);
  return true;
}

function addStockIn(data) {
  const sh = ensureSheet_(IN_TAB, IN_HEADERS);
  sh.appendRow([
    data.date || new Date().toLocaleDateString("th-TH"),
    data.doc_no || "",
    data.note || "",
    data.media_id,
    data.qty,
    data.status || "ปกติ",
  ]);
  return true;
}

function addStockOut(data) {
  const sh = ensureSheet_(OUT_TAB, OUT_HEADERS);
  sh.appendRow([
    data.date || new Date().toLocaleDateString("th-TH"),
    data.doc_no || "",
    data.recipient || "",
    data.note || "",
    data.media_id,
    data.qty,
    data.status || "ปกติ",
  ]);
  return true;
}

// ── DASHBOARD ──────────────────────────────────────────────
function getDashboardData() {
  const equip = getEquipment();
  const media = getMediaItems();

  // equipment by category
  const equipByCat = {};
  equip.forEach(e => {
    const k = e.category || "ไม่ระบุ";
    equipByCat[k] = (equipByCat[k] || 0) + 1;
  });

  // equipment by responsible group
  const equipByGroup = {};
  equip.forEach(e => {
    const k = e.responsible || "ไม่ระบุ";
    equipByGroup[k] = (equipByGroup[k] || 0) + 1;
  });

  // total value (only numeric)
  const totalValue = equip.reduce((s, e) => s + (Number(e.unit_price) || 0), 0);

  // media stock status
  const mediaStatus = { "มีเหลือ": 0, "ใกล้หมด": 0, "หมด": 0 };
  media.forEach(m => { mediaStatus[m.status] = (mediaStatus[m.status] || 0) + 1; });

  // media by group
  const mediaByGroup = {};
  media.forEach(m => {
    const k = m.group_raw || "ไม่ระบุ";
    mediaByGroup[k] = (mediaByGroup[k] || 0) + 1;
  });

  return {
    equipCount:    equip.length,
    equipByCat,
    equipByGroup,
    totalValue,
    mediaCount:    media.length,
    mediaStatus,
    mediaByGroup,
    outOfStock:    mediaStatus["หมด"],
    lowStock:      mediaStatus["ใกล้หมด"],
  };
}
