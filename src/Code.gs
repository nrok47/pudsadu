// ============================================================
// ระบบทะเบียนครุภัณฑ์ + คลังสื่อ — Google Apps Script
// Spreadsheet : 12t737wkzyAW9UVUlN4XJS9UpGUdn3b5GdO2tMiJKA0A
// Equipment GID: 1500564948
// ============================================================

const SS_ID     = "12t737wkzyAW9UVUlN4XJS9UpGUdn3b5GdO2tMiJKA0A";
const EQUIP_GID = 1500564948;   // sheet ID from original project

const MEDIA_TAB = "คลังสื่อ";
const IN_TAB    = "รับเข้า";
const OUT_TAB   = "จ่ายออก";

// Column names — must match row 1 of the equipment sheet exactly
const C_FSN      = "เลขครุภัณฑ์ FSN";
const C_ID       = "เลขครุภัณฑ์";
const C_GF       = "เลขครุภัณฑ์ GF";
const C_YEAR     = "ปีงบประมาณ";
const C_TYPE     = "ประเภท";
const C_CATEGORY = "หมวด/ชนิด";
const C_NAME     = "ชื่อเรียก/คุณลักษณะครุภัณฑ์";
const C_LOCATION = "สถานที่ใช้งาน";
const C_OWNER    = "ผู้รับผิดชอบ";
const C_BRAND    = "ยี่ห้อ";
const C_MODEL    = "รุ่น";
const C_PRICE    = "ราคาต่อหน่วย";
const C_UNIT     = "หน่วยนับ";
const C_DATE     = "วันที่รับเข้า";
const C_STATUS   = "สถานะ";

const MEDIA_HEADERS = ["รหัสสื่อ", "กลุ่มงาน", "ชื่อสื่อ", "หน่วย", "ลิงก์"];
const IN_HEADERS    = ["วันที่", "เลขที่เอกสาร", "หมายเหตุ", "รหัสสื่อ", "จำนวน", "สถานะ"];
const OUT_HEADERS   = ["วันที่", "เลขที่เอกสาร", "ผู้รับ", "หมายเหตุ", "รหัสสื่อ", "จำนวน", "สถานะ"];

// ── entry point ──────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ระบบทะเบียนครุภัณฑ์ + คลังสื่อ ศูนย์อนามัยที่ 10")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── helpers ──────────────────────────────────────────────────
function ss_() {
  return SpreadsheetApp.openById(SS_ID);
}

function getEquipSheet_() {
  const sheets = ss_().getSheets();
  for (const sh of sheets) {
    if (sh.getSheetId() === EQUIP_GID) return sh;
  }
  throw new Error("Equipment sheet GID " + EQUIP_GID + " not found");
}

function getHeaders_(sh) {
  const row = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return row.map(h => String(h).trim()).filter(h => h !== "");
}

function ensureSheet_(name, headers) {
  return ensureSheetWb_(ss_(), name, headers);
}

function ensureSheetWb_(wb, name, headers) {
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

// Convert CE date → BE display string ("D/M/YYYY")
// Applies +600 only when year is in the "Thai short year" range (1960–1975 CE = 2503–2518 BE)
function dateStr_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) {
    const y  = v.getFullYear();
    const be = (y >= 1960 && y <= 1975) ? y + 600 : y;
    return `${v.getDate()}/${v.getMonth() + 1}/${be}`;
  }
  const s = String(v).trim();
  // Handle string like "1968-07-17 00:00:00"
  const m = s.match(/^(1\d{3})-(\d{2})-(\d{2})/);
  if (m) {
    const y  = parseInt(m[1]);
    const be = (y >= 1960 && y <= 1975) ? y + 600 : y;
    return `${parseInt(m[3])}/${parseInt(m[2])}/${be}`;
  }
  return s;
}

// ── EQUIPMENT — fast filter options (scan only 3 columns) ────
// Returns years/locations/statuses for populating dropdowns,
// plus latestYear so the client can default to it.
function getFilterOptions() {
  const sh      = getEquipSheet_();
  const headers = getHeaders_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { years: [], locations: [], statuses: [], latestYear: "" };

  const iYear   = headers.indexOf(C_YEAR);
  const iLoc    = headers.indexOf(C_LOCATION);
  const iStatus = headers.indexOf(C_STATUS);
  const maxCol  = Math.max(iYear, iLoc, iStatus) + 1;
  const data    = sh.getRange(2, 1, lastRow - 1, maxCol).getValues();

  const years = {}, locs = {}, statuses = {};
  data.forEach(row => {
    if (iYear   >= 0 && row[iYear])   years[String(row[iYear])]      = 1;
    if (iLoc    >= 0 && row[iLoc])    locs[String(row[iLoc])]        = 1;
    if (iStatus >= 0 && row[iStatus]) statuses[String(row[iStatus])] = 1;
  });

  const sortedYears = Object.keys(years).sort((a, b) => b - a);
  return {
    years:      sortedYears,
    locations:  Object.keys(locs).sort(),
    statuses:   Object.keys(statuses).sort(),
    latestYear: sortedYears[0] || "",
  };
}

// ── EQUIPMENT — read rows (server-side year filter) ──────────
function getEquipment(yearFilter) {
  const sh      = getEquipSheet_();
  const headers = getHeaders_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const iYear = headers.indexOf(C_YEAR);
  const data  = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const rows = [];
  data.forEach((row, idx) => {
    if (yearFilter && iYear >= 0 && String(row[iYear]) !== String(yearFilter)) return;
    const obj = { _row: idx + 2 };
    headers.forEach((h, c) => {
      obj[h] = (h === C_DATE) ? dateStr_(row[c]) : (row[c] ?? "");
    });
    rows.push(obj);
  });
  return rows;
}

// ── EQUIPMENT CRUD ───────────────────────────────────────────
function saveEquipment(data, rowIndex) {
  const sh      = getEquipSheet_();
  const headers = getHeaders_(sh);
  const row     = headers.map(h => (data[h] !== undefined ? data[h] : ""));
  if (rowIndex) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return true;
}

function deleteEquipment(rowIndex) {
  getEquipSheet_().deleteRow(rowIndex);
  return true;
}

// ── EXPORT EXCEL ─────────────────────────────────────────────
function exportEquipExcel() {
  const url  = `https://docs.google.com/spreadsheets/d/${SS_ID}/export?format=xlsx&gid=${EQUIP_GID}`;
  const resp = UrlFetchApp.fetch(url, {
    headers:            { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error("Export failed: HTTP " + resp.getResponseCode());
  }
  return Utilities.base64Encode(resp.getContent());
}

// ── MEDIA ITEMS ──────────────────────────────────────────────
function getMediaItems() {
  // Open spreadsheet once and reuse — avoids 3 separate API calls
  const wb      = ss_();
  const mediaSh = ensureSheetWb_(wb, MEDIA_TAB, MEDIA_HEADERS);
  const inSh    = ensureSheetWb_(wb, IN_TAB,    IN_HEADERS);
  const outSh   = ensureSheetWb_(wb, OUT_TAB,   OUT_HEADERS);

  const items = sheetData_(mediaSh).map(r => ({
    _row:      r._row,
    media_id:  r["รหัสสื่อ"],
    group_raw: r["กลุ่มงาน"],
    name:      r["ชื่อสื่อ"],
    unit:      r["หน่วย"],
    link:      r["ลิงก์"],
  }));

  const inMap = {}, outMap = {};
  sheetData_(inSh).forEach(r => {
    const code = r["รหัสสื่อ"], qty = Number(r["จำนวน"]) || 0;
    if (code && String(r["สถานะ"] || "") !== "ยกเลิก")
      inMap[code] = (inMap[code] || 0) + qty;
  });
  sheetData_(outSh).forEach(r => {
    const code = r["รหัสสื่อ"], qty = Number(r["จำนวน"]) || 0;
    if (code && String(r["สถานะ"] || "") !== "ยกเลิก")
      outMap[code] = (outMap[code] || 0) + qty;
  });

  return items.map(item => {
    const recv = inMap[item.media_id]  || 0;
    const issu = outMap[item.media_id] || 0;
    const bal  = recv - issu;
    let status = "มีเหลือ";
    if (bal <= 0) status = "หมด";
    else if (recv > 0 && bal / recv < 0.2) status = "ใกล้หมด";
    return { ...item, received: recv, issued: issu, balance: bal, status };
  });
}

function saveMediaItem(data, rowIndex) {
  const sh  = ensureSheet_(MEDIA_TAB, MEDIA_HEADERS);
  const row = [data.media_id, data.group_raw, data.name, data.unit, data.link || ""];
  if (rowIndex) sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else          sh.appendRow(row);
  return true;
}

function deleteMediaItem(rowIndex) {
  ensureSheet_(MEDIA_TAB, MEDIA_HEADERS).deleteRow(rowIndex);
  return true;
}

function addStockIn(data) {
  ensureSheet_(IN_TAB, IN_HEADERS).appendRow([
    data.date || "", data.doc_no || "", data.note || "",
    data.media_id, data.qty, data.status || "ปกติ",
  ]);
  return true;
}

function addStockOut(data) {
  ensureSheet_(OUT_TAB, OUT_HEADERS).appendRow([
    data.date || "", data.doc_no || "", data.recipient || "", data.note || "",
    data.media_id, data.qty, data.status || "ปกติ",
  ]);
  return true;
}

// ── DASHBOARD ────────────────────────────────────────────────
function getDashboardData() {
  const equip = getEquipment("");   // all years
  const media = getMediaItems();

  const equipByCat = {}, equipByGroup = {};
  let totalValue = 0;
  equip.forEach(e => {
    const t = e[C_TYPE]  || "ไม่ระบุ";
    const g = e[C_OWNER] || "ไม่ระบุ";
    equipByCat[t]   = (equipByCat[t]   || 0) + 1;
    equipByGroup[g] = (equipByGroup[g] || 0) + 1;
    totalValue += Number(e[C_PRICE]) || 0;
  });

  const mediaStatus  = { "มีเหลือ": 0, "ใกล้หมด": 0, "หมด": 0 };
  const mediaByGroup = {};
  media.forEach(m => {
    mediaStatus[m.status] = (mediaStatus[m.status] || 0) + 1;
    const g = m.group_raw || "ไม่ระบุ";
    mediaByGroup[g] = (mediaByGroup[g] || 0) + 1;
  });

  return {
    equipCount:  equip.length,
    equipByCat,
    equipByGroup,
    totalValue,
    mediaCount:  media.length,
    mediaStatus,
    mediaByGroup,
    outOfStock:  mediaStatus["หมด"],
    lowStock:    mediaStatus["ใกล้หมด"],
  };
}
