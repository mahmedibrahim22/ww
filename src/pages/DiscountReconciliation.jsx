import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────
const CARD_DISCOUNT_RULES = [
  { keywords: ["auc", "guc", "giu", "al-futtaim", "futtaim", "alfuttaim"], pct: 0.10, label: "10%" },
  { keywords: ["d-card", "d card", "dcard", "d.card", "امن وطني", "امن وطني", "وايت كارد", "white card", "whitecard", "مخابرات", "ديكارد", "ديكاد"], pct: 0.15, label: "15%" },
  { keywords: ["staff", "اصطاف", "استاف", "اصطف", "اصطاف"], pct: 0.25, label: "25%" },
];

const OUTLET_SHEET_KEYWORDS = [
  { outlet: ["arkan delivery", "arkan d", "010"], sheet: ["زايد"] },
  { outlet: ["arkan"], sheet: ["اركان", "أركان"] },
  { outlet: ["ard el golf", "ard golf", "el golf", "002"], sheet: ["جولف", "ارض الجولف"] },
  { outlet: ["shorouk", "el shorouk", "003"], sheet: ["شروق"] },
  { outlet: ["almaza", "005"], sheet: ["الماظة", "الماطة"] },
  { outlet: ["madii", "madi", "maadi", "006"], sheet: ["معادي"] },
  { outlet: ["cloud 9", "cloud9", "007"], sheet: ["كلاود", "كلود"] },
  { outlet: ["tagamo3", "tagamoa", "tagamoo3", "008"], sheet: ["رحاب", "الرحاب"] },
  { outlet: ["cairo festival", "009"], sheet: ["كايرو"] },
  { outlet: ["square one", "square", "011"], sheet: ["سكوير"] },
  { outlet: ["golf central", "012"], sheet: ["ج . سنترال", "ج.سنتر", "ج سنتر", "جولف سنتر"] },
  { outlet: ["district5", "district 5", "013"], sheet: ["ديستريكت"] },
  { outlet: ["rehab", "rihab"], sheet: ["رحاب", "الرحاب"] },
  { outlet: ["zamalek"], sheet: ["زمالك", "الزمالك"] },
];

const CC_BRANCH_SHEET = {
  "arkan d": "زايد", "arkan": "اركان", "golf": "جولف", "sherouk": "شروق",
  "shorouk": "شروق", "almaza": "الماظة", "maadi": "معادي", "cloud 9": "كلاود",
  "tagamoo3": "رحاب", "tagamo3": "رحاب", "cairo": "كايرو", "square": "سكوير",
  "golf central": "ج . سنترال", "district5": "ديستريكت", "rehab": "رحاب", "zamalek": "زمالك",
};

const SHEET_NAME_MAP = {
  "arkan delivery": "زايد", "arkan d": "زايد", "arkan": "اركان",
  "ard el golf": "ارض الجولف", "el golf": "ارض الجولف",
  "shorouk": "الشروق", "el shorouk": "الشروق",
  "almaza": "الماظة", "madi": "المعادي", "maadi": "المعادي",
  "cloud 9": "كلاود 9", "tagamo3": "رحاب", "tagamoo3": "رحاب",
  "cairo festival": "كايرو", "square one": "سكوير ون", "square": "سكوير ون",
  "golf central": "جولف سنترال", "district5": "ديستريكت", "district 5": "ديستريكت",
  "rehab": "الرحاب", "zamalek": "الزمالك",
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
function getExpectedPct(reason) {
  if (!reason) return null;
  const lower = String(reason).toLowerCase();
  for (const rule of CARD_DISCOUNT_RULES)
    if (rule.keywords.some(k => lower.includes(k))) return rule.pct;
  return null;
}

// Normalize reason to a canonical label for grouping in Final Reason column
function normalizeReason(reason) {
  if (!reason) return null;
  const lower = String(reason).toLowerCase().trim();
  if (lower.includes("auc")) return "auc";
  if (lower.includes("guc")) return "guc";
  if (lower.includes("giu")) return "giu";
  if (lower.includes("futtaim")) return "al-futtaim";
  if (lower.includes("white card") || lower.includes("وايت كارد") || lower.includes("whitecard")) return "white card";
  if (lower.includes("d-card") || lower.includes("d card") || lower.includes("d.card") ||
      lower.includes("ديكارد") || lower.includes("ديكاد")) return "d.card";
  if (lower.includes("امن وطني") || lower.includes("امن")) return "امن وطني";
  if (lower.includes("مخابرات")) return "مخابرات";
  if (lower.includes("staff") || lower.includes("اصطاف") || lower.includes("اصطف") || lower.includes("استاف")) return "staff";
  if (lower.includes("comp") || lower === "comp") return "comp";
  // Return the original trimmed reason if no keyword matched
  return String(reason).trim();
}

function detectUserType(createdBy) {
  if (!createdBy) return "unknown";
  return /\d/.test(String(createdBy).trim()) ? "waiter" : "callcenter";
}

// KEY FIX: Handles ALL date formats including partial '21-5', '11-5)', '20-5(' etc.
function parseExcelDate(val, fallbackYear = 2026) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === "number" && val > 30000 && val < 70000) {
    const d = new Date((val - 25569) * 86400 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof val === "string") {
    const s = val.trim();
    const sl = s.toLowerCase();

    // Month names: "21-may", "21 may", "21-may-2026"
    if (/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/.test(sl)) {
      const d = new Date(s);
      if (!isNaN(d.getTime())) return d;
    }

    // Strip non-digit and non-separator chars (removes ')', '(' etc.)
    const clean = s.replace(/[^0-9\-/]/g, "");

    // dd/mm/yyyy or d/m/yyyy
    const m1 = clean.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m1) {
      const d = new Date(`${m1[3]}-${m1[2].padStart(2,"0")}-${m1[1].padStart(2,"0")}`);
      if (!isNaN(d.getTime())) return d;
    }

    // dd-mm-yyyy or d-m-yyyy
    const m2 = clean.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
    if (m2) {
      const yr = m2[3].length === 2 ? "20" + m2[3] : m2[3];
      const d = new Date(`${yr}-${m2[2].padStart(2,"0")}-${m2[1].padStart(2,"0")}`);
      if (!isNaN(d.getTime())) return d;
    }

    // KEY FIX: partial date "21-5" or "5-21" — no year, infer fallbackYear
    const m3 = clean.match(/^(\d{1,2})-(\d{1,2})$/);
    if (m3) {
      const a = parseInt(m3[1]), b = parseInt(m3[2]);
      // Determine which is day and which is month
      // If b <= 12, treat as dd-mm; if a <= 12 treat as mm-dd; prefer dd-mm
      let day, month;
      if (b >= 1 && b <= 12) { day = a; month = b; }
      else if (a >= 1 && a <= 12) { day = b; month = a; }
      else return null;
      if (day >= 1 && day <= 31) {
        const d = new Date(fallbackYear, month - 1, day);
        if (!isNaN(d.getTime())) return d;
      }
    }

    // Standard parse fallback
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

// UTC-safe day comparison to avoid timezone off-by-one
function sameDay(a, b) {
  if (!a || !b) return false;
  // Compare using local date parts to avoid timezone issues
  const ad = a instanceof Date ? a : new Date(a);
  const bd = b instanceof Date ? b : new Date(b);
  return ad.getFullYear() === bd.getFullYear() &&
    ad.getMonth() === bd.getMonth() &&
    ad.getDate() === bd.getDate();
}

function formatDateSlash(d) {
  if (!d) return "";
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

function outletToSheetName(outletName) {
  const lower = outletName.toLowerCase();
  const sorted = Object.entries(SHEET_NAME_MAP).sort((a,b) => b[0].length - a[0].length);
  for (const [key, val] of sorted) if (lower.includes(key)) return val;
  return outletName.replace(/^outlet:\s*\d+\s*/i,"").replace(/garnell\s*/i,"").trim();
}

function findCashierSheetForOutlet(outletName, cashierData) {
  const outletLower = String(outletName).toLowerCase();
  const sheetNames = Object.keys(cashierData);
  for (const rule of OUTLET_SHEET_KEYWORDS) {
    if (!rule.outlet.some(k => outletLower.includes(k))) continue;
    for (const sn of sheetNames)
      if (rule.sheet.some(k => sn.toLowerCase().includes(k.toLowerCase()))) return sn;
  }
  return null;
}

function findCashierSheetForCCBranch(ccBranch) {
  const lower = ccBranch.toLowerCase().trim();
  const entries = Object.entries(CC_BRANCH_SHEET).sort((a,b) => b[0].length - a[0].length);
  for (const [key, val] of entries)
    if (lower.includes(key) || key.includes(lower)) return val;
  return null;
}

// ─── Cashier Parser ───────────────────────────────────────────────────────────
function parseCashierWorkbook(wb, targetDate) {
  const allCashierRecords = {};
  const callCenterRecords = [];

  for (const sheetName of wb.SheetNames) {
    const rawData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
    const isCC = sheetName.includes("كول") || sheetName.toLowerCase().includes("call");

    if (isCC) {
      callCenterRecords.push(...parseCallCenterData(rawData, targetDate));
      continue;
    }

    let headerRow = -1, serialCol = -1, dateCol = -1, pctCol = -1, valCol = -1, reasonCol = -1;
    const SH = ["السريال","سيريال","سريال","سيريل","رقم الشيك","رقم شيك","ser no","serial no","serial","check no","check#","check #"];
    const DH = ["التاريخ","تاريخ","date"];
    const PH = ["نسبة الخصم","نسبه الشيك","النسبة","النسبه","نسبة","نسبه","%","disc%","disc %"];
    const VH = ["قيمة الخصم","قيمهة الخصم","القيمة","القيمه","قيمة","قيمه","amount"];
    const RH = ["السبب","سبب","reason"];

    for (let i = 0; i < Math.min(5, rawData.length); i++) {
      const row = rawData[i]; if (!row) continue;
      let fs = -1;
      for (let j = 0; j < row.length; j++) {
        const c = String(row[j] ?? "").trim().toLowerCase();
        if (SH.some(h => c === h.toLowerCase())) { fs = j; break; }
      }
      if (fs >= 0) {
        headerRow = i; serialCol = fs;
        for (let j = 0; j < row.length; j++) {
          const c = String(row[j] ?? "").trim().toLowerCase();
          if (DH.some(h => c === h.toLowerCase())) dateCol = j;
          if (PH.some(h => c === h.toLowerCase())) pctCol = j;
          if (VH.some(h => c === h.toLowerCase())) valCol = j;
          if (RH.some(h => c === h.toLowerCase())) reasonCol = j;
        }
        break;
      }
    }
    if (headerRow === -1) { allCashierRecords[sheetName] = []; continue; }

    const records = [];
    let currentDate = null;
    // KEY FIX: inTargetZone stays TRUE once we hit targetDate,
    // and only turns FALSE when we hit a LATER date (allows same-day entries without date repeat)
    let inTargetZone = false;

    for (let i = headerRow + 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (!row || row.every(c => c == null || c === "")) continue;

      // ─── Date detection ───────────────────────────────────────────────────
      let dateFound = null;

      // Check known date col first
      if (dateCol >= 0 && row[dateCol] != null) {
        const d = parseExcelDate(row[dateCol]);
        if (d && d.getFullYear() > 2000) dateFound = d;
      }

      // Scan all other cols (handles date in السبب col pattern)
      if (!dateFound) {
        for (let j = 0; j < row.length; j++) {
          if (j === serialCol) continue;
          const v = row[j];
          if (v == null) continue;
          const isDateObj = v instanceof Date;
          const isBigSerial = typeof v === "number" && v > 40000 && v < 70000;
          // KEY FIX: also catch string dates like "21-5", "14-5)"
          const isDateStr = typeof v === "string" && (
            /jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/i.test(v) ||
            /\d{1,2}[\/-]\d{1,2}/.test(v.replace(/[^0-9\-/]/g,""))
          );
          if (isDateObj || isBigSerial || isDateStr) {
            const d = parseExcelDate(v);
            if (d && d.getFullYear() > 2000) { dateFound = d; break; }
          }
        }
      }

      if (dateFound) {
        currentDate = dateFound;
        if (sameDay(dateFound, targetDate)) {
          inTargetZone = true;
        } else if (inTargetZone) {
          // We passed the target date — stop collecting
          inTargetZone = false;
        }
      }

      if (!inTargetZone) continue;

      // ─── Extract serial ────────────────────────────────────────────────
      const serialRaw = serialCol >= 0 ? row[serialCol] : null;
      if (serialRaw == null || serialRaw === "") continue;
      const serialStr = String(serialRaw).replace(/[^\d]/g, "");
      const serialNum = parseInt(serialStr, 10);
      if (isNaN(serialNum) || serialNum <= 0) continue;

      let discPct = pctCol >= 0 ? row[pctCol] : null;
      const discVal = valCol >= 0 ? row[valCol] : null;
      let reason    = reasonCol >= 0 ? row[reasonCol] : null;

      // Reason: if it parses as a date, it's a date not a reason
      if (reason != null) {
        const maybeDate = parseExcelDate(reason);
        if (maybeDate && maybeDate.getFullYear() > 2000) reason = null;
      }

      const isComp = String(discPct || "").toLowerCase().includes("comp") || discPct === 100;
      if (!isComp && typeof discPct === "number" && discPct > 1) discPct = discPct / 100;

      records.push({
        serial: serialNum,
        date: currentDate,
        discountPct: isComp ? "comp" : discPct,
        discountVal: isComp ? "comp" : discVal,
        reason: String(reason || "").trim(),
        isComplimentary: isComp,
      });
    }
    allCashierRecords[sheetName] = records;
  }

  return { allCashierRecords, callCenterRecords };
}

function parseCallCenterData(data, targetDate) {
  let headerRow = -1, cSerial = -1, cReason = -1, cDiscount = -1, cBranch = -1, cDate = -1;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i]; if (!row) continue;
    const joined = row.map(c => String(c||"").toLowerCase()).join("|");
    if (joined.includes("serial") || joined.includes("رقم الشيك")) {
      headerRow = i;
      for (let j = 0; j < row.length; j++) {
        const h = String(row[j]||"").trim().toLowerCase();
        if (h === "serial" || h === "رقم الشيك") cSerial = j;
        if (h === "reason") cReason = j;
        if (h === "discount") cDiscount = j;
        if (h === "branch") cBranch = j;
        if (h === "date") cDate = j;
      }
      break;
    }
  }
  if (headerRow === -1) return [];
  const records = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i]; if (!row || row.every(c => !c)) continue;
    const serialRaw = cSerial >= 0 ? row[cSerial] : null;
    if (!serialRaw) continue;
    const serialNum = Number(String(serialRaw).replace(/[^\d]/g,"")); if (isNaN(serialNum) || serialNum <= 0) continue;
    const dateVal = cDate >= 0 ? parseExcelDate(row[cDate]) : null;
    if (targetDate && dateVal && !sameDay(dateVal, targetDate)) continue;
    const discRaw = cDiscount >= 0 ? String(row[cDiscount]||"").trim().toLowerCase() : "";
    const isComp = discRaw === "full" || discRaw === "comp";
    let discPct = isComp ? "comp" : (parseFloat(discRaw) || null);
    if (!isComp && typeof discPct === "number" && discPct > 1) discPct = discPct / 100;
    records.push({
      serial: serialNum,
      date: dateVal,
      discountPct: discPct,
      reason: cReason >= 0 ? String(row[cReason]||"").trim() : "",
      branch: cBranch >= 0 ? String(row[cBranch]||"").trim().toLowerCase() : "",
      isComplimentary: isComp,
    });
  }
  return records;
}

// ─── Discount/Comp System File Parser ─────────────────────────────────────────
function parseSystemFile(data, targetDate) {
  const outlets = [];
  let currentOutlet = null, headerIdx = -1, colMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i]; if (!row) continue;
    const first = String(row[0]||"").trim();

    if (first.startsWith("Outlet:")) {
      currentOutlet = { name: first, rows: [] };
      outlets.push(currentOutlet);
      headerIdx = -1; colMap = {};
      continue;
    }

    if (first === "Date" || (row[1] && String(row[1]).trim() === "Check #")) {
      headerIdx = i; colMap = {};
      for (let j = 0; j < row.length; j++) {
        const h = String(row[j]||"").trim();
        if (h === "Date")            colMap.date        = j;
        if (h === "Check #")         colMap.check       = j;
        if (h === "Created By")      colMap.createdBy   = j;
        if (h === "Closed By")       colMap.closedBy    = j;
        if (h === "Made By")         colMap.madeBy      = j;
        if (h === "Time")            colMap.time        = j;
        if (h === "Item Name")       colMap.itemName    = j;
        if (h === "Discount Type")   colMap.discountType = j;
        if (h === "Price Code")      colMap.priceCode   = j;
        if (/^\s*Qnty\s*$/.test(h)) colMap.qnty        = j;
        if (/^\s*Price\s*$/.test(h)) colMap.price       = j;
        if (/^\s*Total\s*$/.test(h)) colMap.total       = j;
        if (/^\s*Discount\s*$/.test(h)) colMap.discount = j;
        if (/^\s*Net Total\s*$/.test(h)) colMap.netTotal = j;
        if (h === "Disc %" || h === "Disc%") colMap.discPct = j;
      }
      continue;
    }

    if (!currentOutlet || headerIdx === -1) continue;
    const checkVal = colMap.check !== undefined ? row[colMap.check] : null;
    if (!checkVal || isNaN(Number(checkVal))) continue;

    const rowDate = parseExcelDate(row[colMap.date]);
    if (!rowDate || !sameDay(rowDate, targetDate)) continue;

    const createdBy = String(row[colMap.createdBy]||"").trim();
    currentOutlet.rows.push({
      date: rowDate,
      check: Number(checkVal),
      createdBy,
      closedBy:     String(row[colMap.closedBy]||"").trim(),
      madeBy:       String(row[colMap.madeBy]||"").trim(),
      time:         String(row[colMap.time]||"").trim(),
      itemName:     String(row[colMap.itemName]||"").trim(),
      discountType: String(row[colMap.discountType]||"").trim(),
      priceCode:    String(row[colMap.priceCode]||"").trim(),
      qnty:         row[colMap.qnty],
      price:        row[colMap.price],
      total:        row[colMap.total],
      discount:     row[colMap.discount],
      netTotal:     row[colMap.netTotal],
      userType:     detectUserType(createdBy),
    });
  }
  return outlets;
}

// ─── PDF Receipt Text Parser ───────────────────────────────────────────────────
function parsePdfText(pdfText) {
  const receipts = [];
  if (!pdfText) return receipts;
  const lines = pdfText.split(/\n/).map(l => l.trim()).filter(Boolean);
  let cur = null;
  for (const line of lines) {
    const sm = line.match(/Serial#[:\s]*(\d+)/i);
    if (sm) { cur = { serial: parseInt(sm[1]), outlet:null, actualPct:null, cardType:null, expectedPct:null, discrepancy:false }; receipts.push(cur); }
    if (!cur) continue;
    const om = line.match(/Outlet[:\s]+(.+)/i);
    if (om && !cur.outlet) cur.outlet = om[1].trim();
    const dm = line.match(/\*?\s*Discount\s+(\d+)%/i);
    if (dm) cur.actualPct = parseInt(dm[1]) / 100;
    const ll = line.toLowerCase();
    if (!cur.cardType) {
      if (/g\.i\.u|giu/.test(ll)) { cur.cardType="GIU"; cur.expectedPct=0.10; }
      else if (/g\.u\.c|guc/.test(ll)) { cur.cardType="GUC"; cur.expectedPct=0.10; }
      else if (/a\.u\.c|auc/.test(ll)) { cur.cardType="AUC"; cur.expectedPct=0.10; }
      else if (ll.includes("d-card")||ll.includes("امن وطني")) { cur.cardType="D-Card"; cur.expectedPct=0.15; }
      else if (ll.includes("white card")||ll.includes("مخابرات")) { cur.cardType="White Card"; cur.expectedPct=0.15; }
      else if (ll.includes("staff")||ll.includes("اصطاف")) { cur.cardType="Staff"; cur.expectedPct=0.25; }
      else if (ll.includes("futtaim")) { cur.cardType="Al-Futtaim"; cur.expectedPct=0.10; }
    }
  }
  for (const r of receipts)
    if (r.actualPct !== null && r.expectedPct !== null)
      r.discrepancy = Math.abs(r.actualPct - r.expectedPct) > 0.001;
  return receipts;
}

// ─── Reconciliation Engine ────────────────────────────────────────────────────
function reconcile(cashierData, callCenterRecords, systemOutlets, pdfReceipts, targetDate) {
  const pdfMap = new Map();
  pdfReceipts.forEach(r => pdfMap.set(r.serial, r));

  return systemOutlets.map(outlet => {
    const sheetName = findCashierSheetForOutlet(outlet.name, cashierData);
    const cashierRecords = sheetName ? (cashierData[sheetName] || []) : [];

    const cashierBySerial = {};
    for (const rec of cashierRecords) cashierBySerial[rec.serial] = rec;

    const ccBySerial = {};
    for (const rec of callCenterRecords) {
      if (targetDate && rec.date && !sameDay(rec.date, targetDate)) continue;
      const ccSheet = findCashierSheetForCCBranch(rec.branch);
      if (ccSheet && sheetName && ccSheet.toLowerCase() === sheetName.toLowerCase())
        ccBySerial[rec.serial] = rec;
    }

    const rows = outlet.rows.map(row => {
      let matchRec = cashierBySerial[row.check];
      let source = matchRec ? "cashier" : null;
      if (!matchRec) { matchRec = ccBySerial[row.check]; source = matchRec ? "callcenter" : null; }

      const inCashier = !!matchRec;
      const reason = matchRec ? (matchRec.reason || "") : "";
      const discPct = matchRec ? (matchRec.discountPct ?? "") : "";

      const expectedPct = getExpectedPct(reason);
      const actualNum = typeof discPct === "number" ? discPct : parseFloat(discPct);
      const hasPctDiscrepancy = expectedPct !== null && !isNaN(actualNum) && Math.abs(actualNum - expectedPct) > 0.001;

      const pdfMatch = pdfMap.get(row.check);
      let pdfNote = "";
      if (pdfMatch?.discrepancy) pdfNote = `⚠ كارت ${pdfMatch.cardType} متوقع ${Math.round(pdfMatch.expectedPct*100)}% - مسجل ${Math.round(pdfMatch.actualPct*100)}%`;

      return { ...row, inCashier, reason, discPct, source, expectedPct, hasPctDiscrepancy: hasPctDiscrepancy || !!pdfNote, pdfNote };
    });

    return { outletName: outlet.name, matchedSheet: sheetName || "⚠ لا يوجد فرع مطابق", rows };
  });
}

// ─── Build Final Reason / Total Dis summary ───────────────────────────────────
// Returns array of { reason: string, total: number } sorted by total desc
// One entry per unique normalized reason, summing all discount values
function buildReasonSummary(rows) {
  const map = new Map();
  for (const row of rows) {
    if (!row.inCashier || !row.reason) continue;
    const key = normalizeReason(row.reason) || row.reason.trim();
    if (!key) continue;
    const disc = Number(row.discount) || 0;
    map.set(key, (map.get(key) || 0) + disc);
  }
  return [...map.entries()]
    .map(([reason, total]) => ({ reason, total: Math.round(total * 100) / 100 }))
    .sort((a, b) => b.total - a.total);
}

// ─── Excel Export ─────────────────────────────────────────────────────────────
function exportExcel(reconciled, selectedDate) {
  const wb = XLSX.utils.book_new();
  const HEADERS = ["التاريخ","check","Created by","closed by","Made By","Time","Item Name","Discount Type","Price Code"," Qnty","   Price","      Total","   Discount","   Net Total","Disc %","السبب","ملاحظة خصم","","Final Reason","Total Dis"];
  const hFill  = { patternType:"solid", fgColor:{rgb:"FF1F4E79"} };
  const yFill  = { patternType:"solid", fgColor:{rgb:"FFFFFF00"} };
  const sFill  = { patternType:"solid", fgColor:{rgb:"FFD9EAD3"} };
  const dFill  = { patternType:"solid", fgColor:{rgb:"FFFFCCCC"} };
  const sumHdr = { patternType:"solid", fgColor:{rgb:"FFDDEBF7"} }; // light blue for summary header
  const wFont  = { bold:true, color:{rgb:"FFFFFFFF"} };
  const bFont  = { bold:true };

  for (const outlet of reconciled) {
    const sName = outletToSheetName(outlet.outletName).substring(0,31);
    const regRows   = outlet.rows.filter(r => r.inCashier);
    const unregRows = outlet.rows.filter(r => !r.inCashier);

    const checkGroups  = new Map();
    regRows.forEach(r => { if (!checkGroups.has(r.check)) checkGroups.set(r.check,[]); checkGroups.get(r.check).push(r); });
    const unregGroups = new Map();
    unregRows.forEach(r => { if (!unregGroups.has(r.check)) unregGroups.set(r.check,[]); unregGroups.get(r.check).push(r); });

    // Build reason summary for this outlet
    const reasonSummary = buildReasonSummary(outlet.rows);

    const ws_data=[], ws_styles=[], ws_numfmts=[];
    const push = (cells, styles=[], numfmts=[]) => {
      while (styles.length < cells.length) styles.push(null);
      while (numfmts.length < cells.length) numfmts.push(null);
      ws_data.push(cells); ws_styles.push(styles); ws_numfmts.push(numfmts);
    };

    // Header row (20 cols: 0-16 data, 17 empty, 18 Final Reason, 19 Total Dis)
    push(
      HEADERS,
      HEADERS.map((_,i) => i >= 18 ? {fill:sumHdr, font:bFont} : {fill:hFill, font:wFont})
    );

    let summaryRowIdx = 0; // tracks which reason to place on each data row

    for (const [, rows] of checkGroups) {
      for (let ri = 0; ri < rows.length; ri++) {
        const row = rows[ri];
        const isFirst = ri === 0;
        const dateStr = row.date ? formatDateSlash(row.date) : selectedDate;
        const discPctVal = isFirst
          ? (typeof row.discPct==="number" ? row.discPct : (row.discPct==="comp"?"comp":""))
          : null;

        let discNote = row.pdfNote || "";
        if (isFirst && row.hasPctDiscrepancy && !discNote) {
          const exp = row.expectedPct ? `${Math.round(row.expectedPct*100)}%` : "?";
          const act = typeof discPctVal==="number" ? `${Math.round(discPctVal*100)}%` : String(discPctVal);
          discNote = `⚠ متوقع ${exp} - مسجل ${act}`;
        }

        const cells = Array(20).fill(null);
        cells[0]  = dateStr;
        cells[1]  = row.check;
        cells[2]  = row.createdBy;
        cells[3]  = row.closedBy;
        cells[4]  = row.madeBy;
        cells[5]  = row.time;
        cells[6]  = row.itemName;
        cells[7]  = row.discountType;
        cells[8]  = row.priceCode;
        cells[9]  = row.qnty!=null ? Number(row.qnty) : null;
        cells[10] = row.price!=null ? Number(row.price) : null;
        cells[11] = row.total!=null ? Number(row.total) : null;
        cells[12] = row.discount!=null ? Number(row.discount) : null;
        cells[13] = row.netTotal!=null ? Number(row.netTotal) : null;
        cells[14] = discPctVal;
        cells[15] = isFirst ? (row.reason||"") : "";
        cells[16] = isFirst ? discNote : "";
        // Col 17 = empty spacer, cols 18-19 = reason summary
        if (summaryRowIdx < reasonSummary.length) {
          cells[18] = reasonSummary[summaryRowIdx].reason;
          cells[19] = reasonSummary[summaryRowIdx].total;
          summaryRowIdx++;
        }

        const rowFill = (isFirst && row.hasPctDiscrepancy) ? dFill : null;
        const styles = cells.map((_,i) => rowFill && i < 17 ? {fill:rowFill} : null);
        const numfmts = cells.map(() => null);
        if (typeof discPctVal === "number") numfmts[14] = "0%";
        if (cells[19] != null) numfmts[19] = "#,##0.00";
        push(cells, styles, numfmts);
      }

      // Sum row for this check
      const discSum = rows.reduce((s,r) => s + (Number(r.discount)||0), 0);
      const sc = Array(20).fill(null);
      sc[12] = discSum;
      sc[13] = rows[0]?.reason || "";
      if (summaryRowIdx < reasonSummary.length) {
        sc[18] = reasonSummary[summaryRowIdx].reason;
        sc[19] = reasonSummary[summaryRowIdx].total;
        summaryRowIdx++;
      }
      const ss = sc.map((_,i) => i < 17 ? {fill:sFill} : null);
      ss[12] = {fill:sFill, font:bFont};
      ss[13] = {fill:sFill, font:bFont};
      push(sc, ss, sc.map((_,i) => i===19 ? "#,##0.00" : null));
      push(Array(20).fill(null));
    }

    // "الغير مسجلين" section
    if (unregRows.length > 0) {
      const lc = Array(20).fill(null);
      lc[6] = "الغير مسجلين";
      if (summaryRowIdx < reasonSummary.length) {
        lc[18] = reasonSummary[summaryRowIdx].reason;
        lc[19] = reasonSummary[summaryRowIdx].total;
        summaryRowIdx++;
      }
      push(lc, lc.map((_,i) => ({fill: i===18||i===19 ? {patternType:"solid",fgColor:{rgb:"FFDDEBF7"}} : yFill, font:bFont})));
      for (const [,rows] of unregGroups) {
        for (const row of rows) {
          const cells = Array(20).fill(null);
          cells[0]  = row.date ? formatDateSlash(row.date) : selectedDate;
          cells[1]  = row.check;
          cells[2]  = row.createdBy;
          cells[3]  = row.closedBy;
          cells[4]  = row.madeBy;
          cells[5]  = row.time;
          cells[6]  = row.itemName;
          cells[7]  = row.discountType;
          cells[8]  = row.priceCode;
          cells[9]  = row.qnty!=null ? Number(row.qnty) : null;
          cells[10] = row.price!=null ? Number(row.price) : null;
          cells[11] = row.total!=null ? Number(row.total) : null;
          cells[12] = row.discount!=null ? Number(row.discount) : null;
          cells[13] = row.netTotal!=null ? Number(row.netTotal) : null;
          cells[14] = null;
          cells[15] = "غير مسجل";
          if (summaryRowIdx < reasonSummary.length) {
            cells[18] = reasonSummary[summaryRowIdx].reason;
            cells[19] = reasonSummary[summaryRowIdx].total;
            summaryRowIdx++;
          }
          push(cells, cells.map((_,i) => i < 17 ? {fill:yFill} : null));
        }
      }
    }

    // Build worksheet
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    ws_data.forEach((rowData, rIdx) => {
      rowData.forEach((v, cIdx) => {
        const ref = XLSX.utils.encode_cell({r:rIdx, c:cIdx});
        if (!ws[ref]) ws[ref] = { v: v ?? "", t: typeof v === "number" ? "n" : "s" };
        const st = ws_styles[rIdx]?.[cIdx];
        const nf = ws_numfmts[rIdx]?.[cIdx];
        if (st) ws[ref].s = st;
        if (nf) ws[ref].z = nf;
      });
    });
    ws["!cols"] = [
      {wch:12},{wch:8},{wch:14},{wch:14},{wch:10},{wch:10},
      {wch:28},{wch:20},{wch:10},{wch:6},{wch:10},{wch:10},
      {wch:10},{wch:10},{wch:8},{wch:30},{wch:30},
      {wch:4},{wch:18},{wch:14},
    ];
    let fn = sName, sfx = 2;
    while (wb.SheetNames.includes(fn)) fn = sName.substring(0,28) + sfx++;
    XLSX.utils.book_append_sheet(wb, ws, fn);
  }
  return wb;
}

// ─── DropZone Component ───────────────────────────────────────────────────────
function DropZone({ label, accept, onFile, file, colorClass }) {
  const ref = useRef();
  const [drag, setDrag] = useState(false);
  const handleDrop = useCallback(e => { e.preventDefault(); setDrag(false); const f=e.dataTransfer.files[0]; if(f) onFile(f); }, [onFile]);
  return (
    <div
      onClick={() => ref.current.click()}
      onDragOver={e=>{e.preventDefault();setDrag(true);}}
      onDragLeave={() => setDrag(false)}
      onDrop={handleDrop}
      className={`relative flex flex-col items-center justify-center gap-2 p-4 rounded-xl border-2 border-dashed cursor-pointer transition-all min-h-[90px] ${colorClass} ${drag?"scale-95 opacity-70":""}`}
    >
      <input ref={ref} type="file" accept={accept} className="hidden" onChange={e=>e.target.files[0]&&onFile(e.target.files[0])} />
      {file
        ? <><span className="text-lg">✅</span><span className="text-xs font-bold text-center break-all leading-tight">{file.name}</span></>
        : <><span className="text-2xl opacity-50">⬆</span><span className="text-xs font-bold text-center leading-tight opacity-80">{label}</span><span className="text-[10px] opacity-50">Click or drag & drop</span></>
      }
    </div>
  );
}

// ─── Stats Card ───────────────────────────────────────────────────────────────
function StatCard({ label, value, color }) {
  const colors = { green:"border-emerald-500/40 text-emerald-400", red:"border-red-500/40 text-red-400", yellow:"border-yellow-500/40 text-yellow-400", gray:"border-slate-600 text-slate-300" };
  return (
    <div className={`bg-slate-900/60 border rounded-xl p-4 flex flex-col gap-1 ${colors[color]||colors.gray}`}>
      <span className="text-[10px] font-bold uppercase tracking-widest text-slate-500">{label}</span>
      <span className={`text-3xl font-extrabold font-mono ${(colors[color]||colors.gray).split(" ")[1]}`}>{value}</span>
    </div>
  );
}

// ─── Result Row ───────────────────────────────────────────────────────────────
function ResultRow({ row }) {
  const missing = !row.inCashier;
  const discPctDisp = row.total && row.discount
    ? ((Number(row.discount)/Number(row.total))*100).toFixed(1)+"%"
    : "—";
  return (
    <tr className={`border-b border-slate-800 text-xs ${missing?"bg-yellow-900/20":row.hasPctDiscrepancy?"bg-red-900/15":"hover:bg-slate-800/40"}`}>
      <td className="px-3 py-2 font-mono">
        <span className={missing?"text-yellow-300 font-bold":"text-slate-300"}>#{row.check}</span>
        {missing && <span className="ml-1 text-[9px] px-1.5 py-0.5 rounded bg-yellow-900/50 text-yellow-300 border border-yellow-700/40">⚠ غير مسجل</span>}
        {!missing && row.source==="callcenter" && <span className="ml-1 text-[9px] px-1.5 py-0.5 rounded bg-amber-900/40 text-amber-300 border border-amber-700/40">CC</span>}
      </td>
      <td className="px-3 py-2 text-slate-500 font-mono">{row.time}</td>
      <td className="px-3 py-2">
        <div className="text-slate-400 font-mono">{row.createdBy}</div>
        <div className={`text-[9px] font-bold mt-0.5 ${row.userType==="waiter"?"text-sky-400":"text-amber-400"}`}>
          {row.userType==="waiter"?"Waiter":"Call Center"}
        </div>
      </td>
      <td className="px-3 py-2 text-slate-400 max-w-[180px] truncate">{row.itemName}</td>
      <td className="px-3 py-2 text-right font-mono text-slate-300">
        {row.discount!=null?Number(row.discount).toFixed(2):"—"}
      </td>
      <td className={`px-3 py-2 text-right font-mono font-bold ${missing?"text-yellow-300":row.hasPctDiscrepancy?"text-red-400":"text-emerald-400"}`}>
        {discPctDisp}
        {row.hasPctDiscrepancy && row.expectedPct && (
          <div className="text-[9px] text-red-400 font-normal">متوقع {Math.round(row.expectedPct*100)}%</div>
        )}
      </td>
      <td className="px-3 py-2">
        {missing
          ? <span className="text-yellow-600 italic text-[10px]">—</span>
          : row.reason
            ? <span className="text-slate-300">{row.reason}</span>
            : <span className="text-slate-600 italic text-[10px]">لا يوجد سبب</span>}
        {row.pdfNote && <div className="text-[9px] text-red-400 mt-0.5">{row.pdfNote}</div>}
      </td>
    </tr>
  );
}

// ─── Reason Summary Panel (shown in UI per outlet) ────────────────────────────
function ReasonSummaryPanel({ rows }) {
  const summary = buildReasonSummary(rows);
  if (!summary.length) return null;
  const total = summary.reduce((s,r) => s + r.total, 0);
  return (
    <div className="border-t border-slate-800 bg-slate-950/40 px-4 py-3">
      <div className="text-[9px] font-bold text-slate-500 uppercase tracking-widest mb-2">Final Reason / Total Discount Summary</div>
      <div className="flex flex-wrap gap-2">
        {summary.map((item, i) => (
          <div key={i} className="flex items-center gap-2 bg-slate-900/60 border border-slate-700/50 rounded-lg px-3 py-1.5">
            <span className="text-[10px] font-bold text-blue-300 uppercase">{item.reason}</span>
            <span className="text-[10px] font-mono font-bold text-emerald-400">{item.total.toFixed(2)}</span>
          </div>
        ))}
        <div className="flex items-center gap-2 bg-emerald-900/20 border border-emerald-700/30 rounded-lg px-3 py-1.5">
          <span className="text-[10px] font-bold text-emerald-300">TOTAL</span>
          <span className="text-[10px] font-mono font-bold text-emerald-400">{total.toFixed(2)}</span>
        </div>
      </div>
    </div>
  );
}

// ─── Main Component ───────────────────────────────────────────────────────────
export default function DiscountReconciliation() {
  const [files, setFiles] = useState({ cashier:null, discount:null, comp:null, pdf:null });
  const [wbs, setWbs] = useState({ cashier:null, discount:null, comp:null });
  const [pdfText, setPdfText] = useState("");
  const [targetDate, setTargetDate] = useState(new Date().toISOString().split("T")[0]);
  const [cashierSheets, setCashierSheets] = useState([]);
  const [results, setResults] = useState(null);
  const [activeTab, setActiveTab] = useState("discount");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const readWb = file => new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => { try { res(XLSX.read(new Uint8Array(e.target.result),{type:"array"})); } catch(err){rej(err);} };
    r.onerror = rej; r.readAsArrayBuffer(file);
  });

  const handleFile = async (type, file) => {
    setFiles(p => ({...p, [type]:file}));
    if (type === "pdf") {
      try {
        if (!window.pdfjsLib) {
          await new Promise((res,rej) => {
            const s = document.createElement("script");
            s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
            s.onload = res; s.onerror = rej; document.head.appendChild(s);
          });
          window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
        }
        const ab = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({data:ab}).promise;
        let text = "";
        for (let p = 1; p <= pdf.numPages; p++) {
          const page = await pdf.getPage(p);
          const content = await page.getTextContent();
          text += content.items.map(i => i.str).join("\n") + "\n";
        }
        setPdfText(text);
      } catch(e) { setError("خطأ في قراءة PDF: " + e.message); }
      return;
    }
    try {
      const wb = await readWb(file);
      setWbs(p => ({...p, [type]:wb}));
      if (type === "cashier") setCashierSheets(wb.SheetNames);
    } catch(e) { setError("خطأ في قراءة الملف: " + e.message); }
  };

  const handleReconcile = () => {
    if (!wbs.cashier) { setError("من فضلك ارفع شيت الكاشير أولاً"); return; }
    if (!wbs.discount && !wbs.comp) { setError("ارفع ملف ديسكوند أو كومبليمنتري على الأقل"); return; }
    setError(""); setLoading(true);
    try {
      const dateObj = new Date(targetDate + "T00:00:00");
      const { allCashierRecords, callCenterRecords } = parseCashierWorkbook(wbs.cashier, dateObj);

      const parseFile = wb => {
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1 });
        return parseSystemFile(raw, dateObj);
      };

      const discOutlets = wbs.discount ? parseFile(wbs.discount) : [];
      const compOutlets = wbs.comp ? parseFile(wbs.comp) : [];

      const allOutlets = [...discOutlets];
      for (const co of compOutlets) {
        const existing = allOutlets.find(o => o.name === co.name);
        if (existing) existing.rows.push(...co.rows);
        else allOutlets.push(co);
      }

      const pdfReceiptList = parsePdfText(pdfText);
      const discResult  = wbs.discount ? reconcile(allCashierRecords, callCenterRecords, discOutlets, pdfReceiptList, dateObj) : null;
      const compResult  = wbs.comp     ? reconcile(allCashierRecords, callCenterRecords, compOutlets, pdfReceiptList, dateObj) : null;
      const mergedResult = reconcile(allCashierRecords, callCenterRecords, allOutlets, pdfReceiptList, dateObj);

      setResults({ discount:discResult, comp:compResult, merged:mergedResult });
      setActiveTab(discResult ? "discount" : "comp");
    } catch(e) { setError("خطأ في المعالجة: " + e.message); console.error(e); }
    finally { setLoading(false); }
  };

  const handleExport = () => {
    if (!results?.merged) return;
    const wb = exportExcel(results.merged, targetDate.split("-").reverse().join("-"));
    XLSX.writeFile(wb, `discount-${targetDate}.xlsx`);
  };

  const currentData = activeTab === "discount" ? results?.discount : results?.comp;
  const totalRows   = currentData?.reduce((s,o) => s + o.rows.length, 0) || 0;
  const matchedRows = currentData?.reduce((s,o) => s + o.rows.filter(r => r.inCashier).length, 0) || 0;
  const missingRows = totalRows - matchedRows;
  const discrepRows = currentData?.reduce((s,o) => s + o.rows.filter(r => r.hasPctDiscrepancy).length, 0) || 0;

  return (
    <div className="bg-[#0a0f0d] text-slate-300 min-h-screen p-5 font-mono" dir="ltr">
      <div className="max-w-7xl mx-auto space-y-5">

        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-4">
          <div>
            <h1 className="text-white font-bold text-lg tracking-tight">Garnell <span className="text-emerald-400">مطابقة تسويات</span></h1>
            <p className="text-[10px] text-slate-600 mt-0.5">FINANCIAL AUDIT SYSTEM — RECON V2.3</p>
          </div>
          <div className="text-right">
            <div className="text-[10px] text-emerald-400 font-bold">System Online ●</div>
            <div className="text-[10px] text-slate-600">Stable Environment</div>
          </div>
        </div>

        {/* Upload + Controls */}
        <div className="bg-slate-900/50 border border-slate-800 rounded-2xl p-5">
          <div className="text-[10px] font-bold text-emerald-400 uppercase tracking-widest mb-1">STEP 1</div>
          <div className="text-white font-bold text-sm mb-4">Upload Files</div>

          <div className="grid grid-cols-2 gap-3 mb-4">
            <DropZone label="شيت خصومات الفروع (.xlsx)" accept=".xlsx,.xls" onFile={f=>handleFile("cashier",f)} file={files.cashier} colorClass="border-emerald-700/50 bg-emerald-950/20 hover:border-emerald-500/60" />
            <DropZone label="Discounted Items (.xls)" accept=".xlsx,.xls" onFile={f=>handleFile("discount",f)} file={files.discount} colorClass="border-blue-700/50 bg-blue-950/20 hover:border-blue-500/60" />
            <DropZone label="Complimentary Items (.xls)" accept=".xlsx,.xls" onFile={f=>handleFile("comp",f)} file={files.comp} colorClass="border-purple-700/50 bg-purple-950/20 hover:border-purple-500/60" />
            <DropZone label="Receipts PDF (شيكات الكارنيهات)" accept=".pdf" onFile={f=>handleFile("pdf",f)} file={files.pdf} colorClass="border-orange-700/50 bg-orange-950/20 hover:border-orange-500/60" />
          </div>

          {cashierSheets.length > 0 && (
            <div className="flex flex-wrap gap-1.5 mb-4">
              {cashierSheets.map(s => (
                <span key={s} className={`text-[9px] font-bold px-2 py-0.5 rounded border ${s.includes("كول")||s.toLowerCase().includes("call")?"bg-amber-900/20 text-amber-400 border-amber-800/30":"bg-emerald-900/20 text-emerald-400 border-emerald-800/30"}`}>
                  {s.includes("كول")||s.toLowerCase().includes("call")?"📞":"📋"} {s}
                </span>
              ))}
            </div>
          )}

          <div className="flex items-end gap-3 flex-wrap">
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Target Date</label>
              <input type="date" value={targetDate} onChange={e=>setTargetDate(e.target.value)}
                className="bg-slate-950 border border-slate-700 text-slate-300 text-xs rounded-lg px-3 py-2 focus:border-emerald-500 focus:outline-none" />
            </div>
            <button onClick={handleReconcile} disabled={loading||!files.cashier}
              className="flex items-center gap-2 px-5 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800 disabled:text-slate-600 disabled:cursor-not-allowed text-white font-bold text-xs rounded-lg transition-all shadow-[0_0_20px_rgba(52,211,153,0.2)] hover:shadow-[0_0_30px_rgba(52,211,153,0.4)]">
              <svg className={`w-4 h-4 ${loading?"animate-spin":""}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z"/>
              </svg>
              {loading ? "Processing..." : "Reconcile"}
            </button>
            {results && (
              <button onClick={handleExport}
                className="flex items-center gap-2 px-5 py-2 bg-slate-800 hover:bg-slate-700 border border-slate-700 hover:border-emerald-500/40 text-emerald-400 font-bold text-xs rounded-lg transition-all">
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/>
                </svg>
                Export Excel
              </button>
            )}
          </div>

          {error && (
            <div className="mt-3 flex items-center gap-2 bg-red-950/40 border border-red-800/40 text-red-400 text-xs rounded-lg px-4 py-2.5">
              <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
              </svg>
              {error}
            </div>
          )}
        </div>

        {/* Results */}
        {results && (
          <>
            <div className="grid grid-cols-4 gap-3">
              <StatCard label="Total Orders"  value={totalRows}   color="gray" />
              <StatCard label="Matched ✓"     value={matchedRows} color="green" />
              <StatCard label="Missing ⚠"     value={missingRows} color={missingRows>0?"red":"green"} />
              <StatCard label="Disc % Alert"  value={discrepRows} color={discrepRows>0?"yellow":"green"} />
            </div>

            <div className="flex items-center gap-1 bg-slate-950/60 border border-slate-800 rounded-xl p-1 w-fit">
              {results.discount && (
                <button onClick={()=>setActiveTab("discount")}
                  className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${activeTab==="discount"?"bg-emerald-500/20 text-emerald-400 border border-emerald-500/30":"text-slate-500 hover:text-white hover:bg-white/5"}`}>
                  🏷 Discounted Items
                </button>
              )}
              {results.comp && (
                <button onClick={()=>setActiveTab("comp")}
                  className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${activeTab==="comp"?"bg-purple-500/20 text-purple-400 border border-purple-500/30":"text-slate-500 hover:text-white hover:bg-white/5"}`}>
                  🎁 Complimentary
                </button>
              )}
            </div>

            <div className="flex items-center gap-4 flex-wrap text-[10px] text-slate-500">
              <span><span className="inline-block w-3 h-3 rounded bg-yellow-900/60 border border-yellow-700/40 mr-1"></span>غير مسجل في الكاشير</span>
              <span><span className="inline-block w-3 h-3 rounded bg-red-900/40 border border-red-700/40 mr-1"></span>خلاف في نسبة الخصم</span>
              <span className="text-amber-400 font-bold border border-amber-700/40 bg-amber-900/30 px-1 rounded text-[9px]">CC</span><span> مطابق من كول سنتر</span>
            </div>

            {currentData?.map((outlet, oi) => {
              const om = outlet.rows.filter(r=>!r.inCashier).length;
              const od = outlet.rows.filter(r=>r.hasPctDiscrepancy).length;
              return (
                <div key={oi} className="bg-slate-900/40 border border-slate-800 rounded-2xl overflow-hidden">
                  <div className="flex items-center justify-between px-4 py-3 bg-slate-900/80 border-b border-slate-800">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="text-white font-bold text-xs">{outlet.outletName}</span>
                      {outlet.matchedSheet && !outlet.matchedSheet.startsWith("⚠") && (
                        <span className="text-[9px] px-2 py-0.5 rounded bg-emerald-900/30 text-emerald-400 border border-emerald-800/30">📋 {outlet.matchedSheet}</span>
                      )}
                      {outlet.matchedSheet?.startsWith("⚠") && (
                        <span className="text-[9px] px-2 py-0.5 rounded bg-red-900/30 text-red-400 border border-red-800/30">{outlet.matchedSheet}</span>
                      )}
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] text-slate-600">{outlet.rows.length} orders</span>
                      {od>0 && <span className="text-[9px] font-bold px-2 py-0.5 rounded-full bg-yellow-900/40 text-yellow-300 border border-yellow-700/40">⚠ {od} disc%</span>}
                      {om>0 && <span className="text-[9px] font-bold px-2 py-0.5 rounded-full bg-red-900/40 text-red-300 border border-red-700/40">⚠ {om} missing</span>}
                      {om===0&&od===0 && <span className="text-[9px] font-bold px-2 py-0.5 rounded-full bg-emerald-900/30 text-emerald-400 border border-emerald-800/30">✓ All matched</span>}
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-950/50">
                          {["Check #","Time","Created By","Item Name","Discount","Disc %","السبب / Reason"].map(h=>(
                            <th key={h} className="px-3 py-2 text-left text-[9px] font-bold uppercase tracking-widest text-slate-600 border-b border-slate-800">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {outlet.rows.map((row,ri) => <ResultRow key={ri} row={row} />)}
                      </tbody>
                    </table>
                  </div>

                  {/* Final Reason / Total Dis summary panel */}
                  <ReasonSummaryPanel rows={outlet.rows} />
                </div>
              );
            })}
          </>
        )}
      </div>
    </div>
  );
}