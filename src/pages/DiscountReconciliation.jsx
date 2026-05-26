import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────

// Card type → expected discount %
const CARD_DISCOUNT_RULES = [
  { keywords: ["auc", "guc", "giu", "al-futtaim", "futtaim", "alfuttaim"], pct: 0.10, label: "Student/AUC/GUC/GIU/Futtaim 10%" },
  { keywords: ["d-card", "d card", "dcard", "امن وطني", "امن", "وطني", "white card", "whitecard", "مخابرات"], pct: 0.15, label: "D-Card/White Card 15%" },
  { keywords: ["staff", "اصطاف", "استاف", "اصطاف"], pct: 0.25, label: "Staff 25%" },
];

function getExpectedPct(reasonOrLabel) {
  if (!reasonOrLabel) return null;
  const lower = String(reasonOrLabel).toLowerCase();
  for (const rule of CARD_DISCOUNT_RULES) {
    if (rule.keywords.some(k => lower.includes(k))) return rule.pct;
  }
  return null;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function detectUserType(createdBy) {
  if (!createdBy) return "unknown";
  const s = String(createdBy).trim();
  if (/\d/.test(s)) return "waiter";
  return "callcenter";
}

function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === "number") {
    const d = new Date((val - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return d;
  }
  if (typeof val === "string") {
    const dmyMatch = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmyMatch) {
      const d = new Date(`${dmyMatch[3]}-${dmyMatch[2].padStart(2, "0")}-${dmyMatch[1].padStart(2, "0")}`);
      if (!isNaN(d.getTime())) return d;
    }
    const d = new Date(val);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function formatDate(d) {
  if (!d) return "";
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" });
}

function formatDateSlash(d) {
  if (!d) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function sameDay(a, b) {
  if (!a || !b) return false;
  return a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();
}

// ─── Cashier Sheet Parser (شيت خصومات الفروع) ───────────────────────────────
// Multi-sheet file. Each sheet = one branch.
// Date written ONCE above a group of serials. We propagate it downward.
// We also filter by exact target date, returning only those serials.

function parseCashierSheet(data, sheetName) {
  let headerRow = -1;
  let serialCol = -1, dateCol = -1, discPctCol = -1, discValCol = -1, reasonCol = -1;

  // All possible names for each column — Arabic variants + English
  const SERIAL_HEADERS  = ["السريال", "سيريال", "سريال", "سيريل", "رقم الشيك", "رقم شيك",
                            "ser no", "serial no", "serial", "check no", "check#", "check #"];
  const DATE_HEADERS    = ["التاريخ", "تاريخ", "date"];
  const PCT_HEADERS     = ["نسبة الخصم", "نسبه الشيك", "النسبة", "النسبه", "نسبة", "نسبه", "%", "disc%", "disc %"];
  const VAL_HEADERS     = ["قيمة الخصم", "قيمهة الخصم", "القيمة", "القيمه", "قيمة", "قيمه"];
  const REASON_HEADERS  = ["السبب", "سبب", "reason"];

  // Scan up to row 4 (رحاب has header in row 1)
  for (let i = 0; i < Math.min(4, data.length); i++) {
    const row = data[i];
    if (!row) continue;

    // Try to find serial header in this row
    let foundSerial = -1;
    for (let j = 0; j < row.length; j++) {
      const cell = String(row[j] ?? "").trim().toLowerCase();
      if (SERIAL_HEADERS.some(h => cell === h.toLowerCase())) {
        foundSerial = j;
        break;
      }
    }

    if (foundSerial >= 0) {
      headerRow = i;
      serialCol = foundSerial;
      // Now find the other columns in same row
      for (let j = 0; j < row.length; j++) {
        const cell = String(row[j] ?? "").trim().toLowerCase();
        if (DATE_HEADERS.some(h   => cell === h.toLowerCase())) dateCol    = j;
        if (PCT_HEADERS.some(h    => cell === h.toLowerCase())) discPctCol = j;
        if (VAL_HEADERS.some(h    => cell === h.toLowerCase())) discValCol = j;
        if (REASON_HEADERS.some(h => cell === h.toLowerCase())) reasonCol  = j;
      }
      break;
    }
  }

  // سكوير special case: col 0 header is NaN/empty but col 0 contains serials, col 1 = date
  // Detect by: header row has NaN in col 0 AND col 1 is a date keyword
  if (headerRow === -1) {
    for (let i = 0; i < Math.min(4, data.length); i++) {
      const row = data[i];
      if (!row) continue;
      const col0 = String(row[0] ?? "").trim();
      const col1 = String(row[1] ?? "").trim().toLowerCase();
      if ((col0 === "" || col0 === "nan") && DATE_HEADERS.some(h => col1 === h)) {
        headerRow  = i;
        serialCol  = 0;
        dateCol    = 1;
        // Find pct/val/reason in remaining cols
        for (let j = 2; j < row.length; j++) {
          const cell = String(row[j] ?? "").trim().toLowerCase();
          if (PCT_HEADERS.some(h    => cell === h.toLowerCase())) discPctCol = j;
          if (VAL_HEADERS.some(h    => cell === h.toLowerCase())) discValCol = j;
          if (REASON_HEADERS.some(h => cell === h.toLowerCase())) reasonCol  = j;
        }
        break;
      }
    }
  }

  // Last resort fallback: no header found — guess from first numeric row
  if (headerRow === -1) {
    for (let i = 0; i < Math.min(6, data.length); i++) {
      const row = data[i];
      if (!row) continue;
      // Look for a row where col 0 is a number and another col is a date
      const num = Number(row[0]);
      if (!isNaN(num) && num > 0) {
        const d = parseExcelDate(row[1]);
        if (d) { serialCol = 0; dateCol = 1; discPctCol = 2; discValCol = 3; reasonCol = 4; break; }
      }
    }
    if (serialCol === -1) return [];
    // No header, start from row 0
    const records = [];
    let currentDate = null;
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row || row.every(c => !c)) continue;
      const d = parseExcelDate(row[dateCol]);
      if (d) currentDate = d;
      const serialNum = Number(row[serialCol]);
      if (isNaN(serialNum) || serialNum <= 0) continue;
      const discPct = row[discPctCol];
      const isComp = String(discPct || "").toLowerCase().includes("comp") || discPct === 1;
      records.push({
        serial: serialNum, date: currentDate,
        discountPct: isComp ? "comp" : discPct,
        discountVal: isComp ? "comp" : row[discValCol],
        reason: String(row[reasonCol] || "").trim(),
        isComplimentary: isComp, branch: sheetName,
      });
    }
    return records;
  }

  const records = [];
  let currentDate = null;
  const startRow = headerRow + 1;

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c == null || c === "")) continue;

    // ── Date propagation ──────────────────────────────────────────────────────
    // A "date row" has a valid date somewhere and no valid serial number.
    // We check the dateCol first, then scan all non-serial cols.
    let dateFound = null;

    // Check the known date col
    if (dateCol >= 0 && row[dateCol] != null) {
      dateFound = parseExcelDate(row[dateCol]);
    }

    // Scan all cols for a date (handles sheets where date is in col 0 without header match)
    if (!dateFound) {
      for (let j = 0; j < row.length; j++) {
        if (j === serialCol) continue;
        if (row[j] == null) continue;
        const d = parseExcelDate(row[j]);
        // Must be a real calendar date (year > 2000) and not a small number misread as date
        if (d && d.getFullYear() > 2000 && (typeof row[j] !== "number" || row[j] > 40000)) {
          dateFound = d;
          break;
        }
      }
    }

    if (dateFound) currentDate = dateFound;

    // ── Serial extraction ─────────────────────────────────────────────────────
    const serialRaw = serialCol >= 0 ? row[serialCol] : null;
    if (serialRaw == null || serialRaw === "") continue;

    // Handle values like "(9)", "(19)", "101d" — strip non-digits except leading minus
    const serialStr = String(serialRaw).replace(/[^\d]/g, "");
    const serialNum = parseInt(serialStr, 10);
    if (isNaN(serialNum) || serialNum <= 0) continue;

    const discPct = discPctCol >= 0 ? row[discPctCol] : null;
    const discVal = discValCol >= 0 ? row[discValCol] : null;
    const reason  = reasonCol  >= 0 ? row[reasonCol]  : null;
    // "comp" OR value of 1 (100%) often means complimentary in some sheets
    const isComp  = String(discPct || "").toLowerCase().includes("comp") ||
                    discPct === 1 || discPct === "1";

    records.push({
      serial:          serialNum,
      date:            currentDate,
      discountPct:     isComp ? "comp" : discPct,
      discountVal:     isComp ? "comp" : discVal,
      reason:          String(reason || "").trim(),
      isComplimentary: isComp,
      branch:          sheetName,
    });
  }

  return records;
}

// ─── Call Center Sheet Parser (كول سنتر) ─────────────────────────────────────
// Columns: Serial, NUMBER, Reason, Discount, Amount, Date, Branch

function parseCallCenterSheet(data) {
  // Find header row
  let headerRow = -1;
  let colSerial = -1, colReason = -1, colDiscount = -1, colDate = -1, colBranch = -1;

  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    if (!row) continue;
    const joined = row.map(c => String(c || "").toLowerCase()).join("|");
    if (joined.includes("serial") || joined.includes("سيريال")) {
      headerRow = i;
      for (let j = 0; j < row.length; j++) {
        const h = String(row[j] || "").trim().toLowerCase();
        if (h === "serial") colSerial = j;
        if (h === "reason") colReason = j;
        if (h === "discount") colDiscount = j;
        if (h === "date") colDate = j;
        if (h === "branch") colBranch = j;
      }
      break;
    }
  }

  if (headerRow === -1) return [];

  const records = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => !c)) continue;
    const serialRaw = colSerial >= 0 ? row[colSerial] : null;
    if (!serialRaw) continue;
    const serialNum = Number(serialRaw);
    if (isNaN(serialNum)) continue;

    const dateVal = colDate >= 0 ? parseExcelDate(row[colDate]) : null;
    const discountRaw = colDiscount >= 0 ? String(row[colDiscount] || "").trim().toLowerCase() : "";
    const isComp = discountRaw === "full" || discountRaw === "comp";
    const discPct = isComp ? "comp" : (parseFloat(discountRaw) || null);
    const branch = colBranch >= 0 ? String(row[colBranch] || "").trim().toLowerCase() : "";
    const reason = colReason >= 0 ? String(row[colReason] || "").trim() : "";

    records.push({
      serial: serialNum,
      date: dateVal,
      discountPct: discPct,
      reason,
      branch,
      isComplimentary: isComp,
    });
  }
  return records;
}

// ─── Discount / Complimentary File Parser ─────────────────────────────────────

function isNoteRow(row, colMap) {
  const checkVal = colMap.check !== undefined ? row[colMap.check] : null;
  if (checkVal) return false;
  const allText = row.map(c => String(c || "")).join("");
  return /[\u0600-\u06FF]/.test(allText);
}

function parseDiscountFile(data) {
  const outlets = [];
  let currentOutlet = null;
  let headerIdx = -1;
  let colMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;
    const firstCell = String(row[0] || "").trim();

    if (firstCell.startsWith("Outlet:")) {
      currentOutlet = { name: firstCell, rows: [] };
      outlets.push(currentOutlet);
      headerIdx = -1;
      colMap = {};
      continue;
    }

    if (firstCell === "Date" || (row[1] && String(row[1]).trim() === "Check #")) {
      headerIdx = i;
      colMap = {};
      for (let j = 0; j < row.length; j++) {
        const h = String(row[j] || "").trim();
        if (h === "Date") colMap.date = j;
        if (h === "Check #") colMap.check = j;
        if (h === "Created By") colMap.createdBy = j;
        if (h === "Closed By") colMap.closedBy = j;
        if (h === "Made By") colMap.madeBy = j;
        if (h === "Time") colMap.time = j;
        if (h === "Item Name") colMap.itemName = j;
        if (h === "Discount Type") colMap.discountType = j;
        if (h === "Price Code") colMap.priceCode = j;
        if (h.trim() === "Qnty" || h.trim() === " Qnty") colMap.qnty = j;
        if (h.trim() === "Price" || h.trim() === "   Price") colMap.price = j;
        if (h.trim() === "Total" || h.trim() === "      Total") colMap.total = j;
        if (h.trim() === "Discount" || h.trim() === "   Discount") colMap.discount = j;
        if (h.trim() === "Net Total" || h.trim() === "   Net Total") colMap.netTotal = j;
        if (h === "Disc %" || h === "Disc%") colMap.discPct = j;
        if (h === "السبب") colMap.reason = j;
      }
      continue;
    }

    if (!currentOutlet || headerIdx === -1) continue;

    const checkVal = colMap.check !== undefined ? row[colMap.check] : null;
    const firstCellVal = String(row[0] || "").trim();

    if (firstCellVal.startsWith("Total")) continue;
    if (isNoteRow(row, colMap)) continue;
    if (!checkVal) continue;
    if (isNaN(Number(checkVal))) continue;

    const createdBy = String(row[colMap.createdBy] || "").trim();
    const userType = detectUserType(createdBy);

    currentOutlet.rows.push({
      date: parseExcelDate(colMap.date !== undefined ? row[colMap.date] : null),
      check: Number(checkVal),
      createdBy,
      closedBy: String(row[colMap.closedBy] || "").trim(),
      madeBy: String(row[colMap.madeBy] || "").trim(),
      time: String(row[colMap.time] || "").trim(),
      itemName: String(row[colMap.itemName] || "").trim(),
      discountType: String(row[colMap.discountType] || "").trim(),
      priceCode: String(row[colMap.priceCode] || "").trim(),
      qnty: row[colMap.qnty],
      price: row[colMap.price],
      total: row[colMap.total],
      discount: row[colMap.discount],
      netTotal: row[colMap.netTotal],
      userType,
      reason: "",
      originalRowIdx: i,
    });
  }

  return outlets;
}

// ─── PDF Receipt Parser ───────────────────────────────────────────────────────
// Reads text extracted from PDF receipts.
// Each receipt block contains: Serial#, Outlet, Discount %, etc.

function parsePdfReceipts(pdfText) {
  const receipts = [];
  if (!pdfText) return receipts;

  // Split into receipt blocks by looking for "GARNELL" or "Garnell" headers
  const lines = pdfText.split(/\n/).map(l => l.trim()).filter(Boolean);

  let current = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Detect new receipt - look for Serial# line
    const serialMatch = line.match(/Serial#[:\s]*(\d+)/i);
    if (serialMatch) {
      current = {
        serial: parseInt(serialMatch[1]),
        outlet: null,
        discountPct: null,
        cardType: null,
        actualPct: null,
        expectedPct: null,
        discrepancy: false,
        rawLines: [],
      };
      receipts.push(current);
    }

    if (!current) continue;
    current.rawLines.push(line);

    // Outlet
    const outletMatch = line.match(/Outlet[:\s]+(.+)/i);
    if (outletMatch && !current.outlet) {
      current.outlet = outletMatch[1].trim();
    }

    // Discount line e.g. "* Discount 10%" or "* Discount 15%"
    const discMatch = line.match(/\*?\s*Discount\s+(\d+)%/i);
    if (discMatch) {
      current.discountPct = parseInt(discMatch[1]) / 100;
      current.actualPct = current.discountPct;
    }

    // Detect card type from handwritten annotation lines or card names
    const lineLower = line.toLowerCase();
    if (lineLower.includes("auc") && !current.cardType) {
      current.cardType = "AUC"; current.expectedPct = 0.10;
    } else if (lineLower.includes("giu") && !current.cardType) {
      current.cardType = "GIU"; current.expectedPct = 0.10;
    } else if (lineLower.includes("guc") && !current.cardType) {
      current.cardType = "GUC"; current.expectedPct = 0.10;
    } else if ((lineLower.includes("d-card") || lineLower.includes("d card") || lineLower.includes("dcard") || lineLower.includes("امن وطني")) && !current.cardType) {
      current.cardType = "D-Card (امن وطني)"; current.expectedPct = 0.15;
    } else if ((lineLower.includes("white card") || lineLower.includes("whitecard") || lineLower.includes("مخابرات")) && !current.cardType) {
      current.cardType = "White Card (مخابرات)"; current.expectedPct = 0.15;
    } else if ((lineLower.includes("staff") || lineLower.includes("اصطاف") || lineLower.includes("استاف")) && !current.cardType) {
      current.cardType = "Staff"; current.expectedPct = 0.25;
    } else if (lineLower.includes("al-futtaim") || lineLower.includes("futtaim")) {
      current.cardType = "Al-Futtaim"; current.expectedPct = 0.10;
    }
  }

  // Mark discrepancies
  for (const r of receipts) {
    if (r.actualPct !== null && r.expectedPct !== null) {
      r.discrepancy = Math.abs(r.actualPct - r.expectedPct) > 0.001;
    }
  }

  return receipts;
}

// ─── Outlet → Cashier Sheet Mapping ──────────────────────────────────────────

const OUTLET_SHEET_KEYWORDS = [
  { outlet: ["arkan delivery", "010"], sheet: ["زايد"] },
  { outlet: ["arkan"], sheet: ["اركان", "أركان"] },
  { outlet: ["ard el golf", "ard golf", "el golf", "002"], sheet: ["جولف", "ارض الجولف"] },
  { outlet: ["shorouk", "el shorouk", "003"], sheet: ["شروق"] },
  { outlet: ["almaza", "005"], sheet: ["الماظة", "الماطة"] },
  { outlet: ["madii", "madi", "maadi", "006"], sheet: ["معادي"] },
  { outlet: ["cloud 9", "cloud9", "007"], sheet: ["كلاود", "كلود"] },
  { outlet: ["tagamo3", "tagamoa", "tagamoo3", "008"], sheet: ["رحاب"] },
  { outlet: ["cairo festival", "009"], sheet: ["كايرو"] },
  { outlet: ["square one", "square", "011"], sheet: ["سكوير"] },
  { outlet: ["golf central", "012"], sheet: ["ج . سنترال", "ج.سنتر", "ج سنتر", "جولف سنتر"] },
  { outlet: ["district5", "district 5", "013"], sheet: ["ديستريكت"] },
  { outlet: ["rehab", "rihab"], sheet: ["رحاب"] },
  { outlet: ["zamalek"], sheet: ["زمالك"] },
];

// Call center branch name → cashier sheet
const CC_BRANCH_SHEET = {
  "arkan": "اركان",
  "arkan d": "زايد",
  "golf": "جولف",
  "sherouk": "شروق",
  "shorouk": "شروق",
  "almaza": "الماظة",
  "maadi": "معادي",
  "cloud 9": "كلاود",
  "tagamoo3": "رحاب",
  "tagamo3": "رحاب",
  "cairo": "كايرو",
  "square": "سكوير",
  "golf central": "ج . سنترال",
  "district5": "ديستريكت",
  "rehab": "رحاب",
  "zamalek": "زمالك",
};

function findCashierSheetForOutlet(outletName, cashierData) {
  const outletLower = String(outletName).toLowerCase();
  const sheetNames = Object.keys(cashierData);

  for (const rule of OUTLET_SHEET_KEYWORDS) {
    if (!rule.outlet.some(k => outletLower.includes(k))) continue;
    for (const sn of sheetNames) {
      if (rule.sheet.some(k => sn.toLowerCase().includes(k.toLowerCase()))) return sn;
    }
  }

  const numMatch = outletLower.match(/(\d{3})/);
  if (numMatch) {
    for (const sn of sheetNames) if (sn.includes(numMatch[1])) return sn;
  }

  const outletWords = outletLower
    .replace("outlet:", "").replace("garnell", "")
    .split(/\s+/).filter(w => w.length > 2 && !/^\d+$/.test(w));
  let bestMatch = null, bestScore = 0;
  for (const sn of sheetNames) {
    const score = outletWords.filter(w => sn.toLowerCase().includes(w)).length;
    if (score > bestScore) { bestScore = score; bestMatch = sn; }
  }
  return bestScore > 0 ? bestMatch : null;
}

function findCashierSheetForCCBranch(ccBranch) {
  const lower = ccBranch.toLowerCase().trim();
  for (const [key, val] of Object.entries(CC_BRANCH_SHEET)) {
    if (lower.includes(key) || key.includes(lower)) return val;
  }
  return null;
}

// ─── Reconciliation ───────────────────────────────────────────────────────────

function reconcile(cashierData, callCenterRecords, discountOutlets, targetDate) {
  return discountOutlets.map((outlet) => {
    const sheetName = findCashierSheetForOutlet(outlet.name, cashierData);
    const cashierRecords = sheetName ? (cashierData[sheetName] || []) : [];

    // Filter cashier records by target date
    const cashierBySerial = {};
    for (const rec of cashierRecords) {
      if (targetDate && rec.date && !sameDay(rec.date, targetDate)) continue;
      cashierBySerial[rec.serial] = rec;
    }

    // Filter call center records for this outlet by target date
    const ccBySerial = {};
    for (const rec of callCenterRecords) {
      if (targetDate && rec.date && !sameDay(rec.date, targetDate)) continue;
      // Match branch to outlet
      const ccSheet = findCashierSheetForCCBranch(rec.branch);
      if (ccSheet && sheetName && ccSheet.toLowerCase() === sheetName.toLowerCase()) {
        ccBySerial[rec.serial] = rec;
      }
    }

    return {
      outletName: outlet.name,
      matchedSheet: sheetName || "⚠ No sheet matched",
      rows: outlet.rows.map((row) => {
        // Look in cashier first
        let cashierRec = cashierBySerial[row.check];
        let source = cashierRec ? "cashier" : null;

        // If not found in cashier, check call center sheet
        if (!cashierRec) {
          cashierRec = ccBySerial[row.check];
          source = cashierRec ? "callcenter" : null;
        }

        const inCashier = !!cashierRec;
        const reason = cashierRec ? (cashierRec.reason || "") : "";
        const discPct = cashierRec ? (cashierRec.discountPct ?? "") : "";

        // Verify discount % from cashier reason vs actual
        const expectedPct = getExpectedPct(reason);
        const actualPctNum = typeof discPct === "number" ? discPct : parseFloat(discPct);
        const hasPctDiscrepancy = expectedPct !== null && !isNaN(actualPctNum) &&
          Math.abs(actualPctNum - expectedPct) > 0.001;

        return { ...row, inCashier, reason, discPct, cashierRec, source, outletName: outlet.name, expectedPct, hasPctDiscrepancy };
      }),
    };
  });
}

// ─── Excel Export ─────────────────────────────────────────────────────────────

const SHEET_NAME_MAP = {
  "arkan delivery": "زايد",
  "arkan": "اركان",
  "ard el golf": "ارض الجولف",
  "el golf": "ارض الجولف",
  "shorouk": "الشروق",
  "el shorouk": "الشروق",
  "almaza": "الماظة",
  "madi": "المعادي",
  "maadi": "المعادي",
  "cloud 9": "كلاود 9",
  "tagamo3": "رحاب",
  "tagamoo3": "رحاب",
  "cairo festival": "كايرو",
  "square one": "سكوير ون",
  "square": "سكوير ون",
  "golf central": "جولف سنترال",
  "district5": "ديستريكت",
  "district 5": "ديستريكت",
  "rehab": "الرحاب",
  "zamalek": "الزمالك",
};

function outletToSheetName(outletName) {
  const lower = outletName.toLowerCase();
  // Check longer matches first
  const sorted = Object.entries(SHEET_NAME_MAP).sort((a, b) => b[0].length - a[0].length);
  for (const [key, val] of sorted) {
    if (lower.includes(key)) return val;
  }
  return outletName.replace(/^outlet:\s*\d+\s*/i, "").replace(/garnell\s*/i, "").trim();
}

function generateOutputExcel(reconciled, fileType, selectedDate) {
  const wb = XLSX.utils.book_new();

  const HEADERS = [
    "التاريخ", "check", "Created by", "closed by", "Made By",
    "Time", "Item Name", "Discount Type", "Price Code",
    " Qnty", "   Price", "      Total", "   Discount", "   Net Total",
    "Disc %", "السبب", "ملاحظة خصم",
  ];

  const yellowFill = { patternType: "solid", fgColor: { rgb: "FFFFFF00" } };
  const headerFill = { patternType: "solid", fgColor: { rgb: "FF1F4E79" } };
  const missingHeaderFill = { patternType: "solid", fgColor: { rgb: "FFFFFF00" } };
  const sumFill = { patternType: "solid", fgColor: { rgb: "FFD9EAD3" } };
  const discrepancyFill = { patternType: "solid", fgColor: { rgb: "FFFFCCCC" } };
  const pctFmt = "0%";
  const boldFont = { bold: true };
  const whiteFont = { bold: true, color: { rgb: "FFFFFFFF" } };

  for (const outlet of reconciled) {
    const sheetName = outletToSheetName(outlet.outletName).substring(0, 31);

    const registeredRows = outlet.rows.filter(r => r.inCashier);
    const unregisteredRows = outlet.rows.filter(r => !r.inCashier);

    // Group registered rows by check#
    const checkGroups = new Map();
    for (const row of registeredRows) {
      if (!checkGroups.has(row.check)) checkGroups.set(row.check, []);
      checkGroups.get(row.check).push(row);
    }

    const unregGroups = new Map();
    for (const row of unregisteredRows) {
      if (!unregGroups.has(row.check)) unregGroups.set(row.check, []);
      unregGroups.get(row.check).push(row);
    }

    const ws_data = [];
    const ws_styles = [];
    const ws_numfmts = [];

    function pushRow(cells, styles = [], numfmts = []) {
      while (styles.length < cells.length) styles.push(null);
      while (numfmts.length < cells.length) numfmts.push(null);
      ws_data.push(cells);
      ws_styles.push(styles);
      ws_numfmts.push(numfmts);
    }

    // Header row
    pushRow(
      HEADERS,
      HEADERS.map(() => ({ fill: headerFill, font: whiteFont })),
    );

    // Registered check groups
    for (const [checkNum, rows] of checkGroups) {
      const groupStartDataRow = ws_data.length;

      for (let ri = 0; ri < rows.length; ri++) {
        const row = rows[ri];
        const isFirst = ri === 0;

        const dateStr = row.date ? formatDateSlash(row.date) : (selectedDate || "");
        const discPctVal = isFirst
          ? (typeof row.discPct === "number" ? row.discPct
            : (row.discPct === "comp" ? "comp"
              : (row.cashierRec?.discountPct ?? "")))
          : null;

        // Discount discrepancy note
        let discNote = "";
        if (isFirst && row.hasPctDiscrepancy) {
          const expected = row.expectedPct ? `${Math.round(row.expectedPct * 100)}%` : "?";
          const actual = typeof discPctVal === "number" ? `${Math.round(discPctVal * 100)}%` : String(discPctVal);
          discNote = `⚠ متوقع ${expected} - مسجل ${actual}`;
        }

        const cells = [
          dateStr,
          row.check,
          row.createdBy,
          row.closedBy,
          row.madeBy,
          row.time,
          row.itemName,
          row.discountType,
          row.priceCode,
          row.qnty != null ? Number(row.qnty) : null,
          row.price != null ? Number(row.price) : null,
          row.total != null ? Number(row.total) : null,
          row.discount != null ? Number(row.discount) : null,
          row.netTotal != null ? Number(row.netTotal) : null,
          discPctVal,
          isFirst ? (row.reason || "") : "",
          isFirst ? discNote : "",
        ];

        const rowFill = (isFirst && row.hasPctDiscrepancy) ? discrepancyFill : null;
        const styles = cells.map(() => rowFill ? { fill: rowFill } : null);
        const numfmts = cells.map(() => null);
        if (discPctVal !== null && typeof discPctVal === "number") numfmts[14] = pctFmt;

        pushRow(cells, styles, numfmts);
      }

      // SUM row
      const discSum = rows.reduce((s, r) => s + (Number(r.discount) || 0), 0);
      const sumReason = rows[0]?.reason || "";

      const sumCells = Array(17).fill(null);
      sumCells[12] = discSum;
      sumCells[13] = sumReason;
      const sumStyles = sumCells.map(() => ({ fill: sumFill }));
      sumStyles[12] = { fill: sumFill, font: boldFont };
      sumStyles[13] = { fill: sumFill, font: boldFont };
      pushRow(sumCells, sumStyles);

      // Empty row
      pushRow(Array(17).fill(null));
    }

    // "الغير مسجلين" section
    if (unregisteredRows.length > 0) {
      const labelCells = Array(17).fill(null);
      labelCells[6] = "الغير مسجلين";
      pushRow(
        labelCells,
        labelCells.map(() => ({ fill: missingHeaderFill, font: boldFont })),
      );

      for (const [checkNum, rows] of unregGroups) {
        for (const row of rows) {
          const dateStr = row.date ? formatDateSlash(row.date) : (selectedDate || "");
          const cells = [
            dateStr,
            row.check,
            row.createdBy,
            row.closedBy,
            row.madeBy,
            row.time,
            row.itemName,
            row.discountType,
            row.priceCode,
            row.qnty != null ? Number(row.qnty) : null,
            row.price != null ? Number(row.price) : null,
            row.total != null ? Number(row.total) : null,
            row.discount != null ? Number(row.discount) : null,
            row.netTotal != null ? Number(row.netTotal) : null,
            null,
            "غير مسجل",
            "",
          ];
          pushRow(
            cells,
            cells.map(() => ({ fill: yellowFill })),
          );
        }
      }
    }

    // Build worksheet
    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    ws_data.forEach((rowData, rIdx) => {
      rowData.forEach((cellVal, cIdx) => {
        const ref = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
        if (!ws[ref]) ws[ref] = { v: cellVal ?? "", t: typeof cellVal === "number" ? "n" : "s" };
        const st = ws_styles[rIdx]?.[cIdx];
        const nf = ws_numfmts[rIdx]?.[cIdx];
        if (st) ws[ref].s = st;
        if (nf) ws[ref].z = nf;
      });
    });

    ws["!cols"] = [
      { wch: 12 }, { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 10 },
      { wch: 10 }, { wch: 28 }, { wch: 20 }, { wch: 10 }, { wch: 6 },
      { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 8 },
      { wch: 30 }, { wch: 30 },
    ];

    let finalName = sheetName;
    let suffix = 2;
    while (wb.SheetNames.includes(finalName)) {
      finalName = sheetName.substring(0, 28) + suffix++;
    }

    XLSX.utils.book_append_sheet(wb, ws, finalName);
  }

  return wb;
}

// ─── Sub-components ───────────────────────────────────────────────────────────

function DropZone({ label, accept, onFile, file, icon, color = "green" }) {
  const ref = useRef();
  const [drag, setDrag] = useState(false);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDrag(false);
    const f = e.dataTransfer.files[0];
    if (f) onFile(f);
  }, [onFile]);

  const colors = {
    green: {
      border: file ? "border-[#10b981]/60 bg-[#10b981]/5" : drag ? "border-[#10b981]/80 bg-[#10b981]/10" : "border-[#1a3d2f] bg-[#0d2a1f]/30 hover:border-[#10b981]/40",
      icon: file ? "text-[#10b981]" : "text-[#4b6b5d]",
      label: file ? "text-[#10b981]" : "text-[#648b7a]",
    },
    blue: {
      border: file ? "border-blue-500/60 bg-blue-500/5" : drag ? "border-blue-500/80 bg-blue-500/10" : "border-[#1a3d2f] bg-[#0d2a1f]/30 hover:border-blue-500/40",
      icon: file ? "text-blue-400" : "text-[#4b6b5d]",
      label: file ? "text-blue-400" : "text-[#648b7a]",
    },
    purple: {
      border: file ? "border-purple-500/60 bg-purple-500/5" : drag ? "border-purple-500/80 bg-purple-500/10" : "border-[#1a3d2f] bg-[#0d2a1f]/30 hover:border-purple-500/40",
      icon: file ? "text-purple-400" : "text-[#4b6b5d]",
      label: file ? "text-purple-400" : "text-[#648b7a]",
    },
    orange: {
      border: file ? "border-orange-500/60 bg-orange-500/5" : drag ? "border-orange-500/80 bg-orange-500/10" : "border-[#1a3d2f] bg-[#0d2a1f]/30 hover:border-orange-500/40",
      icon: file ? "text-orange-400" : "text-[#4b6b5d]",
      label: file ? "text-orange-400" : "text-[#648b7a]",
    },
  };
  const c = colors[color] || colors.green;

  return (
    <div
      onClick={() => ref.current.click()}
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={handleDrop}
      className={`relative flex flex-col items-center justify-center gap-2 p-4 rounded-xl border-2 border-dashed cursor-pointer transition-all duration-200 min-h-[90px] group ${c.border}`}
    >
      <input ref={ref} type="file" accept={accept} className="hidden"
        onChange={(e) => e.target.files[0] && onFile(e.target.files[0])} />
      <svg className={`w-6 h-6 transition-colors ${c.icon}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
        {icon}
      </svg>
      <span className={`text-xs font-bold text-center leading-tight transition-colors ${c.label}`}>
        {file ? file.name : label}
      </span>
      {!file && <span className="text-[10px] text-[#4b6b5d]">Click or drag & drop</span>}
      {file && (
        <span className="absolute top-2 left-2">
          <svg className={`w-4 h-4 ${c.icon}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
        </span>
      )}
    </div>
  );
}

function UserBadge({ type }) {
  if (type === "waiter") return (
    <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-sky-900/40 text-sky-400 border border-sky-800/40">
      <span className="w-1.5 h-1.5 rounded-full bg-sky-400 inline-block"></span>Waiter
    </span>
  );
  if (type === "callcenter") return (
    <span className="inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full bg-amber-900/30 text-amber-400 border border-amber-800/40">
      <span className="w-1.5 h-1.5 rounded-full bg-amber-400 inline-block"></span>Call Center
    </span>
  );
  return null;
}

function StatCard({ label, value, color }) {
  const colors = {
    green: "border-[#10b981]/30 text-[#10b981]",
    red: "border-red-500/30 text-red-400",
    yellow: "border-yellow-500/30 text-yellow-400",
    gray: "border-[#1a3d2f] text-gray-300",
  };
  return (
    <div className={`bg-[#0d2a1f]/40 border rounded-xl p-4 flex flex-col gap-1 ${colors[color] || colors.gray}`}>
      <span className="text-[10px] font-bold uppercase tracking-widest text-[#648b7a]">{label}</span>
      <span className={`text-3xl font-extrabold font-mono ${colors[color]?.split(" ")[1]}`}>{value}</span>
    </div>
  );
}

function ResultRow({ row }) {
  const missing = !row.inCashier;
  const discPct = row.total && row.discount
    ? ((Number(row.discount) / Number(row.total)) * 100).toFixed(1) + "%"
    : "—";

  return (
    <tr className={`border-b border-[#1a3d2f] transition-colors ${missing ? "bg-yellow-900/25" : row.hasPctDiscrepancy ? "bg-red-900/20" : "hover:bg-[#0d2a1f]/40"}`}>
      <td className="px-3 py-2.5 text-xs font-mono">
        <span className={missing ? "text-yellow-300 font-bold" : "text-gray-300"}>#{row.check}</span>
        {missing && (
          <span className="ml-2 text-[10px] px-1.5 py-0.5 rounded bg-yellow-900/40 text-yellow-300 border border-yellow-700/40 font-bold">⚠ Missing</span>
        )}
        {row.source === "callcenter" && !missing && (
          <span className="ml-2 text-[10px] px-1.5 py-0.5 rounded bg-amber-900/40 text-amber-300 border border-amber-700/40 font-bold">CC</span>
        )}
      </td>
      <td className="px-3 py-2.5 text-[11px] text-[#648b7a] font-mono">{formatDate(row.date)}</td>
      <td className="px-3 py-2.5">
        <div className="flex flex-col gap-1">
          <span className="text-[11px] text-gray-300 font-mono">{row.createdBy}</span>
          <UserBadge type={row.userType} />
        </div>
      </td>
      <td className="px-3 py-2.5 text-[11px] text-[#648b7a]">{row.itemName}</td>
      <td className="px-3 py-2.5 text-[11px] text-right font-mono text-gray-300">
        {row.discount != null ? Number(row.discount).toFixed(2) : "—"}
      </td>
      <td className="px-3 py-2.5 text-[11px] text-right font-mono text-gray-300">
        {row.total != null ? Number(row.total).toFixed(2) : "—"}
      </td>
      <td className={`px-3 py-2.5 text-[11px] text-right font-mono font-bold ${missing ? "text-yellow-300" : row.hasPctDiscrepancy ? "text-red-400" : "text-[#10b981]"}`}>
        {discPct}
        {row.hasPctDiscrepancy && row.expectedPct && (
          <div className="text-[9px] text-red-400">متوقع {Math.round(row.expectedPct * 100)}%</div>
        )}
      </td>
      <td className="px-3 py-2.5 text-[11px]">
        {missing
          ? <span className="text-yellow-500/70 italic">Not in cashier</span>
          : row.reason
            ? <span className="text-gray-300">{row.reason}</span>
            : <span className="text-[#4b6b5d] italic">No reason recorded</span>
        }
      </td>
    </tr>
  );
}

// PDF Receipts Display
function PdfReceiptCard({ receipt }) {
  const hasDisc = receipt.actualPct !== null;
  const hasCard = receipt.cardType !== null;

  return (
    <div className={`border rounded-xl p-3 text-xs ${receipt.discrepancy ? "border-red-500/40 bg-red-900/10" : "border-[#1a3d2f] bg-[#0d2a1f]/30"}`}>
      <div className="flex items-center justify-between mb-2">
        <span className="font-bold text-white">Serial #{receipt.serial}</span>
        {receipt.discrepancy && (
          <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-red-900/40 text-red-300 border border-red-700/40">⚠ Discrepancy</span>
        )}
        {!receipt.discrepancy && hasCard && (
          <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-[#10b981]/10 text-[#10b981] border border-[#10b981]/30">✓ OK</span>
        )}
      </div>
      <div className="text-[#648b7a] space-y-1">
        {receipt.outlet && <div>Outlet: <span className="text-gray-300">{receipt.outlet}</span></div>}
        {hasCard && <div>Card: <span className="text-blue-300">{receipt.cardType}</span></div>}
        {hasDisc && <div>Applied: <span className={receipt.discrepancy ? "text-red-400 font-bold" : "text-[#10b981]"}>{Math.round(receipt.actualPct * 100)}%</span></div>}
        {hasCard && <div>Expected: <span className="text-gray-300">{Math.round(receipt.expectedPct * 100)}%</span></div>}
      </div>
    </div>
  );
}

// ─── Main Component ───────────────────────────────────────────────────────────

export default function DiscountReconciliation() {
  const [cashierFile, setCashierFile] = useState(null);
  const [discountFile, setDiscountFile] = useState(null);
  const [compFile, setCompFile] = useState(null);
  const [pdfFile, setPdfFile] = useState(null);
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split("T")[0]);
  const [cashierSheets, setCashierSheets] = useState([]);
  const [cashierData, setCashierData] = useState(null);
  const [callCenterRecords, setCallCenterRecords] = useState([]);
  const [discountData, setDiscountData] = useState(null);
  const [compData, setCompData] = useState(null);
  const [pdfReceipts, setPdfReceipts] = useState([]);
  const [results, setResults] = useState(null);
  const [activeTab, setActiveTab] = useState("discount");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const readExcelFile = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

  const handleCashierFile = async (file) => {
    setCashierFile(file); setCashierSheets([]); setCashierData(null); setCallCenterRecords([]);
    try {
      const wb = await readExcelFile(file);
      const sheets = wb.SheetNames;
      setCashierSheets(sheets);
      const allData = {};
      let ccRecs = [];
      for (const s of sheets) {
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[s], { header: 1, raw: true });
        // Check if this is the call center sheet
        const isCC = s.includes("كول") || s.includes("call") || s.toLowerCase().includes("cc");
        if (isCC) {
          ccRecs = parseCallCenterSheet(raw);
        } else {
          allData[s] = parseCashierSheet(raw, s);
        }
      }
      setCashierData(allData);
      setCallCenterRecords(ccRecs);
    } catch (e) { setError("Error reading cashier file: " + e.message); }
  };

  const handleDiscountFile = async (file) => {
    setDiscountFile(file); setDiscountData(null);
    try {
      const wb = await readExcelFile(file);
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
      setDiscountData(parseDiscountFile(raw));
    } catch (e) { setError("Error reading discount file: " + e.message); }
  };

  const handleCompFile = async (file) => {
    setCompFile(file); setCompData(null);
    try {
      const wb = await readExcelFile(file);
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true });
      setCompData(parseDiscountFile(raw));
    } catch (e) { setError("Error reading complimentary file: " + e.message); }
  };

  const handlePdfFile = async (file) => {
    setPdfFile(file); setPdfReceipts([]);
    try {
      // Read PDF as text using FileReader - we'll extract text client-side
      // We use the PDF.js approach via fetch if available, else use ArrayBuffer text extraction
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          // Try to load PDF.js dynamically
          if (!window.pdfjsLib) {
            await new Promise((resolve, reject) => {
              const script = document.createElement("script");
              script.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
              script.onload = resolve;
              script.onerror = reject;
              document.head.appendChild(script);
            });
            window.pdfjsLib.GlobalWorkerOptions.workerSrc =
              "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
          }
          const pdf = await window.pdfjsLib.getDocument({ data: e.target.result }).promise;
          let fullText = "";
          for (let p = 1; p <= pdf.numPages; p++) {
            const page = await pdf.getPage(p);
            const content = await page.getTextContent();
            fullText += content.items.map(i => i.str).join("\n") + "\n";
          }
          const receipts = parsePdfReceipts(fullText);
          setPdfReceipts(receipts);
        } catch (err) {
          setError("Error parsing PDF: " + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (e) { setError("Error reading PDF file: " + e.message); }
  };

  const handleReconcile = () => {
    if (!cashierData) { setError("Please upload cashier file."); return; }
    if (!discountData && !compData) { setError("Please upload at least one discount or complimentary file."); return; }
    setError(""); setLoading(true);
    try {
      const targetDate = selectedDate ? new Date(selectedDate) : null;

      const discountResults = discountData
        ? reconcile(cashierData, callCenterRecords, discountData, targetDate)
        : null;
      const compResults = compData
        ? reconcile(cashierData, callCenterRecords, compData, targetDate)
        : null;

      setResults({ discount: discountResults, comp: compResults });
      setActiveTab(discountResults ? "discount" : "comp");
    } catch (e) { setError("Reconciliation error: " + e.message); }
    finally { setLoading(false); }
  };

  const handleExport = (type) => {
    const data = type === "discount" ? results?.discount : results?.comp;
    if (!data) return;
    const wb = generateOutputExcel(data, type, selectedDate);
    const dateLabel = selectedDate.replace(/-/g, "-");
    XLSX.writeFile(wb, `${type === "discount" ? "Discounted_Items" : "Complimentary"}_${dateLabel}.xlsx`);
  };

  const currentResults = activeTab === "discount" ? results?.discount : results?.comp;
  const totalRows = currentResults ? currentResults.reduce((s, o) => s + o.rows.length, 0) : 0;
  const missingRows = currentResults ? currentResults.reduce((s, o) => s + o.rows.filter(r => !r.inCashier).length, 0) : 0;
  const discrepancyRows = currentResults ? currentResults.reduce((s, o) => s + o.rows.filter(r => r.hasPctDiscrepancy).length, 0) : 0;
  const pdfDiscrepancies = pdfReceipts.filter(r => r.discrepancy).length;

  return (
    <div className="bg-[#020c08] text-gray-300 min-h-screen p-6" dir="ltr">
      <div className="max-w-6xl mx-auto">

        {/* ── Upload Card ───────────────────────────────────────── */}
        <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] rounded-2xl p-6 mb-6 relative overflow-hidden">
          <div className="absolute -top-16 -left-16 w-48 h-48 bg-[#10b981]/5 rounded-full blur-3xl pointer-events-none"></div>

          <div className="flex items-center gap-2 mb-5 border-r-4 border-[#10b981] pr-3">
            <div>
              <p className="text-xs text-[#10b981] font-bold uppercase tracking-widest">Step 1</p>
              <p className="text-white font-bold">Upload Files</p>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-4 mb-3">
            {/* Row 1: Cashier + Discount */}
            <DropZone
              label="شيت خصومات الفروع (.xlsx)"
              accept=".xlsx,.xls"
              onFile={handleCashierFile}
              file={cashierFile}
              color="green"
              icon={<path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M3 10h18M3 14h18M10 3v18M6 3h12a2 2 0 012 2v14a2 2 0 01-2 2H6a2 2 0 01-2-2V5a2 2 0 012-2z" />}
            />
            <DropZone
              label="Discounted Items (.xls)"
              accept=".xlsx,.xls"
              onFile={handleDiscountFile}
              file={discountFile}
              color="blue"
              icon={<path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A2 2 0 013 12V7a4 4 0 014-4z" />}
            />
          </div>

          <div className="grid grid-cols-2 gap-4 mb-5">
            {/* Row 2: Complimentary + PDF */}
            <DropZone
              label="Complimentary Items (.xls)"
              accept=".xlsx,.xls"
              onFile={handleCompFile}
              file={compFile}
              color="purple"
              icon={<path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M12 8v13m0-13V6a4 4 0 00-4-4H5.45a1 1 0 00-.9 1.45l1.5 3A1 1 0 007 7.5h5M12 8V6a4 4 0 014-4h2.55a1 1 0 01.9 1.45l-1.5 3A1 1 0 0117 7.5h-5" />}
            />
            <DropZone
              label="Receipts PDF (شيكات الكارنيهات)"
              accept=".pdf"
              onFile={handlePdfFile}
              file={pdfFile}
              color="orange"
              icon={<path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />}
            />
          </div>

          {/* Cashier sheets preview */}
          {cashierSheets.length > 0 && (
            <div className="mb-4 flex flex-wrap gap-2">
              {cashierSheets.map(s => (
                <span key={s} className={`text-[10px] font-bold px-2 py-1 rounded-lg border ${s.includes("كول") || s.toLowerCase().includes("call") ? "bg-amber-900/20 text-amber-400 border-amber-800/30" : "bg-[#10b981]/10 text-[#10b981] border-[#10b981]/20"}`}>
                  {s.includes("كول") || s.toLowerCase().includes("call") ? "📞" : "📋"} {s}
                </span>
              ))}
            </div>
          )}

          {/* PDF receipts preview */}
          {pdfReceipts.length > 0 && (
            <div className="mb-4 p-3 rounded-xl bg-orange-900/10 border border-orange-800/30">
              <div className="flex items-center justify-between mb-2">
                <span className="text-[11px] font-bold text-orange-400">📄 PDF Receipts Parsed: {pdfReceipts.length}</span>
                {pdfDiscrepancies > 0 && (
                  <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-red-900/40 text-red-300 border border-red-700/40">
                    ⚠ {pdfDiscrepancies} discount discrepancy
                  </span>
                )}
              </div>
              <div className="grid grid-cols-2 md:grid-cols-3 gap-2 max-h-48 overflow-y-auto">
                {pdfReceipts.map((r, i) => <PdfReceiptCard key={i} receipt={r} />)}
              </div>
            </div>
          )}

          {/* Controls */}
          <div className="flex items-end gap-4 flex-wrap">
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-[#648b7a] uppercase tracking-widest">Target Date</label>
              <input
                type="date"
                value={selectedDate}
                onChange={e => setSelectedDate(e.target.value)}
                className="bg-[#020c08] border border-[#1a3d2f] text-gray-300 text-xs rounded-lg px-3 py-2 focus:border-[#10b981]/50 focus:outline-none"
              />
            </div>

            <button
              onClick={handleReconcile}
              disabled={loading || !cashierFile}
              className="flex items-center gap-2 px-5 py-2 bg-[#10b981] hover:bg-[#0ea371] disabled:bg-[#0d2a1f] disabled:text-[#4b6b5d] disabled:cursor-not-allowed text-[#020c08] font-bold text-xs rounded-lg transition-all duration-200 shadow-[0_0_20px_rgba(16,185,129,0.2)] hover:shadow-[0_0_30px_rgba(16,185,129,0.4)]"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
              </svg>
              {loading ? "Processing..." : "Reconcile"}
            </button>
          </div>

          {error && (
            <div className="mt-4 flex items-center gap-2 bg-red-900/20 border border-red-800/40 text-red-400 text-xs rounded-lg px-4 py-3">
              <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              {error}
            </div>
          )}
        </div>

        {/* ── Results ───────────────────────────────────────────── */}
        {results && (
          <>
            {/* Stats */}
            <div className="grid grid-cols-4 gap-4 mb-6">
              <StatCard label="Total Orders" value={totalRows} color="gray" />
              <StatCard label="Matched in Cashier" value={totalRows - missingRows} color="green" />
              <StatCard label="Not in Cashier" value={missingRows} color={missingRows > 0 ? "red" : "green"} />
              <StatCard label="Disc % Discrepancy" value={discrepancyRows} color={discrepancyRows > 0 ? "yellow" : "green"} />
            </div>

            {/* Tabs + Export */}
            <div className="flex items-center justify-between mb-4 flex-wrap gap-3">
              <div className="flex items-center gap-1 bg-[#020c08]/60 border border-[#1a3d2f] rounded-xl p-1">
                {results.discount && (
                  <button onClick={() => setActiveTab("discount")}
                    className={`flex items-center gap-1.5 px-4 py-1.5 rounded-lg text-xs font-bold transition-all duration-200
                      ${activeTab === "discount" ? "bg-[#10b981]/20 text-[#10b981] border border-[#10b981]/30" : "text-[#4b6b5d] hover:text-white hover:bg-white/5"}`}>
                    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A2 2 0 013 12V7a4 4 0 014-4z" />
                    </svg>
                    Discounted Items
                  </button>
                )}
                {results.comp && (
                  <button onClick={() => setActiveTab("comp")}
                    className={`flex items-center gap-1.5 px-4 py-1.5 rounded-lg text-xs font-bold transition-all duration-200
                      ${activeTab === "comp" ? "bg-[#10b981]/20 text-[#10b981] border border-[#10b981]/30" : "text-[#4b6b5d] hover:text-white hover:bg-white/5"}`}>
                    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v13m0-13V6a4 4 0 00-4-4H5.45a1 1 0 00-.9 1.45l1.5 3A1 1 0 007 7.5h5M12 8V6a4 4 0 014-4h2.55a1 1 0 01.9 1.45l-1.5 3A1 1 0 0117 7.5h-5" />
                    </svg>
                    Complimentary
                  </button>
                )}
              </div>

              <button
                onClick={() => handleExport(activeTab)}
                className="flex items-center gap-2 px-4 py-2 bg-[#0d2a1f]/60 hover:bg-[#0d2a1f] border border-[#1a3d2f] hover:border-[#10b981]/40 text-[#10b981] text-xs font-bold rounded-lg transition-all duration-200"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                Export Updated File
              </button>
            </div>

            {/* Legend */}
            <div className="flex items-center gap-4 mb-4 flex-wrap">
              <div className="flex items-center gap-2 text-[11px] text-[#648b7a]">
                <span className="w-3 h-3 rounded bg-yellow-900/60 border border-yellow-700/40"></span> Not registered in cashier
              </div>
              <div className="flex items-center gap-2 text-[11px] text-[#648b7a]">
                <span className="w-3 h-3 rounded bg-red-900/40 border border-red-700/40"></span> Discount % discrepancy
              </div>
              <div className="flex items-center gap-2 text-[11px] text-[#648b7a]">
                <span className="text-amber-400 font-bold text-[10px] border border-amber-700/40 bg-amber-900/30 px-1 rounded">CC</span> Matched via Call Center sheet
              </div>
            </div>

            {/* Tables per outlet */}
            {currentResults && currentResults.map((outlet, oi) => {
              const outletMissing = outlet.rows.filter(r => !r.inCashier).length;
              const outletDiscrep = outlet.rows.filter(r => r.hasPctDiscrepancy).length;
              return (
                <div key={oi} className="mb-5 bg-[#0d2a1f]/30 border border-[#1a3d2f] rounded-2xl overflow-hidden">
                  <div className="flex items-center justify-between px-4 py-3 bg-[#0d2a1f]/60 border-b border-[#1a3d2f]">
                    <div className="flex items-center gap-2 flex-wrap">
                      <svg className="w-4 h-4 text-[#10b981] shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" />
                      </svg>
                      <span className="text-xs font-bold text-white">{outlet.outletName}</span>
                      {outlet.matchedSheet && !outlet.matchedSheet.startsWith("⚠") && (
                        <span className="text-[10px] px-2 py-0.5 rounded bg-[#10b981]/10 text-[#10b981] border border-[#10b981]/20">
                          📋 {outlet.matchedSheet}
                        </span>
                      )}
                      {outlet.matchedSheet?.startsWith("⚠") && (
                        <span className="text-[10px] px-2 py-0.5 rounded bg-red-900/30 text-red-400 border border-red-800/30">
                          {outlet.matchedSheet}
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-3">
                      <span className="text-[10px] text-[#648b7a]">{outlet.rows.length} orders</span>
                      {outletDiscrep > 0 && (
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-yellow-900/40 text-yellow-300 border border-yellow-700/40">
                          ⚠ {outletDiscrep} disc% mismatch
                        </span>
                      )}
                      {outletMissing > 0 && (
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-red-900/40 text-red-300 border border-red-700/40">
                          ⚠ {outletMissing} missing
                        </span>
                      )}
                      {outletMissing === 0 && outletDiscrep === 0 && (
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-[#10b981]/10 text-[#10b981] border border-[#10b981]/30">
                          ✓ All matched
                        </span>
                      )}
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse">
                      <thead>
                        <tr className="bg-[#020c08]/40">
                          {["Check #", "Date", "Created By", "Item Name", "Discount", "Total", "Disc %", "السبب / Reason"].map(h => (
                            <th key={h} className="px-3 py-2.5 text-left text-[10px] font-bold uppercase tracking-widest text-[#4b6b5d] border-b border-[#1a3d2f]">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {outlet.rows.map((row, ri) => <ResultRow key={ri} row={row} />)}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}
          </>
        )}
      </div>
    </div>
  );
}