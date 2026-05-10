import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';

// ══════════════════════════════════════════════════════════════════════════
//  خريطة الفروع
// ══════════════════════════════════════════════════════════════════════════
const BRANCH_MAP = [
  {
    daily_sales_key:  'Cloud 9',
    revenue_keywords: ['cloud 9', 'cloud9'],
    spot_keywords:    ['cloud 9', 'cloud9'],
    display_name:     'Cloud 9',
  },
  {
    daily_sales_key:  'El Rehab',
    revenue_keywords: ['tagamo', 'tagamooa', 'tagamo3', 'rehab', 'hd'],
    spot_keywords:    ['tagamo', 'tagamo3', 'rehab'],
    display_name:     'El Rehab / التجمع',
  },
  {
    daily_sales_key:  'CFC',
    revenue_keywords: ['cairo festival', 'festival city', 'cfc'],
    spot_keywords:    ['cairo festival', 'festival'],
    display_name:     'CFC (Cairo Festival City)',
  },
  {
    daily_sales_key:  'Zamalek',
    revenue_keywords: ['zamalek'],
    spot_keywords:    ['zamalek'],
    display_name:     'Zamalek',
  },
  {
    daily_sales_key:  'EL Gulf',
    revenue_keywords: ['ard el golf'],
    spot_keywords:    ['ard el golf'],
    display_name:     'EL Gulf / ارض الجولف',
    exclude_revenue:  ['golf central'],
    exclude_spot:     ['golf central'],
  },
  {
    daily_sales_key:  'Maadi',
    revenue_keywords: ['madii', 'maadi', "ma'adi"],
    spot_keywords:    ['madii', 'maadi'],
    display_name:     'Maadi',
  },
  {
    daily_sales_key:  'Almaza',
    revenue_keywords: ['almaza'],
    spot_keywords:    ['almaza'],
    display_name:     'Almaza',
  },
  {
    daily_sales_key:  'El shrouk',
    revenue_keywords: ['shorouk'],
    spot_keywords:    ['shorouk'],
    display_name:     'El Shorouk',
  },
  {
    daily_sales_key:  'Arkan Mall',
    revenue_keywords: ['arkan'],
    spot_keywords:    ['arkan'],
    display_name:     'Arkan Mall',
    exclude_revenue:  ['arkan delivery'],
    exclude_spot:     ['arkan delivery'],
  },
  {
    daily_sales_key:  'Arkan Delivery',
    revenue_keywords: ['arkan delivery'],
    spot_keywords:    ['arkan delivery'],
    display_name:     'Arkan Delivery',
  },
  {
    daily_sales_key:  'Square One',
    revenue_keywords: ['square one', 'square'],
    spot_keywords:    ['square one', 'square'],
    display_name:     'Square One',
  },
  {
    daily_sales_key:  'D5',
    revenue_keywords: ['district5', 'district 5', 'd5'],
    spot_keywords:    ['district5', 'district 5', 'd5'],
    display_name:     'D5 (District 5)',
  },
  {
    daily_sales_key:  'Golf Center',
    revenue_keywords: ['golf central'],
    spot_keywords:    ['golf central'],
    display_name:     'Golf Center',
    exclude_revenue:  ['ard el golf'],
    exclude_spot:     ['ard el golf'],
  },
];

// ── مطابقة اسم الفرع بكلمات مفتاحية ──────────────────────────────────────
const matchBranch = (name, keywords, excludes = []) => {
  const lower = String(name || '').toLowerCase();
  if (excludes.some(ex => lower.includes(ex.toLowerCase()))) return false;
  return keywords.some(kw => lower.includes(kw.toLowerCase()));
};

// ══════════════════════════════════════════════════════════════════════════
//  قراءة Daily Revenue
//  القيم في col index 1 دايماً
// ══════════════════════════════════════════════════════════════════════════
const parseRevenueFile = (data) => {
  const rows = Array.isArray(data) ? data : [];
  const result = {};
  let currentBranch = null;

  rows.forEach((row) => {
    const c0  = String(row[0] || '').trim();
    const val = parseFloat(row[1]) || 0;

    if (c0.toLowerCase().startsWith('outlet:')) {
      const outletName = c0.replace(/outlet:/i, '').trim();
      currentBranch = null;
      for (const b of BRANCH_MAP) {
        if (matchBranch(outletName, b.revenue_keywords, b.exclude_revenue || [])) {
          if (!result[b.daily_sales_key]) {
            result[b.daily_sales_key] = { food: 0, beverage: 0, other: 0 };
          }
          currentBranch = b.daily_sales_key;
          break;
        }
      }
      return;
    }

    if (!currentBranch) return;

    if      (c0 === 'Total Food')    result[currentBranch].food     += val;
    else if (c0 === 'Total Bevarge') result[currentBranch].beverage += val;
    else if (c0 === 'Total Other')   result[currentBranch].other    += val;
  });

  return result;
};

// ══════════════════════════════════════════════════════════════════════════
//  قراءة Spot Check
//
//  هيكل الملف الحقيقي (بعد التحقق):
//    'Eat In Sales'    → col1 = value,  col2 = 'N Chks'
//    'Take Away Sales' → col1 = value,  col2 = 'N Chks'
//    'Delivery Sales'  → col1 = value,  col2 = 'N Chks'
//    'Delivery Charge' → col1 = null,   col2 = value   ← مهم!
//    'Service Charge'  → col1 = null,   col2 = value   ← مهم!
//    'VAT 14%'         → col1 = null,   col2 = value   ← مهم!
//    'Total Discount'  → col1 = null,   col2 = value   ← مهم!
//    'Grand Total'     → col1 = value,  col2 = value
// ══════════════════════════════════════════════════════════════════════════
const parseChks = (val) => {
  if (val == null) return 0;
  const s = String(val).trim();
  const m = s.match(/^(\d+)\s*[Cc]hks?/);
  return m ? parseInt(m[1], 10) : 0;
};

const parseSpotCheckFile = (data) => {
  const rows = Array.isArray(data) ? data : [];
  const result = {};
  let currentBranch = null;

  rows.forEach((row) => {
    const c0 = String(row[0] || '').trim();
    const c1 = row[1];   // القيم الرئيسية: Eat In / Take Away / Delivery / Grand Total
    const c2 = row[2];   // Chks للـ sales rows، والقيم الحقيقية لـ Service/VAT/Discount/DeliveryCharge

    if (c0.toLowerCase().startsWith('outlet:')) {
      const outletName = c0.replace(/outlet:/i, '').trim();
      currentBranch = null;

      for (const b of BRANCH_MAP) {
        if (matchBranch(outletName, b.spot_keywords, b.exclude_spot || [])) {
          if (!result[b.daily_sales_key]) {
            result[b.daily_sales_key] = {
              dine_in: 0,       dine_in_chks: 0,
              take_away: 0,     take_away_chks: 0,
              delivery: 0,      delivery_chks: 0,
              delivery_charge: 0,
              discounts: 0,
              service_charge: 0,
              vat: 0,
              grand_total: 0,
            };
          }
          currentBranch = b.daily_sales_key;
          break;
        }
      }
      return;
    }

    if (!currentBranch) return;
    const r = result[currentBranch];

    if (c0 === 'Eat In Sales') {
      r.dine_in      = parseFloat(c1) || 0;
      r.dine_in_chks = parseChks(c2);
    } else if (c0 === 'Take Away Sales') {
      r.take_away       = parseFloat(c1) || 0;
      r.take_away_chks  = parseChks(c2);
    } else if (c0 === 'Delivery Sales') {
      r.delivery        = parseFloat(c1) || 0;
      r.delivery_chks   = parseChks(c2);
    } else if (c0 === 'Delivery Charge') {
      // القيمة في col2 وليس col1 في ملف الـ Spot Check الحقيقي
      r.delivery_charge = parseFloat(c2) || 0;
    } else if (c0 === 'Total Discount') {
      // القيمة في col2
      r.discounts       = parseFloat(c2) || 0;
    } else if (c0 === 'Service Charge') {
      // القيمة في col2
      r.service_charge  = parseFloat(c2) || 0;
    } else if (c0 === 'VAT 14%') {
      // القيمة في col2
      r.vat             = parseFloat(c2) || 0;
    } else if (c0 === 'Grand Total') {
      r.grand_total     = parseFloat(c1) || 0;
    }
  });

  return result;
};

// ── قراءة ملف XLSX/XLS في المتصفح ─────────────────────────────────────────
const readFileAsRows = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb   = XLSX.read(data, { type: 'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

// ══════════════════════════════════════════════════════════════════════════
//  تصدير Daily Sales Report
// ══════════════════════════════════════════════════════════════════════════
const exportFilledExcel = (revenueData, spotData, date) => {
  const wb   = XLSX.utils.book_new();
  const rows = [];

  const branchOrder = [
    'Cloud 9', 'El Rehab', 'CFC', 'Zamalek', 'EL Gulf',
    'Maadi', 'Almaza', 'El shrouk', 'Arkan Mall', 'Arkan Delivery',
    'Square One', 'D5', 'Golf Center',
  ];

  const fmt = (n) => (n != null && n !== 0) ? Math.round(n * 100) / 100 : null;

  // Header Row 1
  rows.push([
    'Restaurant',
    'Sales Chanel', null, null, null, null, null, null, null, null,
    null,
    'Sales Type', null, null, null,
  ]);

  // Header Row 2
  rows.push([
    null, null,
    'Dine in', 'Take away', 'Home Delivery', 'Delivery Charges',
    'Discounts', 'Service Charges 12%', 'VAT 14%', 'Total Sales',
    null,
    'Food', 'Beverage', 'Aggregator Serv. Fees', 'Total',
  ]);

  const dataRowStart = 3; // 1-based row index of first data row (after 2 header rows)

  branchOrder.forEach((key, idx) => {
    const spot    = spotData[key]    || {};
    const revenue = revenueData[key] || {};

    // كل فرع = 4 صفوف: Sales / % / Trans / Trans%
    const salesRowNum = dataRowStart + idx * 4;
    const pctRowNum   = salesRowNum + 1;
    const transRowNum = salesRowNum + 2;
    const transPctNum = salesRowNum + 3;

    const dineIn          = fmt(spot.dine_in)        ?? null;
    const takeAway        = fmt(spot.take_away)       ?? null;
    const homeDelivery    = fmt(spot.delivery)        ?? null;
    const deliveryCharges = fmt(spot.delivery_charge) || null;
    const discounts       = fmt(spot.discounts)       || null;
    const serviceCharges  = fmt(spot.service_charge)  ?? null;
    const vat             = fmt(spot.vat)             ?? null;
    const food            = fmt(revenue.food)         ?? null;
    const beverage        = fmt(revenue.beverage)     ?? null;
    const other           = fmt(revenue.other)        || null;

    const dineInChks    = spot.dine_in_chks   || null;
    const takeAwayChks  = spot.take_away_chks || null;
    const deliveryChks  = spot.delivery_chks  || null;

    // حساب الـ totals مسبقاً لاستخدامها في النسب
    const totalSalesVal   = (spot.dine_in || 0) + (spot.take_away || 0) + (spot.delivery || 0)
                          + (spot.delivery_charge || 0) + (spot.discounts || 0)
                          + (spot.service_charge || 0) + (spot.vat || 0);
    const revTotalVal     = (revenue.food || 0) + (revenue.beverage || 0) + (revenue.other || 0);
    const totalTransVal   = (spot.dine_in_chks || 0) + (spot.take_away_chks || 0) + (spot.delivery_chks || 0);

    // دالة النسبة - بترجع نص مثل "70%" أو null
    const pct = (num, den) => den ? Math.round((num / den) * 100) / 100 : null;

    const totalSalesCols = `C${salesRowNum}:I${salesRowNum}`;

    // Sales row
    rows.push([
      key, 'Sales',
      dineIn, takeAway, homeDelivery, deliveryCharges, discounts, serviceCharges, vat,
      { f: `SUM(${totalSalesCols})` },
      null,
      food, beverage, other,
      { f: `+L${salesRowNum}+M${salesRowNum}+N${salesRowNum}` },
    ]);

    // % row — Sales Channel %  (نسب محسوبة مسبقاً لتفادي مشكلة الـ formula recalc)
    rows.push([
      null, '%',
      pct(spot.dine_in        || 0, totalSalesVal),
      pct(spot.take_away      || 0, totalSalesVal),
      pct(spot.delivery       || 0, totalSalesVal),
      pct(spot.delivery_charge|| 0, totalSalesVal),
      pct(spot.discounts      || 0, totalSalesVal),
      pct(spot.service_charge || 0, totalSalesVal),
      pct(spot.vat            || 0, totalSalesVal),
      totalSalesVal ? 1 : null,   // 100% = 1 في Excel
      null,
      // Sales Type %
      pct(revenue.food      || 0, revTotalVal),
      pct(revenue.beverage  || 0, revTotalVal),
      pct(revenue.other     || 0, revTotalVal),
      revTotalVal ? 1 : null,
    ]);

    // Trans row
    rows.push([
      null, 'Trans ',
      dineInChks,
      takeAwayChks,
      deliveryChks,
      null, null, null, null,
      totalTransVal || null,   // قيمة مباشرة بدل formula
      null,
      null, null, null, null,
    ]);

    // Trans % row — نسب محسوبة مسبقاً
    rows.push([
      null, '%',
      pct(spot.dine_in_chks   || 0, totalTransVal),
      pct(spot.take_away_chks || 0, totalTransVal),
      pct(spot.delivery_chks  || 0, totalTransVal),
      null, null, null, null,
      totalTransVal ? 1 : null,
      null, null, null, null, null,
    ]);
  });

  // ── حساب التوتالات مسبقاً ──────────────────────────────────────────────
  const totDineIn    = branchOrder.reduce((s,k) => s+(spotData[k]?.dine_in||0), 0);
  const totTakeAway  = branchOrder.reduce((s,k) => s+(spotData[k]?.take_away||0), 0);
  const totDelivery  = branchOrder.reduce((s,k) => s+(spotData[k]?.delivery||0), 0);
  const totDelCharge = branchOrder.reduce((s,k) => s+(spotData[k]?.delivery_charge||0), 0);
  const totDiscount  = branchOrder.reduce((s,k) => s+(spotData[k]?.discounts||0), 0);
  const totService   = branchOrder.reduce((s,k) => s+(spotData[k]?.service_charge||0), 0);
  const totVat       = branchOrder.reduce((s,k) => s+(spotData[k]?.vat||0), 0);
  const totSales     = totDineIn+totTakeAway+totDelivery+totDelCharge+totDiscount+totService+totVat;
  const totFood      = branchOrder.reduce((s,k) => s+(revenueData[k]?.food||0), 0);
  const totBeverage  = branchOrder.reduce((s,k) => s+(revenueData[k]?.beverage||0), 0);
  const totOther     = branchOrder.reduce((s,k) => s+(revenueData[k]?.other||0), 0);
  const totRevenue   = totFood+totBeverage+totOther;
  const totDineInChks   = branchOrder.reduce((s,k) => s+(spotData[k]?.dine_in_chks||0), 0);
  const totTakeAwayChks = branchOrder.reduce((s,k) => s+(spotData[k]?.take_away_chks||0), 0);
  const totDelivChks    = branchOrder.reduce((s,k) => s+(spotData[k]?.delivery_chks||0), 0);
  const totTrans        = totDineInChks+totTakeAwayChks+totDelivChks;
  const pctT = (num, den) => den ? Math.round((num/den)*100)/100 : null;
  const fmtT = (n) => (n != null && n !== 0) ? Math.round(n*100)/100 : null;

  // Total row

  rows.push([
    'Total', 'Sales',
    fmtT(totDineIn), fmtT(totTakeAway), fmtT(totDelivery),
    fmtT(totDelCharge), fmtT(totDiscount), fmtT(totService), fmtT(totVat),
    fmtT(totSales),
    null,
    fmtT(totFood), fmtT(totBeverage), fmtT(totOther), fmtT(totRevenue),
  ]);

  // Total %
  rows.push([
    null, '%',
    pctT(totDineIn, totSales), pctT(totTakeAway, totSales), pctT(totDelivery, totSales),
    pctT(totDelCharge, totSales), pctT(totDiscount, totSales),
    pctT(totService, totSales), pctT(totVat, totSales),
    totSales ? 1 : null,
    null,
    pctT(totFood, totRevenue), pctT(totBeverage, totRevenue),
    pctT(totOther, totRevenue), totRevenue ? 1 : null,
  ]);

  // Total Trans
  rows.push([
    null, 'Trans',
    totDineInChks || null, totTakeAwayChks || null, totDelivChks || null,
    null, null, null, null,
    totTrans || null,
    null, null, null, null, null,
  ]);

  // Total Trans %
  rows.push([
    null, '%',
    pctT(totDineInChks, totTrans), pctT(totTakeAwayChks, totTrans), pctT(totDelivChks, totTrans),
    null, null, null, null,
    totTrans ? 1 : null,
    null, null, null, null, null,
  ]);

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // ── تطبيق format النسبة على صفوف الـ % ──────────────────────────────────
  const pctFmt = '0%';
  const totalBranches = branchOrder.length;
  const pctRowsIdx = [];
  for (let i = 0; i < totalBranches; i++) {
    pctRowsIdx.push(2 + i * 4 + 1); // Sales %
    pctRowsIdx.push(2 + i * 4 + 3); // Trans %
  }
  pctRowsIdx.push(2 + totalBranches * 4 + 1); // Total Sales %
  pctRowsIdx.push(2 + totalBranches * 4 + 3); // Total Trans %
  const pctCols = ['C','D','E','F','G','H','I','J','L','M','N','O'];
  pctRowsIdx.forEach(ri => {
    pctCols.forEach(col => {
      const addr = col + (ri + 1);
      if (ws[addr] && ws[addr].v != null) { ws[addr].t = 'n'; ws[addr].z = pctFmt; }
    });
  });

  ws['!cols'] = [
    { wch: 18 }, { wch: 8 },
    { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 17 }, { wch: 12 },
    { wch: 20 }, { wch: 10 }, { wch: 14 }, { wch: 2 },
    { wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 14 },
  ];

  ws['!merges'] = [
    { s: { r: 0, c: 0 },  e: { r: 1, c: 0 }  },
    { s: { r: 0, c: 1 },  e: { r: 0, c: 9 }  },
    { s: { r: 0, c: 11 }, e: { r: 0, c: 14 } },
  ];

  XLSX.utils.book_append_sheet(wb, ws, date || 'Report');
  const filename = `Daily_Sales_Report_${date || new Date().toLocaleDateString('en-GB').replace(/\//g, '-')}.xlsx`;
  XLSX.writeFile(wb, filename);
};

// ══════════════════════════════════════════════════════════════════════════
//  مكوّن معاينة الجدول
// ══════════════════════════════════════════════════════════════════════════
const PreviewTable = ({ revenueData, spotData }) => {
  const branchOrder = [
    'Cloud 9', 'El Rehab', 'CFC', 'Zamalek', 'EL Gulf',
    'Maadi', 'Almaza', 'El shrouk', 'Arkan Mall', 'Arkan Delivery',
    'Square One', 'D5', 'Golf Center',
  ];

  const fmt  = (n) => (n != null && n !== 0) ? n.toLocaleString('en-EG', { maximumFractionDigits: 0 }) : '—';
  const fmtN = (n) => (n != null && n !== 0) ? n.toString() : '—';
  const fmtPct = (num, den) => {
    if (!den || !num) return '—';
    return Math.round((num / den) * 100) + '%';
  };

  return (
    <div className="overflow-x-auto rounded-xl border border-emerald-900/40">
      <table className="w-full text-xs border-collapse" style={{ minWidth: '1300px' }}>
        <thead>
          <tr className="bg-[#0a1f14]">
            <th rowSpan={2} className="px-3 py-3 text-left text-emerald-400 font-bold border-b border-emerald-900/30 sticky left-0 bg-[#0a1f14] z-10">
              الفرع
            </th>
            <th rowSpan={2} className="px-2 py-2 text-center text-gray-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20 text-[10px]">
              نوع
            </th>
            <th colSpan={7} className="px-3 py-2 text-center text-cyan-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">
              Sales Channel (Spot Check)
            </th>
            <th className="px-3 py-2 text-center text-emerald-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">
              Total
            </th>
            <th colSpan={4} className="px-3 py-2 text-center text-amber-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">
              Sales Type (Daily Revenue)
            </th>
            <th className="px-3 py-2 text-center text-emerald-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">
              Grand Total ✓
            </th>
          </tr>
          <tr className="bg-[#071510] text-gray-400">
            {['Dine In','Take Away','Home Delivery','Del. Charges','Discounts','Service 12%','VAT 14%'].map(h => (
              <th key={h} className="px-2 py-1.5 text-center border-b border-emerald-900/20 border-l border-emerald-900/10 whitespace-nowrap text-[10px]">{h}</th>
            ))}
            <th className="px-2 py-1.5 text-center border-b border-emerald-900/20 border-l border-emerald-900/10 text-[10px]">Total Sales</th>
            {['Food','Beverage','Other','Total'].map(h => (
              <th key={h} className="px-2 py-1.5 text-center border-b border-emerald-900/20 border-l border-emerald-900/10 text-[10px]">{h}</th>
            ))}
            <th className="px-2 py-1.5 text-center border-b border-emerald-900/20 border-l border-emerald-900/10 text-[10px]">Grand Total</th>
          </tr>
        </thead>
        <tbody>
          {branchOrder.map((key, idx) => {
            const spot    = spotData[key]    || {};
            const revenue = revenueData[key] || {};
            const hasData = Object.keys(spot).length > 0 || Object.keys(revenue).length > 0;

            const totalSales = (spot.dine_in || 0) + (spot.take_away || 0) + (spot.delivery || 0)
              + (spot.delivery_charge || 0) + (spot.discounts || 0)
              + (spot.service_charge  || 0) + (spot.vat || 0);
            const revenueTotal = (revenue.food || 0) + (revenue.beverage || 0) + (revenue.other || 0);
            const grandTotal = spot.grand_total || 0;
            const isMatch = grandTotal > 0 && Math.abs(totalSales - grandTotal) < 1;

            const totalTrans = (spot.dine_in_chks || 0) + (spot.take_away_chks || 0) + (spot.delivery_chks || 0);
            const bgEven = 'bg-[#040d08]';
            const bgOdd  = 'bg-[#060f0a]';
            const bg = idx % 2 === 0 ? bgEven : bgOdd;

            return (
              <React.Fragment key={key}>
                {/* Sales Row */}
                <tr className={`border-b border-emerald-900/10 transition-colors ${bg} hover:bg-[#0d2a1f]/40`}>
                  <td rowSpan={4} className={`px-3 py-2 font-medium sticky left-0 z-10 border-r border-emerald-900/20 align-middle ${bg} ${!hasData ? 'text-gray-600' : 'text-white'}`}>
                    {BRANCH_MAP.find(b => b.daily_sales_key === key)?.display_name || key}
                  </td>
                  <td className="px-2 py-1.5 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">Sales</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.dine_in)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.take_away)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.delivery)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.delivery_charge)}</td>
                  <td className={`px-2 py-1.5 text-center border-l border-emerald-900/10 ${(spot.discounts || 0) < 0 ? 'text-red-400/80' : 'text-cyan-300/80'}`}>
                    {fmt(spot.discounts)}
                  </td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.service_charge)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.vat)}</td>
                  <td className="px-2 py-1.5 text-center text-emerald-300/80 font-semibold border-l border-emerald-900/20">{fmt(totalSales)}</td>
                  {/* Sales Type */}
                  <td className="px-2 py-1.5 text-center text-amber-300/80 border-l border-emerald-900/20">{fmt(revenue.food)}</td>
                  <td className="px-2 py-1.5 text-center text-amber-300/80 border-l border-emerald-900/10">{fmt(revenue.beverage)}</td>
                  <td className="px-2 py-1.5 text-center text-amber-300/80 border-l border-emerald-900/10">{fmt(revenue.other)}</td>
                  <td className="px-2 py-1.5 text-center text-amber-400/80 font-semibold border-l border-emerald-900/10">{fmt(revenueTotal)}</td>
                  {/* Grand Total */}
                  <td rowSpan={4} className={`px-2 py-1.5 text-center font-bold border-l border-emerald-900/20 align-middle ${
                    !hasData  ? 'text-gray-600' :
                    isMatch   ? 'text-emerald-400' : 'text-red-400'
                  }`}>
                    {grandTotal > 0 ? (
                      <span title={isMatch ? '✅ مطابق' : `⚠️ فرق: ${Math.abs(totalSales - grandTotal).toLocaleString()}`}>
                        {isMatch ? '✅' : '⚠️'}<br />{fmt(grandTotal)}
                      </span>
                    ) : '—'}
                  </td>
                </tr>

                {/* % Row — Sales Channel % + Sales Type % */}
                <tr className={`border-b border-emerald-900/5 ${bg} opacity-70`}>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">%</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.dine_in, totalSales)}</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.take_away, totalSales)}</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.delivery, totalSales)}</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.delivery_charge, totalSales)}</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.discounts, totalSales)}</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.service_charge, totalSales)}</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.vat, totalSales)}</td>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/20 text-[10px]">100%</td>
                  {/* Sales Type % — مقسومة على revenueTotal */}
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/20 text-[10px]">{fmtPct(revenue.food, revenueTotal)}</td>
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/10 text-[10px]">{fmtPct(revenue.beverage, revenueTotal)}</td>
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/10 text-[10px]">{fmtPct(revenue.other, revenueTotal)}</td>
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/10 text-[10px]">{revenueTotal > 0 ? '100%' : '—'}</td>
                </tr>

                {/* Trans Row */}
                <tr className={`border-b border-emerald-900/5 ${bg}`}>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">Trans</td>
                  <td className="px-2 py-1 text-center text-purple-300/70 border-l border-emerald-900/10 text-[10px]">{fmtN(spot.dine_in_chks)}</td>
                  <td className="px-2 py-1 text-center text-purple-300/70 border-l border-emerald-900/10 text-[10px]">{fmtN(spot.take_away_chks)}</td>
                  <td className="px-2 py-1 text-center text-purple-300/70 border-l border-emerald-900/10 text-[10px]">{fmtN(spot.delivery_chks)}</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-purple-400/80 font-semibold border-l border-emerald-900/20 text-[10px]">{fmtN(totalTrans)}</td>
                  <td colSpan={4} className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/20 text-[10px]">—</td>
                </tr>

                {/* Trans % Row */}
                <tr className={`border-b border-emerald-900/10 ${bg} opacity-60`}>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">%</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.dine_in_chks, totalTrans)}</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.take_away_chks, totalTrans)}</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.delivery_chks, totalTrans)}</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/20 text-[10px]">{totalTrans > 0 ? '100%' : '—'}</td>
                  <td colSpan={4} className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/20 text-[10px]">—</td>
                </tr>
              </React.Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
};

// ══════════════════════════════════════════════════════════════════════════
//  الصفحة الرئيسية
// ══════════════════════════════════════════════════════════════════════════
export default function DailySalesReportFiller() {
  const [revenueFile, setRevenueFile] = useState(null);
  const [spotFile,    setSpotFile]    = useState(null);
  const [revenueData, setRevenueData] = useState({});
  const [spotData,    setSpotData]    = useState({});
  const [isReady,     setIsReady]     = useState(false);
  const [isLoading,   setIsLoading]   = useState(false);
  const [reportDate,  setReportDate]  = useState('');
  const [errors,      setErrors]      = useState([]);

  const handleRevenue = useCallback(async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRevenueFile(file.name);
    try {
      const rows = await readFileAsRows(file);
      const data = parseRevenueFile(rows);
      setRevenueData(data);
    } catch (err) {
      setErrors(prev => [...prev, `خطأ في قراءة Daily Revenue: ${err.message}`]);
    }
  }, []);

  const handleSpot = useCallback(async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setSpotFile(file.name);
    try {
      const rows = await readFileAsRows(file);
      const data = parseSpotCheckFile(rows);
      setSpotData(data);
    } catch (err) {
      setErrors(prev => [...prev, `خطأ في قراءة Spot Check: ${err.message}`]);
    }
  }, []);

  const handleAnalyze = () => {
    if (!revenueFile || !spotFile) {
      setErrors(['من فضلك ارفع الملفين أولاً']);
      return;
    }
    setErrors([]);
    setIsLoading(true);
    setTimeout(() => { setIsReady(true); setIsLoading(false); }, 500);
  };

  const handleExport = () => {
    exportFilledExcel(revenueData, spotData, reportDate);
  };

  const totalFood     = Object.values(revenueData).reduce((s, b) => s + (b.food     || 0), 0);
  const totalBeverage = Object.values(revenueData).reduce((s, b) => s + (b.beverage || 0), 0);
  const totalOther    = Object.values(revenueData).reduce((s, b) => s + (b.other    || 0), 0);
  const totalGrand    = Object.values(spotData).reduce((s, b)    => s + (b.grand_total || 0), 0);
  const totalTrans    = Object.values(spotData).reduce((s, b)    => s + (b.dine_in_chks || 0) + (b.take_away_chks || 0) + (b.delivery_chks || 0), 0);
  const fmt = (n) => Math.round(n).toLocaleString('en-EG');

  return (
    <div className="max-w-7xl mx-auto px-4 py-8 space-y-8 bg-[#020c08] min-h-screen text-right" dir="rtl">

      {/* Header */}
      <header className="text-center space-y-2">
        <div className="flex items-center justify-center gap-3">
          <span className="text-4xl">📊</span>
          <h1 className="text-3xl font-bold text-[#10b981]">Daily Sales Report Filler</h1>
        </div>
        <p className="text-gray-400">
          ارفع ملف Daily Revenue وSpot Check وهيتملأ الشيت تلقائياً
        </p>
      </header>

      {/* أخطاء */}
      {errors.length > 0 && (
        <div className="bg-red-900/20 border border-red-700/50 rounded-xl p-4 space-y-1">
          {errors.map((e, i) => (
            <p key={i} className="text-red-400 text-sm">⚠️ {e}</p>
          ))}
        </div>
      )}

      {!isReady ? (
        /* Upload Panel */
        <div className="bg-gray-900 p-10 rounded-2xl border border-gray-800 shadow-2xl max-w-2xl mx-auto space-y-6">

          {/* Date */}
          <div className="flex flex-col gap-2">
            <label className="text-emerald-400 font-bold text-sm">تاريخ التقرير (اختياري):</label>
            <input
              type="date"
              value={reportDate}
              onChange={e => setReportDate(e.target.value)}
              className="bg-black/40 border border-emerald-900/40 rounded-lg px-4 py-2 text-white text-sm focus:outline-none focus:border-emerald-500"
            />
          </div>

          {/* Daily Revenue */}
          <div className="flex flex-col gap-2">
            <label className="text-emerald-400 font-bold">📁 Daily Revenue (.xls / .xlsx):</label>
            <input
              type="file"
              accept=".xls,.xlsx"
              onChange={handleRevenue}
              className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg file:cursor-pointer hover:file:bg-emerald-700"
            />
            {revenueFile && <p className="text-emerald-500 text-xs mt-1">✅ {revenueFile} — تم القراءة بنجاح</p>}
          </div>

          {/* Spot Check */}
          <div className="flex flex-col gap-2">
            <label className="text-emerald-400 font-bold">📁 Spot Check (.xls / .xlsx):</label>
            <input
              type="file"
              accept=".xls,.xlsx"
              onChange={handleSpot}
              className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg file:cursor-pointer hover:file:bg-emerald-700"
            />
            {spotFile && <p className="text-emerald-500 text-xs mt-1">✅ {spotFile} — تم القراءة بنجاح</p>}
          </div>

          {/* ملاحظة الفروع */}
          <div className="bg-black/30 border border-emerald-900/20 rounded-xl p-4 text-xs text-gray-500 space-y-1">
            <p className="text-gray-400 font-semibold mb-2">خريطة الفروع المطبقة:</p>
            <p>• EL Gulf = ارض الجولف (Ard El Golf) فقط</p>
            <p>• Golf Center = Golf Central (فرع مستقل)</p>
            <p>• التجمع (Tagamo3) = El Rehab في الشيت</p>
            <p>• D5 / District5 = D5 في الشيت</p>
            <p>• Cairo Festival City = CFC في الشيت</p>
            <p>• Arkan Mall ≠ Arkan Delivery (فرعان منفصلان)</p>
          </div>

          <button
            onClick={handleAnalyze}
            disabled={!revenueFile || !spotFile || isLoading}
            className="w-full py-4 bg-[#10b981] text-black font-black rounded-xl hover:bg-emerald-400 transition-all shadow-lg shadow-emerald-900/20 disabled:opacity-40 disabled:cursor-not-allowed"
          >
            {isLoading ? 'جاري المعالجة...' : '⚡ معالجة البيانات وعرض المعاينة'}
          </button>
        </div>

      ) : (
        /* Results Panel */
        <div className="space-y-6">

          {/* Summary Cards */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            {[
              { label: 'Total Food',     value: fmt(totalFood),     color: 'text-amber-400',   icon: '🍱' },
              { label: 'Total Beverage', value: fmt(totalBeverage), color: 'text-cyan-400',    icon: '🥤' },
              { label: 'Total Other',    value: fmt(totalOther),    color: 'text-purple-400',  icon: '📦' },
              { label: 'Grand Total',    value: fmt(totalGrand),    color: 'text-emerald-400', icon: '💰' },
              { label: 'Total Trans',    value: fmt(totalTrans),    color: 'text-pink-400',    icon: '🧾' },
            ].map(card => (
              <div key={card.label} className="bg-gray-900 rounded-xl border border-gray-800 p-4 text-center">
                <div className="text-2xl mb-1">{card.icon}</div>
                <div className={`text-xl font-bold ${card.color}`}>{card.value}</div>
                <div className="text-gray-500 text-xs mt-1">{card.label}</div>
              </div>
            ))}
          </div>

          {/* Preview Table */}
          <div className="bg-[#0d2a1f] p-4 rounded-xl border border-emerald-900 shadow-xl">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-lg font-bold text-white flex gap-2 items-center">
                <span className="w-2.5 h-2.5 bg-emerald-500 rounded-full animate-pulse"></span>
                معاينة البيانات المستخرجة
              </h2>
              <div className="text-xs text-gray-500">
                ✅ = Grand Total مطابق &nbsp;|&nbsp; ⚠️ = فيه فرق راجع الأرقام
              </div>
            </div>
            <PreviewTable revenueData={revenueData} spotData={spotData} />
          </div>

          {/* Export */}
          <div className="flex gap-4 justify-center flex-wrap">
            <button
              onClick={handleExport}
              className="bg-emerald-600 hover:bg-emerald-500 text-white px-10 py-3 rounded-xl font-bold transition-colors shadow-lg shadow-emerald-900/20"
            >
              📥 تصدير Daily Sales Report (Excel)
            </button>
            <button
              onClick={() => { setIsReady(false); setErrors([]); }}
              className="text-gray-500 hover:text-emerald-400 underline py-3 px-4 transition-colors"
            >
              ارفع ملفات جديدة
            </button>
          </div>

        </div>
      )}
    </div>
  );
}