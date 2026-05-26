import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';

// ══════════════════════════════════════════════════════════════════════════
//  خريطة الفروع - الأداة الأولى (Daily Sales Report)
// ══════════════════════════════════════════════════════════════════════════
const BRANCH_MAP = [
  { daily_sales_key: 'Cloud 9', revenue_keywords: ['cloud 9', 'cloud9'], spot_keywords: ['cloud 9', 'cloud9'], display_name: 'Cloud 9' },
  { daily_sales_key: 'El Rehab', revenue_keywords: ['tagamo', 'tagamooa', 'tagamo3', 'rehab', 'hd'], spot_keywords: ['tagamo', 'tagamo3', 'rehab'], display_name: 'El Rehab / التجمع' },
  { daily_sales_key: 'CFC', revenue_keywords: ['cairo festival', 'festival city', 'cfc'], spot_keywords: ['cairo festival', 'festival'], display_name: 'CFC (Cairo Festival City)' },
  { daily_sales_key: 'Zamalek', revenue_keywords: ['zamalek'], spot_keywords: ['zamalek'], display_name: 'Zamalek' },
  { daily_sales_key: 'EL Gulf', revenue_keywords: ['ard el golf'], spot_keywords: ['ard el golf'], display_name: 'EL Gulf / ارض الجولف', exclude_revenue: ['golf central'], exclude_spot: ['golf central'] },
  { daily_sales_key: 'Maadi', revenue_keywords: ['madii', 'maadi', "ma'adi"], spot_keywords: ['madii', 'maadi'], display_name: 'Maadi' },
  { daily_sales_key: 'Almaza', revenue_keywords: ['almaza'], spot_keywords: ['almaza'], display_name: 'Almaza' },
  { daily_sales_key: 'El shrouk', revenue_keywords: ['shorouk'], spot_keywords: ['shorouk'], display_name: 'El Shorouk' },
  { daily_sales_key: 'Arkan Mall', revenue_keywords: ['arkan'], spot_keywords: ['arkan'], display_name: 'Arkan Mall', exclude_revenue: ['arkan delivery'], exclude_spot: ['arkan delivery'] },
  { daily_sales_key: 'Arkan Delivery', revenue_keywords: ['arkan delivery'], spot_keywords: ['arkan delivery'], display_name: 'Arkan Delivery' },
  { daily_sales_key: 'Square One', revenue_keywords: ['square one', 'square'], spot_keywords: ['square one', 'square'], display_name: 'Square One' },
  { daily_sales_key: 'D5', revenue_keywords: ['district5', 'district 5', 'd5'], spot_keywords: ['district5', 'district 5', 'd5'], display_name: 'D5 (District 5)' },
  { daily_sales_key: 'Golf Center', revenue_keywords: ['golf central'], spot_keywords: ['golf central'], display_name: 'Golf Center', exclude_revenue: ['ard el golf'], exclude_spot: ['ard el golf'] },
];

const BRANCH_ORDER = [
  'Cloud 9','El Rehab','CFC','Zamalek','EL Gulf',
  'Maadi','Almaza','El shrouk','Arkan Mall','Arkan Delivery',
  'Square One','D5','Golf Center',
];

// ══════════════════════════════════════════════════════════════════════════
//  خريطة فروع Garnell - الأداة الثانية
// ══════════════════════════════════════════════════════════════════════════
const GARNELL_BRANCHES = [
  { key: 'Arkan',        no: '01', name: 'اركان',    rev_kw: ['garnell arkan'],       spot_kw: ['garnell arkan'],       exclude_spot: ['arkan delivery'] },
  { key: 'ArdElGolf',   no: '02', name: 'ارض الجولف',          rev_kw: ['ard el golf'],          spot_kw: ['ard el golf'],         exclude_spot: [] },
  { key: 'Shorouk',     no: '03', name: 'الشروق',         rev_kw: ['el shorouk', 'shorouk'],spot_kw: ['el shorouk', 'shorouk'],exclude_spot: [] },
  { key: 'Zamalek',     no: '04', name: 'الزمالك',        rev_kw: ['zamalek'],              spot_kw: ['zamalek'],             exclude_spot: [] },
  { key: 'Almaza',      no: '05', name: 'الماظة',         rev_kw: ['almaza'],               spot_kw: ['almaza'],              exclude_spot: [] },
  { key: 'Maadi',       no: '06', name: 'المعادي',        rev_kw: ['madii', 'maadi'],       spot_kw: ['madii', 'maadi'],      exclude_spot: [] },
  { key: 'Cloud9',      no: '07', name: 'كلاود 9',        rev_kw: ['cloud 9', 'cloud9'],    spot_kw: ['cloud 9', 'cloud9'],   exclude_spot: [] },
  { key: 'Tagamo3',     no: '08', name: 'التجمع دليفري',  rev_kw: ['tagamo3', 'tagamo'],    spot_kw: ['tagamo3', 'tagamo'],   exclude_spot: [] },
  { key: 'CFC',         no: '09', name: 'كايرو فيستفال',  rev_kw: ['cairo festival', 'festival city', 'cfc'], spot_kw: ['cairo festival', 'festival'], exclude_spot: [] },
  { key: 'ArkanDel',    no: '10', name: 'اركان دليفري',   rev_kw: ['arkan delivery'],       spot_kw: ['arkan delivery'],      exclude_spot: [] },
  { key: 'SquareOne',   no: '11', name: 'سكوير 1',        rev_kw: ['square one', 'square'], spot_kw: ['square one', 'square'],exclude_spot: [] },
  { key: 'GolfCentral', no: '12', name: 'جولف سنترال',    rev_kw: ['golf central'],         spot_kw: ['golf central'],        exclude_spot: [] },
  { key: 'District5',   no: '13', name: 'ديستكريت 5',     rev_kw: ['district5', 'district 5', 'd5'], spot_kw: ['district5', 'district 5', 'd5'], exclude_spot: [] },
];

// ══════════════════════════════════════════════════════════════════════════
//  مساعدات مشتركة
// ══════════════════════════════════════════════════════════════════════════
const matchBranch = (name, keywords, excludes = []) => {
  const lower = String(name || '').toLowerCase();
  if (excludes.some(ex => lower.includes(ex.toLowerCase()))) return false;
  return keywords.some(kw => lower.includes(kw.toLowerCase()));
};

const parseChks = (val) => {
  if (val == null) return 0;
  const m = String(val).trim().match(/^(\d+)\s*[Cc]hks?/);
  return m ? parseInt(m[1], 10) : 0;
};

const readFileAsRows = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb   = XLSX.read(data, { type: 'array' });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }));
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

// ══════════════════════════════════════════════════════════════════════════
//  قراءة Daily Revenue (مشترك)
// ══════════════════════════════════════════════════════════════════════════
const parseRevenueFile = (data, branchMap, keyField) => {
  const rows = Array.isArray(data) ? data : [];
  const result = {};
  let currentBranch = null;

  rows.forEach((row) => {
    const c0  = String(row[0] || '').trim();
    const val = parseFloat(String(row[1] || '').replace(/[(),]/g, m => m === '(' ? '-' : m === ')' ? '' : '')) || 0;

    if (c0.toLowerCase().startsWith('outlet:')) {
      const outletName = c0.replace(/outlet:/i, '').trim();
      currentBranch = null;
      for (const b of branchMap) {
        if (matchBranch(outletName, b.rev_kw || b.revenue_keywords, b.exclude_revenue || [])) {
          const k = b[keyField] || b.key;
          if (!result[k]) result[k] = { food: 0, beverage: 0, other: 0 };
          currentBranch = k;
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
//  قراءة Spot Check (الأداة الأولى)
// ══════════════════════════════════════════════════════════════════════════
const parseSpotCheckFile = (data) => {
  const rows = Array.isArray(data) ? data : [];
  const result = {};
  let currentBranch = null;

  rows.forEach((row) => {
    const c0 = String(row[0] || '').trim();
    const c1 = row[1];
    const c2 = row[2];

    if (c0.toLowerCase().startsWith('outlet:')) {
      const outletName = c0.replace(/outlet:/i, '').trim();
      currentBranch = null;
      for (const b of BRANCH_MAP) {
        if (matchBranch(outletName, b.spot_keywords, b.exclude_spot || [])) {
          if (!result[b.daily_sales_key]) {
            result[b.daily_sales_key] = { dine_in: 0, dine_in_chks: 0, take_away: 0, take_away_chks: 0, delivery: 0, delivery_chks: 0, delivery_charge: 0, discounts: 0, service_charge: 0, vat: 0, grand_total: 0 };
          }
          currentBranch = b.daily_sales_key;
          break;
        }
      }
      return;
    }
    if (!currentBranch) return;
    const r = result[currentBranch];
    const v1 = parseFloat(c1) || 0;
    const v2 = parseFloat(c2) || 0;
    if      (c0 === 'Eat In Sales')    { r.dine_in = v1; r.dine_in_chks = parseChks(c2); }
    else if (c0 === 'Take Away Sales') { r.take_away = v1; r.take_away_chks = parseChks(c2); }
    else if (c0 === 'Delivery Sales')  { r.delivery = v1; r.delivery_chks = parseChks(c2); }
    else if (c0 === 'Delivery Charge') { r.delivery_charge = v2; }
    else if (c0 === 'Total Discount')  { r.discounts = v2; }
    else if (c0 === 'Service Charge')  { r.service_charge = v2; }
    else if (c0 === 'VAT 14%')         { r.vat = v2; }
    else if (c0 === 'Grand Total')     { r.grand_total = v1; }
  });
  return result;
};

// ══════════════════════════════════════════════════════════════════════════
//  قراءة Spot Check - Garnell (الأداة الثانية)
// ══════════════════════════════════════════════════════════════════════════
const parseGarnellSpotCheck = (data) => {
  const rows = Array.isArray(data) ? data : [];
  const result = {};
  let currentKey = null;

  rows.forEach((row) => {
    const c0 = String(row[0] || '').trim();
    const c1 = row[1];
    const c2 = row[2];

    if (c0.toLowerCase().startsWith('outlet:')) {
      const outletName = c0.replace(/outlet:/i, '').trim();
      currentKey = null;
      for (const b of GARNELL_BRANCHES) {
        if (matchBranch(outletName, b.spot_kw, b.exclude_spot || [])) {
          if (!result[b.key]) {
            result[b.key] = { dine_in: 0, dine_in_chks: 0, take_away: 0, take_away_chks: 0, delivery: 0, delivery_chks: 0, delivery_charge: 0, discounts: 0, service_charge: 0, vat: 0, grand_total: 0, closed_checks: 0 };
          }
          currentKey = b.key;
          break;
        }
      }
      return;
    }
    if (!currentKey) return;
    const r = result[currentKey];
    const v1 = parseFloat(String(c1 || '').replace(/[(),]/g, m => m === '(' ? '-' : '')) || 0;
    const v2 = parseFloat(String(c2 || '').replace(/[(),]/g, m => m === '(' ? '-' : '')) || 0;

    if      (c0 === 'Eat In Sales')    { r.dine_in = v1; r.dine_in_chks = parseChks(c2); }
    else if (c0 === 'Take Away Sales') { r.take_away = v1; r.take_away_chks = parseChks(c2); }
    else if (c0 === 'Delivery Sales')  { r.delivery = v1; r.delivery_chks = parseChks(c2); }
    else if (c0 === 'Delivery Charge') { r.delivery_charge = v2; }
    else if (c0 === 'Total Discount')  { r.discounts = v2; }
    else if (c0 === 'Service Charge')  { r.service_charge = v2; }
    else if (c0 === 'VAT 14%')         { r.vat = v2; }
    else if (c0 === 'Grand Total')     { r.grand_total = v1; }
    else if (c0 === 'Closed Checks')   { r.closed_checks = parseInt(String(c2 || '0')) || 0; }
  });
  return result;
};

// ══════════════════════════════════════════════════════════════════════════
//  تصدير Daily Sales Report (الأداة الأولى)
// ══════════════════════════════════════════════════════════════════════════
const exportFilledExcel = (revenueData, spotData, date) => {
  const wb  = XLSX.utils.book_new();
  const aoa = [];

  const ROWS_PER_BRANCH = 5;
  const DATA_START      = 3;
  const N               = BRANCH_ORDER.length;

  const sR = (idx) => DATA_START + idx * ROWS_PER_BRANCH;
  const tR = (idx) => sR(idx) + 2;

  const TOT_S  = DATA_START + N * ROWS_PER_BRANCH;
  const TOT_P  = TOT_S + 1;
  const TOT_T  = TOT_S + 2;
  const TOT_PT = TOT_S + 3;

  aoa.push(['Restaurant','Sales Chanel',null,null,null,null,null,null,null,null,null,'Sales Type',null,null,null]);
  aoa.push([null,null,'Dine in','Take away','Home Delivery','Delivery Charges','Discounts','Service Charges 12%','VAT 14%','Total Sales',null,'Food ','Beverage ','Aggregator Serv. Fees','Total']);

  BRANCH_ORDER.forEach((key, idx) => {
    const spot = spotData[key]    || {};
    const rev  = revenueData[key] || {};
    const SR = sR(idx);
    const TR = tR(idx);
    const nv = (v) => (v && v !== 0) ? Math.round(v * 100) / 100 : null;

    const dineIn = nv(spot.dine_in); const takeAway = nv(spot.take_away);
    const delivery = nv(spot.delivery); const delChg = nv(spot.delivery_charge);
    const disc = nv(spot.discounts); const svc = nv(spot.service_charge);
    const vat = nv(spot.vat); const food = nv(rev.food);
    const bev = nv(rev.beverage); const other = nv(rev.other);
    const tDine = spot.dine_in_chks || null; const tTake = spot.take_away_chks || null;
    const tDel  = spot.delivery_chks || null;

    aoa.push([key,'Sales',dineIn,takeAway,delivery,delChg,disc,svc,vat,{f:`SUM(C${SR}:I${SR})`},null,food,bev,other,{f:`+L${SR}+M${SR}+N${SR}`}]);
    aoa.push([null,'%',{f:`IFERROR(C${SR}/$J${SR},"0%")`},{f:`IFERROR(D${SR}/$J${SR},"0%")`},{f:`IFERROR(E${SR}/$J${SR},"0%")`},{f:`IFERROR(F${SR}/$J${SR},"0%")`},{f:`IFERROR(G${SR}/$J${SR},"0%")`},{f:`IFERROR(H${SR}/$J${SR},"0%")`},{f:`IFERROR(I${SR}/$J${SR},"0%")`},{f:`IFERROR(J${SR}/$J${SR},"0%")`},null,null,null,null,null]);
    aoa.push([null,'Trans ',tDine,tTake,tDel,null,null,null,null,{f:`SUM(C${TR}:I${TR})`},null,{f:`IFERROR(L${SR}/$O${SR},"0%")`},{f:`IFERROR(M${SR}/$O${SR},"0%")`},{f:`IFERROR(N${SR}/$O${SR},"0%")`},{f:`IFERROR(O${SR}/$O${SR},"0%")`}]);
    aoa.push([null,'%',{f:`IFERROR(C${TR}/$J${TR},"0%")`},{f:`IFERROR(D${TR}/$J${TR},"0%")`},{f:`IFERROR(E${TR}/$J${TR},"0%")`},{f:`IFERROR(F${TR}/$J${TR},"0%")`},{f:`IFERROR(G${TR}/$J${TR},"0%")`},{f:`IFERROR(H${TR}/$J${TR},"0%")`},{f:`IFERROR(I${TR}/$J${TR},"0%")`},{f:`IFERROR(J${TR}/$J${TR},"0%")`},null,null,null,null,null]);
    aoa.push([null,'avg',tDine?{f:`C${SR}/C${TR}`}:null,tTake?{f:`D${SR}/D${TR}`}:null,tDel?{f:`E${SR}/E${TR}`}:null,null,null,null,null,null,null,null,null,null,null]);
  });

  const salesRefs = BRANCH_ORDER.map((_, i) => sR(i));
  const transRefs = BRANCH_ORDER.map((_, i) => tR(i));
  const sumRef = (col, rows) => rows.map(r => `${col}${r}`).join('+');

  aoa.push(['Total ','Sales',{f:`+${sumRef('C',salesRefs)}`},{f:`+${sumRef('D',salesRefs)}`},{f:`+${sumRef('E',salesRefs)}`},{f:`+${sumRef('F',salesRefs)}`},{f:`+${sumRef('G',salesRefs)}`},{f:`+${sumRef('H',salesRefs)}`},{f:`+${sumRef('I',salesRefs)}`},{f:`SUM(C${TOT_S}:I${TOT_S})`},null,{f:`+${sumRef('L',salesRefs)}`},{f:`+${sumRef('M',salesRefs)}`},{f:`+${sumRef('N',salesRefs)}`},{f:`+${sumRef('O',salesRefs)}`}]);
  aoa.push([null,'%',{f:`IFERROR(C${TOT_S}/$J${TOT_S},"0%")`},{f:`IFERROR(D${TOT_S}/$J${TOT_S},"0%")`},{f:`IFERROR(E${TOT_S}/$J${TOT_S},"0%")`},{f:`IFERROR(F${TOT_S}/$J${TOT_S},"0%")`},{f:`IFERROR(G${TOT_S}/$J${TOT_S},"0%")`},{f:`IFERROR(H${TOT_S}/$J${TOT_S},"0%")`},{f:`IFERROR(I${TOT_S}/$J${TOT_S},"0%")`},{f:`IFERROR(J${TOT_S}/$J${TOT_S},"0%")`},null,null,null,null,null]);
  aoa.push([null,'Trans ',{f:`+${sumRef('C',transRefs)}`},{f:`+${sumRef('D',transRefs)}`},{f:`+${sumRef('E',transRefs)}`},{f:`+${sumRef('F',transRefs)}`},{f:`+${sumRef('G',transRefs)}`},{f:`+${sumRef('H',transRefs)}`},{f:`+${sumRef('I',transRefs)}`},{f:`SUM(C${TOT_T}:I${TOT_T})`},null,{f:`IFERROR(L${TOT_S}/$O${TOT_S},"0%")`},{f:`IFERROR(M${TOT_S}/$O${TOT_S},"0%")`},{f:`IFERROR(N${TOT_S}/$O${TOT_S},"0%")`},{f:`IFERROR(O${TOT_S}/$O${TOT_S},"0%")`}]);
  aoa.push([null,'%',{f:`IFERROR(C${TOT_T}/$J${TOT_T},"0%")`},{f:`IFERROR(D${TOT_T}/$J${TOT_T},"0%")`},{f:`IFERROR(E${TOT_T}/$J${TOT_T},"0%")`},{f:`IFERROR(F${TOT_T}/$J${TOT_T},"0%")`},{f:`IFERROR(G${TOT_T}/$J${TOT_T},"0%")`},{f:`IFERROR(H${TOT_T}/$J${TOT_T},"0%")`},{f:`IFERROR(I${TOT_T}/$J${TOT_T},"0%")`},{f:`IFERROR(J${TOT_T}/$J${TOT_T},"0%")`},null,null,null,null,null]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{wch:18},{wch:8},{wch:14},{wch:14},{wch:16},{wch:17},{wch:12},{wch:20},{wch:10},{wch:14},{wch:2},{wch:12},{wch:12},{wch:20},{wch:14}];

  const merges = [{s:{r:0,c:0},e:{r:1,c:0}},{s:{r:0,c:1},e:{r:0,c:9}},{s:{r:0,c:11},e:{r:0,c:14}}];
  BRANCH_ORDER.forEach((_, idx) => {
    const startR = sR(idx) - 1;
    merges.push({s:{r:startR,c:0},e:{r:startR+4,c:0}});
  });
  merges.push({s:{r:TOT_S-1,c:0},e:{r:TOT_PT-1,c:0}});
  ws['!merges'] = merges;

  const sheetDate = date || new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
  XLSX.utils.book_append_sheet(wb, ws, sheetDate);
  XLSX.writeFile(wb, `Updated Sales sheet ${sheetDate}.xlsx`);
};

// ══════════════════════════════════════════════════════════════════════════
//  تصدير Garnell Revenues Sheet (الأداة الثانية)
//  التنسيق: صفوف 1-3 هيدر (merge) | من صف 3 → 18: بيانات الفروع
//  الأعمدة:
//   A=NO, B=BRANCH, C=Eat In Value, D=Eat In Chks, E=Eat In AVG,
//   F=Take Away Value, G=Take Away Chks, H=Take Away AVG,
//   I=Delivery Value, J=Delivery Chks, K=Delivery AVG,
//   L=Delivery Charge, M=Service Charge, N=VAT 14%,
//   O=Total Discount, P=Grand Total, Q=Closed Checks, R=AVG
// ══════════════════════════════════════════════════════════════════════════
const exportGarnellExcel = (dayEntries, monthLabel) => {
  const wb = XLSX.utils.book_new();

  // شيت واحد لكل يوم
  dayEntries.forEach(({ date, revenueData, spotData }) => {
    const aoa = [];

    // ── Header (3 صفوف) ──────────────────────────────────────────────
    aoa.push([`المبيعات خلال شهر ${monthLabel}`, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]);
    aoa.push([null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]);
    aoa.push([
      'NO.', 'BRANCH',
      'Eat In Sales', null, null,
      'Take Away Sales', null, null,
      'Delivery Sales', null, null,
      'Delivery Charge', 'Service Charge', 'VAT 14%',
      'Total Discount', 'Grand Total', 'Closed Checks', 'AVG.',
    ]);
    // صف sub-header لـ Value/Chks/AVG
    aoa.push([
      null, null,
      'Value', 'Chks', 'AVG.',
      'Value', 'Chks', 'AVG.',
      'Value', 'Chks', 'AVG.',
      null, null, null, null, null, null, null,
    ]);

    // ── بيانات الفروع (13 فرع، صفوف 5-17) ───────────────────────────
    GARNELL_BRANCHES.forEach((b) => {
      const spot = spotData[b.key]    || {};
      const rev  = revenueData[b.key] || {};
      const rowIdx = aoa.length + 1; // Excel 1-indexed

      const dineIn   = spot.dine_in      || 0;
      const takeAway = spot.take_away    || 0;
      const delivery = spot.delivery     || 0;
      const dineChks  = spot.dine_in_chks    || 0;
      const takeChks  = spot.take_away_chks  || 0;
      const delChks   = spot.delivery_chks   || 0;

      const avgDine = dineChks  ? dineIn   / dineChks  : null;
      const avgTake = takeChks  ? takeAway / takeChks  : null;
      const avgDel  = delChks   ? delivery / delChks   : null;

      const grandTotal   = spot.grand_total     || 0;
      const closedChecks = spot.closed_checks   || (dineChks + takeChks + delChks);
      const avgGrand     = closedChecks ? grandTotal / closedChecks : null;

      const nv = (v) => v !== null && v !== 0 ? Math.round(v * 100) / 100 : null;

      aoa.push([
        b.no,
        b.name,
        nv(dineIn),   dineChks || null,  nv(avgDine),
        nv(takeAway), takeChks || null,  nv(avgTake),
        nv(delivery), delChks  || null,  nv(avgDel),
        nv(spot.delivery_charge),
        nv(spot.service_charge),
        nv(spot.vat),
        nv(spot.discounts),
        nv(grandTotal),
        closedChecks || null,
        nv(avgGrand),
      ]);
    });

    // ── صف الإجمالي ───────────────────────────────────────────────────
    const DATA_ROW_START = 5; // Excel row أول فرع (1-indexed)
    const DATA_ROW_END   = DATA_ROW_START + GARNELL_BRANCHES.length - 1; // 17
    const sumCol = (col) => ({ f: `SUM(${col}${DATA_ROW_START}:${col}${DATA_ROW_END})` });
    const avgCol = (valCol, chkCol) => ({
      f: `IFERROR(SUM(${valCol}${DATA_ROW_START}:${valCol}${DATA_ROW_END})/SUM(${chkCol}${DATA_ROW_START}:${chkCol}${DATA_ROW_END}),"")`
    });
    const totalRow = DATA_ROW_END + 1; // صف الإجمالي في Excel

    aoa.push([
      null, 'الاجمالي',
      sumCol('C'), sumCol('D'), avgCol('C','D'),
      sumCol('F'), sumCol('G'), avgCol('F','G'),
      sumCol('I'), sumCol('J'), avgCol('I','J'),
      sumCol('L'),
      sumCol('M'),
      sumCol('N'),
      sumCol('O'),
      sumCol('P'),
      sumCol('Q'),
      avgCol('P','Q'),
    ]);

    // ── بناء الـ Worksheet ────────────────────────────────────────────
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // ── عرض الأعمدة ──────────────────────────────────────────────────
    ws['!cols'] = [
      {wch:5},  // A: NO
      {wch:16}, // B: BRANCH
      {wch:13}, // C: Eat In Value
      {wch:10}, // D: Eat In Chks
      {wch:10}, // E: Eat In AVG
      {wch:13}, // F: Take Away Value
      {wch:10}, // G: Take Away Chks
      {wch:10}, // H: Take Away AVG
      {wch:13}, // I: Delivery Value
      {wch:10}, // J: Delivery Chks
      {wch:10}, // K: Delivery AVG
      {wch:13}, // L: Delivery Charge
      {wch:14}, // M: Service Charge
      {wch:10}, // N: VAT 14%
      {wch:14}, // O: Total Discount
      {wch:13}, // P: Grand Total
      {wch:14}, // Q: Closed Checks
      {wch:10}, // R: AVG
    ];

    // ── Merges ────────────────────────────────────────────────────────
    const merges = [
      // Header title
      { s:{r:0,c:0}, e:{r:0,c:17} },
      // Row 2 (blank spacer merge)
      { s:{r:1,c:0}, e:{r:1,c:17} },
      // Header row 3 group merges
      { s:{r:2,c:0}, e:{r:3,c:0} },  // NO.
      { s:{r:2,c:1}, e:{r:3,c:1} },  // BRANCH
      { s:{r:2,c:2}, e:{r:2,c:4} },  // Eat In Sales
      { s:{r:2,c:5}, e:{r:2,c:7} },  // Take Away Sales
      { s:{r:2,c:8}, e:{r:2,c:10} }, // Delivery Sales
      { s:{r:2,c:11}, e:{r:3,c:11} },// Delivery Charge
      { s:{r:2,c:12}, e:{r:3,c:12} },// Service Charge
      { s:{r:2,c:13}, e:{r:3,c:13} },// VAT 14%
      { s:{r:2,c:14}, e:{r:3,c:14} },// Total Discount
      { s:{r:2,c:15}, e:{r:3,c:15} },// Grand Total
      { s:{r:2,c:16}, e:{r:3,c:16} },// Closed Checks
      { s:{r:2,c:17}, e:{r:3,c:17} },// AVG.
    ];
    ws['!merges'] = merges;

    // ── اسم الشيت = التاريخ ───────────────────────────────────────────
    const sheetName = date || 'Sheet1';
    XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
  });

  XLSX.writeFile(wb, `Garnell Revenues ${monthLabel}.xlsx`);
};

// ══════════════════════════════════════════════════════════════════════════
//  جدول معاينة الأداة الأولى
// ══════════════════════════════════════════════════════════════════════════
const PreviewTable = ({ revenueData, spotData }) => {
  const fmt    = (n) => (n != null && n !== 0) ? n.toLocaleString('en-EG', { maximumFractionDigits: 0 }) : '—';
  const fmtN   = (n) => (n != null && n !== 0) ? n.toString() : '—';
  const fmtPct = (num, den) => (!den || !num) ? '—' : Math.round((num / den) * 100) + '%';

  return (
    <div className="overflow-x-auto rounded-xl border border-emerald-900/40">
      <table className="w-full text-xs border-collapse" style={{ minWidth: '1300px' }}>
        <thead>
          <tr className="bg-[#0a1f14]">
            <th rowSpan={2} className="px-3 py-3 text-left text-emerald-400 font-bold border-b border-emerald-900/30 sticky left-0 bg-[#0a1f14] z-10">الفرع</th>
            <th rowSpan={2} className="px-2 py-2 text-center text-gray-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20 text-[10px]">نوع</th>
            <th colSpan={7} className="px-3 py-2 text-center text-cyan-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">Sales Channel</th>
            <th className="px-3 py-2 text-center text-emerald-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">Total</th>
            <th colSpan={4} className="px-3 py-2 text-center text-amber-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">Sales Type</th>
            <th className="px-3 py-2 text-center text-emerald-400 font-bold border-b border-emerald-900/30 border-l border-emerald-900/20">✓</th>
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
          {BRANCH_ORDER.map((key, idx) => {
            const spot    = spotData[key]    || {};
            const rev     = revenueData[key] || {};
            const hasData = Object.keys(spot).length > 0 || Object.keys(rev).length > 0;
            const totalSales = (spot.dine_in||0)+(spot.take_away||0)+(spot.delivery||0)+(spot.delivery_charge||0)+(spot.discounts||0)+(spot.service_charge||0)+(spot.vat||0);
            const revTotal   = (rev.food||0)+(rev.beverage||0)+(rev.other||0);
            const grandTotal = spot.grand_total || 0;
            const isMatch    = grandTotal > 0 && Math.abs(totalSales - grandTotal) < 1;
            const totalTrans = (spot.dine_in_chks||0)+(spot.take_away_chks||0)+(spot.delivery_chks||0);
            const bg = idx % 2 === 0 ? 'bg-[#040d08]' : 'bg-[#060f0a]';
            const avgDine = spot.dine_in_chks  ? Math.round(spot.dine_in   / spot.dine_in_chks   * 100) / 100 : null;
            const avgTake = spot.take_away_chks ? Math.round(spot.take_away / spot.take_away_chks * 100) / 100 : null;
            const avgDel  = spot.delivery_chks  ? Math.round(spot.delivery  / spot.delivery_chks  * 100) / 100 : null;

            return (
              <React.Fragment key={key}>
                <tr className={`border-b border-emerald-900/5 ${bg} hover:bg-[#0d2a1f]/40`}>
                  <td rowSpan={5} className={`px-3 py-2 font-medium sticky left-0 z-10 border-r border-emerald-900/20 align-middle ${bg} ${!hasData?'text-gray-600':'text-white'}`}>
                    {BRANCH_MAP.find(b=>b.daily_sales_key===key)?.display_name||key}
                  </td>
                  <td className="px-2 py-1.5 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">Sales</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.dine_in)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.take_away)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.delivery)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.delivery_charge)}</td>
                  <td className={`px-2 py-1.5 text-center border-l border-emerald-900/10 ${(spot.discounts||0)<0?'text-red-400/80':'text-cyan-300/80'}`}>{fmt(spot.discounts)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.service_charge)}</td>
                  <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-emerald-900/10">{fmt(spot.vat)}</td>
                  <td className="px-2 py-1.5 text-center text-emerald-300/80 font-semibold border-l border-emerald-900/20">{fmt(totalSales)}</td>
                  <td className="px-2 py-1.5 text-center text-amber-300/80 border-l border-emerald-900/20">{fmt(rev.food)}</td>
                  <td className="px-2 py-1.5 text-center text-amber-300/80 border-l border-emerald-900/10">{fmt(rev.beverage)}</td>
                  <td className="px-2 py-1.5 text-center text-amber-300/80 border-l border-emerald-900/10">{fmt(rev.other)}</td>
                  <td className="px-2 py-1.5 text-center text-amber-400/80 font-semibold border-l border-emerald-900/10">{fmt(revTotal)}</td>
                  <td rowSpan={5} className={`px-2 py-1.5 text-center font-bold border-l border-emerald-900/20 align-middle text-[11px] ${!hasData?'text-gray-600':isMatch?'text-emerald-400':'text-red-400'}`}>
                    {grandTotal>0?(<span title={isMatch?'✅ مطابق':'⚠️ فرق'}>{isMatch?'✅':'⚠️'}<br/>{fmt(grandTotal)}</span>):'—'}
                  </td>
                </tr>
                <tr className={`border-b border-emerald-900/5 ${bg} opacity-70`}>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">%</td>
                  {['dine_in','take_away','delivery','delivery_charge','discounts','service_charge','vat'].map(k=>(
                    <td key={k} className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot[k],totalSales)}</td>
                  ))}
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/20 text-[10px]">{totalSales?'100%':'—'}</td>
                  <td colSpan={4} className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/20 text-[10px]">—</td>
                </tr>
                <tr className={`border-b border-emerald-900/5 ${bg}`}>
                  <td className="px-2 py-1 text-center text-gray-500 border-l border-emerald-900/10 text-[10px]">Trans</td>
                  <td className="px-2 py-1 text-center text-purple-300/70 border-l border-emerald-900/10 text-[10px]">{fmtN(spot.dine_in_chks)}</td>
                  <td className="px-2 py-1 text-center text-purple-300/70 border-l border-emerald-900/10 text-[10px]">{fmtN(spot.take_away_chks)}</td>
                  <td className="px-2 py-1 text-center text-purple-300/70 border-l border-emerald-900/10 text-[10px]">{fmtN(spot.delivery_chks)}</td>
                  {[0,0,0,0].map((_,i)=><td key={i} className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>)}
                  <td className="px-2 py-1 text-center text-purple-400/80 font-semibold border-l border-emerald-900/20 text-[10px]">{fmtN(totalTrans)}</td>
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/20 text-[10px]">{fmtPct(rev.food,revTotal)}</td>
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/10 text-[10px]">{fmtPct(rev.beverage,revTotal)}</td>
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/10 text-[10px]">{fmtPct(rev.other,revTotal)}</td>
                  <td className="px-2 py-1 text-center text-amber-400/50 border-l border-emerald-900/10 text-[10px]">{revTotal?'100%':'—'}</td>
                </tr>
                <tr className={`border-b border-emerald-900/5 ${bg} opacity-60`}>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">%</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.dine_in_chks,totalTrans)}</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.take_away_chks,totalTrans)}</td>
                  <td className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">{fmtPct(spot.delivery_chks,totalTrans)}</td>
                  <td colSpan={5} className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/10 text-[10px]">—</td>
                  <td colSpan={4} className="px-2 py-1 text-center text-gray-600 border-l border-emerald-900/20 text-[10px]">—</td>
                </tr>
                <tr className={`border-b border-emerald-900/10 ${bg} opacity-80`}>
                  <td className="px-2 py-1 text-center text-yellow-600/70 border-l border-emerald-900/10 text-[10px] font-semibold">avg</td>
                  <td className="px-2 py-1 text-center text-yellow-400/70 border-l border-emerald-900/10 text-[10px]">{avgDine!=null?avgDine.toLocaleString('en-EG',{maximumFractionDigits:2}):'—'}</td>
                  <td className="px-2 py-1 text-center text-yellow-400/70 border-l border-emerald-900/10 text-[10px]">{avgTake!=null?avgTake.toLocaleString('en-EG',{maximumFractionDigits:2}):'—'}</td>
                  <td className="px-2 py-1 text-center text-yellow-400/70 border-l border-emerald-900/10 text-[10px]">{avgDel!=null?avgDel.toLocaleString('en-EG',{maximumFractionDigits:2}):'—'}</td>
                  <td colSpan={9} className="px-2 py-1 text-center text-gray-700 border-l border-emerald-900/10 text-[10px]">—</td>
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
//  جدول معاينة Garnell (الأداة الثانية)
// ══════════════════════════════════════════════════════════════════════════
const GarnellPreviewTable = ({ revenueData, spotData }) => {
  const fmt  = (n) => (n != null && n !== 0) ? Math.round(n).toLocaleString('en-EG') : '—';
  const fmtD = (n) => (n != null && n !== 0) ? Math.round(n * 100) / 100 : '—';

  const totals = { dineIn:0, dineChks:0, takeAway:0, takeChks:0, delivery:0, delChks:0, delCharge:0, svc:0, vat:0, disc:0, grand:0, closed:0 };

  GARNELL_BRANCHES.forEach(b => {
    const s = spotData[b.key] || {};
    totals.dineIn   += s.dine_in      || 0;
    totals.dineChks += s.dine_in_chks || 0;
    totals.takeAway += s.take_away     || 0;
    totals.takeChks += s.take_away_chks || 0;
    totals.delivery += s.delivery      || 0;
    totals.delChks  += s.delivery_chks || 0;
    totals.delCharge+= s.delivery_charge || 0;
    totals.svc      += s.service_charge  || 0;
    totals.vat      += s.vat             || 0;
    totals.disc     += s.discounts       || 0;
    totals.grand    += s.grand_total     || 0;
    totals.closed   += s.closed_checks || (s.dine_in_chks||0)+(s.take_away_chks||0)+(s.delivery_chks||0);
  });

  return (
    <div className="overflow-x-auto rounded-xl border border-blue-900/40">
      <table className="w-full text-xs border-collapse" style={{ minWidth: '1200px' }}>
        <thead>
          <tr className="bg-[#0a1020]">
            <th className="px-2 py-2 text-blue-400 font-bold border-b border-blue-900/30">NO.</th>
            <th className="px-3 py-2 text-blue-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20 text-right">BRANCH</th>
            <th colSpan={3} className="px-2 py-2 text-center text-cyan-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20">Eat In Sales</th>
            <th colSpan={3} className="px-2 py-2 text-center text-indigo-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20">Take Away Sales</th>
            <th colSpan={3} className="px-2 py-2 text-center text-violet-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20">Delivery Sales</th>
            <th className="px-2 py-2 text-center text-yellow-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20 text-[10px]">Del. Charge</th>
            <th className="px-2 py-2 text-center text-orange-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20 text-[10px]">Service</th>
            <th className="px-2 py-2 text-center text-orange-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20 text-[10px]">VAT 14%</th>
            <th className="px-2 py-2 text-center text-red-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20 text-[10px]">Discount</th>
            <th className="px-2 py-2 text-center text-emerald-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20">Grand Total</th>
            <th className="px-2 py-2 text-center text-pink-400 font-bold border-b border-blue-900/30 border-l border-blue-900/20 text-[10px]">Closed Chks</th>
            <th className="px-2 py-2 text-center text-yellow-300 font-bold border-b border-blue-900/30 border-l border-blue-900/20 text-[10px]">AVG.</th>
          </tr>
          <tr className="bg-[#070d1a] text-gray-500 text-[10px]">
            <th colSpan={2} className="border-b border-blue-900/20"/>
            {['Value','Chks','AVG'].map(h=><th key={`ei-${h}`} className="px-2 py-1 text-center border-b border-blue-900/20 border-l border-blue-900/10">{h}</th>)}
            {['Value','Chks','AVG'].map(h=><th key={`ta-${h}`} className="px-2 py-1 text-center border-b border-blue-900/20 border-l border-blue-900/10">{h}</th>)}
            {['Value','Chks','AVG'].map(h=><th key={`dl-${h}`} className="px-2 py-1 text-center border-b border-blue-900/20 border-l border-blue-900/10">{h}</th>)}
            {[...Array(7)].map((_,i)=><th key={i} className="border-b border-blue-900/20 border-l border-blue-900/10"/>)}
          </tr>
        </thead>
        <tbody>
          {GARNELL_BRANCHES.map((b, idx) => {
            const s = spotData[b.key] || {};
            const hasData = Object.keys(s).length > 0;
            const bg = idx % 2 === 0 ? 'bg-[#05080f]' : 'bg-[#070c16]';
            const dineChks = s.dine_in_chks || 0;
            const takeChks = s.take_away_chks || 0;
            const delChks  = s.delivery_chks  || 0;
            const closed   = s.closed_checks || (dineChks + takeChks + delChks);
            const avgDine  = dineChks ? (s.dine_in / dineChks).toFixed(2) : '—';
            const avgTake  = takeChks ? (s.take_away / takeChks).toFixed(2) : '—';
            const avgDel   = delChks  ? (s.delivery / delChks).toFixed(2) : '—';
            const avgGrand = closed   ? (s.grand_total / closed).toFixed(2) : '—';

            return (
              <tr key={b.key} className={`border-b border-blue-900/10 ${bg} hover:bg-blue-900/10`}>
                <td className="px-2 py-2 text-center text-gray-500 text-[10px]">{b.no}</td>
                <td className={`px-3 py-2 font-medium text-right border-l border-blue-900/20 ${!hasData?'text-gray-600':'text-white'}`}>{b.name}</td>
                <td className="px-2 py-1.5 text-center text-cyan-300/80 border-l border-blue-900/10">{fmt(s.dine_in)}</td>
                <td className="px-2 py-1.5 text-center text-cyan-300/60 border-l border-blue-900/10 text-[10px]">{dineChks||'—'}</td>
                <td className="px-2 py-1.5 text-center text-yellow-400/60 border-l border-blue-900/10 text-[10px]">{avgDine}</td>
                <td className="px-2 py-1.5 text-center text-indigo-300/80 border-l border-blue-900/10">{fmt(s.take_away)}</td>
                <td className="px-2 py-1.5 text-center text-indigo-300/60 border-l border-blue-900/10 text-[10px]">{takeChks||'—'}</td>
                <td className="px-2 py-1.5 text-center text-yellow-400/60 border-l border-blue-900/10 text-[10px]">{avgTake}</td>
                <td className="px-2 py-1.5 text-center text-violet-300/80 border-l border-blue-900/10">{fmt(s.delivery)}</td>
                <td className="px-2 py-1.5 text-center text-violet-300/60 border-l border-blue-900/10 text-[10px]">{delChks||'—'}</td>
                <td className="px-2 py-1.5 text-center text-yellow-400/60 border-l border-blue-900/10 text-[10px]">{avgDel}</td>
                <td className="px-2 py-1.5 text-center text-yellow-300/70 border-l border-blue-900/10">{fmt(s.delivery_charge)}</td>
                <td className="px-2 py-1.5 text-center text-orange-300/70 border-l border-blue-900/10">{fmt(s.service_charge)}</td>
                <td className="px-2 py-1.5 text-center text-orange-300/70 border-l border-blue-900/10">{fmt(s.vat)}</td>
                <td className={`px-2 py-1.5 text-center border-l border-blue-900/10 ${(s.discounts||0)<0?'text-red-400/80':'text-gray-500'}`}>{fmt(s.discounts)}</td>
                <td className="px-2 py-1.5 text-center text-emerald-300 font-bold border-l border-blue-900/20">{fmt(s.grand_total)}</td>
                <td className="px-2 py-1.5 text-center text-pink-300/70 border-l border-blue-900/10 text-[10px]">{closed||'—'}</td>
                <td className="px-2 py-1.5 text-center text-yellow-300/70 border-l border-blue-900/10 text-[10px]">{avgGrand}</td>
              </tr>
            );
          })}
          {/* Total Row */}
          <tr className="bg-[#0a1530] border-t-2 border-blue-700/50 font-bold">
            <td className="px-2 py-2 text-center text-blue-400 text-[10px]">—</td>
            <td className="px-3 py-2 text-right text-blue-300 border-l border-blue-900/20">الاجمالي</td>
            <td className="px-2 py-2 text-center text-cyan-300 border-l border-blue-900/10">{fmt(totals.dineIn)}</td>
            <td className="px-2 py-2 text-center text-cyan-300/70 border-l border-blue-900/10 text-[10px]">{totals.dineChks||'—'}</td>
            <td className="px-2 py-2 text-center text-yellow-400/70 border-l border-blue-900/10 text-[10px]">{totals.dineChks?Math.round(totals.dineIn/totals.dineChks):'—'}</td>
            <td className="px-2 py-2 text-center text-indigo-300 border-l border-blue-900/10">{fmt(totals.takeAway)}</td>
            <td className="px-2 py-2 text-center text-indigo-300/70 border-l border-blue-900/10 text-[10px]">{totals.takeChks||'—'}</td>
            <td className="px-2 py-2 text-center text-yellow-400/70 border-l border-blue-900/10 text-[10px]">{totals.takeChks?Math.round(totals.takeAway/totals.takeChks):'—'}</td>
            <td className="px-2 py-2 text-center text-violet-300 border-l border-blue-900/10">{fmt(totals.delivery)}</td>
            <td className="px-2 py-2 text-center text-violet-300/70 border-l border-blue-900/10 text-[10px]">{totals.delChks||'—'}</td>
            <td className="px-2 py-2 text-center text-yellow-400/70 border-l border-blue-900/10 text-[10px]">{totals.delChks?Math.round(totals.delivery/totals.delChks):'—'}</td>
            <td className="px-2 py-2 text-center text-yellow-300/80 border-l border-blue-900/10">{fmt(totals.delCharge)}</td>
            <td className="px-2 py-2 text-center text-orange-300/80 border-l border-blue-900/10">{fmt(totals.svc)}</td>
            <td className="px-2 py-2 text-center text-orange-300/80 border-l border-blue-900/10">{fmt(totals.vat)}</td>
            <td className="px-2 py-2 text-center text-red-400/80 border-l border-blue-900/10">{fmt(totals.disc)}</td>
            <td className="px-2 py-2 text-center text-emerald-300 border-l border-blue-900/20">{fmt(totals.grand)}</td>
            <td className="px-2 py-2 text-center text-pink-300/80 border-l border-blue-900/10 text-[10px]">{totals.closed||'—'}</td>
            <td className="px-2 py-2 text-center text-yellow-300/80 border-l border-blue-900/10 text-[10px]">{totals.closed?Math.round(totals.grand/totals.closed):'—'}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
};

// ══════════════════════════════════════════════════════════════════════════
//  Tab 1: Daily Sales Report Filler
// ══════════════════════════════════════════════════════════════════════════
function DailySalesTab() {
  const [revenueFile, setRevenueFile] = useState(null);
  const [spotFile,    setSpotFile]    = useState(null);
  const [revenueData, setRevenueData] = useState({});
  const [spotData,    setSpotData]    = useState({});
  const [isReady,     setIsReady]     = useState(false);
  const [isLoading,   setIsLoading]   = useState(false);
  const [reportDate,  setReportDate]  = useState('');
  const [errors,      setErrors]      = useState([]);

  const handleRevenue = useCallback(async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setRevenueFile(file.name);
    try {
      const rows = await readFileAsRows(file);
      setRevenueData(parseRevenueFile(rows, BRANCH_MAP, 'daily_sales_key'));
    } catch (err) { setErrors(prev => [...prev, `خطأ في Daily Revenue: ${err.message}`]); }
  }, []);

  const handleSpot = useCallback(async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setSpotFile(file.name);
    try {
      const rows = await readFileAsRows(file);
      setSpotData(parseSpotCheckFile(rows));
    } catch (err) { setErrors(prev => [...prev, `خطأ في Spot Check: ${err.message}`]); }
  }, []);

  const totalFood     = Object.values(revenueData).reduce((s,b) => s+(b.food||0), 0);
  const totalBeverage = Object.values(revenueData).reduce((s,b) => s+(b.beverage||0), 0);
  const totalOther    = Object.values(revenueData).reduce((s,b) => s+(b.other||0), 0);
  const totalGrand    = Object.values(spotData).reduce((s,b) => s+(b.grand_total||0), 0);
  const totalTrans    = Object.values(spotData).reduce((s,b) => s+(b.dine_in_chks||0)+(b.take_away_chks||0)+(b.delivery_chks||0), 0);
  const fmt = (n) => Math.round(n).toLocaleString('en-EG');

  return (
    <div className="space-y-6">
      {errors.length > 0 && (
        <div className="bg-red-900/20 border border-red-700/50 rounded-xl p-4 space-y-1">
          {errors.map((e, i) => <p key={i} className="text-red-400 text-sm">⚠️ {e}</p>)}
        </div>
      )}

      {!isReady ? (
        <div className="bg-gray-900 p-10 rounded-2xl border border-gray-800 shadow-2xl max-w-2xl mx-auto space-y-6">
          <div className="flex flex-col gap-2">
            <label className="text-emerald-400 font-bold text-sm">📅 تاريخ التقرير:</label>
            <input type="date" value={reportDate} onChange={e=>setReportDate(e.target.value)}
              className="bg-black/40 border border-emerald-900/40 rounded-lg px-4 py-2 text-white text-sm focus:outline-none focus:border-emerald-500" />
          </div>
          <div className="flex flex-col gap-2">
            <label className="text-emerald-400 font-bold">📁 Daily Revenue (.xls / .xlsx):</label>
            <input type="file" accept=".xls,.xlsx" onChange={handleRevenue}
              className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg file:cursor-pointer hover:file:bg-emerald-700" />
            {revenueFile && <p className="text-emerald-500 text-xs mt-1">✅ {revenueFile}</p>}
          </div>
          <div className="flex flex-col gap-2">
            <label className="text-emerald-400 font-bold">📁 Spot Check (.xls / .xlsx):</label>
            <input type="file" accept=".xls,.xlsx" onChange={handleSpot}
              className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg file:cursor-pointer hover:file:bg-emerald-700" />
            {spotFile && <p className="text-emerald-500 text-xs mt-1">✅ {spotFile}</p>}
          </div>
          <div className="bg-black/30 border border-emerald-900/20 rounded-xl p-4 text-xs text-gray-500 space-y-1">
            <p className="text-gray-400 font-semibold mb-2">خريطة الفروع:</p>
            <p>• EL Gulf = ارض الجولف (Ard El Golf) فقط</p>
            <p>• Golf Center = Golf Central (فرع مستقل)</p>
            <p>• التجمع / Tagamo3 = El Rehab | D5 = District5 | CFC = Cairo Festival City</p>
            <p>• Arkan Mall ≠ Arkan Delivery (فرعان منفصلان)</p>
          </div>
          <button onClick={() => { if (!revenueFile||!spotFile) { setErrors(['من فضلك ارفع الملفين أولاً']); return; } setErrors([]); setIsLoading(true); setTimeout(() => { setIsReady(true); setIsLoading(false); }, 500); }}
            disabled={!revenueFile||!spotFile||isLoading}
            className="w-full py-4 bg-[#10b981] text-black font-black rounded-xl hover:bg-emerald-400 transition-all disabled:opacity-40 disabled:cursor-not-allowed">
            {isLoading ? 'جاري المعالجة...' : '⚡ معالجة البيانات وعرض المعاينة'}
          </button>
        </div>
      ) : (
        <div className="space-y-6">
          <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
            {[
              { label:'Total Food',     value:fmt(totalFood),     color:'text-amber-400',   icon:'🍱' },
              { label:'Total Beverage', value:fmt(totalBeverage), color:'text-cyan-400',    icon:'🥤' },
              { label:'Total Other',    value:fmt(totalOther),    color:'text-purple-400',  icon:'📦' },
              { label:'Grand Total',    value:fmt(totalGrand),    color:'text-emerald-400', icon:'💰' },
              { label:'Total Trans',    value:fmt(totalTrans),    color:'text-pink-400',    icon:'🧾' },
            ].map(card => (
              <div key={card.label} className="bg-gray-900 rounded-xl border border-gray-800 p-4 text-center">
                <div className="text-2xl mb-1">{card.icon}</div>
                <div className={`text-xl font-bold ${card.color}`}>{card.value}</div>
                <div className="text-gray-500 text-xs mt-1">{card.label}</div>
              </div>
            ))}
          </div>
          <div className="bg-[#0d2a1f] p-4 rounded-xl border border-emerald-900 shadow-xl">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-lg font-bold text-white flex gap-2 items-center">
                <span className="w-2.5 h-2.5 bg-emerald-500 rounded-full animate-pulse"></span>
                معاينة البيانات (5 صفوف لكل فرع)
              </h2>
              <div className="text-xs text-gray-500">✅ مطابق &nbsp;|&nbsp; ⚠️ فيه فرق</div>
            </div>
            <PreviewTable revenueData={revenueData} spotData={spotData} />
          </div>
          <div className="flex gap-4 justify-center flex-wrap">
            <button onClick={() => exportFilledExcel(revenueData, spotData, reportDate)}
              className="bg-emerald-600 hover:bg-emerald-500 text-white px-10 py-3 rounded-xl font-bold transition-colors shadow-lg shadow-emerald-900/20">
              📥 تصدير Updated Sales Sheet (Excel)
            </button>
            <button onClick={() => { setIsReady(false); setErrors([]); }}
              className="text-gray-500 hover:text-emerald-400 underline py-3 px-4 transition-colors">
              ارفع ملفات جديدة
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//  Tab 2: Update Garnell Revenues Sheet
// ══════════════════════════════════════════════════════════════════════════
function GarnellRevenuesTab() {
  const [entries,     setEntries]     = useState([]); // [{date, revenueData, spotData}]
  const [curRevFile,  setCurRevFile]  = useState(null);
  const [curSpotFile, setCurSpotFile] = useState(null);
  const [curRevData,  setCurRevData]  = useState(null);
  const [curSpotData, setCurSpotData] = useState(null);
  const [curDate,     setCurDate]     = useState('');
  const [monthLabel,  setMonthLabel]  = useState('');
  const [errors,      setErrors]      = useState([]);
  const [preview,     setPreview]     = useState(null); // entry being previewed

  const handleRevenue = useCallback(async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setCurRevFile(file.name);
    try {
      const rows = await readFileAsRows(file);
      setCurRevData(parseRevenueFile(rows, GARNELL_BRANCHES, 'key'));
    } catch (err) { setErrors(prev => [...prev, `خطأ في Daily Revenue: ${err.message}`]); }
  }, []);

  const handleSpot = useCallback(async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setCurSpotFile(file.name);
    try {
      const rows = await readFileAsRows(file);
      setCurSpotData(parseGarnellSpotCheck(rows));
    } catch (err) { setErrors(prev => [...prev, `خطأ في Spot Check: ${err.message}`]); }
  }, []);

  const handleAddDay = () => {
    if (!curRevData || !curSpotData || !curDate) {
      setErrors(['من فضلك ارفع الملفين وحدد التاريخ أولاً']); return;
    }
    // Format date as DD-MM-YYYY for sheet name
    const d = new Date(curDate);
    const dateLabel = `${String(d.getDate()).padStart(2,'0')}-${String(d.getMonth()+1).padStart(2,'0')}-${d.getFullYear()}`;
    // Check duplicate
    if (entries.find(e => e.date === dateLabel)) {
      setErrors([`اليوم ${dateLabel} موجود بالفعل - احذفه الأول لو عايز تعدله`]); return;
    }
    setEntries(prev => [...prev, { date: dateLabel, revenueData: curRevData, spotData: curSpotData }].sort((a,b)=>a.date.localeCompare(b.date)));
    setErrors([]);
    setCurRevFile(null); setCurSpotFile(null); setCurRevData(null); setCurSpotData(null); setCurDate('');
    // reset file inputs
    document.querySelectorAll('.garnell-file-input').forEach(el => { el.value = ''; });
  };

  const removeEntry = (date) => {
    setEntries(prev => prev.filter(e => e.date !== date));
    if (preview?.date === date) setPreview(null);
  };

  const totalGrand = entries.reduce((sum, en) =>
    sum + Object.values(en.spotData).reduce((s,b) => s+(b.grand_total||0), 0), 0);

  return (
    <div className="space-y-6">
      {errors.length > 0 && (
        <div className="bg-red-900/20 border border-red-700/50 rounded-xl p-4 space-y-1">
          {errors.map((e, i) => <p key={i} className="text-red-400 text-sm">⚠️ {e}</p>)}
          <button onClick={() => setErrors([])} className="text-xs text-gray-500 underline mt-1">إخفاء</button>
        </div>
      )}

      {/* Upload Card */}
      <div className="bg-gray-900 p-8 rounded-2xl border border-blue-900/40 shadow-2xl">
        <h3 className="text-blue-400 font-bold text-lg mb-6 flex items-center gap-2">
          <span className="text-2xl">📅</span> إضافة يوم جديد
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="flex flex-col gap-2">
            <label className="text-blue-300 font-bold text-sm">📅 تاريخ اليوم:</label>
            <input type="date" value={curDate} onChange={e => setCurDate(e.target.value)}
              className="bg-black/40 border border-blue-900/40 rounded-lg px-4 py-2 text-white text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div className="flex flex-col gap-2">
            <label className="text-blue-300 font-bold text-sm">📁 Daily Revenue:</label>
            <input type="file" accept=".xls,.xlsx" onChange={handleRevenue}
              className="garnell-file-input text-sm text-gray-400 file:bg-blue-900 file:text-white file:border-none file:px-3 file:py-2 file:rounded-lg file:cursor-pointer hover:file:bg-blue-700" />
            {curRevFile && <p className="text-blue-400 text-xs">✅ {curRevFile}</p>}
          </div>
          <div className="flex flex-col gap-2">
            <label className="text-blue-300 font-bold text-sm">📁 Spot Check:</label>
            <input type="file" accept=".xls,.xlsx" onChange={handleSpot}
              className="garnell-file-input text-sm text-gray-400 file:bg-blue-900 file:text-white file:border-none file:px-3 file:py-2 file:rounded-lg file:cursor-pointer hover:file:bg-blue-700" />
            {curSpotFile && <p className="text-blue-400 text-xs">✅ {curSpotFile}</p>}
          </div>
        </div>
        <button onClick={handleAddDay} disabled={!curRevData||!curSpotData||!curDate}
          className="mt-6 w-full py-3 bg-blue-600 text-white font-black rounded-xl hover:bg-blue-500 transition-all disabled:opacity-40 disabled:cursor-not-allowed">
          ➕ إضافة اليوم إلى القائمة
        </button>
      </div>

      {/* Days list */}
      {entries.length > 0 && (
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-white font-bold text-lg">
              📋 الأيام المضافة ({entries.length} يوم)
              <span className="text-blue-400 text-sm mr-3">
                إجمالي المبيعات: {Math.round(totalGrand).toLocaleString('en-EG')} جنيه
              </span>
            </h3>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-3">
            {entries.map(en => {
              const dayGrand = Object.values(en.spotData).reduce((s,b) => s+(b.grand_total||0), 0);
              const isActive = preview?.date === en.date;
              return (
                <div key={en.date} className={`rounded-xl border p-3 text-center cursor-pointer transition-all ${isActive?'border-blue-500 bg-blue-900/30':'border-blue-900/30 bg-[#0a1020] hover:border-blue-700/60'}`}
                  onClick={() => setPreview(isActive ? null : en)}>
                  <div className="text-blue-300 font-bold text-sm">{en.date}</div>
                  <div className="text-emerald-400 text-xs mt-1">{Math.round(dayGrand).toLocaleString('en-EG')}</div>
                  <button onClick={e => { e.stopPropagation(); removeEntry(en.date); }}
                    className="mt-2 text-red-500/60 hover:text-red-400 text-[10px] transition-colors">
                    🗑 حذف
                  </button>
                </div>
              );
            })}
          </div>

          {/* Preview */}
          {preview && (
            <div className="bg-[#07102a] p-4 rounded-xl border border-blue-900 shadow-xl">
              <div className="flex justify-between items-center mb-4">
                <h4 className="text-lg font-bold text-white">معاينة يوم {preview.date}</h4>
                <button onClick={() => setPreview(null)} className="text-gray-500 hover:text-white text-xs">✕ إغلاق</button>
              </div>
              <GarnellPreviewTable revenueData={preview.revenueData} spotData={preview.spotData} />
            </div>
          )}

          {/* Export section */}
          <div className="bg-gray-900 p-6 rounded-2xl border border-blue-900/40">
            <h4 className="text-blue-300 font-bold mb-4">⚙️ إعدادات التصدير</h4>
            <div className="flex flex-col md:flex-row gap-4 items-end">
              <div className="flex flex-col gap-2 flex-1">
                <label className="text-blue-300 text-sm font-bold">اسم الشهر في الهيدر (مثال: 5 - 2026):</label>
                <input type="text" value={monthLabel} onChange={e => setMonthLabel(e.target.value)}
                  placeholder="5 - 2026"
                  className="bg-black/40 border border-blue-900/40 rounded-lg px-4 py-2 text-white text-sm focus:outline-none focus:border-blue-500" />
              </div>
              <button
                onClick={() => {
                  if (entries.length === 0) { setErrors(['أضف يوم واحد على الأقل أولاً']); return; }
                  exportGarnellExcel(entries, monthLabel || '5 - 2026');
                }}
                className="bg-blue-600 hover:bg-blue-500 text-white px-10 py-3 rounded-xl font-bold transition-colors shadow-lg shadow-blue-900/20 whitespace-nowrap">
                📥 تصدير Garnell Revenues ({entries.length} يوم)
              </button>
            </div>
            <p className="text-gray-600 text-xs mt-3">
              * كل يوم هيتحول لشيت منفصل في نفس الملف، بنفس تنسيق الصورة المرفقة (صفوف 3-18)
            </p>
          </div>
        </div>
      )}

      {entries.length === 0 && (
        <div className="text-center py-16 text-gray-600">
          <div className="text-5xl mb-4">📂</div>
          <p>ارفع ملفات اليوم الأول وابدأ تبني الشيت الشهري</p>
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//  الصفحة الرئيسية - Tabs
// ══════════════════════════════════════════════════════════════════════════
export default function App() {
  const [activeTab, setActiveTab] = useState('daily');

  return (
    <div className="max-w-7xl mx-auto px-4 py-8 text-right" dir="rtl">

      {/* Header */}
      <header className="text-center mb-8">
        <div className="flex items-center justify-center gap-3 mb-2">
          <span className="text-4xl">📊</span>
          <h1 className="text-3xl font-bold text-white">Garnell Sales Tools</h1>
        </div>
        <p className="text-gray-500 text-sm">أدوات تحليل وتصدير بيانات المبيعات اليومية</p>
      </header>

      {/* Tabs */}
      <div className="flex gap-2 mb-8 border-b border-gray-800">
        <button
          onClick={() => setActiveTab('daily')}
          className={`px-6 py-3 rounded-t-xl font-bold text-sm transition-all border-b-2 ${
            activeTab === 'daily'
              ? 'border-emerald-500 text-emerald-400 bg-emerald-900/20'
              : 'border-transparent text-gray-500 hover:text-gray-300'
          }`}>
          📊 Daily Sales Report Filler
        </button>
        <button
          onClick={() => setActiveTab('garnell')}
          className={`px-6 py-3 rounded-t-xl font-bold text-sm transition-all border-b-2 ${
            activeTab === 'garnell'
              ? 'border-blue-500 text-blue-400 bg-blue-900/20'
              : 'border-transparent text-gray-500 hover:text-gray-300'
          }`}>
          🏢 Update Garnell Revenues Sheet
        </button>
      </div>

      {/* Tab Content */}
      {activeTab === 'daily' ? <DailySalesTab /> : <GarnellRevenuesTab />}
    </div>
  );
}