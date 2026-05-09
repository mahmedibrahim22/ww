import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import SummaryDashboard from '../components/SummaryDashboard.jsx';
import DiscrepancyTable from '../components/DiscrepancyTable.jsx';

export default function ReconciliationEngine() {
  const [talabatData, setTalabatData] = useState([]);
  const [checkData, setCheckData] = useState([]);
  const [summary, setSummary] = useState(null);
  const [discrepancies, setDiscrepancies] = useState([]);
  const [missingInTalabat, setMissingInTalabat] = useState([]); // مفقودات ملف talabat
  const [missingInPos, setMissingInPos] = useState([]); // مفقودات ملف check details
  const [isAnalyzed, setIsAnalyzed] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  // دالة النسخ
  const copyToClipboard = (text) => {
    navigator.clipboard.writeText(text);
    alert("تم النسخ: " + text);
  };

  const exportToExcel = () => {
    if (discrepancies.length === 0) return;
    const exportData = discrepancies.map(item => ({
      "Order ID": item.id,
      "Branch (Talabat)": item.branchName,
      "Branch (Comsys)": item.comsysBranch,
      "Items": item.items,
      "POS Net": item.posSub,
      "Talabat Net": item.talabatSub,
      "Net Diff Item": item.diffSub,
      "POS Delivery": item.posDeliv,
      "Talabat Delivery": item.talabatDeliv,
      "Delivery Diff": item.diffDeliv,
      "إجمالي الفرق": item.totalDiff, // العمود الجديد المطلوب
      "Talabat Voucher (150)": item.talabatVoucher,
      "Reason Item": item.mainReasonItem,
      "POS Item Price": item.reasonItemPosPrice,
      "Talabat Item Net": item.reasonItemTalabatPrice,
      "Price Diff": item.reasonItemDiff,
      "POS VAT": item.posVAT,
      "Talabat VAT": item.talabatVAT,
      "POS Final Total": item.posTotal,
      "Talabat Final Total": item.talabatTotal,
      "Total Diff": item.diffFinal,
      "Notes": item.note
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Reconciliation_Report");
    XLSX.writeFile(workbook, `Reconciliation_Report_${new Date().toLocaleDateString()}.xlsx`);
  };

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    if (file.name.endsWith('.csv')) {
      reader.onload = (event) => {
        if (type === 'check') setCheckData(event.target.result);
        else {
          Papa.parse(event.target.result, { header: true, skipEmptyLines: true, complete: (results) => setTalabatData(results.data) });
        }
      };
      reader.readAsText(file);
    } else {
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        if (type === 'check') setCheckData(json);
        else setTalabatData(XLSX.utils.sheet_to_json(sheet));
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const handleAnalyze = () => {
    if (talabatData.length === 0 || checkData.length === 0) {
      alert("من فضلك ارفع ملف طلبات وتقرير كومسيس أولاً");
      return;
    }
    setIsLoading(true);

    setTimeout(() => {
      const foundResults = [];
      const rows = Array.isArray(checkData) ? checkData : checkData.split('\n').map(r => r.split(','));
      const outletRow = rows.find(r => String(r).includes("Outlet:"));
      const comsysBranch = outletRow ? String(outletRow).split(":")[1]?.trim() : "فرع غير معروف";
      
      const branchSearchKey = comsysBranch.replace(/Garnell/i, '').trim().toLowerCase();

      const talabatMap = new Map();
      const talabatBranchOrderIds = [];

      talabatData.forEach(item => {
        const fullRestaurantName = String(item["Restaurant name"] || "").toLowerCase();
        let rawId = item["Order ID"] || item["order_id"] || "";
        const id = String(rawId).replace(/"/g, '').trim();
        
        if (id && fullRestaurantName.includes(branchSearchKey)) {
          talabatMap.set(id, item);
          talabatBranchOrderIds.push(id);
        }
      });

      let currentCheck = null;
      let allChecksFound = [];
      let allPosVoucherIds = [];

      rows.forEach((row) => {
        const rowStr = row.join(" ").replace(/\s+/g, ' ').trim();
        
        if (rowStr.toLowerCase().includes("voucher:")) {
          const vMatch = rowStr.match(/voucher:\s*(\d+)/i);
          const vId = vMatch ? vMatch[1] : null;
          
          if (vId) {
            allPosVoucherIds.push(vId);
            currentCheck = { id: vId, items: [], itemsPrices: {}, posSub: 0, posDeliv: 0, posVAT: 0, talabatVoucher: 0 };
          }
        }

        if (currentCheck) {
          const col1 = String(row[1] || "").trim();
          const col4 = String(row[4] || "").trim();
          const col5 = String(row[5] || "").trim();
          const val8 = Math.abs(parseFloat(String(row[8] || "0").replace(/,/g, '')) || 0);

          // استخراج القيم الأساسية بناءً على مسميات الأعمدة
          if (col1 === "Sub Total") currentCheck.posSub = val8;
          if (col5 === "Delivery Charge") currentCheck.posDeliv = val8;
          if (col5 === "VAT") currentCheck.posVAT = val8;

          // استخراج الفتحور (كود 150)
          if (col4 === "150.0" || col4 === "150") {
            currentCheck.talabatVoucher = val8;
          }

          // استخراج الأصناف والكميات
          const rowCleaned = row.map(c => String(c).trim());
          const qnty = parseFloat(rowCleaned[6]);
          const price = parseFloat(rowCleaned[7]?.replace(/,/g, ''));
          const itemName = rowCleaned[5];

          if (!isNaN(qnty) && itemName && itemName !== "" && col1 !== "Sub Total" && col5 !== "VAT" && col5 !== "Delivery Charge") {
            currentCheck.items.push(`${itemName} (x${qnty})`);
            currentCheck.itemsPrices[itemName] = price;
          }

          if (/closing\s*time/i.test(rowStr)) {
            allChecksFound.push(currentCheck);
            currentCheck = null;
          }
        }
      });

      const missingInT = allPosVoucherIds.filter(id => !talabatMap.has(id));
      const posVoucherSet = new Set(allPosVoucherIds);
      const missingInP = talabatBranchOrderIds.filter(id => !posVoucherSet.has(id));

      setMissingInTalabat([...new Set(missingInT)]);
      setMissingInPos([...new Set(missingInP)]);

      allChecksFound.forEach(check => {
        const tItem = talabatMap.get(check.id);
        if (tItem) {
          const round = (n) => Math.round(n * 100) / 100;
          const parseSafe = (val) => parseFloat(String(val || 0).replace(/,/g, '').replace(/"/g, '')) || 0;

          const tSubRaw = parseSafe(tItem["Subtotal"]);
          const tDelivRaw = parseSafe(tItem["Delivery Fee"]);
          const tSubComp = round(tSubRaw / 1.14);
          const tDelivComp = round(tDelivRaw / 1.14);

          const diffSub = round(Math.abs(check.posSub - tSubComp));
          const diffDeliv = round(Math.abs(check.posDeliv - tDelivComp));

          let mainReasonItem = "مطابق";
          let resPosPrice = 0;
          let resTalabatPriceNet = 0;
          let resDiff = 0;

          if (diffSub > 0.1) {
            const orderItemsRaw = tItem["Order Items"] || "";
            for (let itemName in check.itemsPrices) {
              const posPrice = check.itemsPrices[itemName];
              const regex = new RegExp(itemName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ".*?(\\d+\\.?\\d*)", "i");
              const match = orderItemsRaw.match(regex);
              
              if (match) {
                const talabatGross = parseFloat(match[1]);
                const talabatNet = round(talabatGross / 1.14);
                if (Math.abs(posPrice - talabatNet) > 0.5) {
                  mainReasonItem = itemName;
                  resPosPrice = posPrice;
                  resTalabatPriceNet = talabatNet;
                  resDiff = round(Math.abs(posPrice - talabatNet));
                  break;
                }
              }
            }
            if (mainReasonItem === "مطابق") mainReasonItem = "تحقق يدوي";
          }

          // الإجمالي النهائي في كومسيس = (الصافي + التوصيل + الضريبة) - الفتحور المستخرج
          const posFinalTotal = round(check.posSub + check.posDeliv + check.posVAT - check.talabatVoucher);
          const talabatFinalTotal = round(tSubRaw + tDelivRaw);

          foundResults.push({
            id: check.id,
            branchName: tItem["Restaurant name"],
            comsysBranch: comsysBranch,
            items: check.items.join(" | "),
            posSub: check.posSub,
            posDeliv: check.posDeliv,
            posVAT: check.posVAT,
            talabatVoucher: check.talabatVoucher,
            posTotal: posFinalTotal,
            talabatSub: tSubComp,
            talabatDeliv: tDelivComp,
            talabatVAT: parseSafe(tItem["Tax Amount"]),
            talabatTotal: talabatFinalTotal,
            diffSub: diffSub,
            diffDeliv: diffDeliv,
            totalDiff: round(diffSub + diffDeliv), // حساب إجمالي الفرق المطلوب
            diffFinal: round(Math.abs(posFinalTotal - talabatFinalTotal)),
            mainReasonItem: mainReasonItem,
            reasonItemPosPrice: resPosPrice,
            reasonItemTalabatPrice: resTalabatPriceNet,
            reasonItemDiff: resDiff,
            note: check.talabatVoucher > 0 ? `خصم فوتشر: ${check.talabatVoucher}` : `فرق أصناف: ${diffSub}`
          });
        }
      });

      setSummary({
        posTotal: foundResults.reduce((acc, c) => acc + c.posTotal, 0).toFixed(2),
        settlementTotal: foundResults.reduce((acc, c) => acc + c.talabatTotal, 0).toFixed(2),
        diffTotal: foundResults.reduce((acc, c) => acc + c.diffFinal, 0).toFixed(2),
        checksCount: foundResults.length,
        missingCount: missingInT.length + missingInP.length 
      });

      setDiscrepancies(foundResults);
      setIsAnalyzed(true);
      setIsLoading(false);
    }, 800);
  };

  return (
    <div className="max-w-7xl mx-auto px-4 py-8 space-y-8 bg-[#020c08] min-h-screen text-right" dir="rtl">
      <header className="text-center">
        <h1 className="text-3xl font-bold text-[#10b981]">محرك المطابقة المطور - V5.3</h1>
        <p className="text-gray-400 mt-2">نظام حصر مفقودات وتحليل الفواتير والـ Voucher</p>
      </header>

      {!isAnalyzed ? (
        <div className="bg-gray-900 p-10 rounded-2xl border border-gray-800 shadow-2xl max-w-2xl mx-auto">
            <div className="space-y-6">
              <div className="flex flex-col gap-2">
                <label className="text-emerald-400 font-bold">ملف كومسيس (Check Details):</label>
                <input type="file" onChange={(e) => handleFileUpload(e, 'check')} className="text-sm text-gray-500 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg" />
              </div>
              <div className="flex flex-col gap-2">
                <label className="text-emerald-400 font-bold">ملف طلبات (Talabat):</label>
                <input type="file" onChange={(e) => handleFileUpload(e, 'talabat')} className="text-sm text-gray-500 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg" />
              </div>
              <button onClick={handleAnalyze} className="w-full py-4 bg-[#10b981] text-black font-black rounded-xl hover:bg-emerald-400 transition-all">
                {isLoading ? "جاري التحليل..." : "تحليل البيانات وحصر المفقودات"}
              </button>
            </div>
        </div>
      ) : (
        <div className="space-y-6">
            <SummaryDashboard data={summary} />
            <div className="bg-[#0d2a1f] p-4 rounded-xl border border-emerald-900">
              <div className="flex justify-between items-center mb-4">
                 <h2 className="text-xl font-bold text-white">تفاصيل الفروقات والأصناف والـ Voucher</h2>
                 <button onClick={exportToExcel} className="bg-emerald-600 hover:bg-emerald-500 text-white px-6 py-2 rounded-lg font-bold">تصدير Excel (EN)</button>
              </div>
              <DiscrepancyTable data={discrepancies} />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
               <div className="bg-gray-900 p-6 rounded-xl border border-red-900 shadow-lg">
                 <div className="flex justify-between items-center mb-4">
                   <h3 className="text-red-400 font-bold">مفقودات ملف talabat ({missingInTalabat.length})</h3>
                   <button onClick={() => copyToClipboard(missingInTalabat.join('\n'))} className="bg-red-800 hover:bg-red-700 text-white text-xs px-3 py-1 rounded">نسخ الكل</button>
                 </div>
                 <div className="grid grid-cols-3 gap-2 max-h-60 overflow-y-auto p-2 bg-black rounded-lg">
                   {missingInTalabat.length > 0 ? missingInTalabat.map(id => (
                     <div key={id} onClick={() => copyToClipboard(id)} className="bg-gray-800 p-2 rounded text-center text-xs text-red-200 border border-red-900/50 cursor-pointer hover:bg-red-900 transition-colors">{id}</div>
                   )) : <p className="text-gray-600 text-xs col-span-3 text-center">لا يوجد مفقودات</p>}
                 </div>
               </div>

               <div className="bg-gray-900 p-6 rounded-xl border border-orange-900 shadow-lg">
                 <div className="flex justify-between items-center mb-4">
                   <h3 className="text-orange-400 font-bold">مفقودات ملف check details ({missingInPos.length})</h3>
                   <button onClick={() => copyToClipboard(missingInPos.join('\n'))} className="bg-orange-800 hover:bg-orange-700 text-white text-xs px-3 py-1 rounded">نسخ الكل</button>
                 </div>
                 <div className="grid grid-cols-3 gap-2 max-h-60 overflow-y-auto p-2 bg-black rounded-lg">
                   {missingInPos.length > 0 ? missingInPos.map(id => (
                     <div key={id} onClick={() => copyToClipboard(id)} className="bg-gray-800 p-2 rounded text-center text-xs text-orange-200 border border-orange-900/50 cursor-pointer hover:bg-orange-900 transition-colors">{id}</div>
                   )) : <p className="text-gray-600 text-xs col-span-3 text-center">لا يوجد مفقودات</p>}
                 </div>
               </div>
            </div>
            <button onClick={() => setIsAnalyzed(false)} className="block mx-auto text-gray-500 underline">إعادة ضبط</button>
        </div>
      )}
    </div>
  );
}