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
  const [missingInTalabat, setMissingInTalabat] = useState([]); 
  const [missingInPos, setMissingInPos] = useState([]); 
  const [isAnalyzed, setIsAnalyzed] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [paymentFilter, setPaymentFilter] = useState('Both');

  // خريطة ربط محسنة جداً لضمان تغطية كافة الفروع ومسمياتها المختلفة
  const branchMapping = {
    "golf": "garnell sushi & poke, ard el golf",
    "almaza": "garnell sushi & poke, ard el golf",
    "ard el golf": "garnell sushi & poke, ard el golf",
    "maadi": "garnell sushi & poke, maadi",
    "madii": "garnell sushi & poke, maadi", // معالجة خطأ إملائي وارد في الصور
    "zayed": "garnell sushi & poke, sheikh zayed",
    "sheikh zayed": "garnell sushi & poke, sheikh zayed",
    "arkan": "garnell sushi & poke, sheikh zayed",
    "zamalek": "garnell sushi & poke, zamalek - 26 july",
    "tagammoa": "garnell sushi & poke, tagammoa 5 - north investors",
    "square": "garnell sushi & poke, tagammoa 5 - north investors",
    "shorouk": "garnell sushi & poke, el shorouk - 5th district",
    "dokki": "garnell sushi & poke, dokki"
  };

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
      "إجمالي الفرق": item.totalDiff,
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
    XLSX.writeFile(workbook, `Garnell_Recon_${discrepancies[0]?.comsysBranch}_${new Date().toLocaleDateString()}.xlsx`);
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
      
      // 1. استخراج اسم الفرع بذكاء
      const outletRow = rows.find(r => String(r).toLowerCase().includes("outlet:"));
      const comsysBranchFull = outletRow ? String(outletRow).split(":")[1]?.trim() : "Unknown";
      const comsysLower = comsysBranchFull.toLowerCase();

      // تحديد الكلمة المفتاحية للربط
      let targetTalabatMapping = "";
      let activeKeyword = "";
      for (const key in branchMapping) {
        if (comsysLower.includes(key)) {
          targetTalabatMapping = branchMapping[key];
          activeKeyword = key;
          break;
        }
      }

      const talabatMap = new Map();
      const talabatBranchOrderIds = [];

      talabatData.forEach(item => {
        const restaurantName = String(item["Restaurant name"] || "").toLowerCase();
        const id = String(item["Order ID"] || item["order_id"] || "").replace(/"/g, '').trim();
        const status = String(item["Order status"] || "").toLowerCase();
        const payType = String(item["Payment type"] || "").toLowerCase();
        
        // تجاهل الملغيات والفلترة حسب نوع الدفع
        if (status === 'cancelled' || status === 'rejected') return;
        if (paymentFilter === 'Cash' && payType !== 'cash') return;
        if (paymentFilter === 'Credit' && (payType !== 'online' && payType !== 'credit card')) return;

        // مطابقة الفرع: ابحث عن اسم المطعم في طلبات باستخدام الكلمة المفتاحية المستخرجة
        const isMatch = (targetTalabatMapping && restaurantName.includes(activeKeyword)) || 
                        restaurantName.includes(comsysLower.replace(/garnell|sushi|poke|&/g, '').trim());

        if (id && isMatch) {
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
          const vId = rowStr.match(/voucher:\s*(\d+)/i)?.[1];
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

          if (col1 === "Sub Total") currentCheck.posSub = val8;
          if (col5 === "Delivery Charge") currentCheck.posDeliv = val8;
          if (col5 === "VAT") currentCheck.posVAT = val8;
          if (col4 === "150" || col4 === "150.0" || (rowStr.toLowerCase().includes("voucher") && val8 > 0)) {
            currentCheck.talabatVoucher = val8;
          }

          const qnty = parseFloat(row[6]);
          const itemName = String(row[5] || "").trim();
          if (!isNaN(qnty) && itemName && !["VAT", "Sub Total", "Delivery Charge"].includes(itemName)) {
            currentCheck.items.push(`${itemName} (x${qnty})`);
            currentCheck.itemsPrices[itemName] = parseFloat(String(row[7]).replace(/,/g, '')) || 0;
          }

          if (/closing\s*time/i.test(rowStr)) {
            allChecksFound.push(currentCheck);
            currentCheck = null;
          }
        }
      });

      const missingInT = allPosVoucherIds.filter(id => !talabatMap.has(id));
      const missingInP = talabatBranchOrderIds.filter(id => !new Set(allPosVoucherIds).has(id));

      allChecksFound.forEach(check => {
        const tItem = talabatMap.get(check.id);
        if (tItem) {
          const round = (n) => Math.round(n * 100) / 100;
          const tSubRaw = parseFloat(String(tItem["Subtotal"] || 0).replace(/,/g, '')) || 0;
          const tDelivRaw = parseFloat(String(tItem["Delivery Fee"] || 0).replace(/,/g, '')) || 0;
          
          const tSubComp = round(tSubRaw / 1.14);
          const tDelivComp = round(tDelivRaw / 1.14);

          const posTotal = round(check.posSub + check.posDeliv + check.posVAT - check.talabatVoucher);
          const talabatTotal = round(tSubRaw + tDelivRaw);

          foundResults.push({
            id: check.id,
            branchName: tItem["Restaurant name"],
            comsysBranch: comsysBranchFull,
            items: check.items.join(" | "),
            posSub: check.posSub,
            posDeliv: check.posDeliv,
            posVAT: check.posVAT,
            talabatVoucher: check.talabatVoucher,
            posTotal: posTotal,
            talabatSub: tSubComp,
            talabatDeliv: tDelivComp,
            talabatVAT: parseFloat(tItem["Tax Amount"]) || 0,
            talabatTotal: talabatTotal,
            diffSub: round(Math.abs(check.posSub - tSubComp)),
            diffDeliv: round(Math.abs(check.posDeliv - tDelivComp)),
            totalDiff: round(Math.abs(check.posSub - tSubComp) + Math.abs(check.posDeliv - tDelivComp)),
            diffFinal: round(Math.abs(posTotal - talabatTotal)),
            mainReasonItem: check.talabatVoucher > 0 ? "يوجد فوتشر" : "مطابق أصناف",
            reasonItemPosPrice: 0,
            reasonItemTalabatPrice: 0,
            reasonItemDiff: 0,
            note: check.talabatVoucher > 0 ? `فوتشر: ${check.talabatVoucher}` : "مراجعة يدوية"
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
      setMissingInTalabat([...new Set(missingInT)]);
      setMissingInPos([...new Set(missingInP)]);
      setIsAnalyzed(true);
      setIsLoading(false);
    }, 800);
  };

  return (
    <div className="max-w-7xl mx-auto px-4 py-8 space-y-8 bg-[#020c08] min-h-screen text-right" dir="rtl">
      <header className="text-center">
        <h1 className="text-3xl font-bold text-[#10b981]">Garnell Reconciliation - V6.0</h1>
        <p className="text-gray-400 mt-2">نظام المطابقة الموحد لجميع الفروع (دعم المعادي وأركان والغولف)</p>
      </header>

      {!isAnalyzed ? (
        <div className="bg-gray-900 p-10 rounded-2xl border border-gray-800 shadow-2xl max-w-2xl mx-auto">
          <div className="space-y-6">
            <div className="flex flex-col gap-2">
              <label className="text-emerald-400 font-bold">تقرير كومسيس (Check Details):</label>
              <input type="file" onChange={(e) => handleFileUpload(e, 'check')} className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg" />
            </div>
            <div className="flex flex-col gap-2">
              <label className="text-emerald-400 font-bold">تقرير طلبات (Excel/CSV):</label>
              <input type="file" onChange={(e) => handleFileUpload(e, 'talabat')} className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg" />
            </div>

            <div className="bg-black/40 p-4 rounded-xl border border-emerald-900/30">
              <p className="text-white mb-3 font-bold text-sm">تصفية نوع الدفع:</p>
              <div className="flex gap-6">
                {['Both', 'Cash', 'Credit'].map(type => (
                  <label key={type} className="flex items-center gap-2 text-gray-300 cursor-pointer">
                    <input type="radio" name="payFilter" value={type} checked={paymentFilter === type} onChange={(e) => setPaymentFilter(e.target.value)} className="accent-emerald-500" /> 
                    {type === 'Both' ? 'الكل' : type === 'Credit' ? 'فيزا' : 'كاش'}
                  </label>
                ))}
              </div>
            </div>

            <button onClick={handleAnalyze} className="w-full py-4 bg-[#10b981] text-black font-black rounded-xl hover:bg-emerald-400 transition-all shadow-lg shadow-emerald-900/20">
              {isLoading ? "جاري تحليل البيانات..." : "بدء عملية المطابقة الذكية"}
            </button>
          </div>
        </div>
      ) : (
        <div className="space-y-6 animate-in fade-in duration-500">
          <SummaryDashboard data={summary} />
          <div className="bg-[#0d2a1f] p-4 rounded-xl border border-emerald-900 shadow-xl">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-xl font-bold text-white flex gap-2 items-center">
                <span className="w-3 h-3 bg-emerald-500 rounded-full animate-pulse"></span>
                نتائج فرع: {discrepancies[0]?.comsysBranch}
              </h2>
              <button onClick={exportToExcel} className="bg-emerald-600 hover:bg-emerald-500 text-white px-8 py-2 rounded-lg font-bold transition-colors">تصدير التقرير النهائي (Excel)</button>
            </div>
            <DiscrepancyTable data={discrepancies} />
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="bg-gray-900 p-6 rounded-xl border border-red-900/50">
              <h3 className="text-red-400 font-bold mb-4 flex justify-between">فواتير كومسيس غير موجودة في طلبات <span>({missingInTalabat.length})</span></h3>
              <div className="grid grid-cols-3 gap-2 max-h-48 overflow-y-auto p-2 bg-black/50 rounded-lg border border-gray-800">
                {missingInTalabat.map(id => <div key={id} onClick={() => copyToClipboard(id)} className="bg-gray-800 p-2 text-center text-xs text-red-200 rounded cursor-pointer hover:bg-red-900/30 transition-colors">{id}</div>)}
              </div>
            </div>
            <div className="bg-gray-900 p-6 rounded-xl border border-orange-900/50">
              <h3 className="text-orange-400 font-bold mb-4 flex justify-between">فواتير طلبات غير موجودة في كومسيس <span>({missingInPos.length})</span></h3>
              <div className="grid grid-cols-3 gap-2 max-h-48 overflow-y-auto p-2 bg-black/50 rounded-lg border border-gray-800">
                {missingInPos.map(id => <div key={id} onClick={() => copyToClipboard(id)} className="bg-gray-800 p-2 text-center text-xs text-orange-200 rounded cursor-pointer hover:bg-orange-900/30 transition-colors">{id}</div>)}
              </div>
            </div>
          </div>
          <button onClick={() => setIsAnalyzed(false)} className="block mx-auto text-gray-500 hover:text-emerald-400 underline py-4 transition-colors">العودة لرفع ملفات جديدة</button>
        </div>
      )}
    </div>
  );
}