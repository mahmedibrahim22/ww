import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import SummaryDashboard from '../components/SummaryDashboard.jsx';
import DiscrepancyTable from '../components/DiscrepancyTable.jsx';
import ExportButton from '../components/ExportButton.jsx';

export default function ReconciliationEngine() {
  const [talabatData, setTalabatData] = useState([]);
  const [checkData, setCheckData] = useState([]);
  const [summary, setSummary] = useState(null);
  const [discrepancies, setDiscrepancies] = useState([]);
  const [missingVouchers, setMissingVouchers] = useState([]);
  const [missingFromComsys, setMissingFromComsys] = useState([]);
  const [isAnalyzed, setIsAnalyzed] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    
    if (file.name.endsWith('.csv')) {
      reader.onload = (event) => {
        if (type === 'check') setCheckData(event.target.result);
        else {
          Papa.parse(event.target.result, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => setTalabatData(results.data)
          });
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
      const missing = [];
      const notInComsys = [];
      let totalAbsoluteDiffAccumulator = 0; 

      const rows = Array.isArray(checkData) ? checkData : checkData.split('\n').map(r => r.split(','));
      const outletRow = rows.find(r => String(r).includes("Outlet:"));
      const comsysBranch = outletRow ? String(outletRow).split(":")[1]?.trim() : "فرع غير معروف";

      const talabatMap = new Map();
      talabatData.forEach(item => {
        const status = (item["Order status"] || item["order_status"] || "").toLowerCase();
        const payment = (item["Payment method"] || item["payment_method"] || "").toLowerCase();
        const id = String(item["Order ID"] || item["order_id"] || "").replace(/"/g, '').trim();
        if (id && status !== "cancelled" && payment.includes("credit")) {
          talabatMap.set(id, item);
        }
      });

      let currentCheck = null;
      const matchedIds = new Set();

      rows.forEach((row) => {
        const rowStr = row.join(" ");

        // 1. اكتشاف بداية الشيك (Voucher ID)
        if (rowStr.includes("Voucher:")) {
          const vId = rowStr.match(/Voucher:\s*(\d+)/)?.[1];
          currentCheck = { 
            id: vId, 
            items: [], 
            posSub: 0, 
            posDeliv: 0, 
            posTotalAdjusted: 0, 
            talabatVoucher: 0, 
            talabatService: 0 
          };
        }

        if (currentCheck) {
          // جلب رقم "الكود" والقيمة اللي قدامه (في الغالب العمود 4 هو الكود و 8 هو المبلغ)
          const itemID = String(row[4] || "").trim();
          const itemValue = parseFloat(row[8]) || 0;

          // 2. سحب البيانات بناءً على الأكواد اللي حددتها
          if (itemID === "132") {
            currentCheck.posTotalAdjusted = itemValue; // إجمالي كومسيس المصفى
          } else if (itemID === "150") {
            currentCheck.talabatVoucher = Math.abs(itemValue); // طلبات فواتشر
          } else if (itemID === "2104") {
            currentCheck.talabatService = Math.abs(itemValue); // طلبات سيرفيس
          }

          // سحب الصافي والدليفري العاديين للمقارنة
          if (rowStr.includes("Sub Total")) currentCheck.posSub = itemValue;
          if (rowStr.includes("Delivery Charge")) currentCheck.posDeliv = itemValue;
          
          // تجميع محتويات الشيك (الأصناف العادية)
          if (row[5] && !isNaN(row[6]) && !["132", "150", "2104"].includes(itemID)) {
            currentCheck.items.push(`${row[5]} (x${row[6]})`);
          }
        }

        // 3. نهاية الشيك والمطابقة مع شيت طلبات
        if (rowStr.includes("Balance Due") && currentCheck) {
          const tItem = talabatMap.get(currentCheck.id);
          if (tItem) {
            matchedIds.add(currentCheck.id);
            const tSubRaw = parseFloat(String(tItem["Subtotal"] || 0).replace(/,/g, ''));
            const tDelivRaw = parseFloat(String(tItem["Delivery Fee"] || 0).replace(/,/g, ''));
            const round = (num) => Math.round(num * 100) / 100;

            const tSubCompared = round(tSubRaw / 1.14);
            const tDelivCompared = round(tDelivRaw / 1.14);

            const dSub = Math.abs(round(currentCheck.posSub) - tSubCompared);
            const dDeliv = Math.abs(round(currentCheck.posDeliv) - tDelivCompared);
            const finalCheckDiff = Math.abs(dSub + dDeliv);

            foundResults.push({
              id: currentCheck.id,
              branchName: tItem["Restaurant name"] || "طلب مجهول",
              comsysBranch: comsysBranch,
              items: currentCheck.items.join(" | "),
              posSub: currentCheck.posSub,
              talabatSub: tSubCompared,
              diffSub: dSub, 
              posDeliv: currentCheck.posDeliv,
              talabatDeliv: tDelivCompared,
              diffDeliv: dDeliv,
              posTotal: currentCheck.posTotalAdjusted, // الرقم اللي قدام 132
              talabatTotal: round(tSubRaw + tDelivRaw),
              diffFinal: finalCheckDiff,
              talabatVoucher: currentCheck.talabatVoucher, // الرقم اللي قدام 150
              talabatService: currentCheck.talabatService, // الرقم اللي قدام 2104
              note: `كود 132: ${currentCheck.posTotalAdjusted} | كود 150: ${currentCheck.talabatVoucher}`
            });
            totalAbsoluteDiffAccumulator += finalCheckDiff;
          } else {
            missing.push({ id: currentCheck.id, posTotal: currentCheck.posTotalAdjusted });
          }
          currentCheck = null;
        }
      });

      // الفواتير الموجودة في طلبات ومش في كومسيس
      talabatMap.forEach((item, id) => {
        if (!matchedIds.has(id)) {
          notInComsys.push({ id, talabatTotal: parseFloat(item["Subtotal"] || 0) + parseFloat(item["Delivery Fee"] || 0) });
        }
      });

      setSummary({
        posTotal: foundResults.reduce((acc, curr) => acc + curr.posTotal, 0).toFixed(2),
        settlementTotal: foundResults.reduce((acc, curr) => acc + curr.talabatTotal, 0).toFixed(2),
        diffTotal: totalAbsoluteDiffAccumulator.toFixed(2), 
        checksCount: foundResults.length,
        missingCount: missing.length + notInComsys.length
      });
      setDiscrepancies(foundResults);
      setMissingVouchers(missing);
      setMissingFromComsys(notInComsys);
      setIsAnalyzed(true);
      setIsLoading(false);
    }, 1000);
  };

  return (
    <div className="max-w-7xl mx-auto px-4 py-8 space-y-8 bg-[#020c08] min-h-screen text-[#d1fae5] text-right" dir="rtl">
      <header className="text-center space-y-2">
        <h1 className="text-3xl font-bold text-[#10b981]">نظام مطابقة جرنيال - استخراج الأكواد (132, 150, 2104)</h1>
        <p className="text-[#648b7a]">يتم سحب الأرقام مباشرة من الخانة المقابلة للكود في تقرير كومسيس</p>
      </header>

      {!isAnalyzed ? (
        <div className="bg-gray-900 border border-gray-800 rounded-2xl p-8 max-w-4xl mx-auto shadow-2xl">
          <div className="grid md:grid-cols-2 gap-6 text-center">
            <div className="p-6 border border-dashed border-emerald-900 rounded-xl bg-[#0d2a1f]">
              <span className="block mb-4 text-emerald-400 font-bold">تقرير كومسيس (Excel/CSV)</span>
              <input type="file" onChange={(e) => handleFileUpload(e, 'check')} className="text-xs text-gray-400" />
            </div>
            <div className="p-6 border border-dashed border-emerald-900 rounded-xl bg-[#0d2a1f]">
              <span className="block mb-4 text-emerald-400 font-bold">شيت طلبات (Excel/CSV)</span>
              <input type="file" onChange={(e) => handleFileUpload(e, 'talabat')} className="text-xs text-gray-400" />
            </div>
          </div>
          <button onClick={handleAnalyze} className="w-full mt-8 py-4 bg-[#10b981] text-black font-bold rounded-xl hover:bg-emerald-400 transition-all shadow-lg">
            {isLoading ? "جاري قراءة الأكواد 132 و 150..." : "ابدأ التحليل والمطابقة"}
          </button>
        </div>
      ) : (
        <div className="space-y-8 animate-fade-in">
          <SummaryDashboard data={summary} />
          <div className="bg-[#0d2a1f] p-6 rounded-xl border border-emerald-900 shadow-xl">
            <h2 className="text-[#10b981] font-bold mb-4">📊 نتائج تحليل الأكواد المحاسبية</h2>
            <DiscrepancyTable data={discrepancies} />
            <div className="mt-6 flex justify-end gap-2"><ExportButton data={discrepancies} /></div>
          </div>
          <button onClick={() => setIsAnalyzed(false)} className="text-gray-500 hover:text-white underline text-sm block mx-auto py-4">تصفير ورفع ملفات جديدة</button>
        </div>
      )}
    </div>
  );
}