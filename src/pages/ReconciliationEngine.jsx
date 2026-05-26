import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import SummaryDashboard from '../components/SummaryDashboard.jsx';
import DiscrepancyTable from '../components/DiscrepancyTable.jsx';

export default function ReconciliationEngine() {
  const [talabatData, setTalabatData]             = useState([]);
  const [checkData, setCheckData]                 = useState([]);
  const [summary, setSummary]                     = useState(null);
  const [discrepancies, setDiscrepancies]         = useState([]);
  const [missingInTalabat, setMissingInTalabat]   = useState([]);
  const [missingInPos, setMissingInPos]           = useState([]);
  const [isAnalyzed, setIsAnalyzed]               = useState(false);
  const [isLoading, setIsLoading]                 = useState(false);
  const [paymentFilter, setPaymentFilter]         = useState('Both');

  // ══════════════════════════════════════════════════════════════════════════
  //  خريطة الفروع: كلمة مفتاحية من Outlet في كومسيس → keyword للبحث في طلبات
  //
  //  أسماء الفروع الحقيقية في طلبات (من الملف الفعلي):
  //  - "garnell sushi & poke, tagammoa 5 - north investors"
  //  - "garnell sushi & poke, ard el golf"
  //  - "garnell sushi & poke, el sheikh zayed - zayed 2000"
  //  - "garnell sushi & poke, el shorouk - 5th district"
  //  - "garnell sushi & poke, maadi old - el nahda street"
  //  - "garnell sushi & poke, zamalek - 26 july"
  // ══════════════════════════════════════════════════════════════════════════
  const BRANCH_MAP = [
    { keywords: ["ard el golf", "almaza", "golf"],  talabat: "ard el golf"   },
    { keywords: ["maadi", "madii", "ma'adi"],        talabat: "maadi"         },
    { keywords: ["arkan", "sheikh zayed", "zayed"],  talabat: "sheikh zayed"  },
    { keywords: ["zamalek"],                         talabat: "zamalek"       },
    { keywords: ["tagammoa", "square"],              talabat: "tagammoa"      },
    { keywords: ["shorouk"],                         talabat: "shorouk"       },
    { keywords: ["dokki"],                           talabat: "dokki"         },
  ];

  // ── مطابقة اسم الفرع من كومسيس بالخريطة ─────────────────────────────────
  const resolveTalabatBranch = (comsysOutlet) => {
    const lower = comsysOutlet.toLowerCase();
    for (const entry of BRANCH_MAP) {
      if (entry.keywords.some(kw => lower.includes(kw))) {
        return entry.talabat;
      }
    }
    return null;
  };

  // ── هل الطلب ينتمي للفرع المطلوب؟ ───────────────────────────────────────
  const branchMatches = (restaurantName, talabatKeyword) => {
    if (!talabatKeyword) return false;
    return restaurantName.toLowerCase().includes(talabatKeyword);
  };

  // ── هل نوع الدفع يناسب الفلتر المختار؟ ──────────────────────────────────
  // Payment method: "Credit Card" | "CASH" | "Talabat Credit" | nan
  // Payment type:   "Online" | "Cash"
  const paymentMatches = (item) => {
    if (paymentFilter === 'Both') return true;
    const method = String(item['Payment method'] || '').toLowerCase().trim();
    const type   = String(item['Payment type']   || '').toLowerCase().trim();
    if (paymentFilter === 'Credit') {
      return method === 'credit card' || method === 'talabat credit' || type === 'online';
    }
    if (paymentFilter === 'Cash') {
      return method === 'cash' || type === 'cash';
    }
    return true;
  };

  // ══════════════════════════════════════════════════════════════════════════
  //  منطق Reason Item
  //
  //  القاعدة:
  //  - لو posSub == talabatSub (فرق < 0.1) → "✅ مطابق كصنف"
  //    والـ 3 خلايا (POS Item Price, Talabat Item Net, Price Diff) = "" فاضية
  //
  //  - لو مش مطابق أو في صنف مخالف → "⚠️ تحقق يدويا"
  //    والـ 3 خلايا = "" فاضية (المحاسب يملأها يدوياً)
  //
  //  ملاحظة: الفوتشر (كود 150) يظل محفوظاً في بيانات الشيك للحسابات
  //  لكن لا يؤثر على منطق Reason Item - يطبق نفس القاعدة العادية
  // ══════════════════════════════════════════════════════════════════════════
  const resolveReason = (check, tItem, round) => {
    const tSubRaw  = parseFloat(String(tItem['Subtotal'] || 0).replace(/,/g, '')) || 0;
    const tSubComp = round(tSubRaw / 1.14);
    const diffSub  = round(Math.abs(check.posSub - tSubComp));

    // ── مطابق: posSub == talabatSub (فرق < 0.1) ─────────────────────────
    if (diffSub < 0.1) {
      return {
        mainReasonItem:         '✅ مطابق كصنف',
        reasonItemPosPrice:     '',   // فاضي
        reasonItemTalabatPrice: '',   // فاضي
        reasonItemDiff:         '',   // فاضي
        note:                   'مطابقة كاملة'
      };
    }

    // ── غير مطابق أو صنف مخالف → تحقق يدوي ─────────────────────────────
    // (يشمل الحالات التي يوجد فيها فوتشر أيضاً - لا يُكتب "يوجد فوتشر" هنا)
    return {
      mainReasonItem:         '⚠️ تحقق يدويا',
      reasonItemPosPrice:     '',   // فاضي - المحاسب يملأه
      reasonItemTalabatPrice: '',   // فاضي - المحاسب يملأه
      reasonItemDiff:         '',   // فاضي - المحاسب يملأه
      note: `فرق صافي: ${diffSub} ج | يحتاج مراجعة يدوية`
    };
  };

  const copyToClipboard = (text) => {
    navigator.clipboard.writeText(text);
    alert("تم النسخ: " + text);
  };

  const exportToExcel = () => {
    if (discrepancies.length === 0) return;
    const exportData = discrepancies.map(item => ({
      'Order ID':               item.id,
      'Branch (Talabat)':       item.branchName,
      'Branch (Comsys)':        item.comsysBranch,
      'Items':                  item.items,
      'POS Net':                item.posSub,
      'Talabat Net':            item.talabatSub,
      'Net Diff':               item.diffSub,
      'POS Delivery':           item.posDeliv,
      'Talabat Delivery':       item.talabatDeliv,
      'Delivery Diff':          item.diffDeliv,
      'إجمالي الفرق':           item.totalDiff,
      'Talabat Voucher (150)':  item.talabatVoucher,
      'Talabat Service (2104)': item.talabatService,
      'Reason Item':            item.mainReasonItem,
      'POS Item Price':         item.reasonItemPosPrice,
      'Talabat Item Net':       item.reasonItemTalabatPrice,
      'Price Diff':             item.reasonItemDiff,
      'POS VAT':                item.posVAT,
      'Talabat VAT':            item.talabatVAT,
      'POS Final Total':        item.posTotal,
      'Talabat Final Total':    item.talabatTotal,
      'Total Diff':             item.diffFinal,
      'Notes':                  item.note
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook  = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Reconciliation_Report');
    XLSX.writeFile(
      workbook,
      `Garnell_Recon_${discrepancies[0]?.comsysBranch}_${new Date().toLocaleDateString('en-GB').replace(/\//g, '-')}.xlsx`
    );
  };

  // ─── رفع الملفات ──────────────────────────────────────────────────────────
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
        const data     = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet    = workbook.Sheets[workbook.SheetNames[0]];
        const json     = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        if (type === 'check') setCheckData(json);
        else setTalabatData(XLSX.utils.sheet_to_json(sheet));
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // ─── التحليل الرئيسي ──────────────────────────────────────────────────────
  const handleAnalyze = () => {
    if (talabatData.length === 0 || checkData.length === 0) {
      alert("من فضلك ارفع ملف طلبات وتقرير كومسيس أولاً");
      return;
    }
    setIsLoading(true);

    setTimeout(() => {
      const round = (n) => Math.round(n * 100) / 100;

      const rows = Array.isArray(checkData)
        ? checkData
        : checkData.split('\n').map(r => r.split(','));

      // ── 1. استخراج اسم الفرع من سطر Outlet ────────────────────────────
      // سطر Outlet: col[0] = "Outlet: Garnell Ard El Golf"
      const outletRow        = rows.find(r => String(r[0] || '').toLowerCase().includes('outlet:'));
      const comsysBranchFull = outletRow
        ? String(outletRow[0]).replace(/outlet:/i, '').trim()
        : 'Unknown';

      // ── 2. تحديد الفرع المقابل في طلبات ───────────────────────────────
      const talabatKeyword = resolveTalabatBranch(comsysBranchFull);

      // ── 3. بناء خريطة طلبات (فلترة الفرع + الدفع + الحالة) ────────────
      const talabatMap            = new Map();
      const talabatBranchOrderIds = [];

      talabatData.forEach(item => {
        const restaurantName = String(item['Restaurant name'] || '');
        const id             = String(item['Order ID'] || item['order_id'] || '').replace(/"/g, '').trim();
        const status         = String(item['Order status'] || '').toLowerCase().trim();

        if (status === 'cancelled' || status === 'rejected') return;
        if (!branchMatches(restaurantName, talabatKeyword)) return;
        if (!paymentMatches(item)) return;

        if (id) {
          talabatMap.set(id, item);
          talabatBranchOrderIds.push(id);
        }
      });

      // ── 4. قراءة شيكات كومسيس ──────────────────────────────────────────
      // الهيكل الحقيقي للملف:
      //   col[0]=Time  col[1]=MadeBy  col[4]=Item(code)  col[5]=Name
      //   col[6]=Qnty  col[7]=Price   col[8]=Total
      //
      //   ترتيب الصفوف:
      //   Voucher: XXXXXX   → col[0]
      //   أصناف (type=Delivery) + 2104
      //   Sub Total         → col[1]="Sub Total",       col[8]=القيمة
      //   Delivery Charge   → col[5]="Delivery Charge", col[8]=القيمة
      //   VAT               → col[5]="VAT",             col[8]=القيمة
      //   Balance Due       → col[1]="Balance Due"
      //   132               → col[4]=132, col[8]=القيمة (بعد Balance Due)
      //   150               → col[4]=150, col[8]=القيمة (بعد Balance Due)
      //   Closing Time
      //   Total             → col[1]="Total"  ← نهاية الشيك

      let currentCheck    = null;
      const allChecks     = [];
      const allVoucherIds = [];

      const finalizeCheck = () => {
        if (currentCheck && currentCheck.id) {
          allChecks.push({ ...currentCheck });
        }
        currentCheck = null;
      };

      rows.forEach((row) => {
        const c0   = String(row[0] || '');
        const c1   = String(row[1] || '').trim();
        const c4   = String(row[4] || '').trim();
        const c5   = String(row[5] || '').trim();
        const val8 = parseFloat(row[8]) || 0;

        // ── نهاية الشيك ───────────────────────────────────────────────
        if (c1 === 'Total') {
          finalizeCheck();
          return;
        }

        // ── بداية شيك جديد ────────────────────────────────────────────
        if (c0.includes('Voucher:')) {
          if (currentCheck) finalizeCheck();

          const vMatch = c0.match(/Voucher:\s*(\d+)/);
          if (!vMatch) return; // Voucher غير رقمي → تجاهل

          const vId = vMatch[1];
          allVoucherIds.push(vId);
          currentCheck = {
            id:               vId,
            items:            [],
            itemsPrices:      {},   // { "اسم الصنف": سعره } محفوظ للمرجعية
            posSub:           0,
            posDeliv:         0,
            posVAT:           0,
            posTotalAdjusted: 0,   // كود 132 (مرجعي)
            talabatVoucher:   0,   // كود 150 - محفوظ للحسابات والعمود الخاص به
            talabatService:   0    // كود 2104
          };
          return;
        }

        if (!currentCheck) return;

        // ── الأكواد المحاسبية (بعد Balance Due) ──────────────────────
        if (c4 === '132') {
          currentCheck.posTotalAdjusted = Math.abs(val8);
          return;
        }
        if (c4 === '150' || c4 === '150.0') {
          currentCheck.talabatVoucher = Math.abs(val8);
          return;
        }
        if (c4 === '2104') {
          currentCheck.talabatService = Math.abs(val8);
          return;
        }

        // ── السطور الحسابية ───────────────────────────────────────────
        if (c1 === 'Sub Total') {
          currentCheck.posSub = Math.abs(val8);
          return;
        }
        if (c5 === 'Delivery Charge') {
          currentCheck.posDeliv = Math.abs(val8);
          return;
        }
        if (c5 === 'VAT') {
          currentCheck.posVAT = Math.abs(val8);
          return;
        }
        if (c1 === 'Balance Due') return; // لا ننهي الشيك هنا

        // ── الأصناف الحقيقية ──────────────────────────────────────────
        const isSystemLine =
          c0.toLowerCase().includes('opening time') ||
          c1.toLowerCase().includes('opening time') ||
          c0.toLowerCase().includes('printing time') ||
          c1.toLowerCase().includes('printing time') ||
          c0.toLowerCase().includes('closing time')  ||
          c1.toLowerCase().includes('closing time')  ||
          c1 === 'Balance Due' || c1 === 'Sub Total' ||
          c1 === 'Total'       || c5 === 'Delivery Charge' ||
          c5 === 'VAT'         || c0 === 'Time';

        const itemName = c5;
        const itemQty  = row[6];
        const itemCode = c4;
        const isAccountingCode = ['132', '150', '150.0', '2104'].includes(itemCode);
        const typeCol  = String(row[3] || '').trim();

        if (
          itemName &&
          !isAccountingCode &&
          !isSystemLine &&
          itemQty !== undefined && itemQty !== null && itemQty !== '' &&
          !isNaN(Number(itemQty)) &&
          typeCol === 'Delivery'
        ) {
          currentCheck.items.push(`${itemName} (x${itemQty})`);
          currentCheck.itemsPrices[itemName] =
            parseFloat(String(row[7] || '0').replace(/,/g, '')) || 0;
        }
      });

      if (currentCheck) finalizeCheck();

      // ── 5. الفواتير الناقصة ───────────────────────────────────────────
      const posVoucherSet     = new Set(allVoucherIds);
      const missingInT        = allVoucherIds.filter(id => !talabatMap.has(id));
      const missingInP        = talabatBranchOrderIds.filter(id => !posVoucherSet.has(id));

      // ── 6. بناء نتائج المطابقة ────────────────────────────────────────
      const foundResults = [];

      allChecks.forEach(check => {
        const tItem = talabatMap.get(check.id);
        if (!tItem) return;

        const tSubRaw   = parseFloat(String(tItem['Subtotal']     || 0).replace(/,/g, '')) || 0;
        const tDelivRaw = parseFloat(String(tItem['Delivery Fee'] || 0).replace(/,/g, '')) || 0;
        const tTotal    = round(tSubRaw + tDelivRaw);

        const tSubComp   = round(tSubRaw   / 1.14);
        const tDelivComp = round(tDelivRaw / 1.14);
        const tVAT       = round(tTotal * 14 / 114);

        // ✅ صافي كومسيس = Sub Total + Delivery + VAT - Voucher(150)
        const posTotal = round(
          check.posSub + check.posDeliv + check.posVAT - check.talabatVoucher
        );

        const diffSub   = round(Math.abs(check.posSub   - tSubComp));
        const diffDeliv = round(Math.abs(check.posDeliv - tDelivComp));
        const totalDiff = round(diffSub + diffDeliv);
        const diffFinal = round(Math.abs(posTotal - tTotal));

        // ── منطق Reason Item ─────────────────────────────────────────
        // الفوتشر محفوظ في check.talabatVoucher للحسابات
        // لكن عمود Reason Item يطبق المنطق العادي فقط (مطابق / تحقق يدويا)
        const reason = resolveReason(check, tItem, round);

        foundResults.push({
          id:                     check.id,
          branchName:             tItem['Restaurant name'],
          comsysBranch:           comsysBranchFull,
          items:                  check.items.join(' | '),
          // ── كومسيس ──────────────────────────────────────────────────
          posSub:                 check.posSub,
          posDeliv:               check.posDeliv,
          posVAT:                 check.posVAT,
          talabatVoucher:         check.talabatVoucher,
          talabatService:         check.talabatService,
          posTotal,
          posTotal132:            check.posTotalAdjusted,
          // ── طلبات ───────────────────────────────────────────────────
          talabatSub:             tSubComp,
          talabatDeliv:           tDelivComp,
          talabatVAT:             tVAT,
          talabatTotal:           tTotal,
          // ── فروق ────────────────────────────────────────────────────
          diffSub,
          diffDeliv,
          totalDiff,
          diffFinal,
          // ── Reason Item ──────────────────────────────────────────────
          mainReasonItem:         reason.mainReasonItem,
          reasonItemPosPrice:     reason.reasonItemPosPrice,     // '' دايماً
          reasonItemTalabatPrice: reason.reasonItemTalabatPrice, // '' دايماً
          reasonItemDiff:         reason.reasonItemDiff,         // '' دايماً
          note:                   reason.note
        });
      });

      setSummary({
        posTotal:        foundResults.reduce((acc, c) => acc + c.posTotal,     0).toFixed(2),
        settlementTotal: foundResults.reduce((acc, c) => acc + c.talabatTotal, 0).toFixed(2),
        diffTotal:       foundResults.reduce((acc, c) => acc + c.diffFinal,    0).toFixed(2),
        checksCount:     foundResults.length,
        missingCount:    missingInT.length + missingInP.length
      });

      setDiscrepancies(foundResults);
      setMissingInTalabat([...new Set(missingInT)]);
      setMissingInPos([...new Set(missingInP)]);
      setIsAnalyzed(true);
      setIsLoading(false);
    }, 800);
  };

  // ─── الواجهة ──────────────────────────────────────────────────────────────
  return (
    <div className="max-w-7xl mx-auto px-4 py-8 space-y-8 bg-[#020c08] min-h-screen text-right" dir="rtl">
      <header className="text-center">
        <h1 className="text-3xl font-bold text-[#10b981]">Garnell Reconciliation - V6.2</h1>
        <p className="text-gray-400 mt-2">نظام المطابقة الموحد لجميع الفروع (دعم المعادي وأركان والغولف)</p>
      </header>

      {!isAnalyzed ? (
        <div className="bg-gray-900 p-10 rounded-2xl border border-gray-800 shadow-2xl max-w-2xl mx-auto">
          <div className="space-y-6">

            <div className="flex flex-col gap-2">
              <label className="text-emerald-400 font-bold">تقرير كومسيس (Check Details):</label>
              <input
                type="file"
                onChange={(e) => handleFileUpload(e, 'check')}
                className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg"
              />
            </div>

            <div className="flex flex-col gap-2">
              <label className="text-emerald-400 font-bold">تقرير طلبات (Excel/CSV):</label>
              <input
                type="file"
                onChange={(e) => handleFileUpload(e, 'talabat')}
                className="text-sm text-gray-400 file:bg-emerald-900 file:text-white file:border-none file:px-4 file:py-2 file:rounded-lg"
              />
            </div>

            <div className="bg-black/40 p-4 rounded-xl border border-emerald-900/30">
              <p className="text-white mb-3 font-bold text-sm">تصفية نوع الدفع:</p>
              <div className="flex gap-6">
                {['Both', 'Cash', 'Credit'].map(type => (
                  <label key={type} className="flex items-center gap-2 text-gray-300 cursor-pointer">
                    <input
                      type="radio"
                      name="payFilter"
                      value={type}
                      checked={paymentFilter === type}
                      onChange={(e) => setPaymentFilter(e.target.value)}
                      className="accent-emerald-500"
                    />
                    {type === 'Both' ? 'الكل' : type === 'Credit' ? 'فيزا / أونلاين' : 'كاش'}
                  </label>
                ))}
              </div>
              <p className="text-gray-500 text-xs mt-2">
                فيزا = Credit Card + Talabat Credit &nbsp;|&nbsp; كاش = CASH
              </p>
            </div>

            <button
              onClick={handleAnalyze}
              className="w-full py-4 bg-[#10b981] text-black font-black rounded-xl hover:bg-emerald-400 transition-all shadow-lg shadow-emerald-900/20"
            >
              {isLoading ? 'جاري تحليل البيانات...' : 'بدء عملية المطابقة الذكية'}
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
              <button
                onClick={exportToExcel}
                className="bg-emerald-600 hover:bg-emerald-500 text-white px-8 py-2 rounded-lg font-bold transition-colors"
              >
                تصدير التقرير النهائي (Excel)
              </button>
            </div>
            <DiscrepancyTable data={discrepancies} />
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="bg-gray-900 p-6 rounded-xl border border-red-900/50">
              <h3 className="text-red-400 font-bold mb-4 flex justify-between">
                فواتير كومسيس غير موجودة في طلبات
                <span>({missingInTalabat.length})</span>
              </h3>
              <div className="grid grid-cols-3 gap-2 max-h-48 overflow-y-auto p-2 bg-black/50 rounded-lg border border-gray-800">
                {missingInTalabat.map(id => (
                  <div
                    key={id}
                    onClick={() => copyToClipboard(id)}
                    className="bg-gray-800 p-2 text-center text-xs text-red-200 rounded cursor-pointer hover:bg-red-900/30 transition-colors"
                  >
                    {id}
                  </div>
                ))}
              </div>
            </div>
            <div className="bg-gray-900 p-6 rounded-xl border border-orange-900/50">
              <h3 className="text-orange-400 font-bold mb-4 flex justify-between">
                فواتير طلبات غير موجودة في كومسيس
                <span>({missingInPos.length})</span>
              </h3>
              <div className="grid grid-cols-3 gap-2 max-h-48 overflow-y-auto p-2 bg-black/50 rounded-lg border border-gray-800">
                {missingInPos.map(id => (
                  <div
                    key={id}
                    onClick={() => copyToClipboard(id)}
                    className="bg-gray-800 p-2 text-center text-xs text-orange-200 rounded cursor-pointer hover:bg-orange-900/30 transition-colors"
                  >
                    {id}
                  </div>
                ))}
              </div>
            </div>
          </div>

          <button
            onClick={() => setIsAnalyzed(false)}
            className="block mx-auto text-gray-500 hover:text-emerald-400 underline py-4 transition-colors"
          >
            العودة لرفع ملفات جديدة
          </button>
        </div>
      )}
    </div>
  );
}