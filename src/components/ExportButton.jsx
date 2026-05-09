import React from 'react';
import * as XLSX from 'xlsx';

/**
 * 1. مكون زرار التصدير المطور
 * يدعم تصدير كافة بنود التشريح (صافي، ضريبة، دليفري) لضمان دقة التقرير
 */
export const ExportButton = ({ data, fileName = 'Garnial_Reconciliation_Report.xlsx' }) => {
  const handleExport = () => {
    if (!data || data.length === 0) {
      alert("لا توجد بيانات لتصديرها");
      return;
    }

    // تجهيز البيانات لتقرير الإكسيل النهائي بناءً على هيكلة التشريح الجديدة
    const exportData = data.map((item) => ({
      'Voucher ID': item.id || '',
      'الفرع (كومسيس)': item.comsysBranch || '',
      'الفرع (طلبات)': item.branchName || '',
      'صافي كومسيس': item.posSub || 0,
      'صافي طلبات': item.talabatSub || 0,
      'فرق الصافي': item.diffSub || 0,
      'ضريبة كومسيس': item.posTax || 0,
      'ضريبة طلبات': item.talabatTax || 0,
      'فرق الضريبة': item.diffTax || 0,
      'دليفري كومسيس': item.posDeliv || 0,
      'دليفري طلبات': item.talabatDeliv || 0,
      'فرق الدليفري': item.diffDeliv || 0,
      'إجمالي كومسيس': item.posTotal || 0,
      'إجمالي طلبات': item.talabatTotal || 0,
      'الفرق المجمع النهائي': item.diffFinal || 0,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'تقرير المطابقة التفصيلي');
    XLSX.writeFile(workbook, fileName);
  };

  return (
    <button
      onClick={handleExport}
      className="flex items-center space-x-2 space-x-reverse bg-[#0d2a1f] border border-[#1a3d2f] hover:bg-[#1a3d2f] text-[#10b981] font-bold px-5 py-2 rounded-lg transition-all duration-200 shadow-lg group"
    >
      <svg className="w-5 h-5 group-hover:scale-110 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
      </svg>
      <span>تصدير تقرير الإكسيل التشريحي</span>
    </button>
  );
};

/**
 * 2. مكون جدول الفروقات الرئيسي
 */
export default function DiscrepancyTable({ data, hideDiff = false }) {
  const rows = Array.isArray(data) ? data : [];

  const formatNum = (val) => {
    const num = parseFloat(val);
    return isNaN(num) ? "0.00" : num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };

  return (
    <div className="bg-[#020c08] text-gray-300 p-2 rounded-xl font-sans text-right" dir="rtl">
      
      {/* Header مع زر التصدير */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-4 mb-2">
        <div>
          <h2 className="text-xl font-bold text-[#10b981] mb-1">رادار المطابقة (التشريح الرباعي)</h2>
          <p className="text-[#648b7a] text-xs">تحليل الفروقات بناءً على بنود الصافي والضريبة والخدمة</p>
        </div>
        {!hideDiff && <ExportButton data={rows} />}
      </div>

      {/* الجدول الرئيسي */}
      <div className="overflow-x-auto border border-[#1a3d2f] rounded-lg shadow-2xl">
        <table className="w-full text-sm text-right border-collapse">
          <thead className="bg-[#0d2a1f] text-[#10b981] text-xs tracking-wider uppercase font-bold">
            <tr>
              <th className="px-5 py-4 border-b border-[#1a3d2f]">Voucher ID</th>
              <th className="px-5 py-4 border-b border-[#1a3d2f]">الصافي (ك/ط)</th>
              <th className="px-5 py-4 border-b border-[#1a3d2f]">الضريبة (ك/ط)</th>
              <th className="px-5 py-4 border-b border-[#1a3d2f]">الدليفري (ك/ط)</th>
              <th className="px-5 py-4 border-b border-[#1a3d2f]">إجمالي الشيك</th>
              {!hideDiff && <th className="px-5 py-4 border-b border-[#1a3d2f]">الفرق النهائي</th>}
            </tr>
          </thead>
          <tbody className="divide-y divide-[#1a3d2f] text-gray-400">
            {rows.length > 0 ? (
              rows.map((row, index) => (
                <tr key={index} className="hover:bg-[#0d2a1f]/50 transition-colors">
                  <td className="px-5 py-4 font-mono text-[#10b981] font-medium">
                    {row.id}
                  </td>
                  
                  {/* الصافي */}
                  <td className="px-5 py-4">
                    <div className="flex flex-col font-mono text-xs">
                      <span className="text-gray-300">{formatNum(row.posSub)}</span>
                      <span className="text-emerald-500 border-t border-[#1a3d2f]/30">{formatNum(row.talabatSub)}</span>
                    </div>
                  </td>

                  {/* الضريبة */}
                  <td className="px-5 py-4">
                    <div className="flex flex-col font-mono text-xs">
                      <span className="text-gray-400">{formatNum(row.posTax)}</span>
                      <span className="text-blue-400 border-t border-[#1a3d2f]/30">{formatNum(row.talabatTax)}</span>
                    </div>
                  </td>

                  {/* الدليفري */}
                  <td className="px-5 py-4">
                    <div className="flex flex-col font-mono text-xs">
                      <span className="text-gray-400">{formatNum(row.posDeliv)}</span>
                      <span className="text-orange-400 border-t border-[#1a3d2f]/30">{formatNum(row.talabatDeliv)}</span>
                    </div>
                  </td>

                  {/* الإجمالي التراكمي */}
                  <td className="px-5 py-4 font-mono font-bold text-white">
                    <div className="flex flex-col">
                      <span>{formatNum(row.posTotal)}</span>
                      <span className="text-[#10b981] border-t border-white/10">{formatNum(row.talabatTotal)}</span>
                    </div>
                  </td>

                  {/* الفرق النهائي (المجمع) */}
                  {!hideDiff && (
                    <td className="px-5 py-4">
                      <span className={`px-2 py-1 rounded text-[10px] font-bold font-mono ${Math.abs(row.diffFinal) > 1 ? 'bg-red-900/50 text-red-400 border border-red-500/50' : 'bg-emerald-900/30 text-emerald-400'}`}>
                        {formatNum(row.diffFinal)}
                      </span>
                    </td>
                  )}
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="6" className="px-5 py-16 text-center text-gray-600 italic">
                  لا توجد بيانات للمطابقة حالياً. قم برفع الملفات لبدء التشريح المالي.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* تنبيه محاسبي وتوضيح الألوان */}
      {!hideDiff && (
        <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#0d2a1f]/30 p-4 rounded-lg border border-[#1a3d2f]/50 text-[11px]">
          <div className="flex items-center gap-2 text-[#648b7a]">
            <svg className="w-4 h-4 text-[#10b981]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            <span>يتم حساب الفروقات بناءً على القيمة المطلقة لكل بند (صافي، ضريبة، خدمة) لضمان دقة الرصد.</span>
          </div>
          <div className="flex items-center justify-end gap-3 text-gray-500">
             <span className="flex items-center gap-1"><div className="w-2 h-2 bg-emerald-500 rounded-full"></div> طلبات</span>
             <span className="flex items-center gap-1"><div className="w-2 h-2 bg-white rounded-full"></div> كومسيس</span>
             <span className="flex items-center gap-1"><div className="w-2 h-2 bg-blue-400 rounded-full"></div> ضريبة</span>
             <span className="flex items-center gap-1"><div className="w-2 h-2 bg-orange-400 rounded-full"></div> دليفري</span>
          </div>
        </div>
      )}
    </div>
  );
}