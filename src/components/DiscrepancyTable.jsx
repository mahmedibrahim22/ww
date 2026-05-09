import React, { useState } from 'react';
import * as XLSX from 'xlsx';

// 1. مكون زرار التصدير المطور - مع تنسيق كامل للإكسيل
export const ExportButton = ({ data, fileName = 'Garnell_Detailed_Reconciliation_V4.xlsx' }) => {
  const handleExport = () => {
    if (!data || data.length === 0) {
      alert("لا توجد بيانات لتصديرها");
      return;
    }

    const exportData = data.map((item) => ({
      'Voucher ID': item.id || '',
      'فرع كومسيس': item.comsysBranch || '',
      'فرع طلبات': item.branchName || '',
      'محتويات الشيك': item.items || '',
      'صافي الأصناف (Sub Total)': item.posSub || 0,
      'دليفري كومسيس': item.posDeliv || 0,
      'VAT كومسيس': item.posVAT || 0,
      'قيمة Talabat Voucher (150)': item.talabatVoucher || 0,
      'إجمالي صافي كومسيس (Sub+Deliv+VAT-150)': item.posTotal || 0,
      'كود 132 (مرجعي)': item.posTotal132 || 0,
      'صافي طلبات (بعد التقسيم 1.14)': item.talabatSub || 0,
      'فرق الصافي': Math.abs(item.diffSub || 0),
      'دليفري طلبات (بعد التقسيم 1.14)': item.talabatDeliv || 0,
      'فرق الدليفري': Math.abs(item.diffDeliv || 0),
      'قيمة Talabat Services (2104)': item.talabatService || 0,
      'إجمالي طلبات (كامل شامل الضريبة)': item.talabatTotal || 0,
      'ضريبة طلبات (14%)': item.talabatVAT || 0,
      'الفرق النهائي (بناءً على الصافي)': Math.abs(item.diffFinal || 0),
      'ملاحظات التشريح': item.note || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();

    const headerKeys = Object.keys(exportData[0]);
    worksheet['!cols'] = headerKeys.map((key) => {
      if (key === 'محتويات الشيك') return { wch: 60 };
      if (key === 'ملاحظات التشريح') return { wch: 50 };
      if (key.includes('فرع') || key.includes('Voucher ID')) return { wch: 22 };
      return { wch: 28 };
    });

    worksheet['!freeze'] = { xSplit: 0, ySplit: 1 };

    XLSX.utils.book_append_sheet(workbook, worksheet, 'تحليل المطابقة النهائي');
    XLSX.writeFile(workbook, fileName);
  };

  return (
    <button
      onClick={handleExport}
      className="flex items-center space-x-2 space-x-reverse bg-[#10b981] hover:bg-[#059669] text-[#020c08] font-bold px-4 py-2 rounded-lg transition-all duration-200 shadow-lg group"
    >
      <svg className="w-5 h-5 group-hover:scale-110 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
      </svg>
      <span>تصدير التقرير النهائي (Excel)</span>
    </button>
  );
};

// 2. مكون جدول الفروقات المطور
const DiscrepancyTable = ({ data }) => {
  const [selectedInvoice, setSelectedInvoice] = useState(null);
  const checks = Array.isArray(data) ? data : [];

  const formatNum = (val) => {
    const num = parseFloat(val);
    return isNaN(num) ? "0.00" : num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };

  const parseItems = (itemsStr) => {
    if (!itemsStr) return [];
    return itemsStr.split(' | ').filter(Boolean);
  };

  return (
    <div className="bg-[#020c08] text-gray-300 p-2 rounded-xl border border-[#1a3d2f] shadow-2xl font-sans text-right" dir="rtl">

      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-4 mb-2">
        <div>
          <h2 className="text-xl font-bold text-[#10b981] mb-1">رادار المطابقة (الإصدار المحاسبي المعدل - V5)</h2>
          <p className="text-[#648b7a] text-xs">تحليل الضرائب (VAT) وفروقات الدليفري والأصناف</p>
        </div>
      </div>

      {/* الجدول الرئيسي */}
      <div className="overflow-x-auto border border-[#1a3d2f] rounded-lg">
        <table className="w-full text-[11px] text-right border-collapse">
          <thead className="bg-[#0d2a1f] text-[#10b981] sticky top-0 z-10 font-bold">
            <tr>
              <th className="px-3 py-4 border-b border-[#1a3d2f]">Voucher</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f]">فرع (ك / ط)</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center bg-black/20">الصافي والضريبة</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center">خدمات (2104) / فواتشر (150)</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center bg-black/20">الدليفري (/1.14)</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center text-white">صافي كومسيس / طلبات</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center">الفرق النهائي</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-[#1a3d2f]">
            {checks.length > 0 ? (
              checks.map((check, index) => (
                <tr
                  key={index}
                  onClick={() => setSelectedInvoice(check)}
                  className="hover:bg-[#10b981]/10 cursor-pointer transition-all group"
                >
                  <td className="px-3 py-4 font-mono text-[#10b981] font-bold">{check.id}</td>
                  <td className="px-3 py-4">
                    <div className="flex flex-col text-[9px]">
                      <span className="text-gray-400">POS: {check.comsysBranch}</span>
                      <span className="text-emerald-500 font-bold">TAL: {check.branchName}</span>
                    </div>
                  </td>
                  <td className="px-3 py-4 bg-black/20 text-center">
                    <div className="flex flex-col font-mono">
                      <span className="text-white">Sub: {formatNum(check.posSub)} | {formatNum(check.talabatSub)}</span>
                      <span className="text-blue-300 text-[9px]">VAT: {formatNum(check.posVAT)} | {formatNum(check.talabatVAT)}</span>
                    </div>
                  </td>
                  <td className="px-3 py-4 text-center">
                    <div className="flex flex-col text-[9px]">
                      {check.talabatService > 0 && <span className="text-blue-400 font-bold">Srv (2104): {formatNum(check.talabatService)}</span>}
                      {check.talabatVoucher > 0 && <span className="text-yellow-500 font-bold">Vou (150): {formatNum(check.talabatVoucher)}</span>}
                      {check.talabatService === 0 && check.talabatVoucher === 0 && <span className="text-gray-600">---</span>}
                    </div>
                  </td>
                  <td className="px-3 py-4 bg-black/20 text-center">
                    <div className="flex flex-col font-mono">
                      <span className="text-orange-400">{formatNum(check.posDeliv)} | {formatNum(check.talabatDeliv)}</span>
                      <span className="text-[9px] text-gray-500">Diff: {formatNum(check.diffDeliv)}</span>
                    </div>
                  </td>
                  <td className="px-3 py-4 text-center font-bold text-white">
                    <div className="flex flex-col font-mono">
                      <span className="text-emerald-400 text-[11px] font-bold">POS: {formatNum(check.posTotal)}</span>
                      <span className="text-[#10b981] border-t border-white/5 pt-1">TAL: {formatNum(check.talabatTotal)}</span>
                    </div>
                  </td>
                  <td className="px-3 py-4 text-center">
                    <span className={`px-2 py-1 rounded text-[10px] font-bold font-mono ${Math.abs(check.diffFinal) > 0.5 ? 'bg-red-900/50 text-red-400 border border-red-500/50' : 'bg-emerald-900/30 text-emerald-400'}`}>
                      {formatNum(Math.abs(check.diffFinal))}
                    </span>
                  </td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="7" className="px-5 py-20 text-center text-gray-500 italic">في انتظار رفع البيانات للتحليل...</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* ===== مودال تفاصيل الفاتورة ===== */}
      {selectedInvoice && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <div className="bg-[#0d2a1f] border border-[#10b981]/50 w-full max-w-2xl rounded-2xl overflow-hidden shadow-2xl">
            <div className="p-5 border-b border-[#1a3d2f] flex justify-between items-center bg-[#020c08]/50">
              <div className="flex items-center gap-3 text-right">
                <div className="w-2 h-8 bg-[#10b981] rounded-full"></div>
                <h3 className="font-bold text-white text-lg">
                  تحليل الشيك رقم <span className="text-[#10b981] font-mono">#{selectedInvoice.id}</span>
                </h3>
              </div>
              <button onClick={() => setSelectedInvoice(null)} className="w-10 h-10 rounded-full bg-red-900/20 text-red-500 hover:bg-red-500 hover:text-white transition-all flex items-center justify-center">✕</button>
            </div>

            <div className="p-6 overflow-y-auto max-h-[75vh] space-y-6">
              {/* الصنف المحتمل */}
              <div className="bg-emerald-900/20 p-3 rounded-lg border border-emerald-500/30">
                <p className="text-emerald-400 text-xs font-bold">تحليل الصنف المتسبب:</p>
                <p className="text-white text-sm mt-1">{selectedInvoice.suspectedItem || "تحليل تقريبي للضريبة"}</p>
              </div>

              {/* الأصناف */}
              <div>
                <h4 className="text-[10px] font-bold text-[#10b981] uppercase tracking-[0.2em] border-b border-[#1a3d2f] pb-2 mb-3">🧾 الأصناف في كومسيس</h4>
                <div className="bg-black/30 p-3 rounded-xl border border-[#1a3d2f]">
                  {parseItems(selectedInvoice.items).length > 0 ? (
                    <div className="grid grid-cols-1 gap-1">
                      {parseItems(selectedInvoice.items).map((item, i) => (
                        <div key={i} className="flex items-center gap-2 text-sm py-1 border-b border-[#1a3d2f]/50 last:border-0">
                          <span className="text-[#10b981] font-mono text-[10px] w-5 text-center">{i + 1}</span>
                          <span className="text-gray-300">{item}</span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="text-gray-500 text-xs text-center py-2">لا توجد أصناف مسجلة</p>
                  )}
                </div>
              </div>

              {/* الحسبة المحاسبية */}
              <div className="grid md:grid-cols-2 gap-6">
                <div className="space-y-3 text-right">
                  <h4 className="text-[10px] font-bold text-[#10b981] uppercase tracking-[0.2em] border-b border-[#1a3d2f] pb-2">تفصيل كومسيس (POS)</h4>
                  <div className="bg-black/30 p-4 rounded-xl border border-[#1a3d2f] space-y-2 text-sm font-mono">
                    <div className="flex justify-between"><span>Sub Total:</span><span className="text-white">{formatNum(selectedInvoice.posSub)}</span></div>
                    <div className="flex justify-between"><span>Delivery:</span><span className="text-orange-400">{formatNum(selectedInvoice.posDeliv)}</span></div>
                    <div className="flex justify-between"><span>VAT (14%):</span><span className="text-blue-300">{formatNum(selectedInvoice.posVAT)}</span></div>
                    <div className="flex justify-between text-yellow-500 border-t border-white/10 pt-1">
                      <span>(-) Voucher (150):</span>
                      <span>- {formatNum(selectedInvoice.talabatVoucher)}</span>
                    </div>
                    <div className="flex justify-between font-bold text-[#10b981] text-lg border-t border-[#10b981]/30 pt-2">
                      <span>صافي كومسيس:</span>
                      <span>{formatNum(selectedInvoice.posTotal)}</span>
                    </div>
                  </div>
                </div>

                <div className="space-y-3 text-right">
                  <h4 className="text-[10px] font-bold text-blue-400 uppercase tracking-[0.2em] border-b border-[#1a3d2f] pb-2">تفصيل طلبات (Talabat)</h4>
                  <div className="bg-black/30 p-4 rounded-xl border border-[#1a3d2f] space-y-2 text-sm font-mono">
                    <div className="flex justify-between text-gray-400"><span>صافي الأصناف:</span><span>{formatNum(selectedInvoice.talabatSub)}</span></div>
                    <div className="flex justify-between text-gray-400"><span>صافي الدليفري:</span><span>{formatNum(selectedInvoice.talabatDeliv)}</span></div>
                    <div className="flex justify-between text-purple-400 border-t border-white/10 pt-1"><span>الضريبة (14%):</span><span>{formatNum(selectedInvoice.talabatVAT)}</span></div>
                    <div className="flex justify-between font-bold text-white text-lg border-t border-white/30 pt-2">
                      <span>إجمالي طلبات:</span>
                      <span>{formatNum(selectedInvoice.talabatTotal)}</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="p-4 bg-black/40 flex flex-col items-center gap-2 border-t border-[#1a3d2f]">
              <div className={`text-xs font-bold px-6 py-2 rounded-full border ${Math.abs(selectedInvoice.diffFinal) > 0.5 ? 'text-red-400 border-red-500/50' : 'text-emerald-400 border-emerald-500/50'}`}>
                فرق المطابقة النهائي: {formatNum(Math.abs(selectedInvoice.diffFinal))} ج.م
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Legend */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-4 p-4 bg-[#0d2a1f]/30 rounded-lg border border-[#1a3d2f]/50">
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أبيض = كومسيس</span><div className="w-2 h-2 bg-white rounded-full"></div></div>
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أخضر = طلبات</span><div className="w-2 h-2 bg-emerald-500 rounded-full"></div></div>
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أزرق = ضريبة/خدمات</span><div className="w-2 h-2 bg-blue-400 rounded-full"></div></div>
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أصفر = فواتشر (150)</span><div className="w-2 h-2 bg-yellow-400 rounded-full"></div></div>
      </div>
    </div>
  );
};

export default DiscrepancyTable;