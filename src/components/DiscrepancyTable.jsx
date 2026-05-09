import React, { useState } from 'react';
import * as XLSX from 'xlsx';

// 1. مكون زرار التصدير المطور - يشمل كافة الأعمدة الجديدة والحسابات المعدلة بناءً على الأكواد
export const ExportButton = ({ data, fileName = 'Garnell_Detailed_Reconciliation_V3.xlsx' }) => {
  const handleExport = () => {
    if (!data || data.length === 0) {
      alert("لا توجد بيانات لتصديرها");
      return;
    }

    const exportData = data.map((item) => {
      return {
        'Voucher ID': item.id || '',
        'فرع كومسيس': item.comsysBranch || '',
        'فرع طلبات': item.branchName || '',
        'محتويات الشيك': item.items || '',
        'صافي كومسيس': item.posSub || 0,
        'صافي طلبات (بعد التقسيم 1.14)': item.talabatSub || 0,
        'فرق الصافي': Math.abs(item.diffSub || 0),
        'دليفري كومسيس': item.posDeliv || 0,
        'دليفري طلبات (بعد التقسيم 1.14)': item.talabatDeliv || 0,
        'فرق الدليفري': Math.abs(item.diffDeliv || 0),
        'قيمة Talabat Services (2104)': item.talabatService || 0,
        'قيمة Talabat Voucher (150)': item.talabatVoucher || 0,
        'إجمالي كومسيس المصفى (132)': item.posTotal || 0,
        'إجمالي طلبات (كامل شامل الضريبة)': item.talabatTotal || 0,
        'الفرق النهائي (بناءً على الصافي)': Math.abs(item.diffFinal || 0),
        'ملاحظات التشريح': item.note || ''
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
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
      <span>تصدير التقرير النهائي</span>
    </button>
  );
};

// 2. مكون جدول الفروقات المطور بالتعديلات الجديدة (الأكواد المحاسبية)
const DiscrepancyTable = ({ data }) => {
  const [selectedInvoice, setSelectedInvoice] = useState(null);
  const checks = Array.isArray(data) ? data : [];

  const formatNum = (val) => {
    const num = parseFloat(val);
    return isNaN(num) ? "0.00" : num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };

  return (
    <div className="bg-[#020c08] text-gray-300 p-2 rounded-xl border border-[#1a3d2f] shadow-2xl font-sans text-right" dir="rtl">
      
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-4 mb-2">
        <div>
          <h2 className="text-xl font-bold text-[#10b981] mb-1">رادار المطابقة (الإصدار المحاسبي المعدل - V3)</h2>
          <p className="text-[#648b7a] text-xs">مقارنة الصافي (/1.14) وتتبع الأكواد المحاسبية (150 - 2104 - 132)</p>
        </div>
        <ExportButton data={checks} />
      </div>

      {/* الجدول الرئيسي */}
      <div className="overflow-x-auto border border-[#1a3d2f] rounded-lg">
        <table className="w-full text-[11px] text-right border-collapse">
          <thead className="bg-[#0d2a1f] text-[#10b981] sticky top-0 z-10 font-bold">
            <tr>
              <th className="px-3 py-4 border-b border-[#1a3d2f]">Voucher</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f]">فرع (ك / ط)</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center bg-black/20">الصافي (/1.14)</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center">Services (2104) / Voucher (150)</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center bg-black/20">الدليفري (/1.14)</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center text-white">إجمالي مصفى (132) / طلبات</th>
              <th className="px-3 py-4 border-b border-[#1a3d2f] text-center">الفرق (بناءً على الصافي)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-[#1a3d2f]">
            {checks.length > 0 ? (
              checks.map((check, index) => {
                return (
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

                    {/* الصافي المقسم */}
                    <td className="px-3 py-4 bg-black/20 text-center">
                      <div className="flex flex-col font-mono">
                        <span className="text-white">{formatNum(check.posSub)} | {formatNum(check.talabatSub)}</span>
                        <span className="text-yellow-500 text-[9px] mt-1">فرق: {formatNum(check.diffSub)}</span>
                      </div>
                    </td>

                    {/* خدمات وفاوتشر طلبات - بناءً على الأكواد 2104 و 150 */}
                    <td className="px-3 py-4 text-center">
                      <div className="flex flex-col text-[9px]">
                        {check.talabatService > 0 && <span className="text-blue-400 font-bold">Srv (2104): {formatNum(check.talabatService)}</span>}
                        {check.talabatVoucher > 0 && <span className="text-yellow-500 font-bold">Vou (150): {formatNum(check.talabatVoucher)}</span>}
                        {check.talabatService === 0 && check.talabatVoucher === 0 && <span className="text-gray-600">---</span>}
                      </div>
                    </td>

                    {/* الدليفري المقسم */}
                    <td className="px-3 py-4 bg-black/20 text-center">
                      <div className="flex flex-col font-mono">
                        <span className="text-orange-400">{formatNum(check.posDeliv)} | {formatNum(check.talabatDeliv)}</span>
                        <span className="text-[8px] text-gray-500">فرق: {formatNum(check.diffDeliv)}</span>
                      </div>
                    </td>

                    {/* الإجماليات الكاملة - تشمل كود 132 */}
                    <td className="px-3 py-4 text-center font-bold text-white">
                      <div className="flex flex-col font-mono">
                        <span className="text-gray-400 text-[10px]">كود (132): {formatNum(check.posTotal)}</span>
                        <span className="text-[#10b981] border-t border-white/10">{formatNum(check.talabatTotal)}</span>
                      </div>
                    </td>

                    {/* الفرق النهائي */}
                    <td className="px-3 py-4 text-center">
                      <span className={`px-2 py-1 rounded text-[10px] font-bold font-mono ${Math.abs(check.diffFinal) > 0.1 ? 'bg-red-900/50 text-red-400 border border-red-500/50' : 'bg-emerald-900/30 text-emerald-400'}`}>
                        {formatNum(Math.abs(check.diffFinal))}
                      </span>
                    </td>
                  </tr>
                );
              })
            ) : (
              <tr>
                <td colSpan="7" className="px-5 py-20 text-center text-gray-500 italic">في انتظار رفع البيانات للتحليل...</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* مودال تفاصيل الفاتورة */}
      {selectedInvoice && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <div className="bg-[#0d2a1f] border border-[#10b981]/50 w-full max-w-2xl rounded-2xl overflow-hidden shadow-2xl animate-in zoom-in duration-200">
            <div className="p-5 border-b border-[#1a3d2f] flex justify-between items-center bg-[#020c08]/50">
              <div className="flex items-center gap-3 text-right">
                 <div className="w-2 h-8 bg-[#10b981] rounded-full"></div>
                 <h3 className="font-bold text-white text-lg">تحليل الشيك رقم <span className="text-[#10b981] font-mono">#{selectedInvoice.id}</span></h3>
              </div>
              <button onClick={() => setSelectedInvoice(null)} className="w-10 h-10 rounded-full bg-red-900/20 text-red-500 hover:bg-red-500 transition-all flex items-center justify-center">✕</button>
            </div>

            <div className="p-6 grid md:grid-cols-2 gap-8 overflow-y-auto max-h-[75vh]">
              {/* جانب نظام كومسيس */}
              <div className="space-y-4 text-right">
                <h4 className="text-[10px] font-bold text-[#10b981] uppercase tracking-[0.2em] border-b border-[#1a3d2f] pb-2 text-right">بيانات كومسيس (POS)</h4>
                <div className="bg-black/30 p-4 rounded-xl border border-[#1a3d2f] space-y-3 text-sm">
                  <div className="flex justify-between font-mono"><span>الصافي:</span> <span>{formatNum(selectedInvoice.posSub)}</span></div>
                  <div className="flex justify-between font-mono"><span>الدليفري:</span> <span>{formatNum(selectedInvoice.posDeliv)}</span></div>
                  <div className="h-px bg-[#1a3d2f] my-2"></div>
                  <div className="flex justify-between font-bold text-white border-t border-[#10b981]/20 pt-3 mt-1 text-base">
                    <span>إجمالي (كود 132):</span> 
                    <span>{formatNum(selectedInvoice.posTotal)}</span>
                  </div>
                </div>
              </div>

              {/* جانب نظام طلبات */}
              <div className="space-y-4 text-right">
                <h4 className="text-[10px] font-bold text-blue-400 uppercase tracking-[0.2em] border-b border-[#1a3d2f] pb-2 text-right">بيانات طلبات (مقارنة بالصافي)</h4>
                <div className="bg-black/30 p-4 rounded-xl border border-[#1a3d2f] space-y-3 text-sm">
                  <div className="flex justify-between font-mono text-gray-400"><span>الصافي (/1.14):</span> <span>{formatNum(selectedInvoice.talabatSub)}</span></div>
                  <div className="flex justify-between font-mono text-gray-400"><span>الدليفري (/1.14):</span> <span>{formatNum(selectedInvoice.talabatDeliv)}</span></div>
                  <div className="flex justify-between font-mono text-blue-400"><span>خدمات (2104):</span> <span>{formatNum(selectedInvoice.talabatService)}</span></div>
                  <div className="flex justify-between font-mono text-yellow-500 font-bold border-t border-white/5 pt-2"><span>فاوتشر (150):</span> <span>{formatNum(selectedInvoice.talabatVoucher)}</span></div>
                  <div className="flex justify-between font-bold text-[#10b981] border-t border-blue-400/20 pt-3 mt-1 text-base">
                    <span>إجمالي (كامل):</span> <span>{formatNum(selectedInvoice.talabatTotal)}</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="p-4 bg-black/40 flex flex-col items-center gap-2 border-t border-[#1a3d2f]">
              <div className="text-[11px] text-emerald-400 font-bold bg-emerald-900/20 px-4 py-2 rounded-full border border-emerald-500/20 text-center">
                الفرق عند المقارنة بالصافي = {formatNum(Math.abs(selectedInvoice.diffFinal))} ج.م
              </div>
              <div className="text-[10px] text-gray-500">{selectedInvoice.note}</div>
            </div>
          </div>
        </div>
      )}

      {/* Legend */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-4 p-4 bg-[#0d2a1f]/30 rounded-lg border border-[#1a3d2f]/50">
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أبيض = كومسيس</span><div className="w-2 h-2 bg-white rounded-full"></div></div>
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أخضر = طلبات</span><div className="w-2 h-2 bg-emerald-500 rounded-full"></div></div>
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أزرق = خدمات (2104)</span><div className="w-2 h-2 bg-blue-400 rounded-full"></div></div>
        <div className="flex items-center gap-2 text-[10px] justify-end"><span>أصفر = الفاوتشر (150)</span><div className="w-2 h-2 bg-yellow-400 rounded-full"></div></div>
      </div>
    </div>
  );
};

export default DiscrepancyTable;