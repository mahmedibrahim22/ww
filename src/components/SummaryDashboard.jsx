import React from 'react';

export default function SummaryDashboard({ data }) {
  const summary = {
    totalPOS: Number(data?.totalPOS || 0),
    totalTalabat: Number(data?.totalTalabat || 0),
    totalDifference: Number(data?.totalDifference || 0),
    totalChecks: Number(data?.totalChecks || 0),
    totalTaxDiff: Number(data?.totalTaxDiff || 0),
    totalDelivDiff: Number(data?.totalDelivDiff || 0),
    totalDiscounts: Number(data?.totalDiscounts || 0),
    totalComplimentary: Number(data?.totalComplimentary || 0),
    missingInCashier: Number(data?.missingInCashier || 0),
  };

  const hasDiffData = summary.totalDiscounts > 0 || summary.totalComplimentary > 0 || summary.missingInCashier > 0;

  return (
    <div className="bg-[#020c08] text-gray-300 p-6 rounded-xl border border-[#1a3d2f] max-w-6xl mx-auto shadow-2xl font-sans" dir="rtl">
      
      {/* Header Section */}
      <div className="mb-8 text-right border-r-4 border-[#10b981] pr-4">
        <h1 className="text-2xl font-bold text-white mb-1">
          لوحة تحكم الملخصات (Summary Dashboard)
        </h1>
        <p className="text-[#648b7a] text-sm font-medium">
          تحليل إجماليات المطابقة والتشريح المالي | فرع الجولف وألماظة
        </p>
      </div>

      {/* Main Grid for Cards - Reconciliation */}
      <p className="text-[#648b7a] text-xs font-bold uppercase tracking-widest mb-3 text-right">مطابقة طلبات التوصيل</p>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-8">
        
        {/* Card 1: Total POS */}
        <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] p-5 rounded-lg flex flex-col justify-between h-40 transition-all duration-300 hover:bg-[#0d2a1f]/60 hover:border-[#10b981]/50 group relative overflow-hidden">
          <div className="absolute top-0 left-0 w-1 h-full bg-white opacity-20 group-hover:opacity-100 transition-opacity"></div>
          <div>
            <span className="text-[#648b7a] text-xs uppercase tracking-wider block mb-1 font-bold group-hover:text-[#10b981] transition-colors">
              إجمالي مبيعات المطعم (POS)
            </span>
            <span className="text-2xl font-bold text-white mt-2 block font-mono">
              {summary.totalPOS.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            </span>
          </div>
          <div className="text-[#4b6b5d] text-[10px] flex items-center mt-2 border-t border-[#1a3d2f] pt-2">
            <span>مجموع (الصافي + الضريبة + الخدمة) - كومسيس</span>
          </div>
        </div>

        {/* Card 2: Total Settlement (Talabat) */}
        <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] p-5 rounded-lg flex flex-col justify-between h-40 transition-all duration-300 hover:bg-[#0d2a1f]/60 hover:border-[#10b981]/50 group relative overflow-hidden">
          <div className="absolute top-0 left-0 w-1 h-full bg-[#10b981] opacity-40 group-hover:opacity-100 transition-opacity"></div>
          <div>
            <span className="text-[#648b7a] text-xs uppercase tracking-wider block mb-1 font-bold group-hover:text-[#10b981] transition-colors">
              إجمالي مبيعات طلبات (Talabat)
            </span>
            <span className="text-2xl font-bold text-[#10b981] mt-2 block font-mono">
              {summary.totalTalabat.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            </span>
          </div>
          <div className="text-[#4b6b5d] text-[10px] flex items-center mt-2 border-t border-[#1a3d2f] pt-2">
            <span>إجمالي المبالغ المسجلة في ملف التسوية</span>
          </div>
        </div>

        {/* Card 3: Total Difference */}
        <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] p-5 rounded-lg flex flex-col justify-between h-40 transition-all duration-300 hover:bg-[#0d2a1f]/60 group relative overflow-hidden">
          <div className={`absolute top-0 left-0 w-1 h-full ${Math.abs(summary.totalDifference) > 1 ? 'bg-red-500' : 'bg-emerald-500'} opacity-40 group-hover:opacity-100 transition-opacity`}></div>
          <div>
            <span className="text-[#648b7a] text-xs uppercase tracking-wider block mb-1 font-bold group-hover:text-white transition-colors">
              الفرق الإجمالي المجمع (Gap)
            </span>
            <span className={`text-2xl font-bold mt-2 block font-mono ${Math.abs(summary.totalDifference) > 1 ? 'text-red-400' : 'text-emerald-400'}`}>
              {summary.totalDifference.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            </span>
          </div>
          <div className="text-[#4b6b5d] text-[10px] flex items-center mt-2 border-t border-[#1a3d2f] pt-2">
            <span>مجموع الانحرافات المطلقة المكتشفة</span>
          </div>
        </div>

        {/* Card 4: Count of Operations */}
        <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] p-5 rounded-lg flex flex-col justify-between h-40 transition-all duration-300 hover:bg-[#0d2a1f]/60 group relative overflow-hidden">
          <div className="absolute top-0 left-0 w-1 h-full bg-sky-500 opacity-40 group-hover:opacity-100 transition-opacity"></div>
          <div>
            <span className="text-[#648b7a] text-xs uppercase tracking-wider block mb-1 font-bold group-hover:text-sky-400 transition-colors">
              إجمالي الشيكات / الفواتير
            </span>
            <span className="text-2xl font-bold text-sky-400 mt-2 block font-mono">
              {summary.totalChecks}
            </span>
          </div>
          <div className="text-[#4b6b5d] text-[10px] flex items-center mt-2 border-t border-[#1a3d2f] pt-2">
            <span>عدد الأوردرات التي خضعت للتدقيق</span>
          </div>
        </div>
        
      </div>

      {/* Discount Section - يظهر فقط لو فيه بيانات خصومات */}
      {hasDiffData && (
        <>
          <div className="border-t border-[#1a3d2f] mb-6"></div>
          <p className="text-[#648b7a] text-xs font-bold uppercase tracking-widest mb-3 text-right">مراجعة الخصومات والكمبلمنتري</p>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-6 mb-8">

            {/* Card: Total Discounts */}
            <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] p-5 rounded-lg flex flex-col justify-between h-36 transition-all duration-300 hover:bg-[#0d2a1f]/60 hover:border-yellow-500/30 group relative overflow-hidden">
              <div className="absolute top-0 left-0 w-1 h-full bg-yellow-400 opacity-40 group-hover:opacity-100 transition-opacity"></div>
              <div>
                <span className="text-[#648b7a] text-xs uppercase tracking-wider block mb-1 font-bold group-hover:text-yellow-400 transition-colors">
                  إجمالي أوردرات الخصم
                </span>
                <span className="text-2xl font-bold text-yellow-400 mt-2 block font-mono">
                  {summary.totalDiscounts}
                </span>
              </div>
              <div className="text-[#4b6b5d] text-[10px] mt-2 border-t border-[#1a3d2f] pt-2">
                <span>أوردرات Discounted Items</span>
              </div>
            </div>

            {/* Card: Total Complimentary */}
            <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] p-5 rounded-lg flex flex-col justify-between h-36 transition-all duration-300 hover:bg-[#0d2a1f]/60 hover:border-purple-500/30 group relative overflow-hidden">
              <div className="absolute top-0 left-0 w-1 h-full bg-purple-400 opacity-40 group-hover:opacity-100 transition-opacity"></div>
              <div>
                <span className="text-[#648b7a] text-xs uppercase tracking-wider block mb-1 font-bold group-hover:text-purple-400 transition-colors">
                  إجمالي أوردرات الكمبلمنتري
                </span>
                <span className="text-2xl font-bold text-purple-400 mt-2 block font-mono">
                  {summary.totalComplimentary}
                </span>
              </div>
              <div className="text-[#4b6b5d] text-[10px] mt-2 border-t border-[#1a3d2f] pt-2">
                <span>أوردرات Complimentary</span>
              </div>
            </div>

            {/* Card: Missing in Cashier */}
            <div className="bg-[#0d2a1f]/40 border border-[#1a3d2f] p-5 rounded-lg flex flex-col justify-between h-36 transition-all duration-300 hover:bg-[#0d2a1f]/60 hover:border-red-500/30 group relative overflow-hidden">
              <div className={`absolute top-0 left-0 w-1 h-full ${summary.missingInCashier > 0 ? 'bg-red-500' : 'bg-emerald-500'} opacity-40 group-hover:opacity-100 transition-opacity`}></div>
              <div>
                <span className="text-[#648b7a] text-xs uppercase tracking-wider block mb-1 font-bold group-hover:text-red-400 transition-colors">
                  مش مسجل في شيت الكاشير
                </span>
                <span className={`text-2xl font-bold mt-2 block font-mono ${summary.missingInCashier > 0 ? 'text-red-400' : 'text-emerald-400'}`}>
                  {summary.missingInCashier}
                </span>
              </div>
              <div className="text-[#4b6b5d] text-[10px] mt-2 border-t border-[#1a3d2f] pt-2">
                <span>{summary.missingInCashier > 0 ? '⚠ يحتاج مراجعة' : '✓ كل الأوردرات مسجلة'}</span>
              </div>
            </div>

          </div>
        </>
      )}

      {/* Footer Info Box */}
      <div className="mt-2 bg-[#0d2a1f]/20 border border-[#1a3d2f]/50 p-4 rounded-lg flex items-start space-x-3 space-x-reverse">
        <div className="text-[#10b981] mt-0.5 ml-3">
          <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
          </svg>
        </div>
        <div>
          <h4 className="text-[#10b981] text-sm font-bold mb-1">دليل دقة البيانات (Data Accuracy):</h4>
          <p className="text-[#648b7a] text-[11px] leading-relaxed">
            هذه الأرقام تمثل نتيجة عملية <b>"التشريح الرباعي"</b>. الفرق الإجمالي ليس مجرد طرح رقمين، بل هو مجموع الفروقات المكتشفة في (الصافي، الضريبة، وخدمة التوصيل) لكل فاتورة على حدة. في حالة وجود فروقات كبيرة، يرجى مراجعة الجدول التفصيلي بالأسفل وتصدير تقرير الإكسيل للتحليل العميق.
          </p>
        </div>
      </div>
      
    </div>
  );
}