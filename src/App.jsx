import React from 'react';
import ReconciliationEngine from './utils/reconciliationEngine';

function App() {
  return (
    <div className="min-h-screen bg-[#020c08] text-gray-200 font-sans selection:bg-[#10b981]/30" dir="rtl">
      
      {/* Header احترافي يعكس هوية Garnial */}
      <header className="bg-[#0d2a1f]/80 border-b border-[#1a3d2f] py-4 sticky top-0 z-50 backdrop-blur-md shadow-lg">
        <div className="max-w-7xl mx-auto px-6 flex justify-between items-center">
          
          <div className="flex items-center space-x-4 space-x-reverse">
            {/* الأيقونة المميزة */}
            <div className="w-10 h-10 bg-[#10b981] rounded-xl flex items-center justify-center shadow-[0_0_20px_rgba(16,185,129,0.3)] group transition-transform hover:rotate-12">
              <svg className="w-6 h-6 text-[#020c08]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
              </svg>
            </div>
            
            <div>
              <h1 className="text-xl font-extrabold tracking-tight text-white leading-none">
                نظام مطابقة تسويات <span className="text-[#10b981]">Garnell</span>
              </h1>
              <p className="text-[#4b6b5d] text-[10px] mt-1 font-medium uppercase tracking-widest">
                Financial Audit System - RECON V2
              </p>
            </div>
          </div>

          {/* حالة النظام */}
          <div className="flex items-center gap-3">
            <div className="hidden md:flex flex-col items-end">
              <span className="text-[10px] text-[#10b981] font-bold animate-pulse">● System Online</span>
              <span className="text-[9px] text-[#4b6b5d]">Stable Environment</span>
            </div>
            <div className="text-[10px] text-[#10b981] font-mono bg-[#10b981]/10 px-3 py-1.5 rounded-full border border-[#10b981]/20 shadow-inner">
              V 2.1.0
            </div>
          </div>
          
        </div>
      </header>

      {/* المحتوى الرئيسي للتطبيق */}
      <main className="py-12 px-4 md:px-8">
        <div className="max-w-7xl mx-auto">
          
          {/* قسم الترحيب السريع */}
          <div className="mb-10 text-right">
            <h2 className="text-[#10b981] text-sm font-bold mb-2">أهلاً بك في وحدة المراجعة المالي</h2>
            <p className="text-gray-400 text-xs max-w-2xl leading-relaxed">
              قم برفع ملفات "Comsys" وملفات "Talabat Settlement" لبدء عملية التشريح الرباعي المتقدمة للضريبة والصافي وخدمة التوصيل.
            </p>
          </div>

          {/* استدعاء المحرك الرئيسي من المسار: ./utils/reconciliationEngine */}
          <section className="relative">
            <div className="absolute -top-20 -right-20 w-64 h-64 bg-[#10b981]/5 rounded-full blur-3xl pointer-events-none"></div>
            <ReconciliationEngine />
          </section>

        </div>
      </main>

      {/* Footer بسيط واحترافي */}
      <footer className="py-10 border-t border-[#1a3d2f] bg-[#010805]">
        <div className="max-w-7xl mx-auto px-6 flex flex-col md:flex-row justify-between items-center gap-4">
          <p className="text-[#4b6b5d] text-[11px] font-medium">
            &copy; {new Date().getFullYear()} Garnell Sushi & Poke - قسم تكنولوجيا المعلومات والمراجعة المالية
          </p>
          <div className="flex items-center gap-6">
            <span className="text-[10px] text-[#4b6b5d] hover:text-[#10b981] cursor-help transition-colors">Privacy Policy</span>
            <span className="text-[10px] text-[#4b6b5d] hover:text-[#10b981] cursor-help transition-colors">Internal Audit Tool</span>
          </div>
        </div>
      </footer>

      {/* لمسة جمالية خلفية (Glow effect) */}
      <div className="fixed bottom-0 left-0 w-full h-1 bg-gradient-to-r from-transparent via-[#10b981]/50 to-transparent shadow-[0_0_20px_rgba(16,185,129,0.5)]"></div>
    </div>
  );
}

export default App;