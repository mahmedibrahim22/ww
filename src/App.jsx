import React from 'react';
import { BrowserRouter, Routes, Route, NavLink, useLocation } from 'react-router-dom';
import ReconciliationEngine from './utils/reconciliationEngine';
import DailySalesReportFiller from './pages/DailySalesReportFiller';
import DiscountReconciliation from "./pages/DiscountReconciliation";
// ══════════════════════════════════════════════════════════════════════════
//  الهيدر الرئيسي + ناف بار التنقل (بنفس الاستايل الأصلي)
// ══════════════════════════════════════════════════════════════════════════
function Header() {
  const location = useLocation();

  const navLinks = [
    {
      to:    '/',
      label: 'مطابقة طلبات',
      icon: (
        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2"
            d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
        </svg>
      ),
    },
    {
      to:    '/daily',
      label: 'Daily Sales',
      icon: (
        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2"
            d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
        </svg>
      ),
    },
    {
      to:    '/discounts',
      label: 'الخصومات',
      icon: (
        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2"
            d="M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A2 2 0 013 12V7a4 4 0 014-4z" />
        </svg>
      ),
    },
  ];

  return (
    <header className="bg-[#0d2a1f]/80 border-b border-[#1a3d2f] py-4 sticky top-0 z-50 backdrop-blur-md shadow-lg">
      <div className="max-w-7xl mx-auto px-6 flex justify-between items-center">

        {/* الشعار - نفس الأصل */}
        <div className="flex items-center space-x-4 space-x-reverse">
          <div className="w-10 h-10 bg-[#10b981] rounded-xl flex items-center justify-center shadow-[0_0_20px_rgba(16,185,129,0.3)] group transition-transform hover:rotate-12">
            <svg className="w-6 h-6 text-[#020c08]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2"
                d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
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

        {/* اليمين: ناف + حالة النظام */}
        <div className="flex items-center gap-4">

          {/* ── ناف بار التنقل ─────────────────────────────────────── */}
          <nav className="flex items-center gap-1 bg-[#020c08]/60 border border-[#1a3d2f] rounded-xl p-1">
            {navLinks.map(link => {
              const isActive = link.to === '/'
                ? location.pathname === '/'
                : location.pathname.startsWith(link.to);
              return (
                <NavLink
                  key={link.to}
                  to={link.to}
                  end={link.to === '/'}
                  className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-bold transition-all duration-200 ${
                    isActive
                      ? 'bg-[#10b981]/20 text-[#10b981] border border-[#10b981]/30 shadow-inner'
                      : 'text-[#4b6b5d] hover:text-white hover:bg-white/5'
                  }`}
                >
                  {link.icon}
                  <span className="hidden sm:inline">{link.label}</span>
                </NavLink>
              );
            })}
          </nav>

          {/* حالة النظام - نفس الأصل */}
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
      </div>
    </header>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//  صفحة مطابقة طلبات (الصفحة الأولى - نفس الأصل)
// ══════════════════════════════════════════════════════════════════════════
function ReconciliationPage() {
  return (
    <main className="py-12 px-4 md:px-8">
      <div className="max-w-7xl mx-auto">

        {/* قسم الترحيب السريع - نفس الأصل */}
        <div className="mb-10 text-right">
          <h2 className="text-[#10b981] text-sm font-bold mb-2">أهلاً بك في وحدة المراجعة المالي</h2>
          <p className="text-gray-400 text-xs max-w-2xl leading-relaxed">
            قم برفع ملفات "Comsys" وملفات "Talabat Settlement" لبدء عملية التشريح الرباعي المتقدمة للضريبة والصافي وخدمة التوصيل.
          </p>
        </div>

        {/* المحرك الرئيسي - نفس الأصل */}
        <section className="relative">
          <div className="absolute -top-20 -right-20 w-64 h-64 bg-[#10b981]/5 rounded-full blur-3xl pointer-events-none"></div>
          <ReconciliationEngine />
        </section>

      </div>
    </main>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//  صفحة Daily Sales Report (الصفحة الثانية)
// ══════════════════════════════════════════════════════════════════════════
function DailySalesPage() {
  return (
    <main className="py-12 px-4 md:px-8">
      <div className="max-w-7xl mx-auto">

        <div className="mb-10 text-right">
          <h2 className="text-[#10b981] text-sm font-bold mb-2">وحدة تعبئة التقرير اليومي</h2>
          <p className="text-gray-400 text-xs max-w-2xl leading-relaxed">
            ارفع ملف Daily Revenue وSpot Check وسيتم استخراج بيانات كل الفروع وتعبئة الشيت تلقائياً.
          </p>
        </div>

        <section className="relative">
          <div className="absolute -top-20 -left-20 w-64 h-64 bg-[#10b981]/5 rounded-full blur-3xl pointer-events-none"></div>
          <DailySalesReportFiller />
        </section>

      </div>
    </main>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//  صفحة الخصومات والكمبلمنتري (الصفحة الثالثة)
// ══════════════════════════════════════════════════════════════════════════
function DiscountPage() {
  return (
    <main className="py-12 px-4 md:px-8">
      <div className="max-w-7xl mx-auto">

        <div className="mb-10 text-right">
          <h2 className="text-[#10b981] text-sm font-bold mb-2">وحدة مراجعة الخصومات والكمبلمنتري</h2>
          <p className="text-gray-400 text-xs max-w-2xl leading-relaxed">
            ارفع شيت الكاشير وملفات Discounted Items وComplimentary لمطابقة السيريالات وإضافة الأسباب تلقائياً.
          </p>
        </div>

        <section className="relative">
          <div className="absolute -top-20 -right-20 w-64 h-64 bg-[#10b981]/5 rounded-full blur-3xl pointer-events-none"></div>
          <DiscountReconciliation />
        </section>

      </div>
    </main>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//  الفوتر - نفس الأصل
// ══════════════════════════════════════════════════════════════════════════
function Footer() {
  return (
    <footer className="py-10 border-t border-[#1a3d2f] bg-[#010805]">
      <div className="max-w-7xl mx-auto px-6 flex flex-col md:flex-row justify-between items-center gap-4">
        <p className="text-[#4b6b5d] text-[11px] font-medium">
          &copy; {new Date().getFullYear()} Garnell Sushi &amp; Poke - قسم تكنولوجيا المعلومات والمراجعة المالية
        </p>
        <div className="flex items-center gap-6">
          <span className="text-[10px] text-[#4b6b5d] hover:text-[#10b981] cursor-help transition-colors">Privacy Policy</span>
          <span className="text-[10px] text-[#4b6b5d] hover:text-[#10b981] cursor-help transition-colors">Internal Audit Tool</span>
        </div>
      </div>
    </footer>
  );
}

// ══════════════════════════════════════════════════════════════════════════
//  التطبيق الرئيسي
// ══════════════════════════════════════════════════════════════════════════
function App() {
  return (
    <BrowserRouter>
      <div className="min-h-screen bg-[#020c08] text-gray-200 font-sans selection:bg-[#10b981]/30" dir="rtl">

        <Header />

        <Routes>
          <Route path="/"          element={<ReconciliationPage />} />
          <Route path="/daily"     element={<DailySalesPage />} />
          <Route path="/discounts" element={<DiscountPage />} />
        </Routes>

        <Footer />

        {/* لمسة جمالية خلفية - نفس الأصل */}
        <div className="fixed bottom-0 left-0 w-full h-1 bg-gradient-to-r from-transparent via-[#10b981]/50 to-transparent shadow-[0_0_20px_rgba(16,185,129,0.5)]"></div>

      </div>
    </BrowserRouter>
  );
}

export default App;