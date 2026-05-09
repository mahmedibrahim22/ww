import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

export default function FileUpload({ onUpload }) {
  const [talabatFile, setTalabatFile] = useState(null);
  const [checksFile, setChecksFile] = useState(null);
  const [summaryFile, setSummaryFile] = useState(null); // الحالة للملف المجمع
  
  const talabatInputRef = useRef(null);
  const checksInputRef = useRef(null);
  const summaryInputRef = useRef(null); // المرجع للملف المجمع

  const processFile = (file, type) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      const content = e.target.result;

      if (file.name.endsWith('.csv')) {
        Papa.parse(content, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
            onUpload(results.data, type);
          }
        });
      } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        onUpload(json, type);
      }
    };

    if (file.name.endsWith('.csv')) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file); // تم التعديل لـ ArrayBuffer لضمان توافق Excel
    }
  };

  const handleTalabatChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      setTalabatFile(file.name);
      processFile(file, 'talabat');
    }
  };

  const handleChecksChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      setChecksFile(file.name);
      processFile(file, 'check'); // تم توحيد النوع ليتوافق مع الـ Engine
    }
  };

  const handleSummaryChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      setSummaryFile(file.name);
      processFile(file, 'summaryFile');
    }
  };

  return (
    <div className="bg-gray-950 text-gray-300 p-6 rounded-xl border border-gray-900 max-w-6xl mx-auto shadow-2xl font-sans">
      <div className="mb-8 text-right">
        <h1 className="text-2xl font-bold text-[#10b981] mb-2">
          مركز رفع بيانات المراجعة
        </h1>
        <p className="text-gray-500 text-sm">
          قم برفع الملفات الثلاثة المطلوبة لمطابقة بيانات طلبات، كومسيس، والملف المجمع.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        
        {/* 1. رفع شيت الشيكات (كومسيس) */}
        <div className="bg-gray-900/60 border border-emerald-900/30 rounded-lg p-6 flex flex-col items-center justify-center text-center hover:border-emerald-500/50 transition-colors">
          <div className="w-12 h-12 mb-4 bg-emerald-950/30 rounded-full flex items-center justify-center text-emerald-400 border border-emerald-900/50">
            <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
            </svg>
          </div>
          <h2 className="text-sm font-semibold text-gray-200 mb-1">شيت الشيكات (Comsys)</h2>
          <p className="text-xs text-gray-500 mb-4">ملف Check Details</p>
          <input type="file" ref={checksInputRef} className="hidden" accept=".csv, .xlsx, .xls" onChange={handleChecksChange} />
          <button onClick={() => checksInputRef.current.click()} className="w-full py-2 bg-emerald-950/40 border border-emerald-800/50 text-emerald-400 text-xs rounded-md hover:bg-emerald-900/60 transition duration-150">
            {checksFile ? checksFile : "اختر ملف الشيكات"}
          </button>
        </div>

        {/* 2. رفع شيت الطلبات (Talabat) */}
        <div className="bg-gray-900/60 border border-sky-900/30 rounded-lg p-6 flex flex-col items-center justify-center text-center hover:border-sky-500/50 transition-colors">
          <div className="w-12 h-12 mb-4 bg-sky-950/30 rounded-full flex items-center justify-center text-sky-400 border border-sky-900/50">
            <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <h2 className="text-sm font-semibold text-gray-200 mb-1">شيت الطلبات (Talabat)</h2>
          <p className="text-xs text-gray-500 mb-4">ملف Order Details</p>
          <input type="file" ref={talabatInputRef} className="hidden" accept=".csv, .xlsx, .xls" onChange={handleTalabatChange} />
          <button onClick={() => talabatInputRef.current.click()} className="w-full py-2 bg-sky-950/40 border border-sky-800/50 text-sky-400 text-xs rounded-md hover:bg-sky-900/60 transition duration-150">
            {talabatFile ? talabatFile : "اختر ملف الطلبات"}
          </button>
        </div>

        {/* 3. رفع الملف المجمع */}
        <div className="bg-gray-900/60 border border-amber-900/30 rounded-lg p-6 flex flex-col items-center justify-center text-center hover:border-amber-500/50 transition-colors">
          <div className="w-12 h-12 mb-4 bg-amber-950/30 rounded-full flex items-center justify-center text-amber-400 border border-amber-900/50">
            <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
            </svg>
          </div>
          <h2 className="text-sm font-semibold text-gray-200 mb-1">الملف المجمع</h2>
          <p className="text-xs text-gray-500 mb-4">ملف Excel اليومي المجمع</p>
          <input type="file" ref={summaryInputRef} className="hidden" accept=".csv, .xlsx, .xls" onChange={handleSummaryChange} />
          <button onClick={() => summaryInputRef.current.click()} className="w-full py-2 bg-amber-950/40 border border-amber-800/50 text-amber-400 text-xs rounded-md hover:bg-amber-900/60 transition duration-150">
            {summaryFile ? summaryFile : "اختر الملف المجمع"}
          </button>
        </div>

      </div>

      {/* تنبيه أمان البيانات */}
      <div className="mt-8 bg-gray-900/40 border border-gray-800 p-4 rounded-lg flex items-start space-x-3 space-x-reverse text-right">
        <div className="text-gray-500 mt-0.5 ml-3">
          <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
        </div>
        <div className="flex-1">
          <p className="text-gray-500 text-xs leading-relaxed">
            تتم عملية المعالجة بالكامل داخل المتصفح. تذكر مطابقة اسم الفرع في شيت الشيكات مع شيت طلبات لضمان دقة التحليل.
          </p>
        </div>
      </div>
    </div>
  );
}