import React from 'react';
import { NavLink } from 'react-router-dom';

export default function NavBar() {
  const links = [
    { to: '/',       label: 'مطابقة طلبات',      icon: '🔄' },
    { to: '/daily',  label: 'Daily Sales Report', icon: '📊' },
  ];

  return (
    <nav className="sticky top-0 z-50 bg-[#020c08] border-b border-emerald-900/40 shadow-lg shadow-black/30">
      <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between" dir="rtl">
        {/* Logo */}
        <div className="flex items-center gap-2">
          <span className="text-2xl">🍣</span>
          <span className="text-[#10b981] font-black text-lg tracking-tight">Garnell</span>
          <span className="text-gray-500 text-xs font-medium mt-1">Financial System</span>
        </div>

        {/* Links */}
        <div className="flex items-center gap-2">
          {links.map(link => (
            <NavLink
              key={link.to}
              to={link.to}
              end
              className={({ isActive }) =>
                `flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all duration-200 ${
                  isActive
                    ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30'
                    : 'text-gray-400 hover:text-white hover:bg-white/5'
                }`
              }
            >
              <span>{link.icon}</span>
              <span>{link.label}</span>
            </NavLink>
          ))}
        </div>
      </div>
    </nav>
  );
}