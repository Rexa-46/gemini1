import React, { useState, useEffect } from 'react';
import { 
  Plus, Minus, ChevronLeft, ChevronRight, ChevronUp, ChevronDown,
  Zap, Droplet, DollarSign, Car, ShoppingBag, Shirt, Wifi, CreditCard,
  Building, Gamepad2, Fuel, Utensils, Bell, Wallet, Home, PlusCircle, Wrench
} from 'lucide-react';

// لیست میانبرهای پیش‌فرض با آیکن و رنگ‌بندی شبیه تصویر
const DEFAULT_SHORTCUTS = [
  { id: '1', title: 'قبض برق', icon: Home, bg: 'bg-amber-50', color: 'text-amber-700' },
  { id: '2', title: 'قبض آب', icon: Home, bg: 'bg-amber-50', color: 'text-amber-700' },
  { id: '3', title: 'حقوق ماهیانه', icon: DollarSign, bg: 'bg-cyan-50', color: 'text-cyan-600' },
  { id: '4', title: 'کرایه تاکسی', icon: Car, bg: 'bg-amber-100', color: 'text-amber-600' },
  { id: '5', title: 'سوپر مارکت', icon: Utensils, bg: 'bg-amber-100', color: 'text-amber-600' },
  { id: '6', title: 'خرید پوشاک', icon: Shirt, bg: 'bg-cyan-50', color: 'text-cyan-600' },
  { id: '7', title: 'اینترنت', icon: Wifi, bg: 'bg-pink-50', color: 'text-pink-500' },
  { id: '8', title: 'کارمزد بانک', icon: Building, bg: 'bg-purple-50', color: 'text-purple-600' },
  { id: '9', title: 'وام بانکی', icon: DollarSign, bg: 'bg-pink-50', color: 'text-red-500' },
  { id: '10', title: 'بازی و سرگرمی', icon: Gamepad2, bg: 'bg-pink-50', color: 'text-pink-500' },
  { id: '11', title: 'بنزین', icon: Car, bg: 'bg-amber-100', color: 'text-amber-600' },
  { id: '12', title: 'غذای بیرون', icon: Utensils, bg: 'bg-amber-100', color: 'text-amber-600' },
  { id: '13', title: 'تعمیرات ساختمان', icon: Home, bg: 'bg-amber-100', color: 'text-amber-700' },
  { id: '14', title: 'مهر (شهدا) ۶۴۱', sub: '۹۷۰۰۵۰۴۱۱۹۱۱', icon: Wallet, bg: 'bg-green-100', color: 'text-green-600' }
];

export default function App() {
  // ---------------- کلیدها و حالت‌های برنامه ----------------
  const [shortcuts, setShortcuts] = useState(() => {
    try {
      const saved = localStorage.getItem('app_shortcuts');
      return saved ? JSON.parse(saved) : DEFAULT_SHORTCUTS;
    } catch (e) {
      return DEFAULT_SHORTCUTS;
    }
  });

  const [reminders, setReminders] = useState(() => {
    try {
      const saved = localStorage.getItem('app_reminders');
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });

  // مودال ثبت تراکنش سریع با میانبر
  const [selectedShortcut, setSelectedShortcut] = useState(null);
  const [amount, setAmount] = useState('');
  const [selectedBank, setSelectedBank] = useState('بانک ملی');
  
  // مودال افزودن میانبر جدید
  const [showAddShortcut, setShowAddShortcut] = useState(null);
  const [newShortcutTitle, setNewShortcutTitle] = useState('');

  // ---------------- ذخیره‌سازی داده‌ها ----------------
  useEffect(() => {
    try {
      localStorage.setItem('app_shortcuts', JSON.stringify(shortcuts));
    } catch (e) {
      console.error('Error saving shortcuts', e);
    }
  }, [shortcuts]);

  useEffect(() => {
    try {
      localStorage.setItem('app_reminders', JSON.stringify(reminders));
    } catch (e) {
      console.error('Error saving reminders', e);
    }
  }, [reminders]);

  // کلیک روی میانبر جهت ثبت تراکنش سریع
  const handleShortcutClick = (item) => {
    setSelectedShortcut(item);
    setAmount('');
  };

  // افزودن میانبر جدید
  const handleAddShortcut = () => {
    if (!newShortcutTitle.trim()) return;
    const newItem = {
      id: Date.now().toString(),
      title: newShortcutTitle,
      icon: PlusCircle,
      bg: 'bg-cyan-50',
      color: 'text-cyan-600'
    };
    setShortcuts([...shortcuts, newItem]);
    setNewShortcutTitle('');
    setShowAddShortcut(false);
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 pb-20 font-vazir dir-rtl">
      {/* ---------------- ۱. کادر نمودار و درآمد/هزینه ---------------- */}
      <div className="p-4 max-w-md mx-auto">
        <div className="bg-white rounded-3xl p-5 shadow-sm border border-slate-100">
          {/* هدر تاریخ سررسیدها */}
          <div className="flex items-center justify-between mb-4">
            <button className="p-1 rounded-full text-slate-400 hover:bg-slate-100">
              <ChevronLeft className="w-5 h-5" />
            </button>
            <div className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-purple-600"></span>
              <h2 className="text-sm font-bold text-slate-700">سررسیدهای، شنبه ۱۶ فروردین</h2>
            </div>
            <button className="p-1 rounded-full text-slate-400 hover:bg-slate-100">
              <ChevronRight className="w-5 h-5" />
            </button>
          </div>

          {/* نمودار دایره‌ای (Pie Chart) دقیقا شبیه تصویر */}
          <div className="relative w-56 h-56 mx-auto my-3 flex items-center justify-center">
            <svg viewBox="0 0 100 100" className="w-full h-full transform -rotate-90">
              {/* بخش آبی */}
              <circle cx="50" cy="50" r="38" fill="transparent" stroke="#00b0ff" strokeWidth="22" strokeDasharray="60 180" strokeDashoffset="0" />
              {/* بخش قرمز */}
              <circle cx="50" cy="50" r="38" fill="transparent" stroke="#ff3d00" strokeWidth="22" strokeDasharray="75 180" strokeDashoffset="-62" />
              {/* بخش بنفش */}
              <circle cx="50" cy="50" r="38" fill="transparent" stroke="#7c4dff" strokeWidth="22" strokeDasharray="45 180" strokeDashoffset="-139" />
              {/* بخش سبز */}
              <circle cx="50" cy="50" r="38" fill="transparent" stroke="#76ff03" strokeWidth="22" strokeDasharray="50 180" strokeDashoffset="-186" />
            </svg>
            {/* مرکز سفید نمودار */}
            <div className="absolute w-12 h-12 bg-white rounded-full border-4 border-red-100 shadow-inner"></div>
          </div>

          {/* کادرهای جمع درآمد و جمع هزینه (کاملا استاندارد و بدون تداخل) */}
          <div className="grid grid-cols-2 gap-3 mt-5">
            {/* کادر جمع درآمدها */}
            <div className="flex items-center justify-between p-3 rounded-2xl border-2 border-slate-300 bg-white">
              <div className="flex items-center gap-1.5">
                <div className="w-7 h-7 rounded-full bg-slate-100 text-slate-500 flex items-center justify-center">
                  <Plus className="w-4 h-4" />
                </div>
                <div>
                  <span className="text-[11px] text-slate-400 block font-medium">جمع درآمدها</span>
                  <span className="text-sm font-black text-slate-800">۳۸,۰۰۰,۰۰۰</span>
                </div>
              </div>
              <span className="text-[10px] text-slate-400 -rotate-90 font-bold mr-[-6px]">تومان</span>
            </div>

            {/* کادر جمع هزینه‌ها */}
            <div className="flex items-center justify-between p-3 rounded-2xl border-2 border-red-400 bg-white">
              <div className="flex items-center gap-1.5">
                <div className="w-7 h-7 rounded-full bg-red-100 text-red-500 flex items-center justify-center">
                  <Minus className="w-4 h-4" />
                </div>
                <div>
                  <span className="text-[11px] text-red-400 block font-medium">جمع هزینه‌ها</span>
                  <span className="text-sm font-black text-red-500">۱۶,۶۵۰,۰۰۰</span>
                </div>
              </div>
              <span className="text-[10px] text-red-400 -rotate-90 font-bold mr-[-6px]">تومان</span>
            </div>
          </div>
        </div>

        {/* ---------------- ۲. کارت بانکی (۱۶ رقم کامل + تاریخ انقضا سمت راست پایین) ---------------- */}
        <div className="mt-4 bg-gradient-to-r from-emerald-600 to-teal-800 text-white p-5 rounded-3xl shadow-md relative overflow-hidden h-44 flex flex-col justify-between">
          <div className="flex justify-between items-center">
            <span className="text-xs font-semibold opacity-80">بانک پاسارگاد</span>
            <CreditCard className="w-6 h-6 opacity-80" />
          </div>

          {/* شماره کارت کامل ۱۶ رقمی */}
          <div className="text-center my-auto">
            <span className="text-lg tracking-[0.25em] font-mono font-bold block dir-ltr">
              6037 9918 1234 5678
            </span>
          </div>

          {/* تاریخ انقضا دقیقاً سمت راست پایین */}
          <div className="flex justify-between items-end w-full">
            <span className="text-xs font-medium opacity-90">رضا شهدفر</span>
            <div className="text-right">
              <span className="text-[9px] block opacity-75">تاریخ انقضا</span>
              <span className="text-xs font-mono font-bold">05/28</span>
            </div>
          </div>
        </div>

        {/* ---------------- ۳. میانبر تراکنش‌ها (دقیقاً شبیه تصویر) ---------------- */}
        <div className="mt-4 bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
          {/* هدر میانبرها */}
          <div className="bg-rose-900/70 text-white p-3 px-4 flex justify-between items-center">
            <span className="text-sm font-bold">میانبر تراکنش ها</span>
            <ChevronUp className="w-5 h-5 cursor-pointer" />
          </div>

          {/* شبکه دکمه‌های میانبر (۴ ستونه) */}
          <div className="p-4 grid grid-cols-4 gap-3 text-center">
            {shortcuts.map((item) => {
              const IconComp = item.icon || ShoppingBag;
              return (
                <button
                  key={item.id}
                  onClick={() => handleShortcutClick(item)}
                  className="flex flex-col items-center group active:scale-95 transition-transform"
                >
                  <div className={`w-14 h-14 rounded-2xl ${item.bg} flex items-center justify-center mb-1.5 shadow-sm`}>
                    <IconComp className={`w-7 h-7 ${item.color}`} />
                  </div>
                  <span className="text-[11px] font-bold text-slate-700 leading-tight block">{item.title}</span>
                  {item.sub && <span className="text-[9px] text-slate-400 block mt-0.5">{item.sub}</span>}
                </button>
              );
            })}

            {/* دکمه افزودن میانبر جدید */}
            <button
              onClick={() => setShowAddShortcut(true)}
              className="flex flex-col items-center active:scale-95 transition-transform"
            >
              <div className="w-14 h-14 rounded-2xl bg-cyan-50 flex items-center justify-center mb-1.5 shadow-sm">
                <Plus className="w-7 h-7 text-cyan-600" />
              </div>
              <span className="text-[11px] font-bold text-slate-700">افزودن میانبر</span>
            </button>
          </div>
        </div>

        {/* ---------------- ۴. بخش یادآوری (ایمن‌سازی‌شده بدون کرش/صفحه سفید) ---------------- */}
        <div className="mt-4 bg-white rounded-3xl p-4 shadow-sm border border-slate-100">
          <div className="flex items-center gap-2 mb-3">
            <Bell className="w-5 h-5 text-amber-500" />
            <h3 className="text-sm font-bold text-slate-700">یادآوری‌های فعال</h3>
          </div>

          {/* چک کردن امن داده‌ها جهت جلوگیری از سفید شدن صفحه */}
          {(!reminders || !Array.isArray(reminders) || reminders.length === 0) ? (
            <div className="text-center py-6 text-slate-400 text-xs bg-slate-50 rounded-2xl border border-dashed border-slate-200">
              هیچ یادآوری برای این دوره ثبت نشده است.
            </div>
          ) : (
            <div className="space-y-2">
              {reminders.map((rem, idx) => (
                <div key={rem?.id || idx} className="flex justify-between items-center p-2.5 bg-slate-50 rounded-xl">
                  <span className="text-xs font-medium text-slate-700">{rem?.title || 'یادآوری بدون عنوان'}</span>
                  <span className="text-[10px] text-slate-400">{rem?.date || 'امروز'}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* ---------------- ۵. منوی پایین صفحه (ثابت و مقاوم در برابر باز شدن کیبورد) ---------------- */}
      <nav className="fixed bottom-0 inset-x-0 bg-white border-t border-slate-200 py-2.5 px-6 flex justify-around items-center z-50 h-16 max-w-md mx-auto">
        <button className="flex flex-col items-center text-rose-900 font-bold">
          <Home className="w-5 h-5" />
          <span className="text-[10px] mt-1">خانه</span>
        </button>
        <button className="flex flex-col items-center text-slate-400 hover:text-slate-600">
          <Wallet className="w-5 h-5" />
          <span className="text-[10px] mt-1">کیف پول</span>
        </button>
        <button className="flex flex-col items-center text-slate-400 hover:text-slate-600">
          <Bell className="w-5 h-5" />
          <span className="text-[10px] mt-1">یادآوری</span>
        </button>
      </nav>

      {/* ---------------- مودال ثبت تراکنش سریع میانبر ---------------- */}
      {selectedShortcut && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-end sm:items-center justify-center p-4">
          <div className="bg-white rounded-3xl p-5 w-full max-w-sm animate-in fade-in slide-in-from-bottom">
            <h3 className="text-base font-bold text-slate-800 mb-2 flex items-center gap-2">
              ثبت تراکنش برای: <span className="text-rose-900">{selectedShortcut.title}</span>
            </h3>
            
            <div className="mt-4 space-y-3">
              <div>
                <label className="text-xs text-slate-500 font-medium block mb-1">مبلغ (تومان)</label>
                <input 
                  type="number" 
                  value={amount}
                  onChange={(e) => setAmount(e.target.value)}
                  placeholder="مثلا ۵۰,۰۰۰"
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm outline-none focus:border-rose-800"
                />
              </div>

              <div>
                <label className="text-xs text-slate-500 font-medium block mb-1">انتخاب کارت / بانک</label>
                <select 
                  value={selectedBank} 
                  onChange={(e) => setSelectedBank(e.target.value)}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm outline-none"
                >
                  <option value="بانک ملی">بانک ملی (۶۰۳۷)</option>
                  <option value="بانک پاسارگاد">بانک پاسارگاد (۵۰۲۲)</option>
                  <option value="بانک ملت">بانک ملت (۶۱۰۴)</option>
                </select>
              </div>

              <div className="flex gap-2 pt-2">
                <button 
                  onClick={() => {
                    alert(`تراکنش ${selectedShortcut.title} به مبلغ ${amount} تومان با موفقیت ثبت شد.`);
                    setSelectedShortcut(null);
                  }}
                  className="flex-1 py-3 bg-rose-900 text-white rounded-2xl text-sm font-bold hover:bg-rose-950 active:scale-98 transition-all"
                >
                  ثبت تراکنش
                </button>
                <button 
                  onClick={() => setSelectedShortcut(null)}
                  className="px-4 py-3 bg-slate-100 text-slate-600 rounded-2xl text-sm font-bold"
                >
                  انصراف
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ---------------- مودال افزودن میانبر جدید ---------------- */}
      {showAddShortcut && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl p-5 w-full max-w-sm">
            <h3 className="text-sm font-bold text-slate-800 mb-3">افزودن میانبر جدید</h3>
            <input 
              type="text" 
              value={newShortcutTitle}
              onChange={(e) => setNewShortcutTitle(e.target.value)}
              placeholder="عنوان میانبر (مثلا: خرید میوه)"
              className="w-full p-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm outline-none mb-4"
            />
            <div className="flex gap-2">
              <button 
                onClick={handleAddShortcut}
                className="flex-1 py-3 bg-cyan-600 text-white rounded-2xl text-sm font-bold"
              >
                ذخیره میانبر
              </button>
              <button 
                onClick={() => setShowAddShortcut(false)}
                className="px-4 py-3 bg-slate-100 text-slate-600 rounded-2xl text-sm font-bold"
              >
                انصراف
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
