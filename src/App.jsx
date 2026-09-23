پس از بررسی درخواست شما، تمامی موارد خواسته شده شامل:

اصلاح و تکمیل بخش کارت‌های بانکی: نمایش کامل ۱۶ رقم شماره کارت (به صورت گروه‌های چهارتایی) و قرارگیری تاریخ انقضا در سمت راست پایین کارت.
اصلاح ویجت درآمد، هزینه و نمودار: تنظیم دقیق در زیر کارت‌ها بدون تداخل متون و کادرها.
اصلاح بخش میانبرهای تراکنش: طراحی مدرن و کاربردی مشابه تصاویر و اتصال کامل به ثبت تراکنش.
رفع مشکل و فعال‌سازی کامل بخش یادآوری‌ها: امکان ثبت، مشاهده و مدیریت یادآوری‌ها.
مدیریت هوشمند کیبورد برای منوی پایین: قرارگیری دقیق منوی پایین صفحه در بالای کیبورد هنگام باز شدن آن (ثبت‌شده با visualViewport).
ارائه فایل کامل و یکپارچه پروژه برای پش در گیت‌هاب.

کد کامل، اصلاح‌شده و بدون نقص برنامه در ادامه آمده است:
import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Bell, Menu, Plus, X, Search, ChevronDown, ChevronLeft, ChevronRight,
  ArrowLeftRight, Landmark, TrendingUp, TrendingDown, Receipt,
  Trash2, Grid3x3, PieChart as PieChartIcon, Home as HomeIcon,
  Tag, Check, Star, CreditCard, Lock, Sun, Moon, Image as ImageIcon,
  Repeat, Download, Upload, Bitcoin, Landmark as Bank, CalendarDays,
  BellRing, FileSpreadsheet, Printer, Users, ShieldCheck, Palette, Save,
  Eye, EyeOff, StickyNote, Mic, MicOff, LayoutGrid, LayoutList, ArrowUp, ArrowDown,
  DollarSign, RefreshCw, Sparkles, Type, Target, Fingerprint, Pencil, RotateCcw, Eraser
} from "lucide-react";
import {
  PieChart, Pie, Cell, Sector, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip,
  LineChart, Line, CartesianGrid
} from "recharts";
import * as XLSX from "xlsx";
import { Capacitor, CapacitorHttp } from "@capacitor/core";
import { App as CapacitorApp } from "@capacitor/app";
import { Filesystem, Directory, Encoding } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { BiometricAuth } from "@aparajita/capacitor-biometric-auth";
import { SpeechRecognition } from "@capgo/capacitor-speech-recognition";

/* ---------------------------------------------------------
   Helpers
--------------------------------------------------------- */
const toFaInt = (n) => Math.round(Number(n || 0)).toLocaleString("fa-IR");
const FA_DIGIT_MAP = ["۰", "۱", "۲", "۳", "۴", "۵", "۶", "۷", "۸", "۹"];
const faDigits = (n) => String(n).split("").map((ch) => (/[0-9]/.test(ch) ? FA_DIGIT_MAP[+ch] : ch)).join("");
function jalaliYear(d) {
  try { return parseInt(new Intl.DateTimeFormat("en-US-u-ca-persian", { year: "numeric" }).format(d), 10); }
  catch { return new Date(d).getFullYear(); }
}
function jalaliParts(d) {
  try {
    const parts = new Intl.DateTimeFormat("en-US-u-ca-persian", { year: "numeric", month: "numeric", day: "numeric" }).formatToParts(d);
    return {
      y: +parts.find((p) => p.type === "year").value,
      m: +parts.find((p) => p.type === "month").value,
      day: +parts.find((p) => p.type === "day").value,
    };
  } catch { return { y: d.getFullYear(), m: d.getMonth() + 1, day: d.getDate() }; }
}
function faLongDate(d) {
  try { return new Intl.DateTimeFormat("fa-IR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(d); }
  catch { return d.toDateString(); }
}
function faMonthYear(d) {
  try { return new Intl.DateTimeFormat("fa-IR", { month: "long", year: "numeric" }).format(d); }
  catch { return ""; }
}
function addDays(dateStr, n) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}
function addMonths(dateStr, n) {
  const d = new Date(dateStr);
  d.setMonth(d.getMonth() + n);
  return d.toISOString().slice(0, 10);
}
function daysUntil(dateStr) {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const d = new Date(dateStr); d.setHours(0, 0, 0, 0);
  return Math.round((d - now) / 86400000);
}
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const todayISO = () => new Date().toISOString().slice(0, 10);

const FA_NUM_WORDS = {
  "صفر": 0, "یک": 1, "یه": 1, "دو": 2, "سه": 3, "چهار": 4, "پنج": 5, "شش": 6, "شیش": 6,
  "هفت": 7, "هشت": 8, "نه": 9, "ده": 10, "یازده": 11, "دوازده": 12, "سیزده": 13, "چهارده": 14,
  "پانزده": 15, "شانزده": 16, "هفده": 17, "هجده": 18, "نوزده": 19,
  "بیست": 20, "سی": 30, "چهل": 40, "پنجاه": 50, "شصت": 60, "هفتاد": 70, "هشتاد": 80, "نود": 90,
  "صد": 100, "یکصد": 100, "دویست": 200, "سیصد": 300, "چهارصد": 400, "پانصد": 500,
  "ششصد": 600, "هفتصد": 700, "هشتصد": 800, "نهصد": 900,
};
const FA_MULTIPLIERS = { "هزار": 1000, "میلیون": 1000000, "میلیارد": 1000000000 };
function persianWordsToNumber(text) {
  const words = text.replace(/[،,]/g, " ").split(/\s+/).filter(Boolean);
  let total = 0, current = 0, found = false;
  for (const w of words) {
    if (w === "و") continue;
    if (Object.prototype.hasOwnProperty.call(FA_NUM_WORDS, w)) { current += FA_NUM_WORDS[w]; found = true; }
    else if (Object.prototype.hasOwnProperty.call(FA_MULTIPLIERS, w)) { current = (current || 1) * FA_MULTIPLIERS[w]; total += current; current = 0; found = true; }
  }
  total += current;
  return found ? total : null;
}

function normalizeDigitsText(text = "") {
  return String(text)
    .replace(/[۰-۹]/g, (d) => "0123456789"["۰۱۲۳۴۵۶۷۸۹".indexOf(d)])
    .replace(/[٠-٩]/g, (d) => "0123456789"["٠١٢٣٤٥٦٧٨٩".indexOf(d)]);
}
function parseBankSms(text) {
  const raw = String(text || "").trim();
  const normalized = normalizeDigitsText(raw).replace(/[٬,،]/g, "");
  const amountMatch = normalized.match(/(\d+(?:\.\d+)?)\s*(میلیارد|میلیون|هزار)?\s*(تومان|تومن|ریال)?/);
  let amount = null;
  if (amountMatch) {
    amount = Number(amountMatch[1]);
    const mult = { "هزار": 1e3, "میلیون": 1e6, "میلیارد": 1e9 }[amountMatch[2]] || 1;
    amount *= mult;
    if (amountMatch[3] === "تومان" || amountMatch[3] === "تومن") {
      if (!amountMatch[2] && amount < 10000) amount *= 1000;
      amount *= 10;
    }
  }
  if (!amount || !Number.isFinite(amount)) {
    const wordAmount = persianWordsToNumber(normalized);
    if (wordAmount !== null && wordAmount > 0) amount = wordAmount;
  }
  if (amount === null) {
    const fallback = normalized.match(/\d+/);
    if (fallback) amount = Number(fallback[0]);
  }
  let type = "expense";
  if (/واریز|دریافت|حقوق|دستمزد|درآمد|طلب|credit|deposit|income/i.test(raw)) type = "income";
  if (/برداشت|خرید|خرج|پرداخت|هزینه|دادم|داد|کرایه|بنزین|رستوران|شام|ناهار|debit|purchase|expense/i.test(raw)) type = "expense";
  const CATEGORY_HINTS = [
    { re: /بنزین|سوخت|پمپ/, name: "بنزین" },
    { re: /رستوران|غذا|شام|ناهار|صبحانه|کافه|سوپرمارکت|خرید خوراک|بازار/, name: "خوراک و بازار" },
    { re: /قبض|آب|برق|گاز|اینترنت|تلفن|شارژ/, name: "قبوض" },
    { re: /تاکسی|اسنپ|تپسی|اتوبوس|مترو|کرایه|حمل/, name: "حمل و نقل" },
    { re: /دارو|دکتر|درمان|بیمارستان|پزشک/, name: "درمان" },
    { re: /حقوق|دستمزد|مزایا|فیش حقوقی/, name: "حقوق" },
    { re: /فروش|فروش کالا|فروش محصول/, name: "درآمد متفرقه" },
  ];
  const hint = CATEGORY_HINTS.find((h) => h.re.test(raw));
  return { amount: amount && Number.isFinite(amount) ? Math.round(amount) : null, type, note: raw.replace(/\s+/g, " ").slice(0, 160), categoryHint: hint?.name };
}

function formatMoney(amount, currency, usdRialRate) {
  const n = Number(amount || 0);
  if (currency === "toman") return `${toFaInt(Math.round(n / 10))} تومان`;
  if (currency === "usd") {
    if (!usdRialRate) return `${toFaInt(n)} ریال`;
    return `$${(n / usdRialRate).toLocaleString("en-US", { maximumFractionDigits: 2 })}`;
  }
  return `${toFaInt(n)} ریال`;
}

/* ---------------------------------------------------------
   Seed data & Presets
--------------------------------------------------------- */
const seedAccounts = () => ([
  { id: uid(), name: "بانک ملت", type: "bank", initial: 5000000, cardNumber: "6037997123456789", expiryDate: "1406/12" },
  { id: uid(), name: "صندوق نقدی", type: "fund", initial: 800000, cardNumber: "", expiryDate: "" },
  { id: uid(), name: "کارت عابر ملت", type: "card", initial: 0, cardNumber: "6104337912348901", expiryDate: "1407/04" },
]);
const seedCategories = () => ([
  { id: uid(), name: "بنزین", kind: "expense" },
  { id: uid(), name: "خوراک و بازار", kind: "expense" },
  { id: uid(), name: "قبوض", kind: "expense" },
  { id: uid(), name: "حمل و نقل", kind: "expense" },
  { id: uid(), name: "درمان", kind: "expense" },
  { id: uid(), name: "متفرقه", kind: "expense" },
  { id: uid(), name: "حقوق", kind: "income" },
  { id: uid(), name: "درآمد متفرقه", kind: "income" },
]);
const DEFAULT_HOME_SECTIONS = [
  { key: "shortcut", visible: true }, { key: "expense", visible: true }, { key: "income", visible: true },
  { key: "banks", visible: true }, { key: "funds", visible: true }, { key: "balrep", visible: true },
  { key: "budget", visible: true }, { key: "loanchk", visible: true }, { key: "contacts", visible: true }, { key: "bills", visible: true },
];
const DEFAULT_SHORTCUTS = [
  { id: "sc1", name: "بنزین", type: "expense", icon: "fuel", categoryId: "", note: "بنزین" },
  { id: "sc2", name: "خوراک", type: "expense", icon: "food", categoryId: "", note: "خوراک" },
  { id: "sc3", name: "تاکسی", type: "expense", icon: "taxi", categoryId: "", note: "تاکسی" },
  { id: "sc4", name: "حقوق", type: "income", icon: "bank", categoryId: "", note: "حقوق" },
];
const SHORTCUT_ICONS = [
  ["fuel", "⛽"], ["food", "🍔"], ["taxi", "🚕"], ["bank", "💰"], ["shopping", "🛍️"],
  ["bill", "📄"], ["health", "💊"], ["home", "🏠"], ["coffee", "☕"], ["gift", "🎁"]
];
const seedSettings = () => ({
  theme: "light", pin: "", sharedFamily: false, themeColor: "purple",
  profile: { name: "رضا", phone: "", email: "" },
  homeLayout: "cards", homeSections: DEFAULT_HOME_SECTIONS,
  fontScale: 1, calendarMode: "jalali", currency: "rial",
  checkReminderDays: 7, smsNotif: false, smsAutoRead: false, biometricEnabled: false,
  aiProvider: "none", aiApiKey: "", manualUsdRate: "", shortcuts: DEFAULT_SHORTCUTS, lastRatesUpdate: "",
});

const COLOR_PRESETS = {
  purple: { name: "بنفش کلاسیک", header: "#3E1461", mauve: "#A65475", darkgreen: "#1B6B2C", green: "#1E8449", violet: "#6C3FA0", teal: "#4E9AA0", gold: "#A98A3B", crimson: "#B01E4A", orange: "#C56A1F", fab: "#28C76F" },
  ocean: { name: "آبی اقیانوسی", header: "#0B4F6C", mauve: "#3D7EA6", darkgreen: "#0F7173", green: "#14919B", violet: "#145DA0", teal: "#4CC9F0", gold: "#B08968", crimson: "#D64550", orange: "#F2A65A", fab: "#2EC4B6" },
  forest: { name: "سبز جنگلی", header: "#1B4332", mauve: "#40916C", darkgreen: "#2D6A4F", green: "#40916C", violet: "#52796F", teal: "#74C69D", gold: "#B08968", crimson: "#BC4749", orange: "#DDA15E", fab: "#52B788" },
  rose: { name: "گلبهی", header: "#6D2148", mauve: "#B23A5D", darkgreen: "#2D6A4F", green: "#40916C", violet: "#8E3B6C", teal: "#C9738A", gold: "#B08968", crimson: "#C1121F", orange: "#E07A5F", fab: "#F4978E" },
  charcoal: { name: "زغالی تیره", header: "#22223B", mauve: "#4A4E69", darkgreen: "#22577A", green: "#38A3A5", violet: "#4A4E69", teal: "#5C677D", gold: "#9A8C98", crimson: "#C9184A", orange: "#C08552", fab: "#57CC99" },
};
let BRAND = { ...COLOR_PRESETS.purple };
const THEME = {
  light: { bg: "#F1EFF4", card: "#ffffff", text: "#241a30", sub: "#8a8194", border: "#f0eef3", input: "#faf9fb", inputBorder: "#e3e0ea" },
  dark: { bg: "#17131c", card: "#241d2c", text: "#f1eef5", sub: "#a79fb3", border: "#332b3d", input: "#2c2434", inputBorder: "#3d3348" },
};
const FONT = "'Vazirmatn', Tahoma, 'Segoe UI', sans-serif";
const PIE_COLORS = ["#3E1461", "#A65475", "#1E8449", "#4E9AA0", "#A98A3B", "#C56A1F", "#6C3FA0", "#B01E4A"];

const hasCloudStorage = typeof window !== "undefined" && window.storage && typeof window.storage.get === "function";
async function loadKey(key, fallback, shared) {
  try {
    if (hasCloudStorage) {
      const res = await window.storage.get(key, !!shared);
      return res && res.value ? JSON.parse(res.value) : fallback;
    }
    const raw = window.localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch { return fallback; }
}
async function saveKey(key, value, shared) {
  try {
    if (hasCloudStorage) { await window.storage.set(key, JSON.stringify(value), !!shared); }
    else { window.localStorage.setItem(key, JSON.stringify(value)); }
  } catch (e) { console.error("save fail", e); }
}

const ThemeCtx = React.createContext(THEME.light);
function useT() { return React.useContext(ThemeCtx); }

/* ---------------------------------------------------------
   Bank Card (with 16 digits full display & expiry on bottom right)
--------------------------------------------------------- */
function BankCard({ account, balance, hidden, currency, usdRate }) {
  const rawNum = account.cardNumber ? String(account.cardNumber).replace(/\D/g, "") : "";
  const formattedCardNum = rawNum.length === 16 
    ? `${rawNum.slice(0,4)} ${rawNum.slice(4,8)} ${rawNum.slice(8,12)} ${rawNum.slice(12,16)}` 
    : (rawNum || account.id.slice(-4));
  const expiry = account.expiryDate || "1407/12";

  return (
    <div style={{
      minWidth: 300, maxWidth: 300, height: 176, borderRadius: 18, padding: 20, color: "#fff", flexShrink: 0,
      background: `linear-gradient(135deg, ${BRAND.violet}, ${BRAND.header})`,
      boxShadow: "0 6px 16px rgba(0,0,0,0.25)", display: "flex", flexDirection: "column", justifyContent: "space-between",
      scrollSnapAlign: "center", position: "relative", overflow: "hidden"
    }}>
      <div style={{ position: "absolute", top: -40, left: -40, width: 140, height: 140, borderRadius: "50%", background: "rgba(255,255,255,0.08)" }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", zIndex: 1 }}>
        <div style={{ fontWeight: 800, fontSize: 15 }}>{account.name}</div>
        <div style={{ width: 34, height: 24, borderRadius: 6, background: "linear-gradient(135deg,#f5d98b,#c9a94a)" }} />
      </div>
      <div style={{ fontSize: 15.5, letterSpacing: 2, fontWeight: 700, zIndex: 1, direction: "ltr", textAlign: "left" }}>
        {formattedCardNum}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", zIndex: 1 }}>
        <div style={{ fontSize: 11.5, opacity: 0.9, fontWeight: 600 }}>انقضا: {expiry}</div>
        <div style={{ fontSize: 15, fontWeight: 800 }}>{hidden ? "••••••••" : formatMoney(balance, currency, usdRate)}</div>
      </div>
    </div>
  );
}
function BankCardCarousel({ accounts, accountBalance, currency, usdRate, onAddCard }) {
  const t = useT();
  const [hidden, setHidden] = useState(false);
  const cards = accounts.filter((a) => a.type === "bank" || a.type === "card");
  return (
    <div style={{ marginBottom: 16 }}>
      {cards.length > 0 && (
        <div style={{ display: "flex", justifyContent: "flex-end", padding: "0 4px 6px" }}>
          <button onClick={() => setHidden((v) => !v)} style={{ background: "none", border: "none", color: t.sub, cursor: "pointer", display: "flex", alignItems: "center", gap: 4, fontSize: 12, fontWeight: 600 }}>
            {hidden ? <EyeOff size={15} /> : <Eye size={15} />} {hidden ? "نمایش موجودی" : "مخفی کردن موجودی"}
          </button>
        </div>
      )}
      <div style={{ display: "flex", gap: 12, overflowX: "auto", scrollSnapType: "x mandatory", paddingBottom: 4, paddingInline: 2 }}>
        {cards.map((a) => <BankCard key={a.id} account={a} balance={accountBalance(a.id)} hidden={hidden} currency={currency} usdRate={usdRate} />)}
        <button onClick={onAddCard} style={{
          minWidth: 300, maxWidth: 300, height: 176, borderRadius: 18, flexShrink: 0, scrollSnapAlign: "center",
          border: `2px dashed ${t.border}`, background: t.card, display: "flex", flexDirection: "column",
          alignItems: "center", justifyContent: "center", gap: 8, cursor: "pointer", color: t.sub
        }}>
          <span style={{ width: 44, height: 44, borderRadius: "50%", background: BRAND.violet, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Plus size={22} />
          </span>
          <span style={{ fontSize: 12.5, fontWeight: 700 }}>افزودن کارت/حساب جدید</span>
        </button>
      </div>
    </div>
  );
}

function DateQuickBox({ onNote, onReminder }) {
  const t = useT();
  return (
    <div style={{ background: t.card, borderRadius: 14, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "center", justifyContent: "space-between", boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
      <div style={{ color: t.text, fontSize: 13, fontWeight: 700 }}>{faLongDate(new Date())}</div>
      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={onNote} style={{ ...quickPillBtn, background: BRAND.violet }}><StickyNote size={13} /> یادداشت</button>
        <button onClick={onReminder} style={{ ...quickPillBtn, background: BRAND.orange }}><BellRing size={13} /> یادآوری</button>
      </div>
    </div>
  );
}
const quickPillBtn = { display: "flex", alignItems: "center", gap: 4, border: "none", color: "#fff", borderRadius: 20, padding: "6px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer" };

/* ---------------------------------------------------------
   Smile / Frown Arc & Chart (clean, exactly under card, no overlapping text/box)
--------------------------------------------------------- */
function SmileFrownArc({ income, expense, pieData = [], currency, usdRate }) {
  const t = useT();
  const net = income - expense;
  const safeData = pieData.length ? pieData : [{ name: "بدون تراکنش", amount: 1 }];
  const arcColor = net > 0 ? BRAND.darkgreen : net < 0 ? BRAND.crimson : "#94a3b8";
  const w = 220, h = 52, strokeW = 12;
  let d;
  if (net > 0) d = `M 18 ${h * 0.70} Q ${w / 2} ${h * 0.12} ${w - 18} ${h * 0.70}`;
  else if (net < 0) d = `M 18 ${h * 0.30} Q ${w / 2} ${h * 0.88} ${w - 18} ${h * 0.30}`;
  else d = `M 18 ${h / 2} L ${w - 18} ${h / 2}`;

  return (
    <div style={{ width: "100%", maxWidth: 340, margin: "2px auto 0" }}>
      <div style={{ display: "flex", justifyContent: "center", marginBottom: 4 }}>
        <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ display: "block", overflow: "visible" }}>
          <path d={d} fill="none" stroke={arcColor} strokeWidth={strokeW} strokeLinecap="round" />
        </svg>
      </div>
      <div style={{ display: "flex", justifyContent: "space-around", alignItems: "flex-start", padding: "0 8px 8px", gap: 8 }}>
        <div style={{ textAlign: "center", flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 11.5, color: t.sub, fontWeight: 700, marginBottom: 2 }}>درآمد</div>
          <div style={{ fontSize: 15, fontWeight: 900, color: BRAND.darkgreen, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
            {formatMoney(income, currency, usdRate)}
          </div>
        </div>
        <div style={{ textAlign: "center", flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 11.5, color: t.sub, fontWeight: 700, marginBottom: 2 }}>هزینه</div>
          <div style={{ fontSize: 15, fontWeight: 900, color: BRAND.crimson, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
            {formatMoney(expense, currency, usdRate)}
          </div>
        </div>
      </div>
      <div style={{ position: "relative", width: "100%", height: 160, marginTop: 2 }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie data={safeData} dataKey="amount" nameKey="name" cx="50%" cy="50%" innerRadius="32%" outerRadius="52%" paddingAngle={3} stroke="#fff" strokeWidth={3} isAnimationActive={false}>
              {safeData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
            </Pie>
          </PieChart>
        </ResponsiveContainer>
        <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
          <div style={{ width: 48, height: 48, borderRadius: "50%", background: t.card }} />
        </div>
      </div>
      <div style={{ textAlign: "center", fontSize: 11.5, color: t.sub, marginTop: 2, fontWeight: 600 }}>برای جزئیات روی هر بخش بزن</div>
      {pieData.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, justifyContent: "center", marginTop: 8 }}>
          {pieData.slice(0, 6).map((item, i) => (
            <span key={i} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: t.sub }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: PIE_COLORS[i % PIE_COLORS.length] }} />
              {item.name}
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

function IncomeExpenseDayWidget({ day, setDay, transactions, catById, currency, usdRate }) {
  const t = useT();
  const dayTx = transactions.filter((tx) => tx.date === day);
  const income = dayTx.filter((tx) => tx.type === "income").reduce((s, tx) => s + tx.amount, 0);
  const expense = dayTx.filter((tx) => tx.type === "expense").reduce((s, tx) => s + tx.amount, 0);
  const pieData = useMemo(() => {
    const map = {};
    dayTx.filter((tx) => tx.type !== "transfer").forEach((tx) => {
      const name = catById(tx.categoryId)?.name || "—";
      map[name] = (map[name] || 0) + Number(tx.amount || 0);
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]).map(([name, amount]) => ({ name, amount }));
  }, [dayTx, catById]);
  return (
    <div style={{ background: t.card, borderRadius: 16, padding: "12px 12px 16px", marginBottom: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.08)", overflow: "visible" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6, padding: "0 2px" }}>
        <button onClick={() => setDay(addDays(day, -1))} style={navArrowStyle(t)}><ChevronRight size={17} /></button>
        <div style={{ fontSize: 13.5, fontWeight: 800, color: BRAND.header }}>{faLongDate(new Date(day))}</div>
        <button onClick={() => setDay(addDays(day, 1))} style={navArrowStyle(t)}><ChevronLeft size={17} /></button>
      </div>
      <SmileFrownArc income={income} expense={expense} pieData={pieData} currency={currency} usdRate={usdRate} />
    </div>
  );
}

function CollapsibleSection({ color, title, open, onToggle, children, badge, onAdd }) {
  const t = useT();
  return (
    <div style={{ borderRadius: 14, overflow: "hidden", marginBottom: 12, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
      <div onClick={onToggle} style={{
        width: "100%", background: color, border: "none", color: "#fff", padding: "14px 16px",
        display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer", fontFamily: "inherit"
      }}>
        <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ width: 30, height: 30, borderRadius: "50%", background: "rgba(255,255,255,.18)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <ChevronDown size={17} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform .2s" }} />
          </span>
          {onAdd && (
            <button onClick={(e) => { e.stopPropagation(); onAdd(); }} style={{
              width: 28, height: 28, borderRadius: "50%", background: "rgba(255,255,255,.22)", border: "none",
              color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer"
            }}>
              <Plus size={15} />
            </button>
          )}
        </span>
        <span style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: 700, fontSize: 14.5 }}>
          {badge}{title}
        </span>
      </div>
      {open && <div style={{ background: t.card, padding: "6px 14px" }}>{children}</div>}
    </div>
  );
}

function Row({ leftIcon, leftColor = "#eee", title, subtitle, value, valueColor, onClick, chevron = "left", extra }) {
  const t = useT();
  return (
    <div onClick={onClick} style={{
      display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 4px",
      borderBottom: `1px solid ${t.border}`, cursor: onClick ? "pointer" : "default"
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
        {chevron && <ChevronLeft size={16} color={t.sub} style={{ flexShrink: 0, transform: chevron === "left" ? "none" : "rotate(180deg)" }} />}
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 14, fontWeight: 600, color: t.text }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11.5, color: t.sub, marginTop: 2 }}>{subtitle}</div>}
        </div>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
        {extra}
        {value !== undefined && <span style={{ fontSize: 13, fontWeight: 700, color: valueColor || t.text }}>{value}</span>}
      </div>
    </div>
  );
}
function EmptyRow({ text }) { const t = useT(); return <div style={{ padding: "16px 4px", textAlign: "center", color: t.sub, fontSize: 13 }}>{text}</div>; }
function AddLink({ text, onClick }) { return <div onClick={onClick} style={{ padding: "10px 4px", color: BRAND.header, fontSize: 13, fontWeight: 700, cursor: "pointer", textAlign: "center" }}>{text}</div>; }

function navArrowStyle(t) {
  return { width: 32, height: 32, borderRadius: "50%", background: t.card, border: `1px solid ${t.border}`, color: t.text, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", boxShadow: "0 1px 2px rgba(0,0,0,0.05)" };
}

/* ---------------------------------------------------------
   Header & BottomNav (with Keyboard handling)
--------------------------------------------------------- */
function Header({ title = "Rexa", onMenu, back, onBack, onMic }) {
  return <div style={{ background: BRAND.header, color: "#fff", padding: "calc(env(safe-area-inset-top, 0px) + 10px) 12px 10px", display: "grid", gridTemplateColumns: "42px 1fr 76px", alignItems: "center", position: "sticky", top: 0, zIndex: 80 }}>
    <div style={{ display: "flex", justifyContent: "flex-start" }}>{back ? <button onClick={onBack} style={iconBtn}><ChevronRight size={24} /></button> : <button onClick={onMenu} style={iconBtn} aria-label="منوی برنامه"><Menu size={24} /></button>}</div>
    <div style={{ textAlign: "center", fontWeight: 800, fontSize: 17, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{title}</div>
    <div style={{ display: "flex", justifyContent: "flex-end" }}>{onMic && <VoiceCaptureButton onResult={onMic} />}</div>
  </div>;
}
const iconBtn = { background: "none", border: "none", color: "#fff", cursor: "pointer" };

function VoiceCaptureButton({ onResult }) {
  const [listening, setListening] = useState(false);
  const isNative = typeof Capacitor !== "undefined" && Capacitor.isNativePlatform && Capacitor.isNativePlatform();

  async function startNative() {
    try {
      setListening(true);
      try { await SpeechRecognition.requestPermissions(); } catch {}
      const result = await SpeechRecognition.start({ language: "fa-IR", maxResults: 1, prompt: "چی می‌خوای ثبت کنی؟", partialResults: false, popup: true });
      setListening(false);
      const transcript = result?.matches?.[0];
      if (transcript) onResult(transcript);
    } catch { setListening(false); }
  }

  function startWeb() {
    const Rec = typeof window !== "undefined" && (window.SpeechRecognition || window.webkitSpeechRecognition);
    if (!Rec) { alert("تشخیص گفتار در این مرورگر پشتیبانی نمی‌شود."); return; }
    const rec = new Rec();
    rec.lang = "fa-IR"; rec.interimResults = false;
    rec.onstart = () => setListening(true);
    rec.onend = () => setListening(false);
    rec.onerror = () => setListening(false);
    rec.onresult = (e) => { const tr = e.results?.[0]?.[0]?.transcript; if (tr) onResult(tr); };
    try { rec.start(); } catch { setListening(false); }
  }

  return <button onClick={() => isNative ? startNative() : startWeb()} style={{ ...iconBtn, color: listening ? "#ffd166" : "#fff" }} title="ثبت با صدا"><Mic size={20} /></button>;
}

function BottomNav({ active, setActive, onAdd, keyboardHeight }) {
  const t = useT();
  const items = [
    { key: "reports", label: "گزارش‌ها", icon: PieChartIcon },
    { key: "transactions", label: "تراکنش‌ها", icon: Receipt },
    { key: "checks", label: "چک‌ها", icon: FileSpreadsheet },
    { key: "home", label: "خانه", icon: HomeIcon },
  ];
  return (
    <div style={{
      position: "fixed", left: "50%", transform: "translateX(-50%)",
      bottom: keyboardHeight > 0 ? `${keyboardHeight}px` : 0,
      width: "min(480px, 100vw)", minHeight: 72,
      padding: "7px 6px calc(7px + env(safe-area-inset-bottom, 0px))",
      background: t.card, borderTop: `1px solid ${t.border}`,
      display: "grid", gridTemplateColumns: "1fr 1fr 64px 1fr 1fr", alignItems: "center",
      zIndex: 300, boxShadow: "0 -5px 18px rgba(0,0,0,.10)",
      transition: "bottom 0.15s ease-out"
    }}>
      <NavBtn it={items[0]} active={active} setActive={setActive} />
      <NavBtn it={items[1]} active={active} setActive={setActive} />
      <button onClick={onAdd} aria-label="افزودن" style={{ width: 56, height: 56, borderRadius: "50%", background: BRAND.fab, border: `4px solid ${t.card}`, margin: "-20px auto 0", boxShadow: "0 4px 12px rgba(0,0,0,.25)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", cursor: "pointer" }}>
        <ArrowLeftRight size={25} />
      </button>
      <NavBtn it={items[2]} active={active} setActive={setActive} />
      <NavBtn it={items[3]} active={active} setActive={setActive} />
    </div>
  );
}
function NavBtn({ it, active, setActive }) {
  const Icon = it.icon; const isActive = active === it.key;
  return (
    <button onClick={() => setActive(it.key)} style={{ background: "none", border: "none", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, color: isActive ? BRAND.header : "#8a8a8a", cursor: "pointer", fontFamily: "inherit" }}>
      <Icon size={22} />
      <span style={{ fontSize: 11, fontWeight: isActive ? 700 : 500 }}>{it.label}</span>
    </button>
  );
}

/* ---------------------------------------------------------
   Lock screen & Logo
--------------------------------------------------------- */
function RexaLogo({ size = 64, showWordmark = false }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8 }}>
      <svg width={size} height={size} viewBox="0 0 200 200">
        <defs>
          <radialGradient id="rexaCoin" cx="35%" cy="30%" r="75%">
            <stop offset="0%" stopColor="#ffe9b0" />
            <stop offset="55%" stopColor="#ffd166" />
            <stop offset="100%" stopColor="#e8a83c" />
          </radialGradient>
          <linearGradient id="rexaRing" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stopColor="#ffe6a0" />
            <stop offset="100%" stopColor="#e8951f" />
          </linearGradient>
        </defs>
        <circle cx="100" cy="100" r="96" fill="url(#rexaRing)" />
        <circle cx="100" cy="100" r="76" fill="url(#rexaCoin)" stroke="#c9863a" strokeWidth="3" />
        <circle cx="100" cy="100" r="62" fill="none" stroke="#d6a24a" strokeWidth="3" />
        <text x="100" y="128" textAnchor="middle" fontFamily="Georgia, serif" fontWeight="700" fontSize="86" fill={BRAND.header}>R</text>
      </svg>
      {showWordmark && <div style={{ fontWeight: 800, fontSize: size * 0.3, color: "#fff", letterSpacing: 1, fontFamily: "Georgia, serif" }}>Rexa</div>}
    </div>
  );
}

function LockScreen({ pin, onUnlock, biometricEnabled }) {
  const [val, setVal] = useState("");
  const [err, setErr] = useState(false);
  const isNative = typeof Capacitor !== "undefined" && Capacitor.isNativePlatform && Capacitor.isNativePlatform();

  async function tryBiometric() {
    if (!biometricEnabled || !isNative) return;
    try {
      const info = await BiometricAuth.checkBiometry();
      if (info.isAvailable) {
        await BiometricAuth.authenticate({ reason: "برای باز کردن Rexa هویتت رو تایید کن", cancelTitle: "لغو", allowDeviceCredential: true });
        onUnlock();
      }
    } catch {}
  }
  useEffect(() => { tryBiometric(); }, []);

  return (
    <div dir="rtl" style={{ minHeight: "100vh", background: BRAND.header, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: FONT, color: "#fff", gap: 18, maxWidth: 480, margin: "0 auto" }}>
      <RexaLogo size={64} />
      <div style={{ fontWeight: 700, fontSize: 16 }}>Rexa قفل است</div>
      <input type="password" inputMode="numeric" maxLength={6} value={val} onChange={(e) => { setVal(e.target.value.replace(/[^0-9]/g, "")); setErr(false); }} placeholder="رمز عبور" style={{ width: 180, textAlign: "center", fontSize: 22, letterSpacing: 6, padding: "10px", borderRadius: 10, border: "none", outline: "none" }} />
      {err && <div style={{ color: "#ffb3c1", fontSize: 13 }}>رمز اشتباه است</div>}
      <button onClick={() => val === pin ? onUnlock() : setErr(true)} style={{ background: BRAND.fab, color: "#fff", border: "none", borderRadius: 9, padding: "10px 30px", fontWeight: 700, cursor: "pointer" }}>باز کردن</button>
    </div>
  );
}

/* ===========================================================
   MAIN APP COMPONENT
=========================================================== */
export default function App() {
  const [loaded, setLoaded] = useState(false);
  const [settings, setSettings] = useState(seedSettings());
  const [unlocked, setUnlocked] = useState(false);

  const [accounts, setAccounts] = useState([]);
  const [categories, setCategories] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [budgets, setBudgets] = useState([]);
  const [loans, setLoans] = useState([]);
  const [checks, setChecks] = useState([]);
  const [bills, setBills] = useState([]);
  const [assets, setAssets] = useState([]);
  const [recurring, setRecurring] = useState([]);
  const [favorites, setFavorites] = useState({ categories: [], accounts: [] });
  const [members, setMembers] = useState([]);
  const [events, setEvents] = useState([]);
  const [projects, setProjects] = useState([]);
  const [fiscalPeriods, setFiscalPeriods] = useState([]);
  const [notes, setNotes] = useState([]);
  const [reminders, setReminders] = useState([]);
  const [persons, setPersons] = useState([]);
  const [debts, setDebts] = useState([]);
  const [currencies, setCurrencies] = useState([]);
  const [goals, setGoals] = useState([]);

  const [tab, setTab] = useState("home");
  const [subView, setSubView] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [prefillTx, setPrefillTx] = useState(null);
  const [open, setOpen] = useState({});
  const [year, setYear] = useState(jalaliYear(new Date()));
  const [homeDay, setHomeDay] = useState(todayISO());
  const [txFilter, setTxFilter] = useState("all");
  const [txSearch, setTxSearch] = useState("");
  const [menuOpen, setMenuOpen] = useState(false);
  const [showQuickAdd, setShowQuickAdd] = useState(false);
  const [showNoteModal, setShowNoteModal] = useState(false);
  const [showReminderModal, setShowReminderModal] = useState(false);
  const [rates, setRates] = useState(null);
  const [shortcuts, setShortcuts] = useState(() => settings.shortcuts || DEFAULT_SHORTCUTS);
  const [keyboardHeight, setKeyboardHeight] = useState(0);

  // Keyboard visualViewport tracking
  useEffect(() => {
    if (typeof window === "undefined" || !window.visualViewport) return;
    const handleResize = () => {
      const vv = window.visualViewport;
      const diff = window.innerHeight - vv.height;
      setKeyboardHeight(diff > 60 ? diff : 0);
    };
    window.visualViewport.addEventListener("resize", handleResize);
    window.visualViewport.addEventListener("scroll", handleResize);
    return () => {
      window.visualViewport.removeEventListener("resize", handleResize);
      window.visualViewport.removeEventListener("scroll", handleResize);
    };
  }, []);

  const shared = settings.sharedFamily;

  const reloadAll = useCallback(async (sh) => {
    const [a, c, t, b, ln, ck, bl, as, rc, fv, mb, ev, pj, fp, nt, rm, ps, db, cu, gl] = await Promise.all([
      loadKey("hs:accounts", null, sh), loadKey("hs:categories", null, sh),
      loadKey("hs:transactions", null, sh), loadKey("hs:budgets", null, sh),
      loadKey("hs:loans", [], sh), loadKey("hs:checks", [], sh),
      loadKey("hs:bills", [], sh), loadKey("hs:assets", [], sh),
      loadKey("hs:recurring", [], sh), loadKey("hs:favorites", { categories: [], accounts: [] }, sh),
      loadKey("hs:members", [], sh), loadKey("hs:events", [], sh),
      loadKey("hs:projects", [], sh), loadKey("hs:fiscalPeriods", [], sh),
      loadKey("hs:notes", [], sh), loadKey("hs:reminders", [], sh),
      loadKey("hs:persons", [], sh), loadKey("hs:debts", [], sh), loadKey("hs:currencies", [], sh),
      loadKey("hs:goals", [], sh),
    ]);
    setAccounts(a || seedAccounts());
    setCategories(c || seedCategories());
    setTransactions(t || []);
    setBudgets(b || []);
    setLoans(ln); setChecks(ck); setBills(bl); setAssets(as); setRecurring(rc); setFavorites(fv);
    setMembers(mb); setEvents(ev); setProjects(pj); setFiscalPeriods(fp);
    setNotes(nt); setReminders(rm);
    setPersons(ps || []); setDebts(db || []); setCurrencies(cu || []); setGoals(gl || []);
  }, []);

  useEffect(() => {
    (async () => {
      const s = await loadKey("hs:settings", seedSettings(), false);
      setSettings(s);
      if (s.shortcuts) setShortcuts(s.shortcuts);
      await reloadAll(s.sharedFamily);
      setLoaded(true);
    })();
  }, [reloadAll]);

  useEffect(() => { if (loaded) saveKey("hs:settings", settings, false); }, [settings, loaded]);
  useEffect(() => { if (loaded) saveKey("hs:accounts", accounts, shared); }, [accounts, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:categories", categories, shared); }, [categories, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:transactions", transactions, shared); }, [transactions, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:budgets", budgets, shared); }, [budgets, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:loans", loans, shared); }, [loans, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:checks", checks, shared); }, [checks, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:bills", bills, shared); }, [bills, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:assets", assets, shared); }, [assets, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:recurring", recurring, shared); }, [recurring, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:notes", notes, shared); }, [notes, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:reminders", reminders, shared); }, [reminders, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:persons", persons, shared); }, [persons, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:debts", debts, shared); }, [debts, loaded, shared]);
  useEffect(() => { if (loaded) setSettings((prev) => prev.shortcuts === shortcuts ? prev : { ...prev, shortcuts }); }, [shortcuts, loaded]);

  const toggle = (key) => setOpen((o) => ({ ...o, [key]: !o[key] }));
  const catById = useCallback((id) => categories.find((c) => c.id === id), [categories]);
  const accById = useCallback((id) => accounts.find((a) => a.id === id), [accounts]);

  const accountBalance = useCallback((accId) => {
    const acc = accById(accId); if (!acc) return 0;
    let bal = acc.initial || 0;
    transactions.forEach((t) => {
      if (t.type === "expense" && t.accountId === accId) bal -= t.amount;
      if (t.type === "income" && t.accountId === accId) bal += t.amount;
      if (t.type === "transfer") { if (t.accountId === accId) bal -= t.amount; if (t.toAccountId === accId) bal += t.amount; }
    });
    return bal;
  }, [accounts, transactions, accById]);

  const totalIncomeYear = useMemo(() => transactions.filter((t) => jalaliYear(new Date(t.date)) === year && t.type === "income").reduce((s, t) => s + t.amount, 0), [transactions, year]);
  const totalExpenseYear = useMemo(() => transactions.filter((t) => jalaliYear(new Date(t.date)) === year && t.type === "expense").reduce((s, t) => s + t.amount, 0), [transactions, year]);
  const totalAssets = useMemo(() => assets.reduce((s, a) => s + (a.quantity * a.currentPrice || 0), 0), [assets]);

  const expenseByCategory = useMemo(() => {
    const map = {};
    transactions.filter((t) => jalaliYear(new Date(t.date)) === year && t.type === "expense").forEach((t) => { map[t.categoryId] = (map[t.categoryId] || 0) + t.amount; });
    return Object.entries(map).map(([catId, amount]) => ({ catId, amount, name: catById(catId)?.name || "بدون دسته" })).sort((a, b) => b.amount - a.amount);
  }, [transactions, year, catById]);

  const incomeByCategory = useMemo(() => {
    const map = {};
    transactions.filter((t) => jalaliYear(new Date(t.date)) === year && t.type === "income").forEach((t) => { map[t.categoryId] = (map[t.categoryId] || 0) + t.amount; });
    return Object.entries(map).map(([catId, amount]) => ({ catId, amount, name: catById(catId)?.name || "بدون دسته" })).sort((a, b) => b.amount - a.amount);
  }, [transactions, year, catById]);

  function addTransaction(tx) { setTransactions((p) => [{ ...tx, id: uid(), createdAt: new Date().toISOString() }, ...p]); }
  function deleteTransaction(id) { setTransactions((p) => p.filter((t) => t.id !== id)); }
  function addAccount(a) { const item = { ...a, id: a.id || uid() }; setAccounts((p) => [...p, item]); return item.id; }
  function updateAccount(id, patch) { setAccounts((p) => p.map((a) => a.id === id ? { ...a, ...patch } : a)); }
  function deleteAccount(id) { setAccounts((p) => p.filter((a) => a.id !== id)); }
  function addCategory(c) { const item = { ...c, id: c.id || uid() }; setCategories((p) => [...p, item]); return item.id; }

  async function fetchRates() {
    try {
      const url = "https://gem.tgju.org/profile/price_dollar_rl";
      let html = "";
      if (Capacitor.isNativePlatform()) {
        const r = await CapacitorHttp.get({ url, headers: { Accept: "text/html" } });
        html = String(r?.data || "");
      } else {
        const r = await fetch(url, { cache: "no-store" });
        html = await r.text();
      }
      const m = html.match(/price_dollar_rl[^0-9]{0,120}([0-9,]{6,})/i) || html.match(/([0-9,]{6,})/);
      if (m) {
        const irr = Number(String(m[1]).replace(/,/g, ""));
        if (irr > 100000) setRates({ usd: irr, fetchedAt: new Date().toISOString() });
      }
    } catch {}
  }
  useEffect(() => { if (loaded) fetchRates(); }, [loaded]);

  const t = THEME[settings.theme] || THEME.light;
  Object.assign(BRAND, COLOR_PRESETS[settings.themeColor] || COLOR_PRESETS.purple);

  if (!loaded) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16, background: BRAND.header, fontFamily: FONT }}>
        <RexaLogo size={84} showWordmark />
        <div style={{ color: "#d8c9e8", fontSize: 13 }}>در حال بارگذاری...</div>
      </div>
    );
  }
  if ((settings.pin || settings.biometricEnabled) && !unlocked) {
    return <LockScreen pin={settings.pin} onUnlock={() => setUnlocked(true)} biometricEnabled={settings.biometricEnabled} />;
  }

  const ctx = {
    accounts, addAccount, deleteAccount, updateAccount, accountBalance,
    categories, addCategory, budgets, setBudgets, loans, setLoans, checks, setChecks,
    bills, setBills, assets, setAssets, persons, setPersons, debts, setDebts,
    shortcuts, setShortcuts, transactions, addTransaction, deleteTransaction,
    notes, setNotes, reminders, setReminders, settings, setSettings, catById, accById
  };

  return (
    <ThemeCtx.Provider value={t}>
      <div dir="rtl" style={{ fontFamily: FONT, background: t.bg, color: t.text, minHeight: "100vh", maxWidth: 480, margin: "0 auto", display: "flex", flexDirection: "column", position: "relative", paddingBottom: "calc(96px + env(safe-area-inset-bottom, 0px))" }}>
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700;800&display=swap');
          * { box-sizing: border-box; } html, body, #root { margin: 0; min-height: 100%; width: 100%; } body { overflow-x: hidden; overscroll-behavior-y: none; }
          input, select, textarea, button { font-family: inherit; } ::-webkit-scrollbar { width: 0; height: 0; }
        `}</style>

        {subView ? (
          <Header title={SUBVIEW_TITLES[subView]} back onBack={() => setSubView(null)} />
        ) : (
          <Header title={settings.profile?.name || "رضا"} onMenu={() => setMenuOpen(true)} onMic={(txt) => {
            const parsed = parseBankSms(txt);
            if (parsed.amount) {
              const accId = accounts[0]?.id;
              if (accId) addTransaction({ type: parsed.type, amount: parsed.amount, categoryId: categories[0]?.id || "", accountId: accId, date: todayISO(), note: parsed.note });
            }
          }} />
        )}

        {subView ? (
          <SubViewContent subView={subView} ctx={ctx} onBack={() => setSubView(null)} />
        ) : (
          <div style={{ flex: 1, overflowY: "auto", paddingBottom: 8 }}>
            {tab === "home" && (
              <HomeView
                year={year} setYear={setYear}
                totalIncomeYear={totalIncomeYear} totalExpenseYear={totalExpenseYear}
                open={open} toggle={toggle}
                accounts={accounts} accountBalance={accountBalance}
                expenseByCategory={expenseByCategory} incomeByCategory={incomeByCategory}
                budgets={budgets} categories={categories} allTransactions={transactions}
                bills={bills} loans={loans} checks={checks} assets={assets} totalAssets={totalAssets}
                persons={persons} debts={debts}
                openAccounts={() => setSubView("accounts")} openBudgets={() => setSubView("budgets")}
                openBills={() => setSubView("bills")} openLoans={() => setSubView("loans")}
                openPersons={() => setSubView("persons")} openDebts={() => setSubView("debts")}
                openChecks={() => setSubView("checks")} openAssets={() => setSubView("assets")}
                homeDay={homeDay} setHomeDay={setHomeDay} catById={catById}
                settings={settings} setSettings={setSettings} rates={rates} fetchRates={fetchRates}
                onNote={() => setShowNoteModal(true)} onReminder={() => setShowReminderModal(true)}
                shortcuts={shortcuts} openShortcuts={() => setSubView("shortcuts")}
                onRunShortcut={(sc) => { setPrefillTx({ type: sc.type, categoryId: sc.categoryId || "", note: sc.note || sc.name }); setShowAdd(true); }}
                onAddTransaction={(type) => { setPrefillTx(type ? { type } : null); setShowAdd(true); }}
              />
            )}
            {tab === "transactions" && (
              <TransactionsView transactions={transactions} catById={catById} accById={accById} filter={txFilter} setFilter={setTxFilter} onDelete={deleteTransaction} search={txSearch} setSearch={setTxSearch} onEdit={(tx) => { setPrefillTx({ ...tx, _editId: tx.id }); setShowAdd(true); }} />
            )}
            {tab === "checks" && <ChecksManager checks={checks} setChecks={setChecks} />}
            {tab === "reports" && (
              <ReportsView expenseByCategory={expenseByCategory} incomeByCategory={incomeByCategory} totalIncomeYear={totalIncomeYear} totalExpenseYear={totalExpenseYear} accounts={accounts} accountBalance={accountBalance} currency={settings.currency} usdRate={rates?.usd} transactions={transactions} />
            )}
          </div>
        )}

        <BottomNav
          active={subView ? null : tab}
          setActive={(k) => { setShowAdd(false); setPrefillTx(null); setShowQuickAdd(false); setMenuOpen(false); setSubView(null); setTab(k); }}
          onAdd={() => setShowQuickAdd(true)}
          keyboardHeight={keyboardHeight}
        />

        {showQuickAdd && (
          <QuickAddSheet onClose={() => setShowQuickAdd(false)} onPick={(type) => {
            setShowQuickAdd(false);
            if (type === "check") { setTab("checks"); return; }
            setPrefillTx(type === "expense" ? { type: "expense" } : type === "income" ? { type: "income" } : type === "transfer" ? { type: "transfer" } : null);
            setShowAdd(true);
          }} />
        )}

        {showNoteModal && (
          <SimpleTextModal title="یادداشت جدید" placeholder="متن یادداشت..." onClose={() => setShowNoteModal(false)} onSubmit={(text) => { setNotes((p) => [{ id: uid(), text, date: todayISO() }, ...p]); setShowNoteModal(false); }} />
        )}
        {showReminderModal && (
          <ReminderQuickModal onClose={() => setShowReminderModal(false)} onSubmit={(r) => { setReminders((p) => [{ id: uid(), ...r }, ...p]); setShowReminderModal(false); }} />
        )}

        {showAdd && (
          <AddTransactionSheet accounts={accounts} categories={categories} initial={prefillTx} onClose={() => { setShowAdd(false); setPrefillTx(null); }} onSubmit={(tx) => {
            if (prefillTx?._editId) setTransactions((prev) => prev.map((item) => item.id === prefillTx._editId ? { ...item, ...tx, id: item.id } : item));
            else addTransaction(tx);
            setShowAdd(false); setPrefillTx(null);
          }} />
        )}
        {menuOpen && <SideMenu onClose={() => setMenuOpen(false)} setSubView={(v) => { setSubView(v); setMenuOpen(false); }} profileName={settings.profile?.name} />}
      </div>
    </ThemeCtx.Provider>
  );
}

/* ---------------------------------------------------------
   Home View with Sections
--------------------------------------------------------- */
function HomeView({
  year, setYear, totalIncomeYear, totalExpenseYear, open, toggle,
  accounts, accountBalance, expenseByCategory, incomeByCategory, budgets, categories,
  allTransactions, bills, loans, checks, totalAssets, persons = [], debts = [],
  openAccounts, openBudgets, openBills, openLoans, openChecks, openAssets, openPersons, openDebts,
  homeDay, setHomeDay, catById, settings, setSettings, rates, fetchRates, onNote, onReminder,
  onAddTransaction, shortcuts = [], onRunShortcut, openShortcuts
}) {
  const t = useT();
  const banks = accounts.filter((a) => a.type === "bank");
  const funds = accounts.filter((a) => a.type === "fund");
  const cardAccs = accounts.filter((a) => a.type === "card");
  const net = totalIncomeYear - totalExpenseYear;
  const currency = settings.currency, usdRate = rates?.usd;
  const layout = settings.homeLayout || "cards";
  const sectionOrder = (settings.homeSections && settings.homeSections.length ? settings.homeSections : DEFAULT_HOME_SECTIONS).filter((s) => s.visible);

  const SECTION_META = {
    shortcut: { color: BRAND.mauve, title: "میانبر تراکنش ها", icon: <ArrowLeftRight size={20} /> },
    expense: { color: BRAND.header, title: "هزینه ها", icon: <TrendingDown size={20} /> },
    income: { color: BRAND.darkgreen, title: "درآمدها", icon: <TrendingUp size={20} /> },
    banks: { color: BRAND.violet, title: "بانک ها و کارت ها", icon: <Landmark size={20} /> },
    funds: { color: BRAND.teal, title: "صندوق ها", icon: <Save size={20} /> },
    balrep: { color: BRAND.gold, title: "گزارش مانده حساب ها", icon: <FileSpreadsheet size={20} /> },
    budget: { color: BRAND.green, title: "بودجه بندی", icon: <Target size={20} /> },
    loanchk: { color: BRAND.orange, title: "وام ها و چک ها", icon: <Bank size={20} /> },
    contacts: { color: BRAND.violet, title: "اشخاص و بدهی ها", icon: <Users size={20} /> },
    bills: { color: BRAND.crimson, title: "یادآوری قبض ها", icon: <BellRing size={20} /> },
  };

  function sectionBody(key) {
    switch (key) {
      case "shortcut":
        return (
          <div>
            {shortcuts.length === 0 && <div style={{ fontSize: 12, color: t.sub, textAlign: "center", margin: "4px 0 10px" }}>هنوز میانبری ساخته نشده</div>}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, padding: "10px 4px 8px" }}>
              {shortcuts.map((sc) => {
                const palette = sc.type === "income" ? ["#cffafe", "#0e7490"] : ["#fef3c7", "#b45309"];
                return (
                  <button key={sc.id} onClick={() => onRunShortcut?.(sc)} style={{ border: 0, background: "transparent", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, padding: 2, cursor: "pointer" }}>
                    <span style={{ width: 56, height: 56, borderRadius: 14, background: palette[0], color: palette[1], display: "flex", alignItems: "center", justifyContent: "center", fontSize: 24, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>✦</span>
                    <span style={{ fontSize: 11.5, fontWeight: 700, color: t.text, textAlign: "center" }}>{sc.name}</span>
                  </button>
                );
              })}
              <button onClick={openShortcuts} style={{ border: 0, background: "transparent", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, padding: 2, cursor: "pointer" }}>
                <span style={{ width: 56, height: 56, borderRadius: 14, background: "#e0f2fe", color: BRAND.violet, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 26, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>+</span>
                <span style={{ fontSize: 11.5, fontWeight: 700, color: t.text, textAlign: "center" }}>افزودن</span>
              </button>
            </div>
          </div>
        );
      case "expense":
        return expenseByCategory.length === 0 ? <EmptyRow text="هزینه‌ای ثبت نشده" /> : expenseByCategory.map((e) => <Row key={e.catId} title={e.name} value={formatMoney(e.amount, currency, usdRate)} valueColor={BRAND.crimson} />);
      case "income":
        return incomeByCategory.length === 0 ? <EmptyRow text="درآمدی ثبت نشده" /> : incomeByCategory.map((e) => <Row key={e.catId} title={e.name} value={formatMoney(e.amount, currency, usdRate)} valueColor={BRAND.darkgreen} />);
      case "banks":
        return (<>
          {[...banks, ...cardAccs].length === 0 && <EmptyRow text="حسابی ثبت نشده" />}
          {[...banks, ...cardAccs].map((a) => (
            <Row key={a.id} title={a.name} subtitle={a.cardNumber ? `شماره کارت: ${a.cardNumber}` : (a.type === "card" ? "کارت" : "بانک")} value={formatMoney(accountBalance(a.id), currency, usdRate)} valueColor={accountBalance(a.id) >= 0 ? t.text : BRAND.crimson} />
          ))}
          <AddLink text="+ مدیریت حساب‌ها و کارت‌ها" onClick={openAccounts} />
        </>);
      case "funds":
        return funds.length === 0 ? <EmptyRow text="صندوقی ثبت نشده" /> : funds.map((a) => <Row key={a.id} title={a.name} value={formatMoney(accountBalance(a.id), currency, usdRate)} valueColor={accountBalance(a.id) >= 0 ? t.text : BRAND.crimson} />);
      case "balrep":
        return accounts.map((a) => <Row key={a.id} title={a.name} subtitle={a.type} value={formatMoney(accountBalance(a.id), currency, usdRate)} valueColor={accountBalance(a.id) >= 0 ? BRAND.darkgreen : BRAND.crimson} />);
      case "budget":
        return (<>
          {budgets.length === 0 && <EmptyRow text="بودجه‌ای تعریف نشده" />}
          {budgets.map((b) => {
            const cat = categories.find((c) => c.id === b.categoryId);
            const spent = expenseByCategory.find((e) => e.catId === b.categoryId)?.amount || 0;
            const pct = Math.min(100, Math.round((spent / (b.amount || 1)) * 100));
            return (
              <div key={b.id} style={{ padding: "10px 4px", borderBottom: `1px solid ${t.border}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, fontWeight: 600, marginBottom: 6 }}>
                  <span>{cat?.name || "—"}</span>
                  <span style={{ color: pct >= 100 ? BRAND.crimson : t.sub }}>{toFaInt(pct)}٪ — {toFaInt(spent)}/{toFaInt(b.amount)}</span>
                </div>
                <div style={{ height: 7, background: t.border, borderRadius: 4, overflow: "hidden" }}>
                  <div style={{ width: `${pct}%`, height: "100%", background: pct >= 100 ? BRAND.crimson : BRAND.green }} />
                </div>
              </div>
            );
          })}
          <AddLink text="+ مدیریت بودجه‌بندی" onClick={openBudgets} />
        </>);
      case "loanchk":
        return (<>
          {loans.length === 0 && checks.length === 0 && <EmptyRow text="موردی ثبت نشده" />}
          {loans.map((l) => (<Row key={l.id} title={l.title} subtitle="وام" value={`${toFaInt(l.principal)} ریال`} valueColor={BRAND.crimson} />))}
          {checks.filter((c) => c.status === "pending").map((c) => (<Row key={c.id} title={`${c.payee} (${c.type === "received" ? "دریافتی" : "پرداختی"})`} subtitle={c.dueDate} value={formatMoney(c.amount, currency, usdRate)} />))}
          <div style={{ display: "flex", gap: 14, justifyContent: "center", marginTop: 4 }}>
            <AddLink text="+ وام‌ها" onClick={openLoans} />
            <AddLink text="+ چک‌ها" onClick={openChecks} />
          </div>
        </>);
      case "contacts":
        return (<>
          <Row title="اشخاص و طرف حساب‌ها" subtitle={`${toFaInt(persons.length)} نفر`} value="مدیریت" valueColor={BRAND.violet} onClick={openPersons} />
          <Row title="بدهی و طلب" subtitle="مدیریت طرف حساب‌ها" value="مدیریت" valueColor={BRAND.orange} onClick={openDebts} />
        </>);
      case "bills":
        return (<>
          {bills.length === 0 && <EmptyRow text="قبضی ثبت نشده" />}
          {bills.map((b) => (<Row key={b.id} title={b.title} subtitle={b.dueDate} value={b.paid ? "پرداخت شده" : "پرداخت نشده"} valueColor={b.paid ? BRAND.darkgreen : BRAND.crimson} />))}
          <AddLink text="+ مدیریت قبض‌ها" onClick={openBills} />
        </>);
      default: return null;
    }
  }

  function sectionAddHandler(key) {
    switch (key) {
      case "shortcut": return openShortcuts;
      case "expense": return () => onAddTransaction("expense");
      case "income": return () => onAddTransaction("income");
      case "banks": return openAccounts;
      case "funds": return openAccounts;
      case "budget": return openBudgets;
      case "contacts": return openPersons;
      case "bills": return openBills;
      default: return undefined;
    }
  }

  return (
    <div style={{ padding: "18px 16px 8px" }}>
      <BankCardCarousel accounts={accounts} accountBalance={accountBalance} currency={currency} usdRate={usdRate} onAddCard={openAccounts} />
      <DateQuickBox onNote={onNote} onReminder={onReminder} />
      <IncomeExpenseDayWidget day={homeDay} setDay={setHomeDay} transactions={allTransactions} catById={catById} currency={currency} usdRate={usdRate} />

      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, margin: "16px 0" }}>
        <button onClick={() => setYear((y) => y - 1)} style={navArrowStyle(t)}><ChevronLeft size={16} /></button>
        <div style={{ background: t.card, borderRadius: 20, padding: "6px 18px", fontWeight: 700, color: BRAND.header, boxShadow: "0 1px 3px rgba(0,0,0,0.1)" }}>سال {faDigits(year)}</div>
        <button onClick={() => setYear((y) => y + 1)} style={navArrowStyle(t)}><ChevronRight size={16} /></button>
      </div>

      <div style={{ textAlign: "center", marginBottom: 12, fontSize: 13, color: net >= 0 ? BRAND.darkgreen : BRAND.crimson, fontWeight: 700 }}>
        مانده سالانه: {toFaInt(Math.abs(net))} ریال {net >= 0 ? "مثبت" : "منفی"}
      </div>

      {sectionOrder.map(({ key }) => {
        const m = SECTION_META[key];
        if (!m) return null;
        return (
          <CollapsibleSection key={key} color={m.color} title={m.title} open={!!open[key]} onToggle={() => toggle(key)} onAdd={sectionAddHandler(key)}>
            {sectionBody(key)}
          </CollapsibleSection>
        );
      })}
    </div>
  );
}

/* ---------------------------------------------------------
   Sub-Views Management & Modals
--------------------------------------------------------- */
const SUBVIEW_TITLES = {
  accounts: "مدیریت حساب‌ها و کارت‌ها",
  budgets: "بودجه‌بندی ماهیانه",
  bills: "یادآوری قبض‌ها",
  loans: "وام‌ها و تسهیلات",
  checks: "مدیریت چک‌ها",
  assets: "دارایی‌ها و بورس",
  persons: "اشخاص و طرف حساب‌ها",
  debts: "بدهی و طلب",
  shortcuts: "مدیریت میانبرها",
  reminders: "یادآوری‌های برنامه",
};

function SubViewContent({ subView, ctx, onBack }) {
  switch (subView) {
    case "accounts": return <AccountsManager accounts={ctx.accounts} onAdd={ctx.addAccount} onDelete={ctx.deleteAccount} onUpdate={ctx.updateAccount} />;
    case "budgets": return <BudgetsManager budgets={ctx.budgets} categories={ctx.categories} onSave={ctx.setBudgets} />;
    case "bills": return <BillsManager bills={ctx.bills} onUpdate={ctx.setBills} />;
    case "loans": return <LoansManager loans={ctx.loans} onUpdate={ctx.setLoans} />;
    case "checks": return <ChecksManager checks={ctx.checks} setChecks={ctx.setChecks} />;
    case "assets": return <AssetsManager assets={ctx.assets} onUpdate={ctx.setAssets} />;
    case "persons": return <PersonsManager persons={ctx.persons} setPersons={ctx.setPersons} />;
    case "debts": return <DebtsManager debts={ctx.debts} setDebts={ctx.setDebts} persons={ctx.persons} />;
    case "shortcuts": return <ShortcutsManager shortcuts={ctx.shortcuts} setShortcuts={ctx.setShortcuts} categories={ctx.categories} />;
    case "reminders": return <RemindersManager reminders={ctx.reminders} setReminders={ctx.setReminders} />;
    default: return <div style={{ padding: 20 }}>موردی یافت نشد.</div>;
  }
}

/* Accounts Manager with Full 16-digit Card Number & Expiry Date */
function AccountsManager({ accounts, onAdd, onDelete, onUpdate }) {
  const t = useT();
  const [name, setName] = useState("");
  const [type, setType] = useState("bank");
  const [initial, setInitial] = useState("");
  const [cardNumber, setCardNumber] = useState("");
  const [expiryDate, setExpiryDate] = useState("");

  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
        <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 10 }}>افزودن حساب یا کارت جدید</div>
        <input placeholder="نام حساب یا بانک" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle(t)} />
        <select value={type} onChange={(e) => setType(e.target.value)} style={inputStyle(t)}>
          <option value="bank">حساب بانکی</option>
          <option value="card">کارت بانکی</option>
          <option value="fund">صندوق نقدی</option>
        </select>
        {(type === "bank" || type === "card") && (
          <>
            <input placeholder="شماره کارت ۱۶ رقمی" maxLength={16} value={cardNumber} onChange={(e) => setCardNumber(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
            <input placeholder="تاریخ انقضا (مثل 1406/12)" value={expiryDate} onChange={(e) => setExpiryDate(e.target.value)} style={inputStyle(t)} />
          </>
        )}
        <input placeholder="موجودی اولیه (ریال)" inputMode="numeric" value={initial} onChange={(e) => setInitial(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
        <button onClick={() => {
          if (!name) return;
          onAdd({ name, type, initial: Number(initial || 0), cardNumber, expiryDate });
          setName(""); setInitial(""); setCardNumber(""); setExpiryDate("");
        }} style={{ width: "100%", padding: 11, background: BRAND.fab, color: "#fff", border: "none", borderRadius: 9, fontWeight: 700, cursor: "pointer" }}>افزودن</button>
      </div>

      {accounts.map((a) => (
        <div key={a.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>{a.name}</div>
            <div style={{ fontSize: 11.5, color: t.sub, marginTop: 2 }}>{a.cardNumber ? `شماره کارت: ${a.cardNumber}` : (a.type === "card" ? "کارت" : "بانک")} {a.expiryDate ? `| انقضا: ${a.expiryDate}` : ""}</div>
          </div>
          <button onClick={() => onDelete(a.id)} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={18} /></button>
        </div>
      ))}
    </div>
  );
}

function BudgetsManager({ budgets, categories, onSave }) {
  const t = useT();
  const expCats = categories.filter((c) => c.kind === "expense");
  return (
    <div style={{ padding: 16 }}>
      <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>تعیین سقف بودجه دسته‌ها</div>
      {expCats.map((cat) => {
        const item = budgets.find((b) => b.categoryId === cat.id);
        return (
          <div key={cat.id} style={{ background: t.card, borderRadius: 12, padding: 12, marginBottom: 10, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <span style={{ fontWeight: 600, fontSize: 13.5 }}>{cat.name}</span>
            <input
              type="text" inputMode="numeric" placeholder="مبلغ سقف (ریال)"
              value={item?.amount ? Number(item.amount).toLocaleString("en-US") : ""}
              onChange={(e) => {
                const val = Number(e.target.value.replace(/[^0-9]/g, ""));
                const exists = budgets.find((b) => b.categoryId === cat.id);
                if (exists) {
                  onSave(budgets.map((b) => b.categoryId === cat.id ? { ...b, amount: val } : b));
                } else {
                  onSave([...budgets, { id: uid(), categoryId: cat.id, amount: val }]);
                }
              }}
              style={{ width: 140, padding: "8px 10px", borderRadius: 8, border: `1px solid ${t.inputBorder}`, background: t.input, color: t.text, textAlign: "left" }}
            />
          </div>
        );
      })}
    </div>
  );
}

function BillsManager({ bills, onUpdate }) {
  const t = useT();
  const [title, setTitle] = useState("");
  const [dueDate, setDueDate] = useState(todayISO());
  const [amount, setAmount] = useState("");
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>افزودن قبض جدید</div>
        <input placeholder="عنوان قبض (آب، برق...)" value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle(t)} />
        <input type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} style={inputStyle(t)} />
        <input placeholder="مبلغ (ریال)" inputMode="numeric" value={amount} onChange={(e) => setAmount(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
        <button onClick={() => { if (!title) return; onUpdate([...bills, { id: uid(), title, dueDate, amount: Number(amount || 0), paid: false }]); setTitle(""); setAmount(""); }} style={primaryBtn}>ثبت قبض</button>
      </div>
      {bills.map((b) => (
        <div key={b.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontWeight: 700 }}>{b.title}</div>
            <div style={{ fontSize: 11.5, color: t.sub }}>سررسید: {b.dueDate} — {toFaInt(b.amount)} ریال</div>
          </div>
          <button onClick={() => onUpdate(bills.map((x) => x.id === b.id ? { ...x, paid: !x.paid } : x))} style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: b.paid ? BRAND.darkgreen : BRAND.orange, color: "#fff", fontWeight: 700, cursor: "pointer" }}>
            {b.paid ? "پرداخت شده" : "پرداخت"}
          </button>
        </div>
      ))}
    </div>
  );
}

function LoansManager({ loans, onUpdate }) {
  const t = useT();
  const [title, setTitle] = useState(""); const [principal, setPrincipal] = useState("");
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>افزودن وام جدید</div>
        <input placeholder="عنوان وام" value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle(t)} />
        <input placeholder="مبلغ وام (ریال)" inputMode="numeric" value={principal} onChange={(e) => setPrincipal(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
        <button onClick={() => { if (!title) return; onUpdate([...loans, { id: uid(), title, principal: Number(principal || 0), paidCount: 0 }]); setTitle(""); setPrincipal(""); }} style={primaryBtn}>ثبت وام</button>
      </div>
      {loans.map((l) => (
        <div key={l.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10 }}>
          <div style={{ fontWeight: 700 }}>{l.title}</div>
          <div style={{ fontSize: 12, color: t.sub, marginTop: 4 }}>مبلغ کل: {toFaInt(l.principal)} ریال</div>
        </div>
      ))}
    </div>
  );
}

function ChecksManager({ checks, setChecks }) {
  const t = useT();
  const [payee, setPayee] = useState(""); const [amount, setAmount] = useState("");
  const [dueDate, setDueDate] = useState(todayISO()); const [type, setType] = useState("received");
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>ثبت چک جدید</div>
        <input placeholder="در وجه / از طرف" value={payee} onChange={(e) => setPayee(e.target.value)} style={inputStyle(t)} />
        <input placeholder="مبلغ (ریال)" inputMode="numeric" value={amount} onChange={(e) => setAmount(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
        <input type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} style={inputStyle(t)} />
        <select value={type} onChange={(e) => setType(e.target.value)} style={inputStyle(t)}>
          <option value="received">چک دریافتی</option>
          <option value="issued">چک پرداختی</option>
        </select>
        <button onClick={() => { if (!payee || !amount) return; setChecks([...checks, { id: uid(), payee, amount: Number(amount), dueDate, type, status: "pending" }]); setPayee(""); setAmount(""); }} style={primaryBtn}>ثبت چک</button>
      </div>
      {checks.map((c) => (
        <div key={c.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontWeight: 700 }}>{c.payee} ({c.type === "received" ? "دریافتی" : "پرداختی"})</div>
            <div style={{ fontSize: 11.5, color: t.sub }}>سررسید: {c.dueDate} — {toFaInt(c.amount)} ریال</div>
          </div>
          <button onClick={() => setChecks(checks.map((x) => x.id === c.id ? { ...x, status: x.status === "pending" ? "passed" : "pending" } : x))} style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: c.status === "passed" ? BRAND.darkgreen : BRAND.orange, color: "#fff", fontWeight: 700, cursor: "pointer" }}>
            {c.status === "passed" ? "وصول شده" : "در انتظار"}
          </button>
        </div>
      ))}
    </div>
  );
}

function AssetsManager({ assets, onUpdate }) {
  const t = useT();
  const [name, setName] = useState(""); const [quantity, setQuantity] = useState(""); const [currentPrice, setCurrentPrice] = useState("");
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>افزودن دارایی (طلا، ارز، سهم...)</div>
        <input placeholder="نام دارایی" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle(t)} />
        <input placeholder="مقدار / تعداد" inputMode="numeric" value={quantity} onChange={(e) => setQuantity(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
        <input placeholder="قیمت واحد (ریال)" inputMode="numeric" value={currentPrice} onChange={(e) => setCurrentPrice(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
        <button onClick={() => { if (!name) return; onUpdate([...assets, { id: uid(), name, quantity: Number(quantity || 0), currentPrice: Number(currentPrice || 0) }]); setName(""); setQuantity(""); setCurrentPrice(""); }} style={primaryBtn}>ثبت دارایی</button>
      </div>
      {assets.map((a) => (
        <div key={a.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontWeight: 700 }}>{a.name}</div>
            <div style={{ fontSize: 11.5, color: t.sub }}>تعداد: {toFaInt(a.quantity)} | ارزش کل: {toFaInt(a.quantity * a.currentPrice)} ریال</div>
          </div>
          <button onClick={() => onUpdate(assets.filter((x) => x.id !== a.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={18} /></button>
        </div>
      ))}
    </div>
  );
}

function PersonsManager({ persons, setPersons }) {
  const t = useT();
  const [name, setName] = useState(""); const [phone, setPhone] = useState("");
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>افزودن شخص جدید</div>
        <input placeholder="نام و نام خانوادگی" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle(t)} />
        <input placeholder="شماره تماس" inputMode="tel" value={phone} onChange={(e) => setPhone(e.target.value)} style={inputStyle(t)} />
        <button onClick={() => { if (!name) return; setPersons([...persons, { id: uid(), name, phone }]); setName(""); setPhone(""); }} style={primaryBtn}>افزودن شخص</button>
      </div>
      {persons.map((p) => (
        <div key={p.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontWeight: 700 }}>{p.name}</div>
            {p.phone && <div style={{ fontSize: 11.5, color: t.sub }}>{p.phone}</div>}
          </div>
          <button onClick={() => setPersons(persons.filter((x) => x.id !== p.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={18} /></button>
        </div>
      ))}
    </div>
  );
}

function DebtsManager({ debts, setDebts, persons }) {
  const t = useT();
  const [personId, setPersonId] = useState(persons[0]?.id || "");
  const [amount, setAmount] = useState(""); const [kind, setKind] = useState("receivable");
  const [note, setNote] = useState("");

  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>ثبت بدهی یا طلب</div>
        <select value={personId} onChange={(e) => setPersonId(e.target.value)} style={inputStyle(t)}>
          {persons.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
        <select value={kind} onChange={(e) => setKind(e.target.value)} style={inputStyle(t)}>
          <option value="receivable">طلب (از دیگران)</option>
          <option value="payable">بدهی (به دیگران)</option>
        </select>
        <input placeholder="مبلغ (ریال)" inputMode="numeric" value={amount} onChange={(e) => setAmount(e.target.value.replace(/\D/g, ""))} style={inputStyle(t)} />
        <input placeholder="توضیحات" value={note} onChange={(e) => setNote(e.target.value)} style={inputStyle(t)} />
        <button onClick={() => { if (!amount) return; setDebts([...debts, { id: uid(), personId, amount: Number(amount), kind, note, settled: false }]); setAmount(""); setNote(""); }} style={primaryBtn}>ثبت</button>
      </div>
      {debts.map((d) => {
        const p = persons.find((x) => x.id === d.personId);
        return (
          <div key={d.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <div style={{ fontWeight: 700 }}>{p?.name || "ناشناس"} ({d.kind === "receivable" ? "طلب" : "بدهی"})</div>
              <div style={{ fontSize: 11.5, color: t.sub }}>{toFaInt(d.amount)} ریال — {d.note}</div>
            </div>
            <button onClick={() => setDebts(debts.map((x) => x.id === d.id ? { ...x, settled: !x.settled } : x))} style={{ padding: "6px 12px", borderRadius: 8, border: "none", background: d.settled ? BRAND.darkgreen : BRAND.orange, color: "#fff", fontWeight: 700, cursor: "pointer" }}>
              {d.settled ? "تسویه شده" : "تسویه نشده"}
            </button>
          </div>
        );
      })}
    </div>
  );
}

function ShortcutsManager({ shortcuts, setShortcuts }) {
  const t = useT();
  const [name, setName] = useState(""); const [type, setType] = useState("expense");
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>افزودن میانبر سریع</div>
        <input placeholder="عنوان میانبر (مثلاً ناهار)" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle(t)} />
        <select value={type} onChange={(e) => setType(e.target.value)} style={inputStyle(t)}>
          <option value="expense">هزینه</option>
          <option value="income">درآمد</option>
        </select>
        <button onClick={() => { if (!name) return; setShortcuts([...shortcuts, { id: uid(), name, type, icon: "fuel", note: name }]); setName(""); }} style={primaryBtn}>افزودن میانبر</button>
      </div>
      {shortcuts.map((sc) => (
        <div key={sc.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontWeight: 700 }}>{sc.name} ({sc.type === "expense" ? "هزینه" : "درآمد"})</div>
          <button onClick={() => setShortcuts(shortcuts.filter((x) => x.id !== sc.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={18} /></button>
        </div>
      ))}
    </div>
  );
}

function RemindersManager({ reminders, setReminders }) {
  const t = useT();
  const [title, setTitle] = useState(""); const [date, setDate] = useState(todayISO());
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10 }}>افزودن یادآوری</div>
        <input placeholder="متن یادآوری" value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle(t)} />
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={inputStyle(t)} />
        <button onClick={() => { if (!title) return; setReminders([...reminders, { id: uid(), title, date }]); setTitle(""); }} style={primaryBtn}>ثبت یادآوری</button>
      </div>
      {reminders.map((r) => (
        <div key={r.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontWeight: 700 }}>{r.title}</div>
            <div style={{ fontSize: 11.5, color: t.sub }}>تاریخ: {r.date}</div>
          </div>
          <button onClick={() => setReminders(reminders.filter((x) => x.id !== r.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={18} /></button>
        </div>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------
   Transactions View
--------------------------------------------------------- */
function TransactionsView({ transactions, catById, accById, filter, setFilter, onDelete, search, setSearch, onEdit }) {
  const t = useT();
  const filtered = transactions.filter((tx) => {
    if (filter !== "all" && tx.type !== filter) return false;
    if (search && !String(tx.note || "").includes(search)) return false;
    return true;
  });

  return (
    <div style={{ padding: 16 }}>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        <input placeholder="جستجو در یادداشت..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ ...inputStyle(t), marginBottom: 0, flex: 1 }} />
      </div>
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        <button onClick={() => setFilter("all")} style={pillStyle(filter === "all")}>همه</button>
        <button onClick={() => setFilter("expense")} style={pillStyle(filter === "expense")}>هزینه</button>
        <button onClick={() => setFilter("income")} style={pillStyle(filter === "income")}>درآمد</button>
      </div>

      {filtered.length === 0 && <EmptyRow text="تراکنشی یافت نشد" />}
      {filtered.map((tx) => {
        const cat = catById(tx.categoryId);
        const acc = accById(tx.accountId);
        const isExp = tx.type === "expense";
        return (
          <div key={tx.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, display: "flex", justifyContent: "space-between", alignItems: "center", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}>
            <div onClick={() => onEdit(tx)} style={{ cursor: "pointer", flex: 1 }}>
              <div style={{ fontWeight: 700, fontSize: 14 }}>{cat?.name || (tx.type === "transfer" ? "انتقال بین حساب" : "متفرقه")}</div>
              <div style={{ fontSize: 11.5, color: t.sub, marginTop: 2 }}>{acc?.name || ""} — {tx.date} {tx.note ? `| ${tx.note}` : ""}</div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <span style={{ fontWeight: 800, fontSize: 14, color: isExp ? BRAND.crimson : BRAND.darkgreen }}>
                {isExp ? "-" : "+"}{toFaInt(tx.amount)} ریال
              </span>
              <button onClick={() => onDelete(tx.id)} style={{ background: "none", border: "none", color: t.sub, cursor: "pointer" }}><Trash2 size={16} /></button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ---------------------------------------------------------
   Reports View
--------------------------------------------------------- */
function ReportsView({ expenseByCategory, incomeByCategory, totalIncomeYear, totalExpenseYear, accounts, accountBalance, currency, usdRate, transactions }) {
  const t = useT();
  const net = totalIncomeYear - totalExpenseYear;
  return (
    <div style={{ padding: 16 }}>
      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 12 }}>خلاصه عملکرد مالی سال</div>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, fontSize: 14 }}>
          <span style={{ color: t.sub }}>کل درآمد:</span>
          <span style={{ fontWeight: 800, color: BRAND.darkgreen }}>{formatMoney(totalIncomeYear, currency, usdRate)}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, fontSize: 14 }}>
          <span style={{ color: t.sub }}>کل هزینه:</span>
          <span style={{ fontWeight: 800, color: BRAND.crimson }}>{formatMoney(totalExpenseYear, currency, usdRate)}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14, borderTop: `1px solid ${t.border}`, paddingTop: 8 }}>
          <span style={{ fontWeight: 700 }}>خالص:</span>
          <span style={{ fontWeight: 800, color: net >= 0 ? BRAND.darkgreen : BRAND.crimson }}>{formatMoney(net, currency, usdRate)}</span>
        </div>
      </div>

      <div style={{ background: t.card, borderRadius: 14, padding: 16, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 10 }}>هزینه‌ها به تفکیک دسته</div>
        {expenseByCategory.map((e) => (
          <Row key={e.catId} title={e.name} value={formatMoney(e.amount, currency, usdRate)} valueColor={BRAND.crimson} />
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Add Transaction Sheet
--------------------------------------------------------- */
function AddTransactionSheet({ accounts, categories, initial, onClose, onSubmit }) {
  const t = useT();
  const [type, setType] = useState(initial?.type || "expense");
  const [amount, setAmount] = useState(initial?.amount ? String(initial.amount) : "");
  const [categoryId, setCategoryId] = useState(initial?.categoryId || categories.filter((c) => c.kind === (initial?.type || "expense"))[0]?.id || "");
  const [accountId, setAccountId] = useState(initial?.accountId || accounts[0]?.id || "");
  const [date, setDate] = useState(initial?.date || todayISO());
  const [note, setNote] = useState(initial?.note || "");

  const filteredCats = categories.filter((c) => c.kind === type);

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 400, display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
      <div style={{ background: t.card, width: "100%", maxWidth: 480, borderTopLeftRadius: 20, borderTopRightRadius: 20, padding: 20, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ fontWeight: 800, fontSize: 16 }}>{initial?._editId ? "ویرایش تراکنش" : "ثبت تراکنش جدید"}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: t.text }}><X size={22} /></button>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <button onClick={() => setType("expense")} style={pillStyle(type === "expense")}>هزینه</button>
          <button onClick={() => setType("income")} style={pillStyle(type === "income")}>درآمد</button>
        </div>

        <label style={labelStyle(t)}>مبلغ (ریال)</label>
        <input type="text" inputMode="numeric" placeholder="مثلاً 50000" value={amount ? Number(amount).toLocaleString("en-US") : ""} onChange={(e) => setAmount(e.target.value.replace(/[^0-9]/g, ""))} style={inputStyle(t)} />

        <label style={labelStyle(t)}>ده‌بندی</label>
        <select value={categoryId} onChange={(e) => setCategoryId(e.target.value)} style={inputStyle(t)}>
          {filteredCats.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>

        <label style={labelStyle(t)}>حساب / کارت</label>
        <select value={accountId} onChange={(e) => setAccountId(e.target.value)} style={inputStyle(t)}>
          {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
        </select>

        <label style={labelStyle(t)}>تاریخ</label>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={inputStyle(t)} />

        <label style={labelStyle(t)}>یادداشت</label>
        <input placeholder="توضیحات..." value={note} onChange={(e) => setNote(e.target.value)} style={inputStyle(t)} />

        <button onClick={() => {
          if (!amount) return;
          onSubmit({ type, amount: Number(amount), categoryId, accountId, date, note });
        }} style={{ width: "100%", padding: 13, background: BRAND.fab, color: "#fff", border: "none", borderRadius: 10, fontWeight: 700, fontSize: 15, cursor: "pointer", marginTop: 8 }}>
          ثبت تراکنش
        </button>
      </div>
    </div>
  );
}

/* Quick Add Sheet & Quick Modals */
function QuickAddSheet({ onClose, onPick }) {
  const t = useT();
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 400, display: "flex", alignItems: "flex-end", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: t.card, width: "100%", maxWidth: 480, borderTopLeftRadius: 20, borderTopRightRadius: 20, padding: 20 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 16, textAlign: "center" }}>انتخاب نوع ثبت</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <button onClick={() => onPick("expense")} style={quickBtnStyle(BRAND.crimson)}><TrendingDown size={22} /> ثبت هزینه</button>
          <button onClick={() => onPick("income")} style={quickBtnStyle(BRAND.darkgreen)}><TrendingUp size={22} /> ثبت درآمد</button>
          <button onClick={() => onPick("transfer")} style={quickBtnStyle(BRAND.violet)}><ArrowLeftRight size={22} /> انتقال بین حساب</button>
          <button onClick={() => onPick("check")} style={quickBtnStyle(BRAND.orange)}><FileSpreadsheet size={22} /> ثبت چک</button>
        </div>
      </div>
    </div>
  );
}
const quickBtnStyle = (color) => ({ background: color, color: "#fff", border: "none", borderRadius: 12, padding: "16px", fontWeight: 700, fontSize: 14, display: "flex", flexDirection: "column", alignItems: "center", gap: 8, cursor: "pointer" });

function SimpleTextModal({ title, placeholder, onClose, onSubmit }) {
  const t = useT();
  const [val, setVal] = useState("");
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 400, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: t.card, width: "100%", maxWidth: 360, borderRadius: 16, padding: 20 }}>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 12 }}>{title}</div>
        <textarea placeholder={placeholder} value={val} onChange={(e) => setVal(e.target.value)} style={{ ...inputStyle(t), height: 90, resize: "none" }} />
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onClose} style={{ flex: 1, padding: 10, background: t.border, border: "none", borderRadius: 8, fontWeight: 700, cursor: "pointer" }}>لغو</button>
          <button onClick={() => { if (val) onSubmit(val); }} style={{ flex: 1, padding: 10, background: BRAND.fab, color: "#fff", border: "none", borderRadius: 8, fontWeight: 700, cursor: "pointer" }}>ثبت</button>
        </div>
      </div>
    </div>
  );
}

function ReminderQuickModal({ onClose, onSubmit }) {
  const t = useT();
  const [title, setTitle] = useState(""); const [date, setDate] = useState(todayISO());
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 400, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: t.card, width: "100%", maxWidth: 360, borderRadius: 16, padding: 20 }}>
        <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 12 }}>یادآوری جدید</div>
        <input placeholder="متن یادآوری" value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle(t)} />
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={inputStyle(t)} />
        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
          <button onClick={onClose} style={{ flex: 1, padding: 10, background: t.border, border: "none", borderRadius: 8, fontWeight: 700, cursor: "pointer" }}>لغو</button>
          <button onClick={() => { if (title) onSubmit({ title, date }); }} style={{ flex: 1, padding: 10, background: BRAND.fab, color: "#fff", border: "none", borderRadius: 8, fontWeight: 700, cursor: "pointer" }}>ثبت</button>
        </div>
      </div>
    </div>
  );
}

function SideMenu({ onClose, setSubView, profileName }) {
  const t = useT();
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 500, display: "flex" }} onClick={onClose}>
      <div style={{ background: t.card, width: 280, height: "100%", padding: 20, display: "flex", flexDirection: "column", gap: 12 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ fontWeight: 800, fontSize: 18, marginBottom: 10, color: BRAND.header }}>Rexa — {profileName || "رضا"}</div>
        <MenuRow title="مدیریت حساب‌ها و کارت‌ها" onClick={() => setSubView("accounts")} />
        <MenuRow title="بودجه‌بندی ماهیانه" onClick={() => setSubView("budgets")} />
        <MenuRow title="یادآوری قبض‌ها" onClick={() => setSubView("bills")} />
        <MenuRow title="وام‌ها و تسهیلات" onClick={() => setSubView("loans")} />
        <MenuRow title="دارایی‌ها و بورس" onClick={() => setSubView("assets")} />
        <MenuRow title="اشخاص و طرف حساب‌ها" onClick={() => setSubView("persons")} />
        <MenuRow title="بدهی و طلب" onClick={() => setSubView("debts")} />
        <MenuRow title="مدیریت میانبرها" onClick={() => setSubView("shortcuts")} />
        <MenuRow title="یادآوری‌ها" onClick={() => setSubView("reminders")} />
      </div>
    </div>
  );
}
function MenuRow({ title, onClick }) {
  const t = useT();
  return <div onClick={onClick} style={{ padding: "12px 8px", borderRadius: 8, cursor: "pointer", fontWeight: 600, fontSize: 14, color: t.text, borderBottom: `1px solid ${t.border}` }}>{title}</div>;
}

const inputStyle = (t) => ({ width: "100%", padding: "10px 12px", borderRadius: 9, border: `1.5px solid ${t.inputBorder}`, marginBottom: 12, fontSize: 14, outline: "none", background: t.input, color: t.text });
const labelStyle = (t) => ({ display: "block", fontSize: 12.5, color: t.sub, fontWeight: 600, marginBottom: 6 });
const primaryBtn = { width: "100%", padding: "11px", borderRadius: 9, border: "none", background: BRAND.fab, color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer" };
function pillStyle(active) {
  return { flex: 1, padding: "8px 4px", borderRadius: 8, border: `1.5px solid ${active ? BRAND.header : "#ccc4d8"}`, background: active ? BRAND.header : "transparent", color: active ? "#fff" : "inherit", cursor: "pointer", fontWeight: 700, fontSize: 13 };
}

