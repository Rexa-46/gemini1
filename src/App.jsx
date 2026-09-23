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
function faTime(d) {
  try { return new Intl.DateTimeFormat("fa-IR", { hour: "2-digit", minute: "2-digit" }).format(d); }
  catch { return ""; }
}
function addMonths(dateStr, n) {
  const d = new Date(dateStr);
  d.setMonth(d.getMonth() + n);
  return d.toISOString().slice(0, 10);
}
function addDays(dateStr, n) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}
function daysUntil(dateStr) {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const d = new Date(dateStr); d.setHours(0, 0, 0, 0);
  return Math.round((d - now) / 86400000);
}
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const todayISO = () => new Date().toISOString().slice(0, 10);

function findJalaliMonthStart(jYear, jMonth) {
  // Farvardin 1 (Jalali New Year) falls around March 21 of (jYear + 621) in the Gregorian
  // calendar, and each subsequent Jalali month is ~30 days later. This anchor gets us close
  // enough that a modest search window reliably lands on the true day 1 of the target month.
  const approxNewYear = new Date(Date.UTC(jYear + 621, 2, 21));
  const anchor = new Date(approxNewYear);
  anchor.setUTCDate(anchor.getUTCDate() + (jMonth - 1) * 30);
  for (let delta = -25; delta <= 25; delta++) {
    const d = new Date(anchor); d.setUTCDate(d.getUTCDate() + delta);
    const p = jalaliParts(d);
    if (p.y === jYear && p.m === jMonth && p.day === 1) return d;
  }
  return null;
}
function getJalaliMonthCells(jYear, jMonth) {
  const start = findJalaliMonthStart(jYear, jMonth);
  if (!start) return [];
  const cells = []; let d = new Date(start);
  for (let i = 0; i < 32; i++) {
    const p = jalaliParts(d);
    if (p.y !== jYear || p.m !== jMonth) break;
    cells.push({ date: new Date(d), day: p.day, weekday: (d.getUTCDay() + 1) % 7 });
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return cells;
}

// Parses Persian number WORDS ("دویست و پنجاه هزار") into a numeric value.
// Needed because voice transcripts contain spoken numbers as words, not digits —
// the plain digit-regex below only catches typed/pasted numbers.
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
  let fromWords = false;
  if (amountMatch) {
    amount = Number(amountMatch[1]);
    const mult = { "هزار": 1e3, "میلیون": 1e6, "میلیارد": 1e9 }[amountMatch[2]] || 1;
    amount *= mult;
    if (amountMatch[3] === "تومان" || amountMatch[3] === "تومن") {
      // In spoken Persian, "۳۵۰ تومن" is commonly shorthand for 350 thousand Toman.
      if (!amountMatch[2] && amount < 10000) amount *= 1000;
      amount *= 10;
    }
  }
  if (!amount || !Number.isFinite(amount)) {
    const wordAmount = persianWordsToNumber(normalized);
    if (wordAmount !== null && wordAmount > 0) { amount = wordAmount; fromWords = true; }
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
  return { amount: amount && Number.isFinite(amount) ? Math.round(amount) : null, type, note: raw.replace(/\s+/g, " ").slice(0, 160), categoryHint: hint?.name, fromWords };
}

// Currency display: base unit stored everywhere internally is always Rial.
// This only affects how numbers are *shown*, so data entry / storage stays consistent.
// usdRialRate = how many Rials one US Dollar costs right now (from the live rates widget,
// or the manual value the user entered in Settings if the live fetch didn't work).
function formatMoney(amount, currency, usdRialRate) {
  const n = Number(amount || 0);
  if (currency === "toman") return `${toFaInt(Math.round(n / 10))} تومان`;
  if (currency === "usd") {
    if (!usdRialRate) return `${toFaInt(n)} ریال`; // no rate yet, fall back to Rial rather than guess
    return `$${(n / usdRialRate).toLocaleString("en-US", { maximumFractionDigits: 2 })}`;
  }
  return `${toFaInt(n)} ریال`;
}

function resizeImage(file, maxSize = 480) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new window.Image();
      img.onload = () => {
        let w = img.width, h = img.height;
        if (w > h) { if (w > maxSize) { h = h * maxSize / w; w = maxSize; } }
        else { if (h > maxSize) { w = w * maxSize / h; h = maxSize; } }
        const canvas = document.createElement("canvas");
        canvas.width = w; canvas.height = h;
        canvas.getContext("2d").drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/jpeg", 0.7));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/* ---------------------------------------------------------
   Seed data
--------------------------------------------------------- */
const seedAccounts = () => ([
  { id: uid(), name: "بانک ملت", type: "bank", initial: 5000000 },
  { id: uid(), name: "صندوق نقدی", type: "fund", initial: 800000 },
  { id: uid(), name: "کارت عابر ملت", type: "card", initial: 0 },
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
const seedSettings = () => ({
  theme: "light", pin: "", sharedFamily: false, themeColor: "purple",
  profile: { name: "alireza shadfar", phone: "", email: "" },
  homeLayout: "cards", homeSections: DEFAULT_HOME_SECTIONS,
  fontScale: 1, calendarMode: "jalali", currency: "rial",
  checkReminderDays: 7, smsNotif: false, smsAutoRead: false, biometricEnabled: false,
  aiProvider: "none", aiApiKey: "", manualUsdRate: "", shortcuts: [], lastRatesUpdate: "",
});

/* ---------------------------------------------------------
   Theme / palette
   BRAND is intentionally mutable: App() re-assigns its keys from the
   chosen preset on every render, so every component below (which reads
   BRAND.xxx directly at render time) automatically reflects the user's
   color choice without needing a context/hook everywhere.
--------------------------------------------------------- */
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

// Storage abstraction: uses the Claude-artifact window.storage API when present
// (always true inside claude.ai), and falls back to localStorage automatically
// when this same code runs as a standalone deployed app / APK build.
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

/* ---------------------------------------------------------
   Small UI atoms (theme aware)
--------------------------------------------------------- */
function useT() { return React.useContext(ThemeCtx); }
const ThemeCtx = React.createContext(THEME.light);

function GaugeCircle({ value, max, color, label }) {
  const t = useT();
  const size = 156, stroke = 10, r = (size - stroke) / 2, c = 2 * Math.PI * r;
  const frac = max > 0 ? Math.min(value / max, 1) : 0;
  const dash = Math.max(frac * c, value > 0 ? 6 : 0);
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8 }}>
      <div style={{ position: "relative", width: size, height: size }}>
        <svg width={size} height={size} style={{ transform: "rotate(-90deg)" }}>
          <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={t.border} strokeWidth={stroke} />
          <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={color} strokeWidth={stroke}
            strokeDasharray={`${dash} ${c - dash}`} strokeLinecap="round" style={{ transition: "stroke-dasharray .6s" }} />
        </svg>
        <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
          <div style={{ fontSize: 18, fontWeight: 800, color: t.text }}>{toFaInt(value)}</div>
          <div style={{ fontSize: 10.5, color: t.sub, marginTop: 2 }}>ریال</div>
        </div>
      </div>
      <div style={{ fontWeight: 700, color, fontSize: 14.5 }}>{label}</div>
    </div>
  );
}

/* ---------------------------------------------------------
   Exploding / stylized pie chart — click a slice to pop it out
   a little and see its label. Recharts doesn't do true 3D, so
   "3D" here means a soft drop-shadow + slightly thicker ring to
   give it some depth rather than a flat chart.
--------------------------------------------------------- */
function renderExplodingSlice(props) {
  const { cx, cy, innerRadius, outerRadius, startAngle, endAngle, fill, payload, isActive } = props;
  const RADIAN = Math.PI / 180;
  const midAngle = (startAngle + endAngle) / 2;
  const offset = isActive ? 14 : 0;
  const ox = Math.cos(-midAngle * RADIAN) * offset;
  const oy = Math.sin(-midAngle * RADIAN) * offset;
  return (
    <g transform={`translate(${ox},${oy})`} style={{ filter: "drop-shadow(0px 3px 4px rgba(0,0,0,0.35))" }}>
      <Sector cx={cx} cy={cy} innerRadius={innerRadius} outerRadius={isActive ? outerRadius + 6 : outerRadius}
        startAngle={startAngle} endAngle={endAngle} fill={fill} stroke="#fff" strokeWidth={2} />
    </g>
  );
}
function ExplodingPie({ data, height = 220, currency, usdRate }) {
  const t = useT();
  const [active, setActive] = useState(null);
  if (!data || data.length === 0) return <EmptyRow text="داده‌ای برای نمایش نیست" />;
  const item = active != null ? data[active] : null;
  return (
    <div>
      <ResponsiveContainer width="100%" height={height}>
        <PieChart>
          <Pie data={data} dataKey="amount" nameKey="name" innerRadius={height * 0.16} outerRadius={height * 0.36}
            paddingAngle={2}
            onClick={(_, i) => setActive(active === i ? null : i)}
            shape={(props) => renderExplodingSlice({ ...props, isActive: active === props.index })}>
            {data.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} cursor="pointer" />)}
          </Pie>
        </PieChart>
      </ResponsiveContainer>
      <div style={{ textAlign: "center", minHeight: 20, fontSize: 13, fontWeight: 700, color: item ? PIE_COLORS[active % PIE_COLORS.length] : t.sub }}>
        {item ? `${item.name} — ${formatMoney(item.amount, currency, usdRate)}` : "برای جزئیات روی هر بخش بزن"}
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, justifyContent: "center", marginTop: 8 }}>
        {data.map((d, i) => (
          <span key={i} onClick={() => setActive(active === i ? null : i)} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: t.sub, cursor: "pointer" }}>
            <span style={{ width: 9, height: 9, borderRadius: "50%", background: PIE_COLORS[i % PIE_COLORS.length] }} />
            {d.name}
          </span>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Half-circle speedometer gauge — e.g. check collection ratio
--------------------------------------------------------- */
function GaugeSpeedometer({ pct, label, colorFrom = "#B01E4A", colorTo = "#1E8449" }) {
  const t = useT();
  const size = 200, stroke = 16;
  const cx = size / 2, cy = size / 2 + 10, r = size / 2 - stroke;
  const clamped = Math.max(0, Math.min(100, pct));
  const angle = -180 + (clamped / 100) * 180; // -180 (left) .. 0 (right)
  const rad = (angle * Math.PI) / 180;
  const needleX = cx + r * 0.86 * Math.cos(rad);
  const needleY = cy + r * 0.86 * Math.sin(rad);
  const arc = (startDeg, endDeg, color) => {
    const s = (startDeg * Math.PI) / 180, e = (endDeg * Math.PI) / 180;
    const x1 = cx + r * Math.cos(s), y1 = cy + r * Math.sin(s);
    const x2 = cx + r * Math.cos(e), y2 = cy + r * Math.sin(e);
    const large = endDeg - startDeg > 180 ? 1 : 0;
    return <path d={`M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}`} stroke={color} strokeWidth={stroke} fill="none" strokeLinecap="round" />;
  };
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
      <svg width={size} height={size / 2 + 30}>
        {arc(-180, 0, t.border)}
        {arc(-180, angle, clamped >= 60 ? colorTo : clamped >= 30 ? "#C56A1F" : colorFrom)}
        <line x1={cx} y1={cy} x2={needleX} y2={needleY} stroke={t.text} strokeWidth={3} strokeLinecap="round" />
        <circle cx={cx} cy={cy} r={7} fill={t.text} />
      </svg>
      <div style={{ fontSize: 20, fontWeight: 800, color: t.text, marginTop: -6 }}>{toFaInt(Math.round(clamped))}٪</div>
      <div style={{ fontSize: 12.5, color: t.sub, fontWeight: 600 }}>{label}</div>
    </div>
  );
}

/* ---------------------------------------------------------
   Bank card carousel — swipeable, with a show/hide-balance eye
--------------------------------------------------------- */
function BankCard({ account, balance, hidden, currency, usdRate }) {
  const last4 = account.cardNumberLast4 || account.id.slice(-4).toUpperCase();
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
      <div style={{ fontSize: 17, letterSpacing: 3, fontWeight: 700, zIndex: 1, direction: "ltr", textAlign: "left" }}>
        •••• •••• •••• {last4}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", zIndex: 1 }}>
        <div style={{ fontSize: 11, opacity: 0.8 }}>{account.type === "card" ? "کارت" : "بانک"}</div>
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

/* ---------------------------------------------------------
   Today box: date + quick note / reminder buttons
--------------------------------------------------------- */
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
   Income/Expense day widget — day navigation + smile/frown arc
   + a clean donut of that day's categories
   (منحنی رو به پایین = ناترازی منفی، رو به بالا = مثبت، خط صاف = تراز)
   هیچ نوشته یا کادری داخل/روی منحنی نمی‌رود.
--------------------------------------------------------- */
function SmileFrownArc({ income, expense, pieData = [], currency, usdRate }) {
  const t = useT();
  const net = income - expense;
  const safeData = pieData.length ? pieData : [{ name: "بدون تراکنش", amount: 1 }];
  // رنگ و شکل منحنی
  const arcColor = net > 0 ? BRAND.darkgreen : net < 0 ? BRAND.crimson : "#94a3b8";
  const w = 220, h = 56, strokeW = 13;
  // path: مثبت = لبخند (وسط بالا)، منفی = اخم (وسط پایین)، صفر = خط صاف
  let d;
  if (net > 0) {
    d = `M 18 ${h * 0.70} Q ${w / 2} ${h * 0.14} ${w - 18} ${h * 0.70}`;
  } else if (net < 0) {
    d = `M 18 ${h * 0.30} Q ${w / 2} ${h * 0.90} ${w - 18} ${h * 0.30}`;
  } else {
    d = `M 18 ${h / 2} L ${w - 18} ${h / 2}`;
  }

  return (
    <div style={{ width: "100%", maxWidth: 340, margin: "4px auto 0" }}>
      {/* فقط منحنی — بدون هیچ نوشته یا کادر روی آن */}
      <div style={{ display: "flex", justifyContent: "center", marginBottom: 6 }}>
        <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ display: "block", overflow: "visible" }}>
          <path d={d} fill="none" stroke={arcColor} strokeWidth={strokeW} strokeLinecap="round" />
        </svg>
      </div>

      {/* درآمد و هزینه — کاملاً زیر منحنی، بدون همپوشانی */}
      <div style={{ display: "flex", justifyContent: "space-around", alignItems: "flex-start", padding: "0 8px 10px", gap: 8 }}>
        <div style={{ textAlign: "center", flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 12, color: t.sub, fontWeight: 700, marginBottom: 4 }}>درآمد</div>
          <div style={{ fontSize: 16, fontWeight: 900, color: BRAND.darkgreen, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
            {formatMoney(income, currency, usdRate)}
          </div>
        </div>
        <div style={{ textAlign: "center", flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 12, color: t.sub, fontWeight: 700, marginBottom: 4 }}>هزینه</div>
          <div style={{ fontSize: 16, fontWeight: 900, color: BRAND.crimson, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
            {formatMoney(expense, currency, usdRate)}
          </div>
        </div>
      </div>

      {/* دونات تمیز دسته‌ها — جدا از منحنی و اعداد */}
      <div style={{ position: "relative", width: "100%", height: 168, marginTop: 2 }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={safeData}
              dataKey="amount"
              nameKey="name"
              cx="50%"
              cy="50%"
              innerRadius="32%"
              outerRadius="52%"
              paddingAngle={3}
              stroke="#fff"
              strokeWidth={4}
              isAnimationActive={false}
            >
              {safeData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
            </Pie>
          </PieChart>
        </ResponsiveContainer>
        <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
          <div style={{ width: 52, height: 52, borderRadius: "50%", background: t.card }} />
        </div>
      </div>
      <div style={{ textAlign: "center", fontSize: 12, color: t.sub, marginTop: 2, fontWeight: 600 }}>
        برای جزئیات روی هر بخش بزن
      </div>
      {pieData.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, justifyContent: "center", marginTop: 8 }}>
          {pieData.slice(0, 6).map((d, i) => (
            <span key={i} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: t.sub }}>
              <span style={{ width: 8, height: 8, borderRadius: "50%", background: PIE_COLORS[i % PIE_COLORS.length] }} />
              {d.name}
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
        width: "100%", background: color, border: "none", color: "#fff", padding: "15px 16px",
        display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer", fontFamily: "inherit"
      }}>
        <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ width: 32, height: 32, borderRadius: "50%", background: "rgba(255,255,255,.18)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <ChevronDown size={17} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform .2s" }} />
          </span>
          {onAdd && (
            <button onClick={(e) => { e.stopPropagation(); onAdd(); }} style={{
              width: 30, height: 30, borderRadius: "50%", background: "rgba(255,255,255,.22)", border: "none",
              color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer"
            }}>
              <Plus size={16} />
            </button>
          )}
        </span>
        <span style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: 700, fontSize: 15 }}>
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
          <div style={{ fontSize: 14.5, fontWeight: 600, color: t.text }}>{title}</div>
          {subtitle && <div style={{ fontSize: 12, color: t.sub, marginTop: 2 }}>{subtitle}</div>}
        </div>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
        {extra}
        {value !== undefined && <span style={{ fontSize: 13.5, fontWeight: 700, color: valueColor || t.text }}>{value}</span>}
        {leftIcon && (
          <span style={{ width: 36, height: 36, borderRadius: 10, background: leftColor, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff" }}>
            {leftIcon}
          </span>
        )}
      </div>
    </div>
  );
}
function EmptyRow({ text }) { const t = useT(); return <div style={{ padding: "16px 4px", textAlign: "center", color: t.sub, fontSize: 13 }}>{text}</div>; }
function AddLink({ text, onClick }) { return <div onClick={onClick} style={{ padding: "10px 4px", color: BRAND.header, fontSize: 13, fontWeight: 700, cursor: "pointer", textAlign: "center" }}>{text}</div>; }
function SectionTitle({ text }) { const t = useT(); return <div style={{ fontWeight: 700, fontSize: 14, color: t.text, margin: "6px 4px 10px" }}>{text}</div>; }
function StatusBadge({ text, color }) { return <span style={{ fontSize: 11, fontWeight: 700, color: "#fff", background: color, padding: "3px 9px", borderRadius: 20 }}>{text}</span>; }

function pillStyle(active) {
  return { flex: 1, padding: "8px 4px", borderRadius: 8, border: `1.5px solid ${active ? BRAND.header : "#ccc4d8"}`, background: active ? BRAND.header : "transparent", color: active ? "#fff" : "inherit", cursor: "pointer", fontWeight: 700, fontSize: 13 };
}
function useStyles() {
  const t = useT();
  return {
    input: { width: "100%", padding: "10px 12px", borderRadius: 9, border: `1.5px solid ${t.inputBorder}`, marginBottom: 10, fontSize: 14, outline: "none", background: t.input, color: t.text },
    primaryBtn: { width: "100%", padding: "11px", borderRadius: 9, border: "none", background: BRAND.fab, color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer" },
    label: { display: "block", fontSize: 12.5, color: t.sub, fontWeight: 600, marginBottom: 6 },
    card: { background: t.card, borderRadius: 14, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" },
  };
}

/* ---------------------------------------------------------
   Amount input — shows a live thousands separator (۱٬۲۳۴٬۵۶۷)
   while typing; the value passed to onChange is always the plain
   numeric string (no commas), so nothing downstream needs to change.
--------------------------------------------------------- */
function AmountInput({ value, onChange, placeholder, style }) {
  const digitsOnly = (s) => s.replace(/[۰-۹]/g, (d) => "0123456789"["۰۱۲۳۴۵۶۷۸۹".indexOf(d)]).replace(/[^0-9]/g, "");
  const display = value ? Number(value).toLocaleString("en-US") : "";
  return (
    <input
      value={display}
      onChange={(e) => onChange(digitsOnly(e.target.value))}
      placeholder={placeholder}
      inputMode="numeric"
      style={{ ...style, direction: "ltr", textAlign: "right" }}
    />
  );
}

/* ---------------------------------------------------------
   Header + BottomNav
--------------------------------------------------------- */
function Header({ title = "Rexa", onMenu, back, onBack, onMic }) {
  return <div style={{ background: BRAND.header, color: "#fff", padding: "calc(env(safe-area-inset-top, 0px) + 10px) 12px 10px", display: "grid", gridTemplateColumns: "42px 1fr 76px", alignItems: "center", position: "sticky", top: 0, zIndex: 80 }}>
    <div style={{ display: "flex", justifyContent: "flex-start" }}>{back ? <button onClick={onBack} style={iconBtn}><ChevronRight size={24} /></button> : <button onClick={onMenu} style={iconBtn} aria-label="منوی برنامه"><Menu size={24} /></button>}</div>
    <div style={{ textAlign: "center", fontWeight: 800, fontSize: 17, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{title}</div>
    <div style={{ display: "flex", justifyContent: "flex-end" }}>{onMic && <VoiceCaptureButton onResult={onMic} />}</div>
  </div>;
}
const iconBtn = { background: "none", border: "none", color: "#fff", cursor: "pointer" };

function VoiceCaptureButton({ onResult, dark }) {
  const [listening, setListening] = useState(false);
  const isNative = typeof Capacitor !== "undefined" && Capacitor.isNativePlatform && Capacitor.isNativePlatform();

  async function startNative() {
    try {
      setListening(true);
      try { await SpeechRecognition.requestPermissions(); } catch { /* some versions prompt inside start() instead */ }
      const result = await SpeechRecognition.start({
        language: "fa-IR", maxResults: 1, prompt: "چی می‌خوای ثبت کنی؟", partialResults: false, popup: true,
      });
      setListening(false);
      const transcript = result?.matches?.[0];
      if (transcript) onResult(transcript);
      else alert("چیزی شنیده نشد، دوباره امتحان کن.");
    } catch (e) {
      setListening(false);
      alert("مشکلی در تشخیص گفتار پیش اومد. دوباره امتحان کن یا از دکمه‌ی سه‌بار لمس صفحه برای ثبت متنی استفاده کن.");
    }
  }

  function startWeb() {
    const Rec = typeof window !== "undefined" && (window.SpeechRecognition || window.webkitSpeechRecognition);
    if (!Rec) {
      alert("تشخیص گفتار روی این مرورگر پشتیبانی نمی‌شود. وقتی برنامه به‌صورت APK نصب بشه، از تشخیص گفتار واقعی گوشی استفاده می‌کنه. فعلاً می‌تونی از دکمه سه‌بار لمس صفحه برای ثبت متنی استفاده کنی.");
      return;
    }
    const rec = new Rec();
    rec.lang = "fa-IR";
    rec.interimResults = false;
    rec.maxAlternatives = 1;
    rec.onstart = () => setListening(true);
    rec.onend = () => setListening(false);
    rec.onerror = () => setListening(false);
    rec.onresult = (e) => {
      const transcript = e.results?.[0]?.[0]?.transcript;
      if (transcript) onResult(transcript);
    };
    try { rec.start(); } catch { setListening(false); }
  }

  function start() { if (isNative) startNative(); else startWeb(); }

  return (
    <button onClick={start} style={{ ...iconBtn, color: dark ? (listening ? BRAND.orange : BRAND.header) : (listening ? "#ffd166" : "#fff") }} title="ثبت با صدا">
      <Mic size={20} />
    </button>
  );
}

function BottomNav({ active, setActive, onAdd }) {
  const t = useT();
  const items = [
    { key: "reports", label: "گزارش‌ها", icon: PieChartIcon },
    { key: "transactions", label: "تراکنش‌ها", icon: Receipt },
    { key: "checks", label: "چک‌ها", icon: FileSpreadsheet },
    { key: "home", label: "خانه", icon: HomeIcon },
  ];
  return <div style={{ position: "fixed", left: "50%", transform: "translateX(-50%)", bottom: 0, width: "min(480px, 100vw)", minHeight: 72, padding: "7px 6px calc(7px + env(safe-area-inset-bottom, 0px))", background: t.card, borderTop: `1px solid ${t.border}`, display: "grid", gridTemplateColumns: "1fr 1fr 64px 1fr 1fr", alignItems: "center", zIndex: 300, boxShadow: "0 -5px 18px rgba(0,0,0,.10)" }}>
    <NavBtn it={items[0]} active={active} setActive={setActive} />
    <NavBtn it={items[1]} active={active} setActive={setActive} />
    <button onClick={onAdd} aria-label="انتقال بین حساب‌ها" style={{ width: 56, height: 56, borderRadius: "50%", background: BRAND.fab, border: `4px solid ${t.card}`, margin: "-20px auto 0", boxShadow: "0 4px 12px rgba(0,0,0,.25)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff" }}>
      <ArrowLeftRight size={25} />
    </button>
    <NavBtn it={items[2]} active={active} setActive={setActive} />
    <NavBtn it={items[3]} active={active} setActive={setActive} />
  </div>;
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
   Lock screen
--------------------------------------------------------- */
function LockScreen({ pin, onUnlock, biometricEnabled }) {
  const [val, setVal] = useState("");
  const [err, setErr] = useState(false);
  const [bioAvailable, setBioAvailable] = useState(false);
  const [bioChecked, setBioChecked] = useState(false);
  const isNative = typeof Capacitor !== "undefined" && Capacitor.isNativePlatform && Capacitor.isNativePlatform();

  async function tryBiometric() {
    if (!biometricEnabled || !isNative) { setBioChecked(true); return; }
    try {
      const info = await BiometricAuth.checkBiometry();
      setBioAvailable(!!info.isAvailable);
      if (info.isAvailable) {
        try {
          await BiometricAuth.authenticate({
            reason: "برای باز کردن Rexa هویتت رو تایید کن",
            cancelTitle: "لغو",
            allowDeviceCredential: true,
          });
          onUnlock();
          return;
        } catch { /* user cancelled or failed — fall back to PIN, shown below */ }
      }
    } catch { /* biometric not available on this device */ }
    setBioChecked(true);
  }
  useEffect(() => { tryBiometric(); /* eslint-disable-next-line */ }, []);

  return (
    <div dir="rtl" style={{ minHeight: "100vh", background: BRAND.header, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: FONT, color: "#fff", gap: 18, maxWidth: 480, margin: "0 auto" }}>
      <RexaLogo size={64} />
      <div style={{ fontWeight: 700, fontSize: 16 }}>Rexa قفل است</div>
      {biometricEnabled && isNative && bioAvailable && (
        <button onClick={tryBiometric} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6, background: "none", border: "none", color: "#fff", cursor: "pointer" }}>
          <span style={{ width: 60, height: 60, borderRadius: "50%", background: "rgba(255,255,255,0.15)", display: "flex", alignItems: "center", justifyContent: "center" }}><Fingerprint size={30} /></span>
          <span style={{ fontSize: 12 }}>ورود با اثر انگشت</span>
        </button>
      )}
      <input
        type="password" inputMode="numeric" maxLength={6} value={val}
        onChange={(e) => { setVal(e.target.value.replace(/[^0-9]/g, "")); setErr(false); }}
        placeholder="رمز عبور"
        style={{ width: 180, textAlign: "center", fontSize: 22, letterSpacing: 6, padding: "10px", borderRadius: 10, border: "none", outline: "none" }}
      />
      {err && <div style={{ color: "#ffb3c1", fontSize: 13 }}>رمز اشتباه است</div>}
      <button onClick={() => (val === pin ? onUnlock() : setErr(true))}
        style={{ background: BRAND.fab, color: "#fff", border: "none", borderRadius: 9, padding: "10px 30px", fontWeight: 700, cursor: "pointer" }}>
        باز کردن
      </button>
      {biometricEnabled && !isNative && bioChecked && (
        <div style={{ fontSize: 11.5, color: "#d8c9e8", marginTop: 6, maxWidth: 260, textAlign: "center" }}>
          قفل اثر انگشت فقط در نسخه‌ی نصب‌شده‌ی APK کار می‌کند؛ در این پیش‌نمایش وب فقط رمز عددی در دسترس است.
        </div>
      )}
    </div>
  );
}

/* ===========================================================
   MAIN APP
=========================================================== */
/* ---------------------------------------------------------
   Rexa logo — fantasy coin-and-monogram mark, matches the app icon.
   Pure inline SVG so it renders inside the app itself (header,
   splash, side menu) without needing an image asset.
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
        <g fill="#fff" opacity="0.9">
          <path d="M162 44 l6 14 14 6 -14 6 -6 14 -6 -14 -14 -6 14 -6 z" />
          <path d="M34 138 l4 9 9 4 -9 4 -4 9 -4 -9 -9 -4 9 -4 z" opacity="0.7" />
        </g>
      </svg>
      {showWordmark && (
        <div style={{ fontWeight: 800, fontSize: size * 0.3, color: "#fff", letterSpacing: 1, fontFamily: "Georgia, serif" }}>Rexa</div>
      )}
    </div>
  );
}

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
  const [showCapture, setShowCapture] = useState(false);
  const [rates, setRates] = useState(null); // { usd: rialsPerUsd, fetchedAt }
  const [shortcuts, setShortcuts] = useState(() => settings.shortcuts || []);
  const tapTimesRef = useRef([]);
  function handleTripleTap() {
    const now = Date.now();
    const recent = [...tapTimesRef.current.filter((tms) => now - tms < 600), now];
    tapTimesRef.current = recent;
    if (recent.length >= 3) {
      tapTimesRef.current = [];
      setShowCapture(true);
    }
  }

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
      await reloadAll(s.sharedFamily);
      setLoaded(true);
    })();
    // eslint-disable-next-line
  }, []);

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
  useEffect(() => { if (loaded) saveKey("hs:favorites", favorites, shared); }, [favorites, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:members", members, shared); }, [members, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:events", events, shared); }, [events, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:projects", projects, shared); }, [projects, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:fiscalPeriods", fiscalPeriods, shared); }, [fiscalPeriods, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:notes", notes, shared); }, [notes, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:reminders", reminders, shared); }, [reminders, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:persons", persons, shared); }, [persons, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:debts", debts, shared); }, [debts, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:currencies", currencies, shared); }, [currencies, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:goals", goals, shared); }, [goals, loaded, shared]);
  useEffect(() => { if (loaded) setSettings((prev) => prev.shortcuts === shortcuts ? prev : { ...prev, shortcuts }); }, [shortcuts, loaded]);

  // process recurring templates once after load
  useEffect(() => {
    if (!loaded) return;
    let changed = false;
    const newTx = [];
    const updated = recurring.map((r) => {
      if (!r.active) return r;
      let next = r.nextDate; let guard = 0; let rr = { ...r };
      while (next <= todayISO() && guard < 24) {
        newTx.push({ id: uid(), type: rr.type, amount: rr.amount, categoryId: rr.categoryId, accountId: rr.accountId, date: next, note: rr.note || "تراکنش تکرارشونده", tags: ["تکرارشونده"], createdAt: new Date().toISOString(), recurringId: rr.id });
        next = rr.interval === "weekly" ? addDays(next, 7) : addMonths(next, 1);
        guard++; changed = true;
      }
      return { ...rr, nextDate: next };
    });
    if (changed) {
      setRecurring(updated);
      setTransactions((prev) => [...newTx, ...prev]);
    }
    // eslint-disable-next-line
  }, [loaded]);

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

  const totalBalance = useMemo(() => accounts.reduce((s, a) => s + accountBalance(a.id), 0), [accounts, accountBalance]);
  const totalAssets = useMemo(() => assets.reduce((s, a) => s + (a.quantity * a.currentPrice || 0), 0), [assets]);

  const yearTx = useMemo(() => transactions.filter((t) => jalaliYear(new Date(t.date)) === year), [transactions, year]);
  const totalIncomeYear = useMemo(() => yearTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0), [yearTx]);
  const totalExpenseYear = useMemo(() => yearTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0), [yearTx]);
  const gaugeMax = Math.max(totalIncomeYear, totalExpenseYear, 1);

  const expenseByCategory = useMemo(() => {
    const map = {};
    yearTx.filter((t) => t.type === "expense").forEach((t) => { map[t.categoryId] = (map[t.categoryId] || 0) + t.amount; });
    return Object.entries(map).map(([catId, amount]) => ({ catId, amount, name: catById(catId)?.name || "بدون دسته" })).sort((a, b) => b.amount - a.amount);
  }, [yearTx, catById]);
  const incomeByCategory = useMemo(() => {
    const map = {};
    yearTx.filter((t) => t.type === "income").forEach((t) => { map[t.categoryId] = (map[t.categoryId] || 0) + t.amount; });
    return Object.entries(map).map(([catId, amount]) => ({ catId, amount, name: catById(catId)?.name || "بدون دسته" })).sort((a, b) => b.amount - a.amount);
  }, [yearTx, catById]);

  function groupBy(idField, list) {
    const map = {};
    yearTx.filter((tx) => tx.type === "expense" && tx[idField]).forEach((tx) => { map[tx[idField]] = (map[tx[idField]] || 0) + tx.amount; });
    return Object.entries(map).map(([id, amount]) => ({ id, amount, name: list.find((x) => x.id === id)?.name || "—" })).sort((a, b) => b.amount - a.amount);
  }
  const expenseByMember = useMemo(() => groupBy("memberId", members), [yearTx, members]);
  const expenseByEvent = useMemo(() => groupBy("eventId", events), [yearTx, events]);
  const expenseByProject = useMemo(() => groupBy("projectId", projects), [yearTx, projects]);

  // net worth trend: last 8 months
  const netWorthTrend = useMemo(() => {
    const points = [];
    for (let i = 7; i >= 0; i--) {
      const d = new Date(); d.setMonth(d.getMonth() - i);
      const cutoff = d.toISOString().slice(0, 10);
      let bal = accounts.reduce((s, a) => s + (a.initial || 0), 0);
      transactions.forEach((t) => {
        if (t.date > cutoff) return;
        if (t.type === "expense") bal -= t.amount;
        if (t.type === "income") bal += t.amount;
      });
      points.push({ name: faMonthYear(d).split(" ")[0], مانده: bal });
    }
    return points;
  }, [accounts, transactions]);

  function addTransaction(tx) { setTransactions((p) => [{ ...tx, id: uid(), createdAt: new Date().toISOString() }, ...p]); }
  function deleteTransaction(id) { setTransactions((p) => p.filter((t) => t.id !== id)); }
  function addAccount(a) { const item = { ...a, id: a.id || uid() }; setAccounts((p) => [...p, item]); return item.id; }
  function deleteAccount(id) { setAccounts((p) => p.filter((a) => a.id !== id)); }
  function updateAccount(id, patch) { setAccounts((p) => p.map((a) => a.id === id ? { ...a, ...patch } : a)); }
  function addCategory(c) { const item = { ...c, id: c.id || uid() }; setCategories((p) => [...p, item]); return item.id; }
  function deleteCategory(id) { setCategories((p) => p.filter((c) => c.id !== id && c.parentId !== id)); }
  function updateCategory(id, patch) { setCategories((p) => p.map((c) => c.id === id ? { ...c, ...patch } : c)); }
  function upsertBudget(categoryId, amount) {
    setBudgets((prev) => prev.find((b) => b.categoryId === categoryId)
      ? prev.map((b) => (b.categoryId === categoryId ? { ...b, amount } : b))
      : [...prev, { id: uid(), categoryId, amount }]);
  }
  function toggleFavorite(kind, id) {
    setFavorites((prev) => {
      const list = prev[kind];
      return { ...prev, [kind]: list.includes(id) ? list.filter((x) => x !== id) : [...list, id] };
    });
  }

  // نرخ دلار بازار آزاد از TGJU؛ در APK با CapacitorHttp مستقیم به اینترنت وصل می‌شویم.
  // اگر TGJU موقتاً در دسترس نباشد، نرخ قبلی/دستی حفظ می‌شود تا عدد اشتباه نمایش داده نشود.
  async function fetchRates() {
    const url = "https://gem.tgju.org/profile/price_dollar_rl";
    try {
      let html = "";
      if (Capacitor.isNativePlatform()) {
        const r = await CapacitorHttp.get({ url, headers: { Accept: "text/html" } });
        html = String(r?.data || "");
      } else {
        const r = await fetch(url, { cache: "no-store" });
        html = await r.text();
      }
      const patterns = [
        /نرخ فعلی[^0-9]{0,80}([0-9,]{6,})/i,
        /نرخ فعلی\s*:\s*([0-9,]{6,})/i,
        /Last[^0-9]{0,40}([0-9,]{6,})/i,
        /price_dollar_rl[^0-9]{0,120}([0-9,]{6,})/i
      ];
      let irr = 0;
      for (const re of patterns) {
        const m = html.match(re);
        if (m) { irr = Number(String(m[1]).replace(/,/g, "")); if (irr > 0) break; }
      }
      if (!Number.isFinite(irr) || irr <= 0) throw new Error("TGJU dollar rate not found");
      const fetchedAt = new Date().toISOString();
      setRates({ usd: irr, fetchedAt, source: "tgju", sourceLabel: "TGJU · دلار آزاد" });
      setSettings((prev) => ({ ...prev, lastRatesUpdate: fetchedAt }));
    } catch {
      const manual = Number(String(settings.manualUsdRate || "").replace(/,/g, ""));
      setRates((prev) => prev?.usd ? { ...prev, source: prev.source || "cached", fetchedAt: prev.fetchedAt || new Date().toISOString() } : { usd: manual > 0 ? manual : null, fetchedAt: new Date().toISOString(), source: manual > 0 ? "manual" : "failed", sourceLabel: manual > 0 ? "دستی" : "" });
    }
  }
  useEffect(() => { if (loaded) fetchRates(); /* eslint-disable-next-line */ }, [loaded]);
  useEffect(() => {
    if (!loaded || settings.manualUsdRate === undefined) return;
    const manual = Number(String(settings.manualUsdRate || "").replace(/,/g, ""));
    if (manual > 0 && rates?.source === "manual") setRates((r) => ({ ...r, usd: manual }));
  }, [settings.manualUsdRate, loaded]);

  const backupState = { backupVersion: 2, exportedAt: new Date().toISOString(), accounts, categories, transactions, budgets, loans, checks, bills, assets, recurring, favorites, settings, shortcuts, members, events, projects, fiscalPeriods, notes, reminders, persons, debts, currencies, goals };
  async function exportBackup() {
    const filename = `rexa-backup-${todayISO()}.json`;
    const json = JSON.stringify(backupState, null, 2);
    if (Capacitor.isNativePlatform()) {
      try {
        const result = await Filesystem.writeFile({ path: filename, data: json, directory: Directory.Cache, encoding: Encoding.UTF8 });
        await Share.share({ title: "پشتیبان Rexa", text: "فایل پشتیبان Rexa", files: [result.uri], dialogTitle: "ذخیره یا ارسال پشتیبان Rexa" });
        return;
      } catch {}
    }
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob); const a = document.createElement("a");
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }
  function importBackup(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        if (data.accounts) setAccounts(data.accounts);
        if (data.categories) setCategories(data.categories);
        if (data.transactions) setTransactions(data.transactions);
        if (data.budgets) setBudgets(data.budgets);
        if (data.loans) setLoans(data.loans);
        if (data.checks) setChecks(data.checks);
        if (data.bills) setBills(data.bills);
        if (data.assets) setAssets(data.assets);
        if (data.recurring) setRecurring(data.recurring);
        if (data.favorites) setFavorites(data.favorites);
        if (data.members) setMembers(data.members);
        if (data.events) setEvents(data.events);
        if (data.projects) setProjects(data.projects);
        if (data.fiscalPeriods) setFiscalPeriods(data.fiscalPeriods);
        if (data.notes) setNotes(data.notes);
        if (data.reminders) setReminders(data.reminders);
        if (data.persons) setPersons(data.persons);
        if (data.debts) setDebts(data.debts);
        if (data.currencies) setCurrencies(data.currencies);
        if (data.goals) setGoals(data.goals);
        if (data.shortcuts) setShortcuts(data.shortcuts);
        if (data.settings) setSettings((s) => ({ ...s, ...data.settings }));
        alert("بازیابی اطلاعات با موفقیت انجام شد");
      } catch { alert("فایل پشتیبان نامعتبر است"); }
    };
    reader.readAsText(file);
  }
  function exportExcel() {
    const rows = transactions.map((t) => ({
      نوع: t.type === "expense" ? "پرداخت" : t.type === "income" ? "دریافت" : "انتقال",
      مبلغ: t.amount, تاریخ: t.date,
      دسته: catById(t.categoryId)?.name || "", حساب: accById(t.accountId)?.name || "",
      برچسب: (t.tags || []).join("، "), یادداشت: t.note || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "تراکنش‌ها");
    XLSX.writeFile(wb, `transactions-${todayISO()}.xlsx`);
  }

  const performExit = useCallback(async () => {
    if (!window.confirm("آیا از خروج از Rexa مطمئن هستید؟")) return;
    const wantBackup = window.confirm("قبل از خروج پشتیبان‌گیری شود؟\n\nبله = پشتیبان‌گیری و خروج\nخیر = خروج بدون پشتیبان");
    if (wantBackup) { try { await exportBackup(); } catch {} }
    try { if (Capacitor.isNativePlatform()) await CapacitorApp.exitApp(); else window.history.back(); } catch {}
  }, [settings, accounts, categories, transactions, budgets, loans, checks, bills, assets, recurring, favorites, members, events, projects, fiscalPeriods, notes, reminders, shortcuts]);

  useEffect(() => {
    if (!loaded || !Capacitor.isNativePlatform()) return;
    let listener;
    (async () => {
      listener = await CapacitorApp.addListener("backButton", async () => {
        if (showAdd) { setShowAdd(false); setPrefillTx(null); return; }
        if (showQuickAdd) { setShowQuickAdd(false); return; }
        if (showCapture) { setShowCapture(false); return; }
        if (menuOpen) { setMenuOpen(false); return; }
        if (subView) { setSubView(null); return; }
        if (tab !== "home") { setTab("home"); return; }
        await performExit();
      });
    })();
    return () => { listener?.remove?.(); };
  }, [loaded, showAdd, showQuickAdd, showCapture, menuOpen, subView, tab, performExit]);

  const registerVoiceTransaction = useCallback((transcript) => {
    const parsed = parseBankSms(transcript);
    if (!parsed.amount) { alert("مبلغ تشخیص داده نشد. مثال: «امروز ۳۵۰ تومن برای شام دادم»."); return; }
    const sameKind = categories.filter((c) => c.kind === parsed.type);
    const hinted = parsed.categoryHint ? sameKind.find((c) => c.name === parsed.categoryHint) : null;
    const fallback = sameKind.find((c) => c.name === (parsed.type === "expense" ? "متفرقه" : "درآمد متفرقه")) || sameKind[0];
    const accountId = accounts[0]?.id;
    if (!accountId) { alert("برای ثبت صوتی ابتدا یک حساب بسازید."); return; }
    addTransaction({ type: parsed.type, amount: parsed.amount, categoryId: hinted?.id || fallback?.id || "", accountId, date: todayISO(), note: `ثبت صوتی: ${parsed.note}`, tags: ["ثبت صوتی"] });
    alert(`${toFaInt(parsed.amount)} ریال بابت «${hinted?.name || fallback?.name || "بدون دسته"}» ثبت شد.`);
  }, [categories, accounts]);
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

  function openWithPrefill(data) {
    setPrefillTx(data);
    setShowAdd(true);
  }

  async function rebuildData() {
    if (!window.confirm("بازسازی اطلاعات انجام شود؟ قبل از بازسازی یک پشتیبان خودکار تهیه می‌شود.")) return;
    await exportBackup();
    const normalize = (arr, defaults = []) => (Array.isArray(arr) ? arr : defaults).map((x) => ({ ...x, id: x?.id || uid() }));
    setAccounts(normalize(accounts, seedAccounts()));
    setCategories(normalize(categories, seedCategories()));
    setTransactions(normalize(transactions));
    setBudgets(normalize(budgets)); setLoans(normalize(loans)); setChecks(normalize(checks));
    setBills(normalize(bills)); setAssets(normalize(assets)); setRecurring(normalize(recurring));
    setMembers(normalize(members)); setEvents(normalize(events)); setProjects(normalize(projects));
    setFiscalPeriods(normalize(fiscalPeriods)); setNotes(normalize(notes)); setReminders(normalize(reminders));
    setPersons(normalize(persons)); setDebts(normalize(debts)); setCurrencies(normalize(currencies)); setGoals(normalize(goals));
    setFavorites({ categories: (favorites?.categories || []).filter((id) => categories.some((c) => c.id === id)), accounts: (favorites?.accounts || []).filter((id) => accounts.some((a) => a.id === id)) });
    alert("بازسازی اطلاعات با موفقیت انجام شد.");
  }

  async function clearAllData() {
    if (!window.confirm("تمام اطلاعات مالی و تراکنش‌ها پاک می‌شود. ابتدا پشتیبان تهیه شود؟")) return;
    await exportBackup();
    if (!window.confirm("تأیید نهایی: اطلاعات مالی این دستگاه خام شود؟ این عملیات پس از اجرا با پشتیبان قابل بازگردانی است.")) return;
    setAccounts(seedAccounts()); setCategories(seedCategories()); setTransactions([]); setBudgets([]); setLoans([]); setChecks([]);
    setBills([]); setAssets([]); setRecurring([]); setFavorites({ categories: [], accounts: [] }); setMembers([]); setEvents([]);
    setProjects([]); setFiscalPeriods([]); setNotes([]); setReminders([]); setPersons([]); setDebts([]); setCurrencies([]); setGoals([]);
    alert("اطلاعات مالی خام شد. پشتیبان قبلی را در صورت نیاز بازیابی کنید.");
  }

  const ctx = {
    accounts, addAccount, deleteAccount, updateAccount, accountBalance,
    categories, addCategory, deleteCategory, updateCategory,
    budgets, upsertBudget, expenseByCategory, incomeByCategory,
    loans, setLoans, checks, setChecks, bills, setBills, assets, setAssets,
    recurring, setRecurring, favorites, toggleFavorite, shortcuts, setShortcuts,
    settings, setSettings, exportBackup, importBackup, exportExcel,
    transactions, catById, accById, totalBalance, totalAssets,
    members, setMembers, events, setEvents, projects, setProjects,
    persons, setPersons, debts, setDebts, currencies, setCurrencies, goals, setGoals,
    fiscalPeriods, setFiscalPeriods, openWithPrefill,
    notes, setNotes, reminders, setReminders, rates, fetchRates,
    addTransaction, rebuildData, clearAllData,
  };

  return (
    <ThemeCtx.Provider value={t}>
      <div dir="rtl" onClick={handleTripleTap} style={{ fontFamily: FONT, background: t.bg, color: t.text, minHeight: "100vh", maxWidth: 480, margin: "0 auto", display: "flex", flexDirection: "column", position: "relative", paddingBottom: "calc(96px + env(safe-area-inset-bottom, 0px))", boxShadow: "0 0 30px rgba(0,0,0,0.08)", zoom: settings.fontScale || 1 }}>
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700;800&display=swap');
          * { box-sizing: border-box; } html, body, #root { margin: 0; min-height: 100%; width: 100%; } body { overflow-x: hidden; overscroll-behavior-y: none; }
          input, select, textarea, button { font-family: inherit; }
          ::-webkit-scrollbar { width: 0; height: 0; }
          @media print { .no-print { display: none !important; } }
        `}</style>

        {subView ? (
          <Header title={SUBVIEW_TITLES[subView]} back onBack={() => setSubView(null)} />
        ) : (
          <Header title={settings.profile?.name || "alireza shadfar"} onMenu={() => setMenuOpen(true)}
            onMic={registerVoiceTransaction} />
        )}

        {subView ? (
          <SubViewContent subView={subView} ctx={ctx} onBack={() => setSubView(null)} />
        ) : (
          <div style={{ flex: 1, overflowY: "auto", paddingBottom: 8 }}>
            {tab === "home" && (
              <HomeView
                year={year} setYear={setYear}
                totalIncomeYear={totalIncomeYear} totalExpenseYear={totalExpenseYear} gaugeMax={gaugeMax}
                open={open} toggle={toggle}
                accounts={accounts} accountBalance={accountBalance}
                expenseByCategory={expenseByCategory} incomeByCategory={incomeByCategory}
                budgets={budgets} categories={categories} transactions={yearTx} allTransactions={transactions}
                bills={bills} loans={loans} checks={checks} assets={assets} totalAssets={totalAssets} persons={persons} debts={debts}
                openAccounts={() => setSubView("accounts")} openBudgets={() => setSubView("budgets")}
                openBills={() => setSubView("bills")} openLoans={() => setSubView("loans")} openPersons={() => setSubView("persons")} openDebts={() => setSubView("debts")}
                openChecks={() => setSubView("checks")} openAssets={() => setSubView("assets")}
                homeDay={homeDay} setHomeDay={setHomeDay} catById={catById}
                settings={settings} setSettings={setSettings} rates={rates} fetchRates={fetchRates}
                onNote={() => setShowNoteModal(true)} onReminder={() => setShowReminderModal(true)}
                shortcuts={shortcuts}
                onRunShortcut={(sc) => { setShowQuickAdd(false); setShowAdd(true); setPrefillTx({ type: sc.type, categoryId: sc.categoryId, accountId: sc.accountId, note: sc.note || `میانبر: ${sc.name}` }); }}
                openShortcuts={() => setSubView("shortcuts")}
                onAddTransaction={(type) => { setPrefillTx(type ? { type } : null); setShowAdd(true); }}
              />
            )}
            {tab === "transactions" && (
              <TransactionsView
                transactions={transactions} catById={catById} accById={accById} checks={checks}
                filter={txFilter} setFilter={setTxFilter} onDelete={deleteTransaction}
                onEdit={(tx) => { setPrefillTx({ ...tx, _editId: tx.id }); setShowAdd(true); }}
                search={txSearch} setSearch={setTxSearch}
              />
            )}
            {tab === "operations" && <OperationsView setSubView={setSubView} onAdd={() => setShowAdd(true)} />}
            {tab === "checks" && <ChecksManager checks={checks} setChecks={setChecks} />}
            {tab === "reports" && (
              <ReportsView
                expenseByCategory={expenseByCategory} incomeByCategory={incomeByCategory}
                totalIncomeYear={totalIncomeYear} totalExpenseYear={totalExpenseYear}
                accounts={accounts} accountBalance={accountBalance}
                netWorthTrend={netWorthTrend} exportExcel={exportExcel}
                expenseByMember={expenseByMember} expenseByEvent={expenseByEvent} expenseByProject={expenseByProject}
                categories={categories} checks={checks} currency={settings.currency} usdRate={rates?.usd}
                transactions={transactions} updateAccount={updateAccount}
              />
            )}
          </div>
        )}

        <BottomNav
          active={subView ? null : tab}
          setActive={(k) => {
            setShowAdd(false); setPrefillTx(null); setShowQuickAdd(false); setShowCapture(false); setMenuOpen(false); setSubView(null);
            setTab(k);
          }}
          onAdd={() => {
            setShowAdd(false); setPrefillTx(null); setShowCapture(false); setMenuOpen(false); setSubView(null);
            setShowQuickAdd(true);
          }}
        />

        {showQuickAdd && (
          <QuickAddSheet
            onClose={() => setShowQuickAdd(false)}
            onPick={(type) => {
              setShowQuickAdd(false);
              if (type === "check") { setTab("checks"); return; }
              setPrefillTx(type === "expense" ? { type: "expense" } : type === "income" ? { type: "income" } : type === "transfer" ? { type: "transfer" } : null);
              setShowAdd(true);
            }}
          />
        )}

        {showNoteModal && (
          <SimpleTextModal title="یادداشت جدید" placeholder="یادداشتت رو بنویس..."
            onClose={() => setShowNoteModal(false)}
            onSubmit={(text) => { setNotes((p) => [{ id: uid(), text, date: todayISO(), createdAt: new Date().toISOString() }, ...p]); setShowNoteModal(false); }} />
        )}
        {showReminderModal && (
          <ReminderQuickModal onClose={() => setShowReminderModal(false)}
            onSubmit={(r) => { setReminders((p) => [{ id: uid(), ...r }, ...p]); setShowReminderModal(false); }} />
        )}
        {showCapture && (
          <SmartCaptureOverlay
            onClose={() => setShowCapture(false)}
            onParsed={(data) => { setShowCapture(false); openWithPrefill(data); }}
          />
        )}

        {showAdd && (
          <AddTransactionSheet
            accounts={accounts} categories={categories} favorites={favorites}
            onAddAccount={addAccount} onAddCategory={addCategory}
            members={members} events={events} projects={projects}
            initial={prefillTx}
            onClose={() => { setShowAdd(false); setPrefillTx(null); }}
            onSubmit={(tx) => {
              if (prefillTx?._editId) setTransactions((prev) => prev.map((item) => item.id === prefillTx._editId ? { ...item, ...tx, id: item.id, createdAt: item.createdAt } : item));
              else addTransaction(tx);
              setShowAdd(false); setPrefillTx(null);
            }}
          />
        )}
        {menuOpen && <SideMenu onClose={() => setMenuOpen(false)} setSubView={(v) => { setSubView(v); setMenuOpen(false); }} profileName={settings.profile?.name} />}
      </div>
    </ThemeCtx.Provider>
  );
}

/* ---------------------------------------------------------
   Home View
--------------------------------------------------------- */
function HomeView({
  year, setYear, totalIncomeYear, totalExpenseYear, gaugeMax, open, toggle,
  accounts, accountBalance, expenseByCategory, incomeByCategory, budgets, categories,
  transactions, allTransactions, bills, loans, checks, assets, totalAssets, persons = [], debts = [],
  openAccounts, openBudgets, openBills, openLoans, openChecks, openAssets, openPersons, openDebts,
  homeDay, setHomeDay, catById, settings, setSettings, rates, fetchRates, onNote, onReminder, onAddTransaction, shortcuts = [], onRunShortcut, openShortcuts
}) {
  const t = useT();
  const banks = accounts.filter((a) => a.type === "bank");
  const funds = accounts.filter((a) => a.type === "fund");
  const cardAccs = accounts.filter((a) => a.type === "card");
  const net = totalIncomeYear - totalExpenseYear;
  const upcomingBills = bills.filter((b) => !b.paid && daysUntil(b.dueDate) <= 5).sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  const upcomingChecks = checks.filter((c) => c.status === "pending" && daysUntil(c.dueDate) <= (settings.checkReminderDays || 7)).sort((a, b) => a.dueDate.localeCompare(b.dueDate));
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
                // پالت‌های رنگی شبیه تصویر میانبرها
                const palette = sc.type === "income"
                  ? ["#cffafe", "#0e7490"]
                  : (sc.icon === "taxi" || sc.icon === "bus" || sc.icon === "car" || sc.icon === "fuel")
                    ? ["#fef3c7", "#b45309"]
                    : (sc.icon === "water" || sc.icon === "electric" || sc.icon === "home")
                      ? ["#fef9c3", "#a16207"]
                      : (sc.icon === "food" || sc.icon === "market")
                        ? ["#fef9c3", "#a16207"]
                        : (sc.icon === "fun" || sc.icon === "internet" || sc.icon === "clothes")
                          ? ["#fce7f3", "#be185d"]
                          : (sc.icon === "bank" || sc.icon === "phone")
                            ? ["#fee2e2", "#b91c1c"]
                            : ["#fff0c8", "#d5a600"];
                const iconText = SHORTCUT_ICONS.find(([k]) => k === sc.icon)?.[1] || "✦";
                return (
                  <button key={sc.id} onClick={() => onRunShortcut?.(sc)} style={{ border: 0, background: "transparent", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, padding: 2, cursor: "pointer" }}>
                    <span style={{ width: 58, height: 58, borderRadius: 14, background: palette[0], color: palette[1], display: "flex", alignItems: "center", justifyContent: "center", fontSize: 28, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>{iconText}</span>
                    <span style={{ fontSize: 11.5, fontWeight: 700, color: t.text, textAlign: "center", lineHeight: 1.3 }}>{sc.name}</span>
                  </button>
                );
              })}
              <button onClick={openShortcuts || (() => {})} style={{ border: 0, background: "transparent", display: "flex", flexDirection: "column", alignItems: "center", gap: 6, padding: 2, cursor: "pointer" }}>
                <span style={{ width: 58, height: 58, borderRadius: 14, background: "#e0f2fe", color: BRAND.violet, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 30, boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>+</span>
                <span style={{ fontSize: 11.5, fontWeight: 700, color: t.text, textAlign: "center" }}>افزودن میانبر</span>
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
            <Row key={a.id} title={a.name} subtitle={a.type === "card" ? "کارت" : "بانک"} value={formatMoney(accountBalance(a.id), currency, usdRate)}
              valueColor={accountBalance(a.id) >= 0 ? t.text : BRAND.crimson} />
          ))}
          <AddLink text="+ مدیریت حساب‌ها و کارت‌ها" onClick={openAccounts} />
        </>);
      case "funds":
        return funds.length === 0 ? <EmptyRow text="صندوقی ثبت نشده" /> : funds.map((a) => <Row key={a.id} title={a.name} value={formatMoney(accountBalance(a.id), currency, usdRate)} valueColor={accountBalance(a.id) >= 0 ? t.text : BRAND.crimson} />);
      case "balrep":
        return accounts.map((a) => (
          <Row key={a.id} title={a.name} subtitle={a.type === "bank" ? "بانک" : a.type === "card" ? "کارت" : "صندوق"}
            value={formatMoney(accountBalance(a.id), currency, usdRate)} valueColor={accountBalance(a.id) >= 0 ? BRAND.darkgreen : BRAND.crimson} />
        ));
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
          {loans.map((l) => (
            <Row key={l.id} title={l.title} subtitle="وام" value={`${toFaInt(l.principal - (l.paidCount || 0) * l.monthlyPayment)} ریال باقی‌مانده`} valueColor={BRAND.crimson} />
          ))}
          {checks.filter((c) => c.status === "pending").map((c) => (
            <Row key={c.id} title={`${c.payee} (${c.type === "received" ? "دریافتی" : "پرداختی"})`} subtitle={faLongDate(new Date(c.dueDate))} value={formatMoney(c.amount, currency, usdRate)} />
          ))}
          <div style={{ display: "flex", gap: 14, justifyContent: "center", marginTop: 4 }}>
            <AddLink text="+ وام‌ها" onClick={openLoans} />
            <AddLink text="+ چک‌ها" onClick={openChecks} />
            <AddLink text="+ دارایی‌ها" onClick={openAssets} />
          </div>
        </>);
      case "contacts":
        return (<><Row title="اشخاص و طرف حساب‌ها" subtitle={`${toFaInt(persons.length)} نفر`} value="مدیریت" valueColor={BRAND.violet} onClick={openPersons} />
          <Row title="بدهی و طلب" subtitle="مدیریت طرف حساب‌ها" value="مدیریت" valueColor={BRAND.orange} onClick={openDebts} /><Row title="طلب‌های باز" value={`${toFaInt(debts.filter((d) => !d.settled && d.kind === "receivable").reduce((s, d) => s + Number(d.amount || 0), 0))} ریال`} valueColor={BRAND.darkgreen} /><Row title="بدهی‌های باز" value={`${toFaInt(debts.filter((d) => !d.settled && d.kind === "payable").reduce((s, d) => s + Number(d.amount || 0), 0))} ریال`} valueColor={BRAND.crimson} /></>);
      case "bills":
        return (<>
          {bills.length === 0 && <EmptyRow text="قبضی ثبت نشده" />}
          {bills.map((b) => (
            <Row key={b.id} title={b.title} subtitle={faLongDate(new Date(b.dueDate))}
              value={b.paid ? "پرداخت شده" : `${toFaInt(daysUntil(b.dueDate))} روز`}
              valueColor={b.paid ? BRAND.darkgreen : daysUntil(b.dueDate) < 0 ? BRAND.crimson : BRAND.orange} />
          ))}
          <AddLink text="+ مدیریت قبض‌ها" onClick={openBills} />
        </>);
      default: return null;
    }
  }

  function sectionAddHandler(key) {
    switch (key) {
      case "shortcut": return openShortcuts || (() => onAddTransaction());
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
      <CurrencyRatesStrip rates={rates} fetchRates={fetchRates} />

      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, margin: "18px 0" }}>
        <button onClick={() => setYear((y) => y - 1)} style={navArrowStyle(t)}><ChevronLeft size={16} /></button>
        <div style={{ background: t.card, borderRadius: 20, padding: "6px 18px", fontWeight: 700, color: BRAND.header, boxShadow: "0 1px 3px rgba(0,0,0,0.1)" }}>سال {faDigits(year)}</div>
        <button onClick={() => setYear((y) => y + 1)} style={navArrowStyle(t)}><ChevronRight size={16} /></button>
        <button onClick={() => setSettings((s) => ({ ...s, homeLayout: layout === "cards" ? "icons" : "cards" }))}
          style={{ ...navArrowStyle(t), width: "auto", padding: "0 10px", borderRadius: 16, display: "flex", alignItems: "center", gap: 4, fontSize: 11 }}>
          {layout === "cards" ? <LayoutGrid size={15} /> : <LayoutList size={15} />}
        </button>
      </div>

      <div style={{ textAlign: "center", marginBottom: 12, fontSize: 13, color: net >= 0 ? BRAND.darkgreen : BRAND.crimson, fontWeight: 700 }}>
        مانده سالانه: {toFaInt(Math.abs(net))} ریال {net >= 0 ? "مثبت" : "منفی"}
      </div>
      {totalAssets > 0 && (
        <div style={{ textAlign: "center", marginBottom: 18, fontSize: 12.5, color: BRAND.orange, fontWeight: 700 }}>
          ارزش دارایی‌های دیجیتال/بورس: {toFaInt(totalAssets)} ریال
        </div>
      )}

      {(upcomingBills.length > 0 || upcomingChecks.length > 0) && (
        <div style={{ background: "#fff6ea", border: "1px solid #f0d9a8", borderRadius: 12, padding: 12, marginBottom: 16, display: "flex", flexDirection: "column", gap: 6 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, color: BRAND.orange, fontWeight: 700, fontSize: 13 }}><BellRing size={16} /> یادآوری‌های نزدیک</div>
          {upcomingBills.map((b) => (
            <div key={b.id} style={{ fontSize: 12.5, color: "#6b4c14" }}>قبض {b.title} — {daysUntil(b.dueDate) < 0 ? "سررسید گذشته" : `${toFaInt(daysUntil(b.dueDate))} روز مانده`}</div>
          ))}
          {upcomingChecks.map((c) => (
            <div key={c.id} style={{ fontSize: 12.5, color: "#6b4c14" }}>چک {c.type === "received" ? "دریافتی" : "پرداختی"} {c.payee} — {daysUntil(c.dueDate) < 0 ? "سررسید گذشته" : `${toFaInt(daysUntil(c.dueDate))} روز مانده`}</div>
          ))}
        </div>
      )}

      {layout === "cards" ? (
        sectionOrder.map(({ key }) => {
          const m = SECTION_META[key];
          if (!m) return null;
          return (
            <CollapsibleSection key={key} color={m.color} title={m.title} open={!!open[key]} onToggle={() => toggle(key)} onAdd={sectionAddHandler(key)}>
              {sectionBody(key)}
            </CollapsibleSection>
          );
        })
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 20 }}>
          {sectionOrder.map(({ key }) => {
            const m = SECTION_META[key];
            if (!m) return null;
            return (
              <button key={key} onClick={() => toggle(key)} style={{
                display: "flex", flexDirection: "column", alignItems: "center", gap: 8, background: t.card, border: "none",
                borderRadius: 14, padding: "16px 6px", cursor: "pointer", boxShadow: "0 1px 3px rgba(0,0,0,0.08)"
              }}>
                <span style={{ width: 44, height: 44, borderRadius: 12, background: m.color, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center" }}>{m.icon}</span>
                <span style={{ fontSize: 11, fontWeight: 700, color: t.text, textAlign: "center" }}>{m.title}</span>
              </button>
            );
          })}
        </div>
      )}
      {layout === "icons" && sectionOrder.map(({ key }) => {
        const m = SECTION_META[key];
        if (!m || !open[key]) return null;
        return (
          <div key={key} style={{ background: t.card, borderRadius: 14, padding: "10px 14px", marginBottom: 14, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
              <div style={{ fontWeight: 700, color: m.color, fontSize: 13.5 }}>{m.title}</div>
              {sectionAddHandler(key) && (
                <button onClick={sectionAddHandler(key)} style={{ width: 26, height: 26, borderRadius: "50%", background: m.color, border: "none", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
                  <Plus size={14} />
                </button>
              )}
            </div>
            {sectionBody(key)}
          </div>
        );
      })}
    </div>
  );
}
const navArrowStyle = (t) => ({ width: 30, height: 30, borderRadius: "50%", border: "none", background: t.card, color: BRAND.header, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 1px 3px rgba(0,0,0,0.1)", cursor: "pointer" });

function CurrencyRatesStrip({ rates, fetchRates }) {
  const t = useT();
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 16, fontSize: 12, color: t.sub }}>
      <DollarSign size={13} />
      {rates?.usd ? (
        <span>هر دلار ≈ {toFaInt(Math.round(rates.usd))} ریال {rates.source === "tgju" ? "· بازار آزاد TGJU" : rates.source === "manual" ? "(دستی)" : ""}</span>
      ) : (
        <span>نرخ ارز در دسترس نیست</span>
      )}
      <button onClick={fetchRates} style={{ background: "none", border: "none", cursor: "pointer", color: t.sub, display: "flex" }}><RefreshCw size={13} /></button>
    </div>
  );
}


/* ---------------------------------------------------------
   Transactions View
--------------------------------------------------------- */
function TransactionSummaryChart({ transactions, onSelect }) {
  const t = useT();
  const [selected, setSelected] = useState(null);
  const summary = useMemo(() => {
    const incomeTx = transactions.filter((x) => x.type === "income");
    const expenseTx = transactions.filter((x) => x.type === "expense");
    return [
      { name: "هزینه", amount: expenseTx.reduce((s, x) => s + x.amount, 0), type: "expense", count: expenseTx.length },
      { name: "درآمد", amount: incomeTx.reduce((s, x) => s + x.amount, 0), type: "income", count: incomeTx.length },
    ];
  }, [transactions]);
  if (!summary.some((x) => x.amount > 0)) return null;
  const detail = selected ? summary.find((x) => x.type === selected) : null;
  return <div style={{ background: t.card, borderRadius: 14, padding: 12, margin: "8px 14px 16px", boxShadow: "0 1px 4px rgba(0,0,0,.07)" }}>
    <div style={{ fontWeight: 800, fontSize: 13.5, marginBottom: 6 }}>خلاصه همه تراکنش‌ها</div>
    <ResponsiveContainer width="100%" height={190}>
      <BarChart data={summary} onClick={(e) => { const p = e?.activePayload?.[0]?.payload; if (p) { setSelected(p.type); onSelect?.(p.type); } }}>
        <XAxis dataKey="name" tick={{ fontFamily: FONT, fontSize: 11 }} />
        <YAxis tick={{ fontFamily: FONT, fontSize: 10 }} width={45} />
        <Tooltip formatter={(v) => `${toFaInt(v)} ریال`} contentStyle={{ fontFamily: FONT, direction: "rtl" }} />
        <Bar dataKey="amount" radius={[8,8,0,0]} cursor="pointer">
          {summary.map((x, i) => <Cell key={i} fill={x.type === "income" ? BRAND.darkgreen : BRAND.crimson} />)}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
    {detail ? (
      <div style={{ background: detail.type === "income" ? "#eef8f1" : "#fff0f3", borderRadius: 10, padding: 10, textAlign: "center", fontSize: 12.5, fontWeight: 700, color: detail.type === "income" ? BRAND.darkgreen : BRAND.crimson }}>
        {detail.name}: {toFaInt(detail.amount)} ریال · {toFaInt(detail.count)} تراکنش
      </div>
    ) : <div style={{ textAlign: "center", color: t.sub, fontSize: 11.5 }}>روی ستون درآمد یا هزینه لمس کن تا جزئیات و همان موارد را ببینی.</div>}
  </div>;
}

function TransactionsView({ transactions, catById, accById, checks = [], filter, setFilter, onDelete, onEdit, search, setSearch }) {
  const t = useT();
  const [preview, setPreview] = useState(null);
  const checkStatusLabel = { pending: "در انتظار", cashed: "نقد شده", bounced: "برگشتی" };
  const checkPseudo = checks.map((c) => ({
    id: "chk-" + c.id, __isCheck: true, type: "check", amount: c.amount, date: c.dueDate,
    note: `چک ${c.type === "received" ? "دریافتی" : "پرداختی"} — ${c.payee}`, checkStatus: c.status,
  }));
  const combined = [...transactions, ...checkPseudo];
  const filtered = combined.filter((tx) => {
    if (filter !== "all" && tx.type !== filter) return false;
    if (search.trim()) {
      const q = search.trim();
      const hay = `${catById(tx.categoryId)?.name || ""} ${tx.note || ""} ${(tx.tags || []).join(" ")} ${tx.amount}`;
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  const grouped = useMemo(() => {
    const map = {};
    filtered.forEach((tx) => { (map[tx.date] = map[tx.date] || []).push(tx); });
    return Object.entries(map).sort((a, b) => new Date(b[0]) - new Date(a[0]));
  }, [filtered]);
  const tabs = [{ key: "all", label: "همه" }, { key: "expense", label: "پرداخت ها" }, { key: "income", label: "دریافت ها" }, { key: "transfer", label: "انتقال ها" }, { key: "check", label: "چک ها" }];

  return (
    <div>
      <div style={{ padding: "10px 14px 0" }}>
        <div style={{ position: "relative" }}>
          <Search size={16} color={t.sub} style={{ position: "absolute", top: 12, right: 12 }} />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="جستجو در تراکنش‌ها، برچسب یا یادداشت..."
            style={{ width: "100%", padding: "10px 36px 10px 12px", borderRadius: 10, border: `1.5px solid ${t.inputBorder}`, background: t.input, color: t.text, fontSize: 13.5, outline: "none", marginBottom: 10 }} />
        </div>
      </div>
      <div style={{ display: "flex", background: BRAND.header, padding: "0 8px 12px", gap: 4 }}>
        {tabs.map((tItem) => (
          <button key={tItem.key} onClick={() => setFilter(tItem.key)} style={{ flex: 1, padding: "8px 4px", border: "none", borderRadius: 8, cursor: "pointer", background: filter === tItem.key ? "rgba(255,255,255,0.18)" : "transparent", color: "#fff", fontWeight: filter === tItem.key ? 700 : 500, fontSize: 12.5 }}>{tItem.label}</button>
        ))}
      </div>
      <div style={{ padding: "12px 14px" }}>
        {grouped.length === 0 && <EmptyRow text="تراکنشی یافت نشد" />}
        {grouped.map(([date, txs]) => {
          const dayTotal = txs.reduce((s, tx) => s + (tx.type === "expense" ? -tx.amount : tx.type === "income" ? tx.amount : 0), 0);
          return (
            <div key={date} style={{ marginBottom: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <span style={{ fontSize: 12.5, color: t.sub, fontWeight: 600 }}>{faLongDate(new Date(date))}</span>
                <span style={{ fontSize: 12.5, fontWeight: 700, color: dayTotal >= 0 ? BRAND.darkgreen : BRAND.crimson }}>{dayTotal >= 0 ? "+" : "-"}{toFaInt(Math.abs(dayTotal))} ریال</span>
              </div>
              {txs.map((tx) => {
                if (tx.__isCheck) {
                  const statusColor2 = { pending: BRAND.gold, cashed: BRAND.darkgreen, bounced: BRAND.crimson }[tx.checkStatus];
                  return (
                    <div key={tx.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, boxShadow: "0 1px 3px rgba(0,0,0,0.08)", border: `1px dashed ${statusColor2}55` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                        <div style={{ display: "flex", gap: 10 }}>
                          <span style={{ width: 38, height: 38, borderRadius: 10, background: statusColor2 + "22", color: statusColor2, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><FileSpreadsheet size={18} /></span>
                          <div>
                            <div style={{ fontWeight: 700, fontSize: 14 }}>{tx.note}</div>
                            <div style={{ fontSize: 12, color: t.sub, marginTop: 3 }}>سررسید چک — <StatusBadge text={checkStatusLabel[tx.checkStatus]} color={statusColor2} /></div>
                          </div>
                        </div>
                        <div style={{ textAlign: "left" }}>
                          <div style={{ fontWeight: 800, color: statusColor2, fontSize: 14.5 }}>{toFaInt(tx.amount)}</div>
                        </div>
                      </div>
                    </div>
                  );
                }
                const cat = catById(tx.categoryId), acc = accById(tx.accountId), toAcc = accById(tx.toAccountId);
                const color = tx.type === "expense" ? BRAND.crimson : tx.type === "income" ? BRAND.darkgreen : BRAND.violet;
                const Icon = tx.type === "expense" ? TrendingDown : tx.type === "income" ? TrendingUp : ArrowLeftRight;
                return (
                  <div key={tx.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                      <div style={{ display: "flex", gap: 10 }}>
                        <span style={{ width: 38, height: 38, borderRadius: 10, background: color + "22", color, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><Icon size={18} /></span>
                        <div>
                          <div style={{ fontWeight: 700, fontSize: 14 }}>{tx.type === "transfer" ? `انتقال به ${toAcc?.name || "—"}` : (cat?.name || "بدون دسته")}</div>
                          <div style={{ fontSize: 12, color: t.sub, marginTop: 3 }}>{tx.type === "expense" ? "از حساب" : tx.type === "income" ? "به حساب" : "از حساب"}: {acc?.name || "—"}</div>
                          {tx.note && <div style={{ fontSize: 12, color: t.sub, marginTop: 2 }}>{tx.note}</div>}
                          {(tx.tags || []).length > 0 && (
                            <div style={{ display: "flex", gap: 5, marginTop: 5, flexWrap: "wrap" }}>
                              {tx.tags.map((tag, i) => <span key={i} style={{ fontSize: 10.5, background: t.border, color: t.sub, padding: "2px 8px", borderRadius: 20 }}>{tag}</span>)}
                            </div>
                          )}
                          {tx.photo && <img src={tx.photo} onClick={() => setPreview(tx.photo)} style={{ width: 44, height: 44, borderRadius: 8, objectFit: "cover", marginTop: 6, cursor: "pointer" }} />}
                        </div>
                      </div>
                      <div style={{ textAlign: "left" }}>
                        <div style={{ fontWeight: 800, color, fontSize: 14.5 }}>{toFaInt(tx.amount)}</div>
                        <div style={{ display: "flex", gap: 4, marginTop: 6 }}>
                          <button onClick={() => onEdit?.(tx)} style={{ background: "none", border: "none", color: BRAND.violet, cursor: "pointer", padding: 3 }} title="ویرایش"><Pencil size={15} /></button>
                          <button onClick={() => { if (window.confirm("این تراکنش حذف شود؟")) onDelete(tx.id); }} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer", padding: 3 }} title="حذف"><Trash2 size={15} /></button>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
      <TransactionSummaryChart transactions={transactions} onSelect={(kind) => setFilter(kind)} />
      {preview && (
        <div onClick={() => setPreview(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 70, maxWidth: 480, margin: "0 auto" }}>
          <img src={preview} style={{ maxWidth: "88%", maxHeight: "70%", borderRadius: 10 }} />
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Operations View
--------------------------------------------------------- */
function OperationsView({ setSubView, onAdd }) {
  const st = useStyles();
  const items = [
    { title: "اهداف مالی و پس‌انداز", icon: <Target size={17} />, color: BRAND.green, key: "goals" },
    { title: "اشخاص و طرف حساب‌ها", icon: <Users size={17} />, color: BRAND.violet, key: "persons" },
    { title: "بدهکاران و بستانکاران", icon: <Receipt size={17} />, color: BRAND.crimson, key: "debts" },
    { title: "واحدهای پولی", icon: <DollarSign size={17} />, color: BRAND.gold, key: "currencies" },
    { title: "ماشین حساب", icon: <Type size={17} />, color: BRAND.teal, key: "calculator" },
    { title: "پشتیبانی و تیکت", icon: <Bell size={17} />, color: BRAND.header, key: "support" },
    { title: "حساب‌ها و کارت‌ها", icon: <Landmark size={17} />, color: BRAND.violet, key: "accounts" },
    { title: "دسته‌بندی‌ها و برچسب‌ها", icon: <Tag size={17} />, color: BRAND.mauve, key: "categories" },
    { title: "بودجه‌بندی", icon: <Save size={17} />, color: BRAND.green, key: "budgets" },
    { title: "تراکنش‌های تکرارشونده", icon: <Repeat size={17} />, color: BRAND.teal, key: "recurring" },
    { title: "چک‌ها", icon: <FileSpreadsheet size={17} />, color: BRAND.gold, key: "checks" },
    { title: "وام و اقساط", icon: <Bank size={17} />, color: BRAND.crimson, key: "loans" },
    { title: "یادآوری قبض‌ها", icon: <BellRing size={17} />, color: BRAND.orange, key: "bills" },
    { title: "دارایی‌ها (ارز دیجیتال / بورس)", icon: <Bitcoin size={17} />, color: "#7a5cff", key: "assets" },
    { title: "اعضای منزل، رویداد و پروژه", icon: <Users size={17} />, color: BRAND.violet, key: "tags" },
    { title: "دوره مالی", icon: <CalendarDays size={17} />, color: BRAND.darkgreen, key: "periods" },
    { title: "پیامک بانکی (افزودن نیمه‌خودکار)", icon: <BellRing size={17} />, color: "#666", key: "sms" },
    { title: "تقویم شمسی", icon: <CalendarDays size={17} />, color: BRAND.header, key: "calendar" },
    { title: "اطلاعات کاربری", icon: <Users size={17} />, color: "#444", key: "profile" },
    { title: "تنظیمات و پشتیبان‌گیری", icon: <ShieldCheck size={17} />, color: "#555", key: "settings" },
  ];
  return (
    <div style={{ padding: "10px 16px" }}>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {items.map((it) => <Row key={it.key} title={it.title} leftIcon={it.icon} leftColor={it.color} onClick={() => setSubView(it.key)} />)}
        <Row title="ثبت تراکنش جدید" leftIcon={<Plus size={17} />} leftColor={BRAND.fab} onClick={onAdd} />
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Reports View
--------------------------------------------------------- */
const PIE_COLORS = ["#B01E4A", "#6C3FA0", "#4E9AA0", "#A98A3B", "#3E1461", "#1E8449", "#A65475", "#555"];
function ReportsView({ categories = [], expenseByCategory, incomeByCategory, totalIncomeYear, totalExpenseYear, accounts, accountBalance, updateAccount, netWorthTrend, exportExcel, expenseByMember, expenseByEvent, expenseByProject, checks = [], currency, usdRate, transactions = [] }) {
  const st = useStyles();
  const [period, setPeriod] = useState("year");
  const [selectedAccount, setSelectedAccount] = useState(null); const [editBalance, setEditBalance] = useState("");
  const periodTx = useMemo(() => {
    const now = new Date(); const today = todayISO();
    if (period === "day") return transactions.filter((x) => x.date === today);
    if (period === "month") return transactions.filter((x) => { const d = new Date(x.date); return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth(); });
    if (period === "week") { const start = new Date(now); start.setDate(now.getDate() - 6); const st = start.toISOString().slice(0,10); return transactions.filter((x) => x.date >= st && x.date <= today); }
    return transactions.filter((x) => jalaliYear(new Date(x.date)) === jalaliYear(now));
  }, [transactions, period]);
  const reportIncome = periodTx.filter((x) => x.type === "income").reduce((s, x) => s + x.amount, 0);
  const reportExpense = periodTx.filter((x) => x.type === "expense").reduce((s, x) => s + x.amount, 0);
  const incomeExpensePie = [{ name: "درآمد", amount: reportIncome }, { name: "هزینه", amount: reportExpense }].filter((d) => d.amount > 0);
  const profit = reportIncome - reportExpense;
  const reportExpenseByCategory = useMemo(() => {
    const map = {};
    periodTx.filter((x) => x.type === "expense").forEach((x) => { map[x.categoryId] = (map[x.categoryId] || 0) + x.amount; });
    return Object.entries(map).map(([catId, amount]) => ({ name: categories.find((c) => c.id === catId)?.name || "بدون دسته", amount, catId })).sort((a,b) => b.amount-a.amount);
  }, [periodTx]);

  const checkStats = useMemo(() => {
    const pending = checks.filter((c) => c.status === "pending").length;
    const cashed = checks.filter((c) => c.status === "cashed").length;
    const bounced = checks.filter((c) => c.status === "bounced").length;
    const resolved = cashed + bounced;
    const pct = resolved > 0 ? (cashed / resolved) * 100 : 0;
    return { pending, cashed, bounced, pct };
  }, [checks]);

  return (
    <div style={{ padding: "16px" }}>
      <div className="no-print" style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button onClick={exportExcel} style={{ ...st.primaryBtn, background: BRAND.green, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><FileSpreadsheet size={16} /> خروجی Excel</button>
        <button onClick={() => window.print()} style={{ ...st.primaryBtn, background: "#555", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Printer size={16} /> چاپ / PDF</button>
      </div>

      <div className="no-print" style={{ display: "flex", gap: 6, marginBottom: 18, background: "#f1eef4", borderRadius: 10, padding: 4 }}>
        {[{ k: "day", l: "روزانه" }, { k: "week", l: "هفتگی" }, { k: "month", l: "ماهیانه" }, { k: "year", l: "سالیانه" }].map((p) => (
          <button key={p.k} onClick={() => setPeriod(p.k)} style={{ flex: 1, padding: "7px 4px", borderRadius: 7, border: "none", cursor: "pointer", background: period === p.k ? BRAND.header : "transparent", color: period === p.k ? "#fff" : "#3E1461", fontWeight: 700, fontSize: 12 }}>{p.l}</button>
        ))}
      </div>
      {period !== "year" && (
        <div style={{ fontSize: 11.5, color: "#8a8194", marginBottom: 14, textAlign: "center" }}>
          توجه: نمودارهای زیر همچنان بر اساس سال انتخابی در صفحه‌ی خانه محاسبه شده‌اند؛ فیلتر {period === "day" ? "روزانه" : period === "week" ? "هفتگی" : "ماهیانه"} روی نمای «روند دارایی خالص» و ویجت خانه اعمال می‌شود.
        </div>
      )}

      <SectionTitle text="سود و زیان سالانه" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <Row title="مجموع درآمد" value={formatMoney(reportIncome, currency, usdRate)} valueColor={BRAND.darkgreen} chevron={null} />
        <Row title="مجموع هزینه" value={formatMoney(reportExpense, currency, usdRate)} valueColor={BRAND.crimson} chevron={null} />
        <Row title="سود / زیان خالص" value={formatMoney(Math.abs(profit), currency, usdRate)} valueColor={profit >= 0 ? BRAND.darkgreen : BRAND.crimson} chevron={null} />
      </div>

      <SectionTitle text="روند دارایی خالص (۸ ماه اخیر)" />
      <div style={{ ...st.card, padding: 12, marginBottom: 18 }}>
        <ResponsiveContainer width="100%" height={180}>
          <LineChart data={netWorthTrend}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" tick={{ fontFamily: FONT, fontSize: 11 }} />
            <YAxis tick={{ fontFamily: FONT, fontSize: 10 }} width={44} />
            <Tooltip formatter={(v) => toFaInt(v)} contentStyle={{ fontFamily: FONT, direction: "rtl" }} />
            <Line type="monotone" dataKey="مانده" stroke={BRAND.violet} strokeWidth={2.5} dot={{ r: 3 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <SectionTitle text="درآمد و هزینه" />
      <div style={{ ...st.card, padding: 12, marginBottom: 18 }}>
        <ExplodingPie data={incomeExpensePie} currency={currency} usdRate={usdRate} />
      </div>

      <SectionTitle text="توزیع هزینه‌ها بر اساس دسته" />
      <div style={{ ...st.card, padding: 12, marginBottom: 18 }}>
        <ExplodingPie data={reportExpenseByCategory} currency={currency} usdRate={usdRate} />
      </div>

      <SectionTitle text="گزارش چک‌ها" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-around", marginBottom: 10 }}>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: BRAND.gold }}>{toFaInt(checkStats.pending)}</div>
            <div style={{ fontSize: 11.5, color: "#8a8194" }}>در جریان وصول</div>
          </div>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: BRAND.darkgreen }}>{toFaInt(checkStats.cashed)}</div>
            <div style={{ fontSize: 11.5, color: "#8a8194" }}>وصول شده</div>
          </div>
          <div style={{ textAlign: "center" }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: BRAND.crimson }}>{toFaInt(checkStats.bounced)}</div>
            <div style={{ fontSize: 11.5, color: "#8a8194" }}>برگشت خورده</div>
          </div>
        </div>
        {(checkStats.cashed + checkStats.bounced) > 0 && (
          <div style={{ display: "flex", justifyContent: "center" }}>
            <GaugeSpeedometer pct={checkStats.pct} label="نسبت چک‌های وصول‌شده به کل تسویه‌شده‌ها" />
          </div>
        )}
        {checks.length === 0 && <EmptyRow text="چکی ثبت نشده" />}
      </div>

      <SectionTitle text="گزارش مانده حساب‌ها" />
      <div style={{ ...st.card, padding: "4px 12px", marginBottom: expenseByMember?.length || expenseByEvent?.length || expenseByProject?.length ? 18 : 0 }}>
        {accounts.map((a) => <Row key={a.id} title={a.name} subtitle={`${a.type === "bank" ? "بانک" : a.type === "card" ? "کارت" : "صندوق"} · لمس برای جزئیات`} value={formatMoney(accountBalance(a.id), currency, usdRate)} valueColor={accountBalance(a.id) >= 0 ? BRAND.darkgreen : BRAND.crimson} onClick={() => { setSelectedAccount(a); setEditBalance(String(accountBalance(a.id))); }} />)}
      </div>

      {expenseByMember?.length > 0 && (
        <>
          <SectionTitle text="گزارش هزینه به‌تفکیک اعضای خانواده" />
          <div style={{ ...st.card, padding: "4px 12px", marginBottom: 18 }}>
            {expenseByMember.map((m) => <Row key={m.id} title={m.name} value={formatMoney(m.amount, currency, usdRate)} valueColor={BRAND.crimson} chevron={null} />)}
          </div>
        </>
      )}
      {expenseByEvent?.length > 0 && (
        <>
          <SectionTitle text="گزارش هزینه به‌تفکیک رویداد" />
          <div style={{ ...st.card, padding: "4px 12px", marginBottom: 18 }}>
            {expenseByEvent.map((m) => <Row key={m.id} title={m.name} value={formatMoney(m.amount, currency, usdRate)} valueColor={BRAND.crimson} chevron={null} />)}
          </div>
        </>
      )}
      {expenseByProject?.length > 0 && (
        <>
          <SectionTitle text="گزارش هزینه به‌تفکیک پروژه" />
          <div style={{ ...st.card, padding: "4px 12px" }}>
            {expenseByProject.map((m) => <Row key={m.id} title={m.name} value={formatMoney(m.amount, currency, usdRate)} valueColor={BRAND.crimson} chevron={null} />)}
          </div>
        </>
      )}
      {selectedAccount && <div onClick={() => setSelectedAccount(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", zIndex: 300, display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
        <div onClick={(e) => e.stopPropagation()} style={{ width: "min(480px,100vw)", background: "#fff", borderRadius: "18px 18px 0 0", padding: 18, paddingBottom: "calc(24px + env(safe-area-inset-bottom,0px))" }}>
          <div style={{ fontWeight: 800, fontSize: 16, marginBottom: 8 }}>جزئیات {selectedAccount.name}</div>
          <div style={{ fontSize: 12, color: "#8a8194", marginBottom: 12 }}>مانده فعلی: {formatMoney(accountBalance(selectedAccount.id), currency, usdRate)}</div>
          <label style={st.label}>اصلاح مانده حساب</label>
          <AmountInput value={editBalance} onChange={setEditBalance} placeholder="مانده جدید" style={st.input} />
          <button onClick={() => { const target = Number(editBalance || 0); updateAccount?.(selectedAccount.id, { initial: (selectedAccount.initial || 0) + target - accountBalance(selectedAccount.id) }); setSelectedAccount(null); }} style={st.primaryBtn}>ذخیره اصلاح مانده</button>
          <button onClick={() => setSelectedAccount(null)} style={{ ...st.primaryBtn, background: "#eee", color: "#333", marginTop: 8 }}>بستن</button>
        </div>
      </div>}
    </div>
  );
}

/* ---------------------------------------------------------
   SubView Router
--------------------------------------------------------- */
const SUBVIEW_TITLES = { goals: "اهداف مالی و پس‌انداز", persons: "اشخاص و طرف حساب‌ها", debts: "بدهکاران و بستانکاران", currencies: "واحدهای پولی", calculator: "ماشین حساب", support: "پشتیبانی", shortcuts: "میانبرهای تراکنش", accounts: "حساب‌ها و کارت‌ها", categories: "دسته‌بندی‌ها و برچسب‌ها", budgets: "بودجه‌بندی", recurring: "تراکنش‌های تکرارشونده", checks: "چک‌ها", loans: "وام و اقساط", bills: "یادآوری قبض‌ها", assets: "دارایی‌ها", calendar: "تقویم شمسی", settings: "تنظیمات و امنیت", profile: "ویرایش اطلاعات کاربری", backup: "پشتیبان‌گیری و بازیابی", access: "مدیریت دسترسی", basic: "تنظیمات پایه", tutorial: "آموزش Rexa", share: "ارسال برنامه به دیگران", rate: "امتیاز به برنامه", about: "درباره Rexa", tags: "اعضا، رویداد و پروژه", periods: "دوره مالی", sms: "پیامک بانکی" };
function SubViewContent({ subView, ctx, onBack }) {
  return (
    <div style={{ flex: 1, overflowY: "auto", padding: "20px 16px 112px", paddingTop: "calc(20px + env(safe-area-inset-top, 0px))" }}>
      {subView === "goals" && <GoalsManager {...ctx} />}
      {subView === "persons" && <PersonsManager {...ctx} />}
      {subView === "debts" && <DebtsManager {...ctx} />}
      {subView === "currencies" && <CurrenciesManager {...ctx} />}
      {subView === "calculator" && <CalculatorView {...ctx} />}
      {subView === "support" && <SupportView {...ctx} />}
      {subView === "shortcuts" && <ShortcutsManager {...ctx} />}
      {subView === "accounts" && <AccountsManager {...ctx} />}
      {subView === "categories" && <CategoriesManager {...ctx} />}
      {subView === "budgets" && <BudgetsManager {...ctx} />}
      {subView === "recurring" && <RecurringManager {...ctx} />}
      {subView === "checks" && <ChecksManager {...ctx} />}
      {subView === "loans" && <LoansManager {...ctx} />}
      {subView === "bills" && <BillsManager {...ctx} />}
      {subView === "assets" && <AssetsManager {...ctx} />}
      {subView === "calendar" && <CalendarViewSub {...ctx} />}
      {subView === "settings" && <SettingsView {...ctx} />}
      {subView === "profile" && <ProfileView {...ctx} />}
      {subView === "backup" && <BackupView {...ctx} />}
      {subView === "access" && <AccessView {...ctx} />}
      {subView === "basic" && <BasicSettingsView {...ctx} />}
      {subView === "tutorial" && <TutorialView {...ctx} />}
      {subView === "share" && <ShareAppView {...ctx} />}
      {subView === "rate" && <RateAppView {...ctx} />}
      {subView === "about" && <AboutView {...ctx} />}
      {subView === "tags" && <MembersEventsProjectsManager {...ctx} />}
      {subView === "periods" && <FiscalPeriodsManager {...ctx} />}
      {subView === "sms" && <BankSmsManager {...ctx} onBack={onBack} />}
    </div>
  );
}

/* ---------------------------------------------------------
   Accounts / Categories / Budgets Managers
--------------------------------------------------------- */
const SHORTCUT_ICONS = [
  ["taxi", "🚕"], ["bus", "🚌"], ["car", "🚗"], ["fuel", "⛽"], ["food", "🍽️"], ["market", "🛒"],
  ["home", "🏠"], ["water", "💧"], ["electric", "💡"], ["internet", "🌐"], ["clothes", "👕"], ["salary", "💵"],
  ["bank", "🏦"], ["fun", "🎭"], ["health", "💊"], ["gift", "🎁"], ["phone", "📱"], ["other", "✦"]
];
function ShortcutsManager({ shortcuts, setShortcuts, categories, accounts }) {
  const st = useStyles();
  const [form, setForm] = useState({ name: "", type: "expense", categoryId: "", accountId: accounts[0]?.id || "", note: "", icon: "other" });
  const cats = categories.filter((c) => c.kind === form.type);
  function add() {
    if (!form.name.trim() || !form.categoryId || !form.accountId) return;
    setShortcuts((p) => [...p, { id: uid(), ...form, name: form.name.trim() }]);
    setForm({ ...form, name: "", categoryId: "", note: "", icon: "other" });
  }
  const iconText = (id) => SHORTCUT_ICONS.find(([k]) => k === id)?.[1] || "✦";
  return <div>
    <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
      <div style={{ fontWeight: 800, marginBottom: 10 }}>ساخت میانبر تراکنش</div>
      <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="مثلا کرایه تاکسی" style={st.input} />
      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
        <button onClick={() => setForm({ ...form, type: "expense", categoryId: "" })} style={pillStyle(form.type === "expense")}>هزینه</button>
        <button onClick={() => setForm({ ...form, type: "income", categoryId: "" })} style={pillStyle(form.type === "income")}>درآمد</button>
      </div>
      <label style={st.label}>آیکن میانبر</label>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 7, marginBottom: 10 }}>
        {SHORTCUT_ICONS.map(([key, icon]) => (
          <button key={key} type="button" onClick={() => setForm({ ...form, icon: key })} style={{ width: 42, height: 42, borderRadius: 11, border: form.icon === key ? `2px solid ${BRAND.header}` : "1px solid #e5e0e9", background: form.icon === key ? BRAND.header + "12" : "#fff", fontSize: 21, cursor: "pointer" }}>{icon}</button>
        ))}
      </div>
      <select value={form.categoryId} onChange={(e) => setForm({ ...form, categoryId: e.target.value })} style={st.input}>
        <option value="">دسته‌بندی را انتخاب کنید</option>{cats.map((c) => <option key={c.id} value={c.id}>{c.parentId ? "↳ " : ""}{c.name}</option>)}
      </select>
      <select value={form.accountId} onChange={(e) => setForm({ ...form, accountId: e.target.value })} style={st.input}>
        <option value="">حساب پیش‌فرض</option>{accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
      <input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="شرح ثابت (اختیاری)" style={st.input} />
      <button onClick={add} style={st.primaryBtn}>افزودن میانبر</button>
    </div>
    <div style={{ ...st.card, padding: "10px 12px" }}>
      {shortcuts.length === 0 && <EmptyRow text="هنوز میانبری ساخته نشده" />}
      {shortcuts.length > 0 && <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
        {shortcuts.map((sc) => <div key={sc.id} style={{ textAlign: "center", position: "relative" }}>
          <button type="button" style={{ width: 62, height: 62, borderRadius: 14, border: "none", background: sc.type === "income" ? "#b9f3ef" : "#fff0c8", fontSize: 31, cursor: "default" }}>{iconText(sc.icon)}</button>
          <div style={{ fontSize: 11.5, fontWeight: 800, marginTop: 6 }}>{sc.name}</div>
          <button onClick={() => setShortcuts((p) => p.filter((x) => x.id !== sc.id))} style={{ position: "absolute", top: -5, right: 1, width: 22, height: 22, borderRadius: "50%", border: "none", background: BRAND.crimson, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center" }}><Trash2 size={12} /></button>
        </div>)}
      </div>}
    </div>
  </div>;
}

function AccountsManager({ accounts, addAccount, deleteAccount, updateAccount, accountBalance, favorites, toggleFavorite }) {
  const st = useStyles();
  const [name, setName] = useState(""); const [type, setType] = useState("bank"); const [initial, setInitial] = useState(""); const [last4, setLast4] = useState("");
  const [editingId, setEditingId] = useState(null); const [editBalance, setEditBalance] = useState("");
  return <div>
    <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
      <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن حساب / کارت جدید</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
        <button onClick={() => setType("bank")} style={pillStyle(type === "bank")}>بانک</button><button onClick={() => setType("fund")} style={pillStyle(type === "fund")}>صندوق</button><button onClick={() => setType("card")} style={pillStyle(type === "card")}>کارت</button>
      </div>
      <input placeholder="نام حساب" value={name} onChange={(e) => setName(e.target.value)} style={st.input} />
      <AmountInput placeholder="موجودی اولیه (ریال)" value={initial} onChange={setInitial} style={st.input} />
      {(type === "bank" || type === "card") && <input placeholder="۴ رقم آخر شماره کارت" value={last4} onChange={(e) => setLast4(e.target.value.replace(/[^0-9]/g, "").slice(0, 4))} style={st.input} inputMode="numeric" />}
      <button onClick={() => { if (!name.trim()) return; addAccount({ name: name.trim(), type, initial: Number(initial || 0), cardNumberLast4: last4 || undefined }); setName(""); setInitial(""); setLast4(""); }} style={st.primaryBtn}>افزودن</button>
    </div>
    <div style={{ ...st.card, padding: "4px 12px" }}>
      {accounts.length === 0 && <EmptyRow text="حسابی ثبت نشده" />}
      {accounts.map((a) => <div key={a.id}>
        <Row title={a.name} subtitle={`${a.type === "bank" ? "بانک" : a.type === "card" ? "کارت" : "صندوق"} · موجودی: ${toFaInt(accountBalance(a.id))} ریال`}
          extra={<div style={{ display: "flex", gap: 4 }}>
            <button onClick={() => { setEditingId(a.id); setEditBalance(String(accountBalance(a.id))); }} style={{ ...miniBtn, color: BRAND.violet }}><Pencil size={13} /></button>
            <button onClick={() => toggleFavorite("accounts", a.id)} style={{ background: "none", border: "none", cursor: "pointer" }}><Star size={16} fill={favorites.accounts.includes(a.id) ? "#f5b301" : "none"} color="#f5b301" /></button>
          </div>}
          leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson} onClick={() => deleteAccount(a.id)} chevron={null} />
        {editingId === a.id && <div style={{ display: "flex", gap: 6, padding: "4px 0 10px" }}>
          <AmountInput value={editBalance} onChange={setEditBalance} placeholder="مانده جدید" style={{ ...st.input, margin: 0, flex: 1 }} />
          <button onClick={() => { const target = Number(editBalance || 0); updateAccount?.(a.id, { initial: (a.initial || 0) + target - accountBalance(a.id) }); setEditingId(null); }} style={{ ...st.primaryBtn, width: "auto", padding: "0 14px" }}>اصلاح</button>
        </div>}
      </div>)}
    </div>
  </div>;
}

function CategoriesManager({ categories, addCategory, deleteCategory, updateCategory, favorites, toggleFavorite }) {
  const st = useStyles();
  const [name, setName] = useState(""); const [kind, setKind] = useState("expense");
  const [subFormFor, setSubFormFor] = useState(null);
  const [subName, setSubName] = useState("");
  const topLevel = (list) => list.filter((c) => !c.parentId);
  const childrenOf = (id) => categories.filter((c) => c.parentId === id);
  const expense = topLevel(categories.filter((c) => c.kind === "expense"));
  const income = topLevel(categories.filter((c) => c.kind === "income"));

  function addSub(parentId, parentKind) {
    if (!subName.trim()) return;
    addCategory({ name: subName.trim(), kind: parentKind, parentId });
    setSubName(""); setSubFormFor(null);
  }

  function renderGroup(list) {
    return list.map((c) => (
      <div key={c.id}>
        <CatRow c={c} onDelete={deleteCategory} onUpdate={updateCategory} fav={favorites.categories.includes(c.id)} onFav={() => toggleFavorite("categories", c.id)}
          onAddSub={() => setSubFormFor(subFormFor === c.id ? null : c.id)} />
        {childrenOf(c.id).map((sub) => (
          <div key={sub.id} style={{ paddingRight: 22 }}>
            <CatRow c={sub} isSub onDelete={deleteCategory} onUpdate={updateCategory} fav={favorites.categories.includes(sub.id)} onFav={() => toggleFavorite("categories", sub.id)} />
          </div>
        ))}
        {subFormFor === c.id && (
          <div style={{ display: "flex", gap: 6, padding: "8px 4px 8px 22px" }}>
            <input autoFocus placeholder={`زیرمجموعه‌ی ${c.name}`} value={subName} onChange={(e) => setSubName(e.target.value)} style={{ ...st.input, margin: 0, flex: 1 }} />
            <button onClick={() => addSub(c.id, c.kind)} style={{ ...st.primaryBtn, width: "auto", padding: "0 16px" }}>افزودن</button>
          </div>
        )}
      </div>
    ));
  }

  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن سرفصل جدید</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setKind("expense")} style={pillStyle(kind === "expense")}>هزینه</button>
          <button onClick={() => setKind("income")} style={pillStyle(kind === "income")}>درآمد</button>
        </div>
        <input placeholder="نام سرفصل" value={name} onChange={(e) => setName(e.target.value)} style={st.input} />
        <button onClick={() => { if (!name.trim()) return; addCategory({ name: name.trim(), kind }); setName(""); }} style={st.primaryBtn}>افزودن</button>
      </div>
      <div style={{ fontSize: 11.5, color: "#8a8194", margin: "0 4px 12px" }}>برای افزودن زیرمجموعه، روی آیکون + کنار هر سرفصل بزن. برای تغییر نام، روی آیکون مداد بزن.</div>
      <SectionTitle text="دسته‌های هزینه" />
      <div style={{ ...st.card, padding: "4px 12px", marginBottom: 16 }}>
        {expense.length === 0 && <EmptyRow text="دسته‌ای ثبت نشده" />}
        {renderGroup(expense)}
      </div>
      <SectionTitle text="دسته‌های درآمد" />
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {income.length === 0 && <EmptyRow text="دسته‌ای ثبت نشده" />}
        {renderGroup(income)}
      </div>
    </div>
  );
}
function CatRow({ c, onDelete, onUpdate, fav, onFav, onAddSub, isSub }) {
  const t = useT();
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(c.name);
  if (editing) {
    return (
      <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 4px", borderBottom: `1px solid ${t.border}` }}>
        <input autoFocus value={val} onChange={(e) => setVal(e.target.value)} style={{ flex: 1, padding: "6px 10px", borderRadius: 7, border: `1.5px solid ${t.inputBorder}`, background: t.input, color: t.text, fontSize: 13 }} />
        <button onClick={() => { if (val.trim()) onUpdate(c.id, { name: val.trim() }); setEditing(false); }} style={{ background: BRAND.green, border: "none", borderRadius: 7, color: "#fff", padding: "6px 10px", cursor: "pointer" }}><Check size={14} /></button>
        <button onClick={() => { setVal(c.name); setEditing(false); }} style={{ background: "none", border: "none", color: t.sub, cursor: "pointer" }}><X size={14} /></button>
      </div>
    );
  }
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 4px", borderBottom: `1px solid ${t.border}` }}>
      <span style={{ fontSize: isSub ? 13 : 14, fontWeight: isSub ? 500 : 600, color: isSub ? t.sub : t.text }}>{isSub && "↳ "}{c.name}</span>
      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={onFav} style={{ background: "none", border: "none", cursor: "pointer" }}><Star size={15} fill={fav ? "#f5b301" : "none"} color="#f5b301" /></button>
        <button onClick={() => setEditing(true)} style={{ background: "none", border: "none", color: BRAND.violet, cursor: "pointer" }}><Pencil size={15} /></button>
        {onAddSub && <button onClick={onAddSub} style={{ background: "none", border: "none", color: BRAND.green, cursor: "pointer" }}><Plus size={15} /></button>}
        <button onClick={() => onDelete(c.id)} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
      </div>
    </div>
  );
}

function BudgetsManager({ categories, budgets, upsertBudget, addCategory }) {
  const st = useStyles();
  const expenseCats = categories.filter((c) => c.kind === "expense");
  const [newName, setNewName] = useState("");
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 12 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 8 }}>افزودن آیتم بودجه جدید</div>
        <div style={{ display: "flex", gap: 6 }}>
          <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="مثلا آموزش" style={{ ...st.input, margin: 0, flex: 1 }} />
          <button onClick={() => { if (!newName.trim()) return; addCategory?.({ name: newName.trim(), kind: "expense" }); setNewName(""); }} style={{ ...st.primaryBtn, width: "auto", padding: "0 14px" }}>افزودن</button>
        </div>
      </div>
      <div style={{ ...st.card, padding: "8px 12px" }}>
      {expenseCats.length === 0 && <EmptyRow text="ابتدا یک دسته هزینه بسازید" />}
      {expenseCats.map((c) => {
        const b = budgets.find((bb) => bb.categoryId === c.id);
        return <BudgetRow key={c.id} category={c} budget={b} upsertBudget={upsertBudget} st={st} />;
      })}
      </div>
    </div>
  );
}
function BudgetRow({ category, budget, upsertBudget, st }) {
  const [val, setVal] = useState(budget ? String(budget.amount) : "");
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 4px", borderBottom: "1px solid #f0eef3", gap: 10 }}>
      <span style={{ fontSize: 13.5, fontWeight: 600, flexShrink: 0 }}>{category.name}</span>
      <AmountInput placeholder="بودجه ماهانه" value={val} onChange={(v) => { setVal(v); upsertBudget(category.id, Number(v || 0)); }} style={{ ...st.input, margin: 0, width: 150 }} />
    </div>
  );
}

/* ---------------------------------------------------------
   Recurring Manager
--------------------------------------------------------- */
function RecurringManager({ recurring, setRecurring, categories, accounts }) {
  const st = useStyles();
  const [form, setForm] = useState({ type: "expense", amount: "", categoryId: "", accountId: accounts[0]?.id || "", interval: "monthly", startDate: todayISO(), note: "" });
  function add() {
    if (!form.amount || !form.categoryId || !form.accountId) return;
    setRecurring((p) => [...p, { id: uid(), ...form, amount: Number(form.amount), nextDate: form.startDate, active: true }]);
    setForm({ ...form, amount: "", note: "" });
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن تراکنش تکرارشونده (مثل حقوق یا اجاره)</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, type: "expense", categoryId: "" })} style={pillStyle(form.type === "expense")}>هزینه</button>
          <button onClick={() => setForm({ ...form, type: "income", categoryId: "" })} style={pillStyle(form.type === "income")}>درآمد</button>
        </div>
        <AmountInput placeholder="مبلغ" value={form.amount} onChange={(v) => setForm({ ...form, amount: v })} style={st.input} />
        <select value={form.categoryId} onChange={(e) => setForm({ ...form, categoryId: e.target.value })} style={st.input}>
          <option value="">دسته را انتخاب کنید</option>
          {categories.filter((c) => c.kind === form.type).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
        <select value={form.accountId} onChange={(e) => setForm({ ...form, accountId: e.target.value })} style={st.input}>
          {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
        </select>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, interval: "monthly" })} style={pillStyle(form.interval === "monthly")}>ماهانه</button>
          <button onClick={() => setForm({ ...form, interval: "weekly" })} style={pillStyle(form.interval === "weekly")}>هفتگی</button>
        </div>
        <label style={st.label}>تاریخ شروع</label>
        <JalaliDateInput value={form.startDate} onChange={(v) => setForm({ ...form, startDate: v })} style={st.input} />
        <input placeholder="یادداشت (اختیاری)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>افزودن</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {recurring.length === 0 && <EmptyRow text="موردی ثبت نشده" />}
        {recurring.map((r) => (
          <Row key={r.id} title={categories.find((c) => c.id === r.categoryId)?.name || "—"}
            subtitle={`${r.interval === "monthly" ? "ماهانه" : "هفتگی"} · تراکنش بعدی: ${faLongDate(new Date(r.nextDate))}`}
            value={`${toFaInt(r.amount)} ریال`}
            extra={<button onClick={() => setRecurring((p) => p.map((x) => x.id === r.id ? { ...x, active: !x.active } : x))} style={{ background: "none", border: "none", cursor: "pointer", color: r.active ? BRAND.green : "#aaa", fontSize: 11, fontWeight: 700 }}>{r.active ? "فعال" : "غیرفعال"}</button>}
            leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson} onClick={() => setRecurring((p) => p.filter((x) => x.id !== r.id))} chevron={null} />
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Checks Manager
--------------------------------------------------------- */
function ChecksManager({ checks, setChecks }) {
  const st = useStyles();
  const [form, setForm] = useState({ type: "received", payee: "", amount: "", dueDate: todayISO(), note: "", sayadId: "", checkNumber: "" });
  function add() {
    if (!form.payee || !form.amount || !form.checkNumber || form.sayadId.length < 16 || form.sayadId.length > 20) return;
    setChecks((p) => [{ id: uid(), ...form, amount: Number(form.amount), status: "pending" }, ...p]);
    setForm({ ...form, payee: "", amount: "", note: "", sayadId: "", checkNumber: "" });
  }
  const canAdd = form.payee && form.amount && form.checkNumber && form.sayadId.length >= 16 && form.sayadId.length <= 20;
  const statusColor = { pending: BRAND.gold, cashed: BRAND.darkgreen, bounced: BRAND.crimson };
  const statusLabel = { pending: "در انتظار", cashed: "نقد شده", bounced: "برگشتی" };
  const sorted = [...checks].sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  const dueSoon = sorted.filter((c) => c.status === "pending" && daysUntil(c.dueDate) <= 7);
  return (
    <div>
      {dueSoon.length > 0 && (
        <div style={{ background: "#fff6ea", border: "1px solid #f0d9a8", borderRadius: 10, padding: 12, marginBottom: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, color: BRAND.orange, fontWeight: 700, fontSize: 13, marginBottom: 6 }}><BellRing size={16} /> یادآوری سررسید چک‌ها</div>
          {dueSoon.map((c) => (
            <div key={c.id} style={{ fontSize: 12.5, color: "#6b4c14" }}>
              {c.payee} — {toFaInt(c.amount)} ریال — {daysUntil(c.dueDate) < 0 ? "سررسید گذشته" : daysUntil(c.dueDate) === 0 ? "امروز سررسید است" : `${toFaInt(daysUntil(c.dueDate))} روز مانده`}
            </div>
          ))}
        </div>
      )}
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>ثبت چک جدید</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, type: "received" })} style={pillStyle(form.type === "received")}>دریافتی</button>
          <button onClick={() => setForm({ ...form, type: "paid" })} style={pillStyle(form.type === "paid")}>پرداختی</button>
        </div>
        <input placeholder="نام طرف حساب" value={form.payee} onChange={(e) => setForm({ ...form, payee: e.target.value })} style={st.input} />
        <AmountInput placeholder="مبلغ" value={form.amount} onChange={(v) => setForm({ ...form, amount: v })} style={st.input} />
        <label style={st.label}>شناسه صیادی (۱۶ رقم)</label>
        <input placeholder="شناسه ۱۶ رقمی روی چک" value={form.sayadId}
          onChange={(e) => setForm({ ...form, sayadId: e.target.value.replace(/[^0-9]/g, "").slice(0, 20) })}
          style={{ ...st.input, direction: "ltr", textAlign: "right", borderColor: form.sayadId && form.sayadId.length < 16 ? BRAND.crimson : st.input.border }} inputMode="numeric" maxLength={20} />
        {form.sayadId && form.sayadId.length < 16 && <div style={{ fontSize: 11, color: BRAND.crimson, marginTop: -6, marginBottom: 10 }}>شناسه صیادی باید حداقل ۱۶ رقم باشد ({toFaInt(form.sayadId.length)}/۱۶)</div>}
        <label style={st.label}>شماره چک</label>
        <input placeholder="شماره سریال چک" value={form.checkNumber} onChange={(e) => setForm({ ...form, checkNumber: e.target.value })} style={{ ...st.input, direction: "ltr", textAlign: "right" }} />
        <label style={st.label}>تاریخ سررسید</label>
        <JalaliDateInput value={form.dueDate} onChange={(v) => setForm({ ...form, dueDate: v })} style={st.input} />
        <input placeholder="یادداشت (اختیاری)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} style={st.input} />
        <button onClick={add} disabled={!canAdd} style={{ ...st.primaryBtn, opacity: canAdd ? 1 : 0.5 }}>ثبت چک</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {checks.length === 0 && <EmptyRow text="چکی ثبت نشده" />}
        {sorted.map((c) => {
          const d = daysUntil(c.dueDate);
          const urgent = c.status === "pending" && d <= 7;
          return (
          <div key={c.id} style={{ padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <span style={{ fontWeight: 700, fontSize: 14 }}>{c.payee} ({c.type === "received" ? "دریافتی" : "پرداختی"})</span>
              <span style={{ fontWeight: 700 }}>{toFaInt(c.amount)} ریال</span>
            </div>
            {(c.sayadId || c.checkNumber) && (
              <div style={{ fontSize: 11, color: "#8a8194", marginBottom: 4, direction: "ltr", textAlign: "right" }}>
                {c.sayadId && `شناسه صیادی: ${c.sayadId}`}{c.sayadId && c.checkNumber ? " · " : ""}{c.checkNumber && `شماره چک: ${c.checkNumber}`}
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 12, color: urgent ? (d < 0 ? BRAND.crimson : BRAND.orange) : "#8a8194", fontWeight: urgent ? 700 : 400 }}>
                سررسید: {faLongDate(new Date(c.dueDate))}{c.status === "pending" ? ` (${d < 0 ? "گذشته" : d === 0 ? "امروز" : `${toFaInt(d)} روز مانده`})` : ""}
              </span>
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <StatusBadge text={statusLabel[c.status]} color={statusColor[c.status]} />
                <select value={c.status} onChange={(e) => setChecks((p) => p.map((x) => x.id === c.id ? { ...x, status: e.target.value } : x))} style={{ fontSize: 11, borderRadius: 6, border: "1px solid #ddd" }}>
                  <option value="pending">در انتظار</option><option value="cashed">نقد شده</option><option value="bounced">برگشتی</option>
                </select>
                <button onClick={() => setChecks((p) => p.filter((x) => x.id !== c.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
              </div>
            </div>
          </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Loans Manager
--------------------------------------------------------- */
function LoansManager({ loans, setLoans }) {
  const st = useStyles();
  const [form, setForm] = useState({ title: "", principal: "", installments: "", monthlyPayment: "", startDate: todayISO() });
  function add() {
    if (!form.title || !form.principal || !form.installments || !form.monthlyPayment) return;
    setLoans((p) => [...p, { id: uid(), title: form.title, principal: Number(form.principal), installments: Number(form.installments), monthlyPayment: Number(form.monthlyPayment), startDate: form.startDate, paidCount: 0 }]);
    setForm({ title: "", principal: "", installments: "", monthlyPayment: "", startDate: todayISO() });
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>ثبت وام جدید</div>
        <input placeholder="عنوان وام (مثلا وام خودرو)" value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} style={st.input} />
        <AmountInput placeholder="مبلغ اصل وام" value={form.principal} onChange={(v) => setForm({ ...form, principal: v })} style={st.input} />
        <input placeholder="تعداد اقساط" value={form.installments} onChange={(e) => setForm({ ...form, installments: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <AmountInput placeholder="مبلغ هر قسط" value={form.monthlyPayment} onChange={(v) => setForm({ ...form, monthlyPayment: v })} style={st.input} />
        <label style={st.label}>تاریخ شروع</label>
        <JalaliDateInput value={form.startDate} onChange={(v) => setForm({ ...form, startDate: v })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>ثبت وام</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {loans.length === 0 && <EmptyRow text="وامی ثبت نشده" />}
        {loans.map((l) => {
          const remaining = l.principal - l.paidCount * l.monthlyPayment;
          const nextDue = addMonths(l.startDate, l.paidCount);
          return (
            <div key={l.id} style={{ padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>{l.title}</span>
                <button onClick={() => setLoans((p) => p.filter((x) => x.id !== l.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
              </div>
              <div style={{ fontSize: 12.5, color: "#8a8194", marginBottom: 4 }}>قسط {toFaInt(l.paidCount)} از {toFaInt(l.installments)} پرداخت شده — سررسید بعدی: {l.paidCount < l.installments ? faLongDate(new Date(nextDue)) : "تسویه شده"}</div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 700, color: remaining > 0 ? BRAND.crimson : BRAND.darkgreen }}>باقی‌مانده: {toFaInt(Math.max(remaining, 0))} ریال</span>
                {l.paidCount < l.installments && (
                  <button onClick={() => setLoans((p) => p.map((x) => x.id === l.id ? { ...x, paidCount: x.paidCount + 1 } : x))} style={{ background: BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>ثبت پرداخت قسط</button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Bills Manager
--------------------------------------------------------- */
const BILL_CATS = ["آب", "برق", "گاز", "اینترنت", "تلفن", "سایر"];
function BillsManager({ bills, setBills }) {
  const st = useStyles();
  const [form, setForm] = useState({ title: "آب", amount: "", dueDate: todayISO(), recurringMonthly: true });
  function add() {
    setBills((p) => [...p, { id: uid(), title: form.title, amount: Number(form.amount || 0), dueDate: form.dueDate, recurringMonthly: form.recurringMonthly, paid: false }]);
    setForm({ ...form, amount: "" });
  }
  function markPaid(b) {
    setBills((prev) => prev.map((x) => x.id === b.id ? (x.recurringMonthly ? { ...x, dueDate: addMonths(x.dueDate, 1), paid: false } : { ...x, paid: true }) : x));
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن قبض</div>
        <div style={{ display: "flex", gap: 6, marginBottom: 10, flexWrap: "wrap" }}>
          {BILL_CATS.map((c) => <button key={c} onClick={() => setForm({ ...form, title: c })} style={{ ...pillStyle(form.title === c), flex: "none", padding: "7px 12px" }}>{c}</button>)}
        </div>
        <AmountInput placeholder="مبلغ (اختیاری)" value={form.amount} onChange={(v) => setForm({ ...form, amount: v })} style={st.input} />
        <label style={st.label}>تاریخ سررسید</label>
        <JalaliDateInput value={form.dueDate} onChange={(v) => setForm({ ...form, dueDate: v })} style={st.input} />
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, marginBottom: 12 }}>
          <input type="checkbox" checked={form.recurringMonthly} onChange={(e) => setForm({ ...form, recurringMonthly: e.target.checked })} /> یادآوری ماهانه تکرار شود
        </label>
        <button onClick={add} style={st.primaryBtn}>افزودن قبض</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {bills.length === 0 && <EmptyRow text="قبضی ثبت نشده" />}
        {bills.sort((a, b) => a.dueDate.localeCompare(b.dueDate)).map((b) => {
          const d = daysUntil(b.dueDate);
          return (
            <div key={b.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{b.title}{b.amount ? ` — ${toFaInt(b.amount)} ریال` : ""}</div>
                <div style={{ fontSize: 12, color: b.paid ? BRAND.darkgreen : d < 0 ? BRAND.crimson : "#8a8194" }}>
                  {faLongDate(new Date(b.dueDate))} · {b.paid ? "پرداخت شده" : d < 0 ? "سررسید گذشته" : `${toFaInt(d)} روز مانده`}
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                {!b.paid && <button onClick={() => markPaid(b)} style={{ background: BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer" }}>ثبت پرداخت</button>}
                <button onClick={() => setBills((p) => p.filter((x) => x.id !== b.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Assets Manager
--------------------------------------------------------- */
function AssetPriceEdit({ asset, setAssets, st }) {
  const [val, setVal] = useState(String(asset.currentPrice));
  return (
    <AmountInput value={val} onChange={(v) => { setVal(v); setAssets((p) => p.map((x) => x.id === asset.id ? { ...x, currentPrice: Number(v || 0) } : x)); }} style={{ ...st.input, margin: 0, width: 110 }} />
  );
}
function AssetsManager({ assets, setAssets }) {
  const [refreshing, setRefreshing] = useState(false);
  const st = useStyles();
  const [form, setForm] = useState({ kind: "crypto", symbol: "", quantity: "", avgPrice: "", currentPrice: "" });
  const total = assets.reduce((s, a) => s + a.quantity * a.currentPrice, 0);
  function add() {
    if (!form.symbol || !form.quantity) return;
    setAssets((p) => [...p, { id: uid(), kind: form.kind, symbol: form.symbol, quantity: Number(form.quantity), avgPrice: Number(form.avgPrice || 0), currentPrice: Number(form.currentPrice || form.avgPrice || 0) }]);
    setForm({ ...form, symbol: "", quantity: "", avgPrice: "", currentPrice: "" });
  }
  return (
    <div>
      <div style={{ background: "#eef8f1", border: "1px solid #cde7d4", borderRadius: 10, padding: 10, marginBottom: 14, fontSize: 12, color: "#28633a", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
        <span>قیمت دارایی‌های شناخته‌شده ارز دیجیتال را می‌توانی لحظه‌ای بروزرسانی کنی.</span>
        <button disabled={refreshing} onClick={async () => {
          setRefreshing(true);
          try {
            const ids = { BTC: "bitcoin", ETH: "ethereum", USDT: "tether", BNB: "binancecoin", XRP: "ripple", SOL: "solana" };
            for (const a of assets.filter((x) => x.kind === "crypto")) {
              const coin = ids[String(a.symbol).toUpperCase()]; if (!coin) continue;
              const data = Capacitor.isNativePlatform()
                ? (await CapacitorHttp.get({ url: `https://api.coingecko.com/api/v3/simple/price?ids=${coin}&vs_currencies=irr` })).data
                : await (await fetch(`https://api.coingecko.com/api/v3/simple/price?ids=${coin}&vs_currencies=irr`)).json();
              const price = Number(data?.[coin]?.irr);
              if (price > 0) setAssets((p) => p.map((x) => x.id === a.id ? { ...x, currentPrice: price, priceUpdatedAt: new Date().toISOString() } : x));
            }
          } catch {} finally { setRefreshing(false); }
        }} style={{ ...miniBtn, width: 38, height: 38, background: BRAND.green, color: "#fff" }}><RefreshCw size={15} /></button>
      </div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن دارایی</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, kind: "crypto" })} style={pillStyle(form.kind === "crypto")}>ارز دیجیتال</button>
          <button onClick={() => setForm({ ...form, kind: "stock" })} style={pillStyle(form.kind === "stock")}>بورس</button>
        </div>
        <input placeholder="نماد (مثلا BTC یا فولاد)" value={form.symbol} onChange={(e) => setForm({ ...form, symbol: e.target.value })} style={st.input} />
        <input placeholder="تعداد / مقدار" value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} style={st.input} inputMode="decimal" />
        <AmountInput placeholder="قیمت خرید (ریال)" value={form.avgPrice} onChange={(v) => setForm({ ...form, avgPrice: v })} style={st.input} />
        <AmountInput placeholder="قیمت فعلی (ریال)" value={form.currentPrice} onChange={(v) => setForm({ ...form, currentPrice: v })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>افزودن دارایی</button>
      </div>
      <SectionTitle text={`ارزش کل: ${toFaInt(total)} ریال`} />
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {assets.length === 0 && <EmptyRow text="دارایی‌ای ثبت نشده" />}
        {assets.map((a) => (
          <div key={a.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14 }}>{a.symbol} <span style={{ fontSize: 11, color: "#8a8194" }}>({a.kind === "crypto" ? "ارز دیجیتال" : "بورس"})</span></div>
              <div style={{ fontSize: 12, color: "#8a8194" }}>تعداد: {a.quantity} · ارزش: {toFaInt(a.quantity * a.currentPrice)} ریال</div>
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <AssetPriceEdit asset={a} setAssets={setAssets} st={st} />
              <button onClick={() => setAssets((p) => p.filter((x) => x.id !== a.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Members / Events / Projects Manager (تگ‌گذاری اعضای خانواده، رویداد و پروژه)
--------------------------------------------------------- */
function MembersEventsProjectsManager({ members, setMembers, events, setEvents, projects, setProjects }) {
  const st = useStyles();
  const [tab, setTab] = useState("members");
  const [name, setName] = useState("");
  const map = {
    members: { list: members, setList: setMembers, label: "عضو خانواده", color: BRAND.violet },
    events: { list: events, setList: setEvents, label: "رویداد", color: BRAND.teal },
    projects: { list: projects, setList: setProjects, label: "پروژه", color: BRAND.gold },
  };
  const cur = map[tab];
  function add() {
    if (!name.trim()) return;
    cur.setList((p) => [...p, { id: uid(), name: name.trim() }]);
    setName("");
  }
  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button onClick={() => setTab("members")} style={pillStyle(tab === "members")}>اعضای خانواده</button>
        <button onClick={() => setTab("events")} style={pillStyle(tab === "events")}>رویدادها</button>
        <button onClick={() => setTab("projects")} style={pillStyle(tab === "projects")}>پروژه‌ها</button>
      </div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن {cur.label}</div>
        <input placeholder={`نام ${cur.label}`} value={name} onChange={(e) => setName(e.target.value)} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>افزودن</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {cur.list.length === 0 && <EmptyRow text={`${cur.label}ی ثبت نشده`} />}
        {cur.list.map((m) => (
          <Row key={m.id} title={m.name} leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson}
            onClick={() => cur.setList((p) => p.filter((x) => x.id !== m.id))} chevron={null} />
        ))}
      </div>
      <div style={{ fontSize: 12, color: "#8a8194", marginTop: 12 }}>
        این موارد را می‌توانید هنگام ثبت تراکنش (بخش «بیشتر») به هر پرداخت یا دریافت نسبت دهید تا در گزارش‌ها به‌تفکیک عضو، رویداد یا پروژه ببینید.
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Fiscal Periods Manager (دوره مالی)
--------------------------------------------------------- */
function FiscalPeriodsManager({ fiscalPeriods, setFiscalPeriods }) {
  const st = useStyles();
  const [form, setForm] = useState({ title: "", startDate: todayISO(), endDate: todayISO() });
  function add() {
    if (!form.title) return;
    setFiscalPeriods((p) => [...p, { id: uid(), ...form }]);
    setForm({ title: "", startDate: todayISO(), endDate: todayISO() });
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>تعریف دوره مالی جدید</div>
        <div style={{ fontSize: 12, color: "#8a8194", marginBottom: 10 }}>مثلا «سال مالی ۱۴۰۴» یا «فصل بهار». در گزارش‌ها می‌توانید بر اساس این بازه فیلتر کنید.</div>
        <input placeholder="عنوان دوره" value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} style={st.input} />
        <label style={st.label}>تاریخ شروع</label>
        <JalaliDateInput value={form.startDate} onChange={(v) => setForm({ ...form, startDate: v })} style={st.input} />
        <label style={st.label}>تاریخ پایان</label>
        <JalaliDateInput value={form.endDate} onChange={(v) => setForm({ ...form, endDate: v })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>افزودن دوره</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {fiscalPeriods.length === 0 && <EmptyRow text="دوره‌ای ثبت نشده" />}
        {fiscalPeriods.map((p) => (
          <Row key={p.id} title={p.title} subtitle={`${faLongDate(new Date(p.startDate))} تا ${faLongDate(new Date(p.endDate))}`}
            leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson}
            onClick={() => setFiscalPeriods((prev) => prev.filter((x) => x.id !== p.id))} chevron={null} />
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Bank SMS quick-add (پیامک بانکی)
--------------------------------------------------------- */
function BankSmsManager({ openWithPrefill, onBack, categories, accounts, addTransaction }) {
  const st = useStyles();
  const t = useT();
  const [text, setText] = useState("");
  const parsed = text.trim() ? parseBankSms(text) : null;
  const isNative = typeof Capacitor !== "undefined" && Capacitor.isNativePlatform && Capacitor.isNativePlatform();
  const [syncing, setSyncing] = useState(false);
  const [syncError, setSyncError] = useState("");
  const [suggestions, setSuggestions] = useState([]); // [{ id, body, date, parsed }]
  const [confirmedIds, setConfirmedIds] = useState([]);
  const defaultAccountId = accounts[0]?.id;

  async function syncFromPhone() {
    setSyncing(true);
    setSyncError("");
    try {
      const mod = await import("capacitor-sms-inbox");
      const plugin = mod.SMSInboxReader || mod.SmsInboxReader || mod.SMSInbox || mod.SmsInbox || mod.default;
      if (!plugin) throw new Error("پلاگین پیامک در این نسخه پیدا نشد");
      let perm = await plugin.checkPermissions();
      if (perm.sms !== "granted") perm = await plugin.requestPermissions();
      if (perm.sms !== "granted") { setSyncError("اجازه‌ی دسترسی به پیامک‌ها داده نشد."); setSyncing(false); return; }
      const thirtyDaysAgo = Date.now() - 30 * 24 * 60 * 60 * 1000;
      const { smsList } = await plugin.getSMSList({ filter: { minDate: thirtyDaysAgo, maxCount: 100 } });
      const candidates = (smsList || [])
        .map((m) => ({ id: String(m.id), body: m.body || "", date: m.date, parsed: parseBankSms(m.body || "") }))
        .filter((m) => m.parsed.amount && /بانک|کارت|حساب|واریز|برداشت|خرید|پرداخت|انتقال/.test(m.body));
      setSuggestions(candidates);
      if (candidates.length === 0) setSyncError("پیامک بانکی جدیدی پیدا نشد (یا همه قبلاً بررسی شده‌اند).");
    } catch (e) {
      setSyncError("خواندن پیامک‌ها ممکن نشد. این قابلیت فقط در اپلیکیشن نصب‌شده (APK) کار می‌کند، نه در مرورگر.");
    }
    setSyncing(false);
  }

  function confirmSuggestion(s) {
    if (!defaultAccountId) { alert("اول یک حساب بانکی بساز."); return; }
    const catMatch = categories.find((c) => c.name === s.parsed.categoryHint && c.kind === s.parsed.type);
    addTransaction({
      type: s.parsed.type, amount: s.parsed.amount, categoryId: catMatch?.id || "",
      accountId: defaultAccountId, date: new Date(s.date).toISOString().slice(0, 10),
      note: s.body.slice(0, 140), tags: ["پیامک بانکی"],
    });
    setConfirmedIds((p) => [...p, s.id]);
  }

  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 8, fontSize: 14, display: "flex", alignItems: "center", gap: 6 }}>
          <BellRing size={16} color={BRAND.header} /> همگام‌سازی خودکار از پیامک‌های گوشی
        </div>
        {!isNative && (
          <div style={{ fontSize: 12, color: "#8a8194", marginBottom: 10 }}>
            این بخش فقط داخل اپلیکیشن نصب‌شده (APK) روی گوشی کار می‌کند، چون به مجوز خواندن پیامک نیاز داره — الان داری از پیش‌نمایش مرورگر استفاده می‌کنی.
          </div>
        )}
        <button onClick={syncFromPhone} disabled={syncing} style={{ ...st.primaryBtn, opacity: syncing ? 0.6 : 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
          <RefreshCw size={16} /> {syncing ? "در حال بررسی پیامک‌ها..." : "بررسی پیامک‌های بانکی گوشی"}
        </button>
        {syncError && <div style={{ fontSize: 12, color: BRAND.crimson, marginTop: 8 }}>{syncError}</div>}
      </div>

      {suggestions.length > 0 && (
        <>
          <SectionTitle text={`${toFaInt(suggestions.filter((s) => !confirmedIds.includes(s.id)).length)} تراکنش پیشنهادی از پیامک`} />
          <div style={{ ...st.card, padding: "4px 12px", marginBottom: 16 }}>
            {suggestions.map((s) => {
              const done = confirmedIds.includes(s.id);
              return (
                <div key={s.id} style={{ padding: "10px 4px", borderBottom: `1px solid ${t.border}`, opacity: done ? 0.45 : 1 }}>
                  <div style={{ fontSize: 12.5, marginBottom: 4 }}>{s.body.slice(0, 90)}</div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontWeight: 700, fontSize: 13, color: s.parsed.type === "income" ? BRAND.darkgreen : BRAND.crimson }}>
                      {toFaInt(s.parsed.amount)} ریال — {s.parsed.type === "income" ? "دریافت" : "پرداخت"}
                    </span>
                    {!done ? (
                      <button onClick={() => confirmSuggestion(s)} style={{ background: BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "5px 12px", fontSize: 11.5, fontWeight: 700, cursor: "pointer" }}>ثبت کن</button>
                    ) : (
                      <span style={{ fontSize: 11.5, color: BRAND.darkgreen, fontWeight: 700 }}>✓ ثبت شد</span>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}

      <SectionTitle text="یا متن پیامک را دستی بچسبان" />
      <div style={{ ...st.card, padding: 14 }}>
        <textarea value={text} onChange={(e) => setText(e.target.value)} placeholder="مثلا: از حساب شما مبلغ 250,000 ریال بابت خرید کسر شد..." rows={4}
          style={{ ...st.input, resize: "vertical", fontFamily: "inherit" }} />
        {parsed && (
          <div style={{ fontSize: 12.5, marginBottom: 10, color: "#8a8194" }}>
            تشخیص داده شده: {parsed.amount ? `${toFaInt(parsed.amount)} ریال` : "مبلغی یافت نشد"} — نوع: {parsed.type === "income" ? "دریافت" : "پرداخت"}
          </div>
        )}
        <button
          disabled={!parsed || !parsed.amount}
          onClick={() => { openWithPrefill({ amount: parsed.amount, type: parsed.type, note: parsed.note, categoryHint: parsed.categoryHint }); onBack(); }}
          style={{ ...st.primaryBtn, opacity: parsed && parsed.amount ? 1 : 0.5 }}>
          ادامه و تکمیل ثبت تراکنش
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Calendar View
--------------------------------------------------------- */
/* ---------------------------------------------------------
   Jalali date picker — replaces native <input type="date">,
   which always shows the Gregorian calendar regardless of the
   phone's language. This renders a Persian-calendar popup instead.
--------------------------------------------------------- */
function JalaliDateInput({ value, onChange, style }) {
  const t = useT();
  const [open, setOpen] = useState(false);
  const d = value ? new Date(value) : new Date();
  const [viewY, setViewY] = useState(jalaliYear(d));
  const [viewM, setViewM] = useState(jalaliParts(d).m);
  const cells = useMemo(() => getJalaliMonthCells(viewY, viewM), [viewY, viewM]);
  function prevMonth() { if (viewM === 1) { setViewM(12); setViewY((y) => y - 1); } else setViewM((m) => m - 1); }
  function nextMonth() { if (viewM === 12) { setViewM(1); setViewY((y) => y + 1); } else setViewM((m) => m + 1); }
  const weekDays = ["ش", "ی", "د", "س", "چ", "پ", "ج"];
  return (
    <div style={{ position: "relative", marginBottom: 10 }}>
      <button type="button" onClick={() => setOpen((v) => !v)} style={{ ...style, margin: 0, textAlign: "right", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span>{value ? faLongDate(new Date(value)) : "انتخاب تاریخ"}</span>
        <CalendarDays size={15} />
      </button>
      {open && (
        <div onClick={(e) => e.stopPropagation()} style={{ position: "absolute", zIndex: 40, top: "100%", marginTop: 6, right: 0, background: t.card, borderRadius: 12, padding: 12, boxShadow: "0 6px 20px rgba(0,0,0,0.25)", width: 260 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
            <button type="button" onClick={prevMonth} style={navArrowStyle(t)}><ChevronLeft size={14} /></button>
            <div style={{ fontWeight: 700, fontSize: 13 }}>{cells[0] ? faMonthYear(cells[0].date) : ""}</div>
            <button type="button" onClick={nextMonth} style={navArrowStyle(t)}><ChevronRight size={14} /></button>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 4 }}>
            {weekDays.map((w) => <div key={w} style={{ textAlign: "center", fontSize: 10, color: t.sub, fontWeight: 700 }}>{w}</div>)}
            {cells[0] && Array.from({ length: cells[0].weekday }).map((_, i) => <div key={"b" + i} />)}
            {cells.map((c) => {
              const iso = c.date.toISOString().slice(0, 10);
              const selected = value === iso;
              return (
                <button type="button" key={c.day} onClick={() => { onChange(iso); setOpen(false); }} style={{
                  aspectRatio: "1", border: "none", borderRadius: 7, cursor: "pointer", fontSize: 11,
                  background: selected ? BRAND.header : "transparent", color: selected ? "#fff" : t.text
                }}>{toFaInt(c.day)}</button>
              );
            })}
          </div>
          <button type="button" onClick={() => { onChange(todayISO()); setOpen(false); }} style={{ marginTop: 8, width: "100%", background: "none", border: "none", color: BRAND.header, fontWeight: 700, fontSize: 12, cursor: "pointer" }}>امروز</button>
        </div>
      )}
    </div>
  );
}

function CalendarViewSub({ transactions, catById }) {
  const t = useT();
  const now = new Date();
  const [jy, setJy] = useState(jalaliYear(now));
  const [jm, setJm] = useState(jalaliParts(now).m);
  const [selDay, setSelDay] = useState(null);
  const cells = useMemo(() => getJalaliMonthCells(jy, jm), [jy, jm]);
  const dayTx = (d) => transactions.filter((tx) => tx.date === d.date.toISOString().slice(0, 10));
  function prevMonth() { if (jm === 1) { setJm(12); setJy((y) => y - 1); } else setJm((m) => m - 1); setSelDay(null); }
  function nextMonth() { if (jm === 12) { setJm(1); setJy((y) => y + 1); } else setJm((m) => m + 1); setSelDay(null); }
  const weekDays = ["ش", "ی", "د", "س", "چ", "پ", "ج"];
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, marginBottom: 16 }}>
        <button onClick={prevMonth} style={navArrowStyle(t)}><ChevronLeft size={16} /></button>
        <div style={{ fontWeight: 700 }}>{cells[0] ? faMonthYear(cells[0].date) : ""}</div>
        <button onClick={nextMonth} style={navArrowStyle(t)}><ChevronRight size={16} /></button>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6, marginBottom: 8 }}>
        {weekDays.map((w) => <div key={w} style={{ textAlign: "center", fontSize: 11, fontWeight: 700, color: t.sub }}>{w}</div>)}
        {cells[0] && Array.from({ length: cells[0].weekday }).map((_, i) => <div key={"b" + i} />)}
        {cells.map((c) => {
          const txs = dayTx(c);
          const isSel = selDay && selDay.day === c.day;
          return (
            <button key={c.day} onClick={() => setSelDay(c)} style={{
              aspectRatio: "1", borderRadius: 9, border: "none", cursor: "pointer",
              background: isSel ? BRAND.header : t.card, color: isSel ? "#fff" : t.text,
              display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", position: "relative", fontSize: 13, fontWeight: 600
            }}>
              {toFaInt(c.day)}
              {txs.length > 0 && <span style={{ width: 5, height: 5, borderRadius: "50%", background: isSel ? "#fff" : BRAND.crimson, position: "absolute", bottom: 5 }} />}
            </button>
          );
        })}
      </div>
      {selDay && (
        <div style={{ marginTop: 16 }}>
          <SectionTitle text={faLongDate(selDay.date)} />
          {dayTx(selDay).length === 0 && <EmptyRow text="تراکنشی در این روز نیست" />}
          {dayTx(selDay).map((tx) => (
            <Row key={tx.id} title={catById(tx.categoryId)?.name || (tx.type === "transfer" ? "انتقال" : "—")} value={`${toFaInt(tx.amount)} ریال`}
              valueColor={tx.type === "expense" ? BRAND.crimson : tx.type === "income" ? BRAND.darkgreen : BRAND.violet} chevron={null} />
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Settings View
--------------------------------------------------------- */
function SettingsView({ settings, setSettings, exportBackup, importBackup, rebuildData, clearAllData }) {
  const st = useStyles();
  const t = useT();
  const fileRef = useRef();
  const [pinInput, setPinInput] = useState("");
  const [profile, setProfile] = useState(settings.profile || { name: "", phone: "", email: "" });
  return (
    <div>
      <SectionTitle text="اطلاعات کاربری" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <input placeholder="نام و نام خانوادگی" value={profile.name} onChange={(e) => setProfile({ ...profile, name: e.target.value })} style={st.input} />
        <input placeholder="شماره تماس" value={profile.phone} onChange={(e) => setProfile({ ...profile, phone: e.target.value })} style={st.input} />
        <input placeholder="ایمیل (اختیاری)" value={profile.email} onChange={(e) => setProfile({ ...profile, email: e.target.value })} style={st.input} />
        <button onClick={() => setSettings((s) => ({ ...s, profile }))} style={st.primaryBtn}>ذخیره اطلاعات کاربری</button>
      </div>

      <SectionTitle text="رنگ‌بندی برنامه" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {Object.entries(COLOR_PRESETS).map(([key, p]) => (
            <button key={key} onClick={() => setSettings((s) => ({ ...s, themeColor: key }))}
              style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer" }}>
              <span style={{
                width: 42, height: 42, borderRadius: "50%", background: p.header,
                border: settings.themeColor === key ? `3px solid ${p.fab}` : "3px solid transparent",
                boxShadow: settings.themeColor === key ? "0 0 0 2px #fff, 0 0 0 3px " + p.header : "none",
                display: "flex", alignItems: "center", justifyContent: "center"
              }}>
                {settings.themeColor === key && <Check size={18} color="#fff" />}
              </span>
              <span style={{ fontSize: 10.5, color: t.text }}>{p.name}</span>
            </button>
          ))}
        </div>
      </div>

      <SectionTitle text="ظاهر برنامه" />
      <div style={{ ...st.card, padding: "4px 12px", marginBottom: 18 }}>
        <Row title="تم روشن / تاریک" leftIcon={settings.theme === "dark" ? <Moon size={16} /> : <Sun size={16} />} leftColor={BRAND.header}
          extra={<button onClick={() => setSettings((s) => ({ ...s, theme: s.theme === "dark" ? "light" : "dark" }))} style={{ background: BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{settings.theme === "dark" ? "روشن کن" : "تاریک کن"}</button>}
          chevron={null} />
      </div>

      <SectionTitle text="اندازه فونت" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 8 }}>
          {[{ v: 0.9, l: "کوچک" }, { v: 1, l: "معمولی" }, { v: 1.15, l: "بزرگ" }, { v: 1.3, l: "خیلی بزرگ" }].map((o) => (
            <button key={o.v} onClick={() => setSettings((s) => ({ ...s, fontScale: o.v }))} style={pillStyle(settings.fontScale === o.v)}>{o.l}</button>
          ))}
        </div>
      </div>

      <SectionTitle text="تقویم و واحد پول" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ fontSize: 12, color: "#8a8194", marginBottom: 8 }}>تقویم</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <button onClick={() => setSettings((s) => ({ ...s, calendarMode: "jalali" }))} style={pillStyle(settings.calendarMode !== "gregorian")}>شمسی</button>
          <button onClick={() => setSettings((s) => ({ ...s, calendarMode: "gregorian" }))} style={pillStyle(settings.calendarMode === "gregorian")}>میلادی</button>
        </div>
        <div style={{ fontSize: 11.5, color: "#8a8194", marginBottom: 14 }}>در این نسخه فقط قالب تاریخ تغییر می‌کند؛ متن رابط کاربری همچنان فارسی می‌ماند.</div>
        <div style={{ fontSize: 12, color: "#8a8194", marginBottom: 8 }}>واحد نمایش پول</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setSettings((s) => ({ ...s, currency: "rial" }))} style={pillStyle(settings.currency === "rial")}>ریال</button>
          <button onClick={() => setSettings((s) => ({ ...s, currency: "toman" }))} style={pillStyle(settings.currency === "toman")}>تومان</button>
          <button onClick={() => setSettings((s) => ({ ...s, currency: "usd" }))} style={pillStyle(settings.currency === "usd")}>دلار</button>
        </div>
        {settings.currency === "usd" && (
          <>
            <div style={{ fontSize: 11.5, color: "#8a8194", marginBottom: 8 }}>اگر نرخ آنلاین در دسترس نبود، نرخ هر دلار به ریال را خودتان وارد کنید:</div>
            <AmountInput placeholder="مثلا 600000" value={settings.manualUsdRate} onChange={(v) => setSettings((s) => ({ ...s, manualUsdRate: v }))} style={st.input} />
          </>
        )}
      </div>

      <SectionTitle text="یادآوری‌ها و اعلان‌ها" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <label style={st.label}>هشدار سررسید چک چند روز قبل؟</label>
        <input value={settings.checkReminderDays} onChange={(e) => setSettings((s) => ({ ...s, checkReminderDays: Number(e.target.value.replace(/[^0-9]/g, "") || 0) }))} style={st.input} inputMode="numeric" />
        <Row title="اعلان تشخیص پیامک بانکی" leftIcon={<Bell size={16} />} leftColor={BRAND.orange}
          extra={<button onClick={() => setSettings((s) => ({ ...s, smsNotif: !s.smsNotif }))} style={{ background: settings.smsNotif ? BRAND.green : "#aaa", color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{settings.smsNotif ? "فعال" : "غیرفعال"}</button>}
          chevron={null} />
        <button onClick={async () => {
          try {
            const mod = await import("capacitor-sms-inbox");
            const plugin = mod.SMSInboxReader || mod.SmsInboxReader || mod.SMSInbox || mod.SmsInbox || mod.default;
            const result = await plugin?.requestPermissions?.();
            if (result?.sms === "granted") { setSettings((s) => ({ ...s, smsAutoRead: true })); alert("مجوز خواندن پیامک فعال شد."); }
            else alert("مجوز خواندن پیامک داده نشد.");
          } catch { alert("این قابلیت فقط در APK اندروید در دسترس است."); }
        }} style={{ ...st.primaryBtn, background: BRAND.header, marginTop: 10 }}>فعال‌سازی دسترسی پیامک گوشی</button>
        <div style={{ fontSize: 11.5, color: "#8a8194", marginTop: 6 }}>Rexa برای خواندن پیامک‌های بانکی به مجوز Android نیاز دارد. این دسترسی برای APK مستقیم قابل استفاده است؛ انتشار در Google Play تابع سیاست‌های SMS آن فروشگاه است.</div>
      </div>

      <SectionTitle text="هوش مصنوعی (اختیاری)" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ fontSize: 11.5, color: "#8a8194", marginBottom: 10 }}>
          تشخیص محلی متن/صدا فعال است و در نسخه‌های Native می‌تواند با سرویس AI شما تکمیل شود. اگر کلید API خودتان (مثلاً از Anthropic) را اینجا وارد کنید، تلاش می‌شود برای دقت بیشتر از آن استفاده شود — ولی چون تماس مستقیم از داخل اپ به سرورهای AI معمولاً با محدودیت CORS مواجه می‌شود، ممکن است نیاز به یک سرور واسط کوچک داشته باشید تا کاملاً قابل‌اعتماد شود.
        </div>
        <input placeholder="کلید API (اختیاری)" value={settings.aiApiKey} onChange={(e) => setSettings((s) => ({ ...s, aiApiKey: e.target.value }))} style={st.input} />
      </div>

      <SectionTitle text="امنیت" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ fontSize: 12.5, color: "#8a8194", marginBottom: 10 }}>اول یک رمز عددی تنظیم کن؛ اثر انگشت به‌عنوان راه سریع‌تر بازکردن، علاوه بر رمز عددی کار می‌کند (رمز عددی همیشه به‌عنوان جایگزین در دسترس می‌ماند).</div>
        <input placeholder="رمز عددی جدید (خالی = بدون قفل)" value={pinInput} onChange={(e) => setPinInput(e.target.value.replace(/[^0-9]/g, "").slice(0, 6))} style={st.input} inputMode="numeric" />
        <button onClick={() => setSettings((s) => ({ ...s, pin: pinInput }))} style={{ ...st.primaryBtn, marginBottom: 14 }}>{pinInput ? "تنظیم رمز" : "حذف رمز"}</button>
        <Row title="باز کردن با اثر انگشت" leftIcon={<Fingerprint size={16} />} leftColor={BRAND.violet}
          extra={<button onClick={() => {
            if (!settings.biometricEnabled && !settings.pin) { alert("ابتدا یک رمز عددی تنظیم کنید تا قفل امن Rexa فعال شود."); return; }
            setSettings((s) => ({ ...s, biometricEnabled: !s.biometricEnabled }));
          }} style={{ background: settings.biometricEnabled ? BRAND.green : "#aaa", color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{settings.biometricEnabled ? "فعال" : "غیرفعال"}</button>}
          chevron={null} />
        <div style={{ fontSize: 11, color: "#8a8194", marginTop: 6 }}>این گزینه از اثر انگشت یا Face ID واقعی گوشی استفاده می‌کند و فقط وقتی برنامه به‌صورت APK نصب شده باشد کار می‌کند (نه در پیش‌نمایش وب).</div>
      </div>

      <SectionTitle text="حساب مشترک خانوادگی" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ fontSize: 12.5, color: "#8a8194", marginBottom: 10 }}>
          با روشن‌کردن این گزینه، اطلاعات مالی (حساب‌ها، تراکنش‌ها، بودجه و...) به‌صورت مشترک ذخیره می‌شود و هر کسی که به همین برنامه دسترسی داشته باشد آن را می‌بیند. این حالت جایگزین ورود واقعی چند کاربره نیست، صرفاً یک دفتر مشترک است.
        </div>
        <Row title="فعال‌سازی حساب مشترک" leftIcon={<Users size={16} />} leftColor={BRAND.violet}
          extra={<button onClick={() => setSettings((s) => ({ ...s, sharedFamily: !s.sharedFamily }))} style={{ background: settings.sharedFamily ? BRAND.crimson : BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{settings.sharedFamily ? "غیرفعال کن" : "فعال کن"}</button>}
          chevron={null} />
      </div>

      <SectionTitle text="پشتیبان‌گیری و بازیابی" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <button onClick={exportBackup} style={{ ...st.primaryBtn, background: BRAND.header, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, marginBottom: 10 }}><Download size={16} /> دریافت فایل پشتیبان</button>
        <input ref={fileRef} type="file" accept="application/json" style={{ display: "none" }} onChange={(e) => e.target.files[0] && importBackup(e.target.files[0])} />
        <button onClick={() => fileRef.current.click()} style={{ ...st.primaryBtn, background: BRAND.gold, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Upload size={16} /> بازیابی از فایل پشتیبان</button>
      </div>

      <SectionTitle text="نگهداری و تعمیر اطلاعات" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <button onClick={rebuildData} style={{ ...st.primaryBtn, background: BRAND.violet, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, marginBottom: 10 }}><RotateCcw size={16} /> بازسازی اطلاعات</button>
        <button onClick={clearAllData} style={{ ...st.primaryBtn, background: BRAND.crimson, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Eraser size={16} /> خام کردن اطلاعات</button>
        <div style={{ fontSize: 11.5, color: "#8a8194", marginTop: 9, lineHeight: 1.8 }}>قبل از هر دو عملیات پشتیبان تهیه می‌شود. «خام کردن» اطلاعات مالی این دستگاه را پاک و حساب‌های پایه را دوباره ایجاد می‌کند.</div>
      </div>

      <SectionTitle text="درباره همگام‌سازی با Google Drive" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18, fontSize: 12.5, color: "#8a8194" }}>
        اتصال مستقیم به Google Drive نیازمند ورود واقعی به حساب گوگل است که در این محیط در دسترس نیست. برای انتقال اطلاعات بین دستگاه‌ها از «دریافت فایل پشتیبان» استفاده کنید و همان فایل را در Drive خودتان نگه دارید یا در دستگاه دیگر «بازیابی» کنید.
      </div>

      <SectionTitle text="شخصی‌سازی نمای صفحه اول" />
      <div style={{ ...st.card, padding: "8px 12px", marginBottom: 18 }}>
        <div style={{ fontSize: 11.5, color: "#8a8194", padding: "6px 4px 10px" }}>هر بخش را می‌توانید نمایش/مخفی کنید یا با فلش‌ها ترتیبش را عوض کنید.</div>
        {(settings.homeSections || DEFAULT_HOME_SECTIONS).map((sec, i, arr) => {
          const meta = HOME_SECTION_LABELS[sec.key] || sec.key;
          function move(dir) {
            const idx = i + dir;
            if (idx < 0 || idx >= arr.length) return;
            const copy = [...arr];
            [copy[i], copy[idx]] = [copy[idx], copy[i]];
            setSettings((s) => ({ ...s, homeSections: copy }));
          }
          return (
            <div key={sec.key} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 4px", borderBottom: "1px solid #f0eef3" }}>
              <span style={{ fontSize: 13, fontWeight: 600, opacity: sec.visible ? 1 : 0.4 }}>{meta}</span>
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <button onClick={() => move(-1)} style={miniBtn}><ArrowUp size={13} /></button>
                <button onClick={() => move(1)} style={miniBtn}><ArrowDown size={13} /></button>
                <button onClick={() => setSettings((s) => ({ ...s, homeSections: arr.map((x) => x.key === sec.key ? { ...x, visible: !x.visible } : x) }))}
                  style={{ ...miniBtn, background: sec.visible ? BRAND.green : "#aaa", color: "#fff" }}>
                  {sec.visible ? <Eye size={13} /> : <EyeOff size={13} />}
                </button>
              </div>
            </div>
          );
        })}
      </div>

      <SectionTitle text="آموزش" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18, fontSize: 13, lineHeight: 2, color: t.text }}>
        <div>• برای ثبت سریع تراکنش، دکمه‌ی سبز + پایین صفحه را بزن.</div>
        <div>• سه‌بار پشت‌سرهم روی هر جای صفحه بزن تا کادر «ثبت سریع» باز شود؛ هرچی تایپ کنی خودش تشخیص می‌ده هزینه بوده یا درآمد.</div>
        <div>• آیکون میکروفون بالای صفحه برای ثبت با صدا است (اگر گوشی‌ات پشتیبانی کند).</div>
        <div>• روی کارت‌های بانکی بالای صفحه چپ‌وراست بکش تا همه‌ی حساب‌هایت را ببینی.</div>
        <div>• از «شخصی‌سازی نمای صفحه اول» بالاتر همین صفحه، می‌توانی ترتیب و نمایش بخش‌های خانه را عوض کنی.</div>
      </div>

      <SectionTitle text="ارسال برنامه به دیگران" />
      <div style={{ ...st.card, padding: 14 }}>
        <button
          onClick={async () => {
            const text = "برنامه حسابداری Rexa رو امتحان کن!";
            if (navigator.share) { try { await navigator.share({ title: "Rexa", text }); } catch {} }
            else { try { await navigator.clipboard.writeText(text); alert("متن معرفی کپی شد."); } catch { alert(text); } }
          }}
          style={{ ...st.primaryBtn, background: BRAND.violet, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
          <Users size={16} /> اشتراک‌گذاری برنامه
        </button>
      </div>
    </div>
  );
}
const miniBtn = { width: 26, height: 26, borderRadius: 7, border: "none", background: "#eee", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" };
const HOME_SECTION_LABELS = {
  shortcut: "میانبر تراکنش‌ها", expense: "هزینه‌ها", income: "درآمدها", banks: "بانک‌ها و کارت‌ها",
  funds: "صندوق‌ها", balrep: "گزارش مانده حساب‌ها", budget: "بودجه‌بندی", loanchk: "وام‌ها و چک‌ها", contacts: "اشخاص و بدهی‌ها", bills: "یادآوری قبض‌ها",
};

/* ---------------------------------------------------------
   Add Transaction Sheet
--------------------------------------------------------- */
/* ---------------------------------------------------------
   Quick-add action sheet (opened from the center + button)
--------------------------------------------------------- */
function QuickAddSheet({ onClose, onPick }) {
  const st = useStyles();
  const options = [
    { key: "expense", label: "ثبت هزینه", icon: <TrendingDown size={20} />, color: BRAND.crimson },
    { key: "income", label: "ثبت درآمد", icon: <TrendingUp size={20} />, color: BRAND.darkgreen },
    { key: "transfer", label: "انتقال بین حساب‌ها", icon: <ArrowLeftRight size={20} />, color: BRAND.violet },
    { key: "check", label: "ثبت چک", icon: <FileSpreadsheet size={20} />, color: BRAND.gold },
  ];
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "flex-end", zIndex: 55, maxWidth: 480, margin: "0 auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", width: "100%", borderRadius: "18px 18px 0 0", padding: "20px 18px calc(env(safe-area-inset-bottom, 0px) + 20px)" }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16, textAlign: "center" }}>چی می‌خوای ثبت کنی؟</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 12 }}>
          {options.map((o) => (
            <button key={o.key} onClick={() => onPick(o.key)} style={{
              display: "flex", flexDirection: "column", alignItems: "center", gap: 8, padding: "18px 8px",
              borderRadius: 14, border: "none", background: "#f7f5fa", cursor: "pointer"
            }}>
              <span style={{ width: 44, height: 44, borderRadius: 12, background: o.color, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center" }}>{o.icon}</span>
              <span style={{ fontSize: 12.5, fontWeight: 700, color: "#241a30" }}>{o.label}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Simple text modal (used for the quick "یادداشت" note capture)
--------------------------------------------------------- */
function SimpleTextModal({ title, placeholder, onClose, onSubmit }) {
  const st = useStyles();
  const [text, setText] = useState("");
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "flex-end", zIndex: 55, maxWidth: 480, margin: "0 auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", width: "100%", borderRadius: "18px 18px 0 0", padding: "20px 18px calc(env(safe-area-inset-bottom, 0px) + 20px)" }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 12 }}>{title}</div>
        <textarea autoFocus value={text} onChange={(e) => setText(e.target.value)} placeholder={placeholder} rows={4}
          style={{ ...st.input, resize: "vertical", fontFamily: "inherit" }} />
        <button disabled={!text.trim()} onClick={() => onSubmit(text.trim())} style={{ ...st.primaryBtn, opacity: text.trim() ? 1 : 0.5 }}>ذخیره</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Reminder quick modal
--------------------------------------------------------- */
function ReminderQuickModal({ onClose, onSubmit }) {
  const st = useStyles();
  const [text, setText] = useState("");
  const [date, setDate] = useState(initial?.date || todayISO());
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "flex-end", zIndex: 55, maxWidth: 480, margin: "0 auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", width: "100%", borderRadius: "18px 18px 0 0", padding: "20px 18px calc(env(safe-area-inset-bottom, 0px) + 20px)" }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 12 }}>یادآوری جدید</div>
        <input autoFocus value={text} onChange={(e) => setText(e.target.value)} placeholder="مثلا: تماس با بانک" style={st.input} />
        <label style={st.label}>تاریخ</label>
        <JalaliDateInput value={date} onChange={(v) => setDate(v)} style={st.input} />
        <button disabled={!text.trim()} onClick={() => onSubmit({ text: text.trim(), date, done: false })} style={{ ...st.primaryBtn, opacity: text.trim() ? 1 : 0.5 }}>ذخیره یادآوری</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Smart capture overlay — opened by triple-tap anywhere on the
   screen. Semi-transparent so the page behind stays visible.
   Text is parsed locally (same engine as the SMS/voice parser).
--------------------------------------------------------- */
function SmartCaptureOverlay({ onClose, onParsed }) {
  const st = useStyles();
  const [text, setText] = useState("");
  const parsed = text.trim() ? parseBankSms(text) : null;
  return (
    <div onClick={onClose} style={{
      position: "fixed", inset: 0, background: "rgba(20,10,30,0.4)", backdropFilter: "blur(2px)",
      display: "flex", alignItems: "center", justifyContent: "center", zIndex: 90
    }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        width: "88%", maxWidth: 380, background: "rgba(255,255,255,0.92)", borderRadius: 18, padding: 18,
        boxShadow: "0 8px 30px rgba(0,0,0,0.3)"
      }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, fontWeight: 700, fontSize: 14, color: BRAND.header }}>
            <Sparkles size={17} /> ثبت سریع
          </div>
          <VoiceCaptureButton dark onResult={(t) => setText((prev) => (prev ? prev + " " : "") + t)} />
        </div>
        <textarea autoFocus value={text} onChange={(e) => setText(e.target.value)} rows={3}
          placeholder="بنویس یا میکروفن رو بزن و بگو چی خرج شد یا چی دریافت کردی..." style={{ ...st.input, resize: "vertical", fontFamily: "inherit", background: "rgba(255,255,255,0.7)" }} />
        {parsed && (
          <div style={{ fontSize: 12, color: "#6a6275", marginBottom: 10 }}>
            تشخیص: {parsed.amount ? `${toFaInt(parsed.amount)} ریال` : "مبلغی پیدا نشد"} — {parsed.type === "income" ? "دریافت" : "پرداخت"}{parsed.categoryHint ? ` — ${parsed.categoryHint}` : ""}
          </div>
        )}
        <button disabled={!parsed?.amount} onClick={() => onParsed(parsed)} style={{ ...st.primaryBtn, opacity: parsed?.amount ? 1 : 0.5 }}>ادامه ثبت</button>
      </div>
    </div>
  );
}

function AddTransactionSheet({ accounts, categories, favorites, members = [], events = [], projects = [], initial, onClose, onSubmit, onAddAccount, onAddCategory }) {
  const st = useStyles();
  const [type, setType] = useState(initial?.type || "expense");
  const [amount, setAmount] = useState(initial?.amount ? String(initial.amount) : "");
  const [categoryId, setCategoryId] = useState(() => {
    if (initial?.categoryId && categories.some((c) => c.id === initial.categoryId)) return initial.categoryId;
    if (initial?.categoryHint) {
      const match = categories.find((c) => c.name === initial.categoryHint && c.kind === (initial.type || "expense"));
      if (match) return match.id;
    }
    return "";
  });
  const [accountId, setAccountId] = useState(initial?.accountId || accounts[0]?.id || "");
  const [toAccountId, setToAccountId] = useState(initial?.toAccountId || accounts[1]?.id || accounts[0]?.id || "");
  const [date, setDate] = useState(initial?.date || todayISO());
  const [note, setNote] = useState(initial?.note || "");
  const [tagsInput, setTagsInput] = useState((initial?.tags || []).join(", "));
  const [photo, setPhoto] = useState(initial?.photo || null);
  const [showMore, setShowMore] = useState(false);
  const [memberId, setMemberId] = useState("");
  const [eventId, setEventId] = useState("");
  const [projectId, setProjectId] = useState("");
  const [inlineAdd, setInlineAdd] = useState(null);
  const [inlineName, setInlineName] = useState("");

  const filteredCats = categories.filter((c) => c.kind === type);
  const favCats = filteredCats.filter((c) => favorites.categories.includes(c.id));
  const canSubmit = amount && Number(amount) > 0 && accountId && (type === "transfer" ? toAccountId && toAccountId !== accountId : categoryId);

  async function handlePhoto(e) {
    const f = e.target.files[0];
    if (!f) return;
    const dataUrl = await resizeImage(f);
    setPhoto(dataUrl);
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 120, maxWidth: 480, margin: "0 auto", paddingBottom: "calc(82px + env(safe-area-inset-bottom, 0px))" }}>
      <div style={{ background: "#fff", width: "100%", borderRadius: "18px", padding: "18px 18px 24px", maxHeight: "calc(88vh - 70px)", overflowY: "auto", boxShadow: "0 -8px 30px rgba(0,0,0,.18)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: "#8a8194" }}><X size={22} /></button>
          <div style={{ fontWeight: 700, fontSize: 16 }}>{initial?._editId ? "ویرایش تراکنش" : "ثبت تراکنش جدید"}</div>
          <div style={{ width: 22 }} />
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          {[{ k: "expense", l: "پرداخت", c: BRAND.crimson }, { k: "income", l: "دریافت", c: BRAND.darkgreen }, { k: "transfer", l: "انتقال", c: BRAND.violet }].map((o) => (
            <button key={o.k} onClick={() => { setType(o.k); setCategoryId(""); }} style={{ flex: 1, padding: "10px 4px", borderRadius: 9, border: `1.5px solid ${type === o.k ? o.c : "#e3e0ea"}`, background: type === o.k ? o.c : "#fff", color: type === o.k ? "#fff" : "#241a30", fontWeight: 700, cursor: "pointer", fontSize: 13.5 }}>{o.l}</button>
          ))}
        </div>

        {favCats.length > 0 && (
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 12 }}>
            {favCats.map((c) => (
              <button key={c.id} onClick={() => setCategoryId(c.id)} style={{ display: "flex", alignItems: "center", gap: 4, background: categoryId === c.id ? BRAND.header : "#f1eef4", color: categoryId === c.id ? "#fff" : "#3E1461", border: "none", borderRadius: 20, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                <Star size={11} fill="#f5b301" color="#f5b301" /> {c.name}
              </button>
            ))}
          </div>
        )}

        <label style={st.label}>مبلغ (ریال)</label>
        <AmountInput value={amount} onChange={setAmount} placeholder="0" style={st.input} />

        {type !== "transfer" ? (
          <>
            <label style={st.label}>دسته‌بندی</label>
            <div style={{ display: "flex", gap: 6 }}>
              <select value={categoryId} onChange={(e) => setCategoryId(e.target.value)} style={{ ...st.input, marginBottom: 8, flex: 1 }}>
                <option value="">انتخاب کنید</option>{filteredCats.map((c) => <option key={c.id} value={c.id}>{c.parentId ? "↳ " : ""}{c.name}</option>)}
              </select>
              <button onClick={() => { setInlineAdd("category"); setInlineName(""); }} style={{ ...miniBtn, width: 44, height: 44, background: BRAND.header, color: "#fff" }}><Plus size={18} /></button>
            </div>
            <label style={st.label}>{type === "expense" ? "از حساب" : "به حساب"}</label>
            <div style={{ display: "flex", gap: 6 }}>
              <select value={accountId} onChange={(e) => setAccountId(e.target.value)} style={{ ...st.input, marginBottom: 8, flex: 1 }}>
                {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
              <button onClick={() => { setInlineAdd("account"); setInlineName(""); }} style={{ ...miniBtn, width: 44, height: 44, background: BRAND.violet, color: "#fff" }}><Plus size={18} /></button>
            </div>
            {inlineAdd && <div style={{ background: "#f7f4fa", borderRadius: 10, padding: 10, marginBottom: 10 }}>
              <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 6 }}>{inlineAdd === "category" ? "دسته جدید" : "حساب جدید"}</div>
              <div style={{ display: "flex", gap: 6 }}>
                <input autoFocus value={inlineName} onChange={(e) => setInlineName(e.target.value)} placeholder={inlineAdd === "category" ? "نام دسته" : "نام حساب"} style={{ ...st.input, margin: 0, flex: 1 }} />
                <button onClick={() => {
                  if (!inlineName.trim()) return;
                  const id = uid();
                  if (inlineAdd === "category") { onAddCategory?.({ id, name: inlineName.trim(), kind: type }); setCategoryId(id); }
                  else { onAddAccount?.({ id, name: inlineName.trim(), type: "bank", initial: 0 }); setAccountId(id); }
                  setInlineAdd(null);
                }} style={{ ...st.primaryBtn, width: "auto", padding: "0 14px" }}>افزودن</button>
              </div>
            </div>}
          </>
        ) : (
          <>
            <label style={st.label}>از حساب</label>
            <select value={accountId} onChange={(e) => setAccountId(e.target.value)} style={st.input}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
            <label style={st.label}>به حساب</label>
            <select value={toAccountId} onChange={(e) => setToAccountId(e.target.value)} style={st.input}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </>
        )}

        <label style={st.label}>تاریخ</label>
        <JalaliDateInput value={date} onChange={(v) => setDate(v)} style={st.input} />

        <label style={st.label}>برچسب‌ها (با کاما جدا کنید)</label>
        <input value={tagsInput} onChange={(e) => setTagsInput(e.target.value)} placeholder="مثلا: سفر, ضروری" style={st.input} />

        <label style={st.label}>یادداشت (اختیاری)</label>
        <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="توضیحات..." style={st.input} />

        <label style={st.label}>پیوست عکس (اختیاری)</label>
        <label style={{ display: "flex", alignItems: "center", gap: 8, border: "1.5px dashed #d8d2e0", borderRadius: 9, padding: "10px 12px", cursor: "pointer", marginBottom: 12, fontSize: 13, color: "#8a8194" }}>
          <ImageIcon size={16} /> {photo ? "عکس انتخاب شد ✓" : "افزودن عکس رسید"}
          <input type="file" accept="image/*" onChange={handlePhoto} style={{ display: "none" }} />
        </label>
        {photo && <img src={photo} style={{ width: 60, height: 60, borderRadius: 8, objectFit: "cover", marginBottom: 12 }} />}

        {(members.length > 0 || events.length > 0 || projects.length > 0) && (
          <div style={{ marginBottom: 12 }}>
            <button onClick={() => setShowMore((v) => !v)} style={{ background: "none", border: "none", color: BRAND.header, fontWeight: 700, fontSize: 12.5, cursor: "pointer", padding: 0, marginBottom: showMore ? 10 : 0 }}>
              {showMore ? "بستن جزئیات بیشتر ▲" : "افزودن عضو / رویداد / پروژه ▼"}
            </button>
            {showMore && (
              <>
                {members.length > 0 && (<>
                  <label style={st.label}>عضو خانواده</label>
                  <select value={memberId} onChange={(e) => setMemberId(e.target.value)} style={st.input}>
                    <option value="">— بدون عضو —</option>
                    {members.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
                  </select>
                </>)}
                {events.length > 0 && (<>
                  <label style={st.label}>رویداد</label>
                  <select value={eventId} onChange={(e) => setEventId(e.target.value)} style={st.input}>
                    <option value="">— بدون رویداد —</option>
                    {events.map((ev) => <option key={ev.id} value={ev.id}>{ev.name}</option>)}
                  </select>
                </>)}
                {projects.length > 0 && (<>
                  <label style={st.label}>پروژه</label>
                  <select value={projectId} onChange={(e) => setProjectId(e.target.value)} style={st.input}>
                    <option value="">— بدون پروژه —</option>
                    {projects.map((pr) => <option key={pr.id} value={pr.id}>{pr.name}</option>)}
                  </select>
                </>)}
              </>
            )}
          </div>
        )}

        <button disabled={!canSubmit}
          onClick={() => onSubmit({
            type, amount: Number(amount), categoryId, accountId,
            toAccountId: type === "transfer" ? toAccountId : undefined,
            date, note, photo: photo || undefined,
            tags: tagsInput.split(",").map((s) => s.trim()).filter(Boolean),
            memberId: memberId || undefined, eventId: eventId || undefined, projectId: projectId || undefined,
          })}
          style={{ ...st.primaryBtn, opacity: canSubmit ? 1 : 0.5, marginTop: 6 }}>
          <span style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Check size={17} /> ثبت تراکنش</span>
        </button>
      </div>
    </div>
  );
}


/* ---------------------------------------------------------
   Parmis-style modules: persons, debts, currencies, calculator, support
--------------------------------------------------------- */
function GoalsManager({ goals, setGoals }) {
  const st = useStyles();
  const [form, setForm] = useState({ title: "", target: "", saved: "", dueDate: "", note: "" });
  function add() {
    const target = Number(form.target || 0), saved = Number(form.saved || 0);
    if (!form.title.trim() || target <= 0) return;
    setGoals((p) => [{ id: uid(), title: form.title.trim(), target, saved: Math.min(saved, target), dueDate: form.dueDate || "", note: form.note.trim(), createdAt: new Date().toISOString() }, ...p]);
    setForm({ title: "", target: "", saved: "", dueDate: "", note: "" });
  }
  function changeSaved(id, delta) {
    setGoals((p) => p.map((g) => g.id === id ? { ...g, saved: Math.max(0, Math.min(g.target, Number(g.saved || 0) + delta)) } : g));
  }
  const totalTarget = goals.reduce((s, g) => s + Number(g.target || 0), 0);
  const totalSaved = goals.reduce((s, g) => s + Number(g.saved || 0), 0);
  return <div>
    <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
      <div style={{ fontWeight: 800, marginBottom: 6 }}>هدف مالی جدید</div>
      <div style={{ fontSize: 12, color: useT().sub, marginBottom: 10 }}>مثلاً خرید خودرو، سفر، پس‌انداز اضطراری یا خرید خانه.</div>
      <input value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} placeholder="نام هدف" style={st.input} />
      <AmountInput placeholder="مبلغ هدف" value={form.target} onChange={(v) => setForm({ ...form, target: v })} style={st.input} />
      <AmountInput placeholder="مبلغ پس‌انداز فعلی (اختیاری)" value={form.saved} onChange={(v) => setForm({ ...form, saved: v })} style={st.input} />
      <label style={st.label}>تاریخ هدف (اختیاری)</label>
      <JalaliDateInput value={form.dueDate || todayISO()} onChange={(v) => setForm({ ...form, dueDate: v })} style={st.input} />
      <input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="توضیحات" style={st.input} />
      <button onClick={add} style={{ ...st.primaryBtn, background: BRAND.green }}>افزودن هدف</button>
    </div>
    <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: useT().sub, marginBottom: 8 }}>
        <span>مجموع اهداف</span><b style={{ color: useT().text }}>{toFaInt(totalTarget)} ریال</b>
      </div>
      <div style={{ height: 10, background: useT().border, borderRadius: 20, overflow: "hidden" }}><div style={{ height: "100%", width: `${totalTarget ? Math.min(100, totalSaved / totalTarget * 100) : 0}%`, background: BRAND.green, borderRadius: 20 }} /></div>
      <div style={{ marginTop: 7, textAlign: "center", fontSize: 12, color: BRAND.green, fontWeight: 700 }}>پس‌انداز شده: {toFaInt(totalSaved)} ریال</div>
    </div>
    <div style={{ ...st.card, padding: "4px 12px" }}>
      {goals.length === 0 && <EmptyRow text="هنوز هدف مالی ثبت نشده" />}
      {goals.map((g) => {
        const pct = g.target ? Math.min(100, Number(g.saved || 0) / g.target * 100) : 0;
        return <div key={g.id} style={{ padding: "14px 2px", borderBottom: "1px solid #f0eef3" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
            <div><div style={{ fontWeight: 800 }}>{g.title}</div>{g.dueDate && <div style={{ fontSize: 11.5, color: useT().sub, marginTop: 3 }}>هدف: {faLongDate(new Date(g.dueDate))}</div>}</div>
            <button onClick={() => setGoals((p) => p.filter((x) => x.id !== g.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15}/></button>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginTop: 10 }}><span>{toFaInt(g.saved)} ریال</span><b>{toFaInt(g.target)} ریال</b></div>
          <div style={{ height: 9, background: useT().border, borderRadius: 20, overflow: "hidden", marginTop: 6 }}><div style={{ height: "100%", width: `${pct}%`, background: BRAND.green, borderRadius: 20 }} /></div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
            <span style={{ fontSize: 11.5, color: BRAND.green, fontWeight: 700 }}>{toFaInt(Math.round(pct))}% تکمیل</span>
            <div style={{ display: "flex", gap: 6 }}><button onClick={() => changeSaved(g.id, -100000)} style={miniBtn}>−۱۰۰هزار</button><button onClick={() => changeSaved(g.id, 100000)} style={{ ...miniBtn, color: BRAND.green }}>+۱۰۰هزار</button></div>
          </div>
        </div>;
      })}
    </div>
  </div>;
}

function PersonsManager({ persons, setPersons }) {
  const st = useStyles();
  const [form, setForm] = useState({ name: "", phone: "", note: "" });
  const [editing, setEditing] = useState(null);
  function save() {
    if (!form.name.trim()) return;
    if (editing) setPersons((p) => p.map((x) => x.id === editing ? { ...x, ...form, name: form.name.trim() } : x));
    else setPersons((p) => [{ id: uid(), ...form, name: form.name.trim() }, ...p]);
    setForm({ name: "", phone: "", note: "" }); setEditing(null);
  }
  return <div>
    <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
      <div style={{ fontWeight: 800, marginBottom: 10 }}>تعریف شخص / طرف حساب</div>
      <input placeholder="نام و نام خانوادگی / فروشگاه" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={st.input} />
      <input placeholder="شماره تماس" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} style={st.input} inputMode="tel" />
      <input placeholder="توضیحات" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} style={st.input} />
      <button onClick={save} style={st.primaryBtn}>{editing ? "ذخیره تغییرات" : "افزودن شخص"}</button>
    </div>
    <div style={{ ...st.card, padding: "4px 12px" }}>
      {persons.length === 0 && <EmptyRow text="شخصی ثبت نشده" />}
      {persons.map((p) => <Row key={p.id} title={p.name} subtitle={[p.phone, p.note].filter(Boolean).join(" · ") || "طرف حساب"}
        extra={<div style={{ display: "flex", gap: 5 }}><button onClick={() => { setEditing(p.id); setForm({ name: p.name, phone: p.phone || "", note: p.note || "" }); }} style={miniBtn}><Pencil size={13}/></button></div>}
        leftIcon={<Trash2 size={15}/>} leftColor={BRAND.crimson} onClick={() => setPersons((all) => all.filter((x) => x.id !== p.id))} chevron={null} />)}
    </div>
    <div style={{ fontSize: 12, color: useT().sub, marginTop: 10 }}>این بخش برای ثبت اشخاص، فروشگاه‌ها و طرف حساب‌هاست و پایه‌ی مدیریت بدهکار/بستانکار را تشکیل می‌دهد.</div>
  </div>;
}

function DebtsManager({ debts, setDebts, persons, accounts }) {
  const st = useStyles();
  const [form, setForm] = useState({ personId: "", kind: "receivable", amount: "", dueDate: todayISO(), note: "" });
  function add() {
    const amount = Number(form.amount || 0); if (!form.personId || !amount) return;
    setDebts((p) => [{ id: uid(), ...form, amount }, ...p]);
    setForm({ personId: "", kind: "receivable", amount: "", dueDate: todayISO(), note: "" });
  }
  const open = debts.filter((d) => !d.settled);
  return <div>
    <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
      <div style={{ fontWeight: 800, marginBottom: 10 }}>ثبت بدهی / طلب</div>
      <select value={form.personId} onChange={(e) => setForm({ ...form, personId: e.target.value })} style={st.input}><option value="">انتخاب طرف حساب</option>{persons.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}</select>
      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}><button onClick={() => setForm({ ...form, kind: "receivable" })} style={pillStyle(form.kind === "receivable")}>طلب از دیگران</button><button onClick={() => setForm({ ...form, kind: "payable" })} style={pillStyle(form.kind === "payable")}>بدهی به دیگران</button></div>
      <AmountInput placeholder="مبلغ" value={form.amount} onChange={(v) => setForm({ ...form, amount: v })} style={st.input} />
      <JalaliDateInput value={form.dueDate} onChange={(v) => setForm({ ...form, dueDate: v })} style={st.input} />
      <input placeholder="توضیحات" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} style={st.input} />
      <button onClick={add} style={st.primaryBtn}>ثبت</button>
    </div>
    <div style={{ ...st.card, padding: "4px 12px" }}>
      {open.length === 0 && <EmptyRow text="بدهی یا طلب باز وجود ندارد" />}
      {open.map((d) => { const person = persons.find((p) => p.id === d.personId); return <Row key={d.id} title={person?.name || "طرف حساب حذف‌شده"} subtitle={`${d.kind === "receivable" ? "طلب" : "بدهی"} · سررسید ${faLongDate(new Date(d.dueDate))}`} value={`${toFaInt(d.amount)} ریال`} valueColor={d.kind === "receivable" ? BRAND.darkgreen : BRAND.crimson}
        extra={<button onClick={() => setDebts((p) => p.map((x) => x.id === d.id ? { ...x, settled: true, settledAt: todayISO() } : x))} style={{ ...miniBtn, color: BRAND.green }}><Check size={14}/></button>} chevron={null} />; })}
    </div>
  </div>;
}

function CurrenciesManager({ currencies, setCurrencies }) {
  const st = useStyles();
  const defaults = [{ code: "IRR", name: "ریال", rate: 1 }, { code: "USD", name: "دلار", rate: "" }, { code: "EUR", name: "یورو", rate: "" }, { code: "AED", name: "درهم", rate: "" }, { code: "GBP", name: "پوند", rate: "" }];
  const [form, setForm] = useState({ code: "", name: "", rate: "" });
  const list = currencies.length ? currencies : defaults;
  function add() { if (!form.code || !form.name) return; setCurrencies((p) => [...p, { id: uid(), code: form.code.toUpperCase(), name: form.name, rate: Number(form.rate || 0) }]); setForm({ code: "", name: "", rate: "" }); }
  return <div>
    <div style={{ ...st.card, padding: 14, marginBottom: 16 }}><div style={{ fontWeight: 800, marginBottom: 10 }}>واحدهای پولی و نرخ تبدیل</div><input placeholder="کد ارز مثل USD" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} style={st.input}/><input placeholder="نام ارز" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={st.input}/><AmountInput placeholder="نرخ هر واحد به ریال" value={form.rate} onChange={(v) => setForm({ ...form, rate: v })} style={st.input}/><button onClick={add} style={st.primaryBtn}>افزودن ارز</button></div>
    <div style={{ ...st.card, padding: "4px 12px" }}>{list.map((c, i) => <Row key={c.id || c.code} title={`${c.name} (${c.code})`} value={c.code === "IRR" ? "۱" : c.rate ? `${toFaInt(c.rate)} ریال` : "نرخ ثبت نشده"} valueColor={BRAND.header} extra={i >= defaults.length ? <button onClick={() => setCurrencies((p) => p.filter((x) => x.id !== c.id))} style={{ ...miniBtn, color: BRAND.crimson }}><Trash2 size={13}/></button> : null} chevron={null} />)}</div>
  </div>;
}

function CalculatorView() {
  const st = useStyles();
  const [expr, setExpr] = useState("");
  const [result, setResult] = useState("");
  function calc() { try { if (!/^[0-9+\-*/().% ]+$/.test(expr)) return; const value = Function(`"use strict"; return (${expr})`)(); if (Number.isFinite(value)) setResult(String(value)); } catch {} }
  const keys = ["7","8","9","/","4","5","6","*","1","2","3","-","0",".","=","+"];
  return <div><div style={{ ...st.card, padding: 16, marginBottom: 14 }}><div style={{ fontSize: 12, color: useT().sub, marginBottom: 6 }}>ماشین حساب</div><div style={{ fontSize: 26, fontWeight: 800, direction: "ltr", textAlign: "left", minHeight: 38 }}>{result || expr || "۰"}</div></div><div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>{keys.map((k) => <button key={k} onClick={() => k === "=" ? calc() : setExpr((x) => x + k)} style={{ height: 52, border: "none", borderRadius: 12, background: k === "=" ? BRAND.header : "#fff", color: k === "=" ? "#fff" : "#241a30", fontSize: 18, fontWeight: 700 }}>{k}</button>)}<button onClick={() => { setExpr(""); setResult(""); }} style={{ gridColumn: "1 / -1", height: 44, border: "none", borderRadius: 12, background: BRAND.crimson, color: "#fff", fontWeight: 700 }}>پاک کردن</button></div></div>;
}

function SupportView() {
  const st = useStyles();
  const [message, setMessage] = useState("");
  const [tickets, setTickets] = useState(() => { try { return JSON.parse(localStorage.getItem("rexa:tickets") || "[]"); } catch { return []; } });
  function send() { if (!message.trim()) return; const next = [{ id: uid(), text: message.trim(), date: new Date().toISOString(), status: "ارسال‌شده" }, ...tickets]; setTickets(next); localStorage.setItem("rexa:tickets", JSON.stringify(next)); setMessage(""); }
  return <div><div style={{ ...st.card, padding: 14, marginBottom: 16 }}><div style={{ fontWeight: 800, marginBottom: 8 }}>پشتیبانی و تیکت</div><div style={{ fontSize: 12, color: useT().sub, marginBottom: 10 }}>پیام خود را ثبت کنید تا در سابقه تیکت‌های این دستگاه نگهداری شود.</div><textarea value={message} onChange={(e) => setMessage(e.target.value)} placeholder="متن درخواست یا مشکل..." style={{ ...st.input, minHeight: 120, resize: "vertical" }}/><button onClick={send} style={st.primaryBtn}>ثبت تیکت</button></div><div style={{ ...st.card, padding: "4px 12px" }}>{tickets.length === 0 ? <EmptyRow text="تیکتی ثبت نشده"/> : tickets.map((x) => <Row key={x.id} title={x.text} subtitle={faLongDate(new Date(x.date))} value={x.status} valueColor={BRAND.teal} chevron={null}/>)}</div></div>;
}

/* ---------------------------------------------------------
   Side menu
--------------------------------------------------------- */
function ProfileView({ settings, setSettings }) { const st=useStyles(); const p=settings.profile||{}; const [f,setF]=useState({name:p.name||"",phone:p.phone||"",email:p.email||""}); return <div><div style={{...st.card,padding:16}}><SectionTitle text="ویرایش اطلاعات کاربری"/><input value={f.name} onChange={e=>setF({...f,name:e.target.value})} placeholder="نام و نام خانوادگی" style={st.input}/><input value={f.phone} onChange={e=>setF({...f,phone:e.target.value})} placeholder="شماره موبایل" style={st.input}/><input value={f.email} onChange={e=>setF({...f,email:e.target.value})} placeholder="ایمیل" style={st.input}/><button onClick={()=>{setSettings(x=>({...x,profile:{...(x.profile||{}),...f}}));alert("اطلاعات کاربری ذخیره شد.")}} style={st.primaryBtn}>ذخیره اطلاعات</button></div></div>; }
function BackupView({ exportBackup, importBackup }) { const st=useStyles(); const ref=useRef(null); return <div><div style={{...st.card,padding:16}}><SectionTitle text="پشتیبان‌گیری و بازیابی"/><div style={{fontSize:12,color:useT().sub,lineHeight:1.9,marginBottom:12}}>از اطلاعات برنامه فایل پشتیبان بگیر و در صورت نیاز آن را بازیابی کن.</div><button onClick={exportBackup} style={{...st.primaryBtn,background:BRAND.header,marginBottom:10}}>دریافت فایل پشتیبان</button><button onClick={()=>ref.current?.click()} style={{...st.primaryBtn,background:BRAND.gold,color:"#241a30"}}>بازیابی از فایل پشتیبان</button><input ref={ref} type="file" accept="application/json,.json" hidden onChange={e=>{const f=e.target.files?.[0];if(f)importBackup(f);e.target.value=""}}/></div></div>; }
function AccessView({ settings, setSettings }) { const st=useStyles(); return <div><div style={{...st.card,padding:16}}><SectionTitle text="مدیریت دسترسی"/><Row title="قفل با اثر انگشت / Face ID" subtitle="احراز هویت واقعی گوشی در APK" leftIcon={<Fingerprint size={17}/>} leftColor={BRAND.violet} extra={<input type="checkbox" checked={!!settings.biometricEnabled} onChange={e=>setSettings(p=>({...p,biometricEnabled:e.target.checked}))}/>} chevron={null}/><Row title="دسترسی پیامک بانکی" subtitle="مجوز خواندن پیامک‌های بانکی در Android" leftIcon={<BellRing size={17}/>} leftColor={BRAND.header} chevron={null}/></div></div>; }
function BasicSettingsView({ settings, setSettings }) { const st=useStyles(); return <div><div style={{...st.card,padding:16}}><SectionTitle text="تنظیمات پایه"/><label style={st.label}>واحد پول اصلی</label><select value={settings.currency||"IRR"} onChange={e=>setSettings(p=>({...p,currency:e.target.value}))} style={st.input}><option value="IRR">ریال</option><option value="TOMAN">تومان</option><option value="USD">دلار</option></select><label style={st.label}>اندازه نوشته‌ها</label><select value={settings.fontScale||1} onChange={e=>setSettings(p=>({...p,fontScale:Number(e.target.value)}))} style={st.input}><option value="0.9">کوچک</option><option value="1">استاندارد</option><option value="1.1">بزرگ</option><option value="1.2">خیلی بزرگ</option></select></div></div>; }
function TutorialView() { const st=useStyles(); const a=["دکمه + برای هزینه، درآمد، چک و انتقال است.","میکروفون و ثبت سریع برای ثبت طبیعی هزینه و درآمد هستند.","منوی سه‌خط دسترسی به حساب‌ها، اشخاص، بودجه، وام، چک و گزارش‌ها را دارد.","برای حفظ اطلاعات مرتب پشتیبان بگیر.","اثر انگشت را از مدیریت دسترسی فعال کن."]; return <div>{a.map((x,i)=><div key={i} style={{...st.card,padding:14,marginBottom:10,display:"flex",gap:10}}><b>{i+1}</b><span style={{fontSize:13,lineHeight:1.9}}>{x}</span></div>)}</div>; }
function ShareAppView() { const st=useStyles(); return <div><div style={{...st.card,padding:20,textAlign:"center"}}><RexaLogo size={70}/><div style={{fontWeight:800,margin:"10px 0"}}>Rexa</div><button onClick={async()=>{try{if(navigator.share)await navigator.share({title:"Rexa",text:"Rexa؛ مدیریت مالی شخصی"});else{await navigator.clipboard.writeText("Rexa؛ مدیریت مالی شخصی");alert("متن معرفی کپی شد.")}}catch{}}} style={{...st.primaryBtn,background:BRAND.violet}}>ارسال برنامه به دیگران</button></div></div>; }
function RateAppView() { const st=useStyles(); const [r,setR]=useState(0); return <div><div style={{...st.card,padding:20,textAlign:"center"}}><div style={{fontWeight:800}}>امتیاز به برنامه</div><div style={{display:"flex",justifyContent:"center",gap:6,margin:"15px 0"}}>{[1,2,3,4,5].map(n=><button key={n} onClick={()=>setR(n)} style={{border:0,background:"none",fontSize:30,opacity:n<=r?1:.3}}>★</button>)}</div><button disabled={!r} onClick={()=>alert("امتیاز شما ثبت شد.")} style={{...st.primaryBtn,opacity:r?1:.5}}>ثبت امتیاز</button></div></div>; }
function AboutView() { const st=useStyles(); return <div><div style={{...st.card,padding:20,textAlign:"center"}}><RexaLogo size={78}/><div style={{fontWeight:900,fontSize:19,marginTop:8}}>Rexa</div><div style={{fontSize:12,color:useT().sub,marginTop:5}}>مدیریت مالی شخصی</div><div style={{borderTop:"1px solid #eee",margin:"18px 0",paddingTop:14,fontSize:12.5,lineHeight:2}}>نسخه ۵.۱.۲<br/>Android / Capacitor<br/>Rexa Personal Finance</div></div></div>; }

/* ---------------------------------------------------------
   Side menu
--------------------------------------------------------- */
function SideMenu({ onClose, setSubView, profileName }) {
  const sections=[
    {title:"حسابداری و مدیریت",items:[["اهداف مالی و پس‌انداز","goals"],["اشخاص و طرف حساب‌ها","persons"],["بدهکاران و بستانکاران","debts"],["حساب‌ها و کارت‌ها","accounts"],["دسته‌بندی‌ها و برچسب‌ها","categories"],["بودجه‌بندی","budgets"],["تراکنش‌های تکرارشونده","recurring"],["چک‌ها","checks"],["وام و اقساط","loans"],["دارایی‌ها","assets"],["واحدهای پولی","currencies"],["یادآوری قبض‌ها","bills"],["اعضای منزل، رویداد و پروژه","tags"],["دوره مالی","periods"],["تقویم شمسی","calendar"]]},
    {title:"ابزارها",items:[["میانبرهای تراکنش","shortcuts"],["پیامک بانکی","sms"],["ماشین حساب","calculator"],["پشتیبانی و تیکت","support"]]},
    {title:"تنظیمات و برنامه",items:[["ویرایش اطلاعات کاربری","profile"],["پشتیبان‌گیری","backup"],["مدیریت دسترسی","access"],["تنظیمات پایه","basic"],["تنظیمات و امنیت","settings"],["آموزش","tutorial"],["ارسال برنامه به دیگران","share"],["امتیاز به برنامه","rate"],["درباره","about"]]}
  ];
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.42)",zIndex:300,maxWidth:480,margin:"0 auto"}} onClick={onClose}><div onClick={e=>e.stopPropagation()} style={{position:"absolute",top:0,bottom:0,left:0,width:"86%",maxWidth:390,background:"#fff",boxShadow:"3px 0 18px rgba(0,0,0,.22)",overflowY:"auto",paddingTop:"env(safe-area-inset-top,0px)",paddingBottom:"calc(28px + env(safe-area-inset-bottom,0px))"}}><div style={{background:BRAND.header,color:"#fff",padding:"16px",position:"sticky",top:0,zIndex:2}}><div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}><div style={{display:"flex",alignItems:"center",gap:10}}><RexaLogo size={42}/><div><div style={{fontWeight:800,fontSize:15}}>{profileName||"کاربر Rexa"}</div><div style={{fontSize:11,color:"#d8c9e8",marginTop:3}}>Rexa · نسخه ۵.۱.۲</div></div></div><button onClick={onClose} style={{background:"rgba(255,255,255,.12)",border:0,borderRadius:9,color:"#fff",width:36,height:36}}><X size={21}/></button></div></div><div style={{padding:"8px 14px 0"}}>{sections.map(sec=><div key={sec.title}><div style={{fontSize:11,fontWeight:800,color:BRAND.violet,padding:"13px 6px 7px"}}>{sec.title}</div><div style={{border:"1px solid #eeeaf2",borderRadius:12,overflow:"hidden",marginBottom:6}}>{sec.items.map(([label,key],i)=><div key={key} onClick={()=>setSubView(key)} style={{padding:"12px 10px",borderBottom:i===sec.items.length-1?"none":"1px solid #f0eef3",fontSize:13.5,fontWeight:600,color:"#241a30",background:"#fff"}}>{label}</div>)}</div></div>)}<div style={{textAlign:"center",color:"#918899",fontSize:10.5,padding:"14px 0 10px"}}>Rexa Personal Finance · نسخه ۵.۱.۲ · Android</div></div></div></div>;
}
