import express from "express";
import path from "path";
import * as XLSX from "xlsx";
import { createServer as createViteServer } from "vite";

// ============================================================
// CONFIGURATION & GOOGLE SHEETS REGISTRY
// ============================================================
const ALL_MONTHS = [
  "জানুয়ারি", "ফেব্রুয়ারি", "মার্চ", "এপ্রিল", "মে", "জুন",
  "জুলাই", "আগস্ট", "সেপ্টেম্বর", "অক্টোবর", "নভেম্বর", "ডিসেম্বর"
];

const NET_BILL_LINKS: Record<string, string> = {
  "এপ্রিল": "https://docs.google.com/spreadsheets/d/1JzUL1DOA_ubXONu9PhpCfo9oKeFj__o4oCGeeksV-pA/export?format=xlsx",
  "মে": "https://docs.google.com/spreadsheets/d/1nSM8fZUwHxUSZQc88u8Oaquch21226fYaC9n85NV3G0/export?format=xlsx",
  "জুন": "https://docs.google.com/spreadsheets/d/1cPxErMlxd24-YJPN6mBmA-3xPUjG_Ql3lMXEL8dFPzY/export?format=xlsx",
  "জুলাই": "https://docs.google.com/spreadsheets/d/1LWGmlA0yJwh6v1hQhpx4RrFeFBb_zW51D3GPehme-7s/export?format=xlsx",
  "আগস্ট": "https://docs.google.com/spreadsheets/d/1gZoIiCAYhUZW5nSQCUb6OoNDiNY-087JafIQGr0JaZA/export?format=xlsx",
  "সেপ্টেম্বর": "https://docs.google.com/spreadsheets/d/13OcS4yagT4bQNlrY-NuG592QVw9-JxIqntDFVO1zNt8/export?format=xlsx",
  "অক্টোবর": "https://docs.google.com/spreadsheets/d/1BeTi5KkxDB9wLM7rclo131xLnMadt4BNU1hO-AkcAHw/export?format=xlsx",
  "নভেম্বর": "https://docs.google.com/spreadsheets/d/1tU_5ybOZR5Qk2ypcmTH0ET6f3mUxOtwSANYIrFgm1Ns/export?format=xlsx"
};

const DIS_BILL_LINKS: Record<string, string> = {
  "এপ্রিল": "https://docs.google.com/spreadsheets/d/1FjkMU78ge41qPdHa40byxscAo7zJHMj-uB6M1jKGtQM/export?format=xlsx",
  "মে": "https://docs.google.com/spreadsheets/d/1ai1RicItkQIl4_hpOhpP-olKc5FZLLrwLSKbYVFaS-M/export?format=xlsx",
  "জুন": "https://docs.google.com/spreadsheets/d/186olqasPT2CV0Feqs4OUOsdkbWSsA6is6E-faFNVqxI/export?format=xlsx",
  "জুলাই": "https://docs.google.com/spreadsheets/d/1Kzz2ac6SbeJgUyanm8YRVHKmMD3Cux-PfOSIsvurzAk/export?format=xlsx",
  "আগস্ট": "https://docs.google.com/spreadsheets/d/1h8v6TiQJeberx9R6VGTcQu7t6LiudOKrBS9WVqA2cZ0/export?format=xlsx",
  "সেপ্টেম্বর": "https://docs.google.com/spreadsheets/d/1XntdO25UdT-40zJcPSONY_pVYborJiQ4YCBYjMM_8uE/export?format=xlsx",
  "অক্টোবর": "https://docs.google.com/spreadsheets/d/1EFOy9pKsVnB7Nfkdn9bkUNtClE1mn3Cw16ycg1__f4I/export?format=xlsx"
};

const NET_CONN_LINKS: Record<string, string> = {
  "মে": "https://docs.google.com/spreadsheets/d/16bT82iDAlejCN86k2PXGzuDdojeLLKgICFVKybNIWJc/export?format=xlsx",
  "জুন": "https://docs.google.com/spreadsheets/d/1MzoMhmIY3PGwIzzJ3Dif7FyNs_sHDXCYAu9oS7SKMtk/export?format=xlsx",
  "জুলাই": "https://docs.google.com/spreadsheets/d/1eWAu9sMK8vrDlV3Ue35jYY6rfNVM0EOE43ecdZtY0yc/export?format=xlsx",
  "আগস্ট": "https://docs.google.com/spreadsheets/d/144Km2kNlKpSdyQh9O1Q7BaIDf7bwOTGubTazHWMGPL0/export?format=xlsx",
  "সেপ্টেম্বর": "https://docs.google.com/spreadsheets/d/1xL-oLgzRaGB-t4xd_AiVyGE42VkovVHYVr3gwa_sws0/export?format=xlsx",
  "অক্টোবর": "https://docs.google.com/spreadsheets/d/1j4KCcQpVO2jsMWPgEaxilQgxXELIv7FDlwUDoa-8UH4/export?format=xlsx",
  "নভেম্বর": "https://docs.google.com/spreadsheets/d/1qZk86u3nFnMv3LPs8XTRADCkVyxlUB7gbP_ik08TDdo/export?format=xlsx"
};

const DUE_BILL_LINKS: Record<string, string> = {
  "জুলাই": "https://docs.google.com/spreadsheets/d/1AWUhQdWBluPN3kGhG0jsLo224Qc0hst7DvM0PYwFVak/export?format=xlsx",
  "আগস্ট": "https://docs.google.com/spreadsheets/d/14uKbaHXq_Apreon8sKQVTbqSmQmvIn69Sh7idTqA0yI/export?format=xlsx",
  "সেপ্টেম্বর": "https://docs.google.com/spreadsheets/d/1GJKWaWknUISnk0g_yXq9WoeJo86Pni6U6biZr7zWwaU/export?format=xlsx",
  "অক্টোবর": "https://docs.google.com/spreadsheets/d/1rGLnOjC51iftP-Y9nIk5E3iv_4K0rT1WTy5yCbd-Dj4/export?format=xlsx"
};

const AREAS: Record<string, string[]> = {
  area1: [
    "চর শিমুল চূড়া", "সরকার বাড়ী", "বলিদে পাড়া", "নবীন নগর", "বালুঘাট",
    "কেরানী বাড়ী", "আমেদ বাড়ী", "কৃষ্ণচূড়া", "আকন্দ পাড়া", "চাংপাড়া",
    "কান্দা পাড়া", "কান্দাপাড়া", "ব্যাংতলা", "বাকশাবাড়ী", "দহের পাড়",
    "ইজারা পাড়া", "দষ্টিপাড়া", "ভেলুয়া", "লক্ষ্মিডাংরী", "লক্ষীডাংরী",
    "চকবন্দী", "কক্সবাজার"
  ],
  area2: [
    "উত্তর পাড়া", "দক্ষিণ পাড়া", "টানা ব্রীজ বাজার", "দত্তপাড়া",
    "ফুলকার চর", "কড়ইতলা", "মোল্লা বাজার", "পেচার চর", "চন্দনপুর"
  ]
};

// Aliases mapping for friendly API requests (accepts English or Bengali)
function normalizeNetType(input?: string): string {
  if (!input) return "নেট বিল";
  const s = input.toLowerCase().trim();
  if (s.includes("নেট বিল") || s === "net" || s === "net-bill" || s === "internet" || s === "bill") return "নেট বিল";
  if (s.includes("ডিস") || s === "dish" || s === "cable" || s === "tv") return "ডিস বিল";
  if (s.includes("সংযোগ") || s.includes("conn") || s === "connection") return "নেট সংযোগ";
  if (s.includes("বকেয়া") || s.includes("due") || s === "arrear") return "বকেয়া বিল";
  return input;
}

function normalizeMonth(input?: string): string {
  if (!input) return "নভেম্বর";
  const s = input.toLowerCase().trim();
  if (s.includes("নভেম্বর") || s.startsWith("nov")) return "নভেম্বর";
  if (s.includes("অক্টোবর") || s.startsWith("oct")) return "অক্টোবর";
  if (s.includes("সেপ্টেম্বর") || s.startsWith("sep")) return "সেপ্টেম্বর";
  if (s.includes("আগস্ট") || s.startsWith("aug")) return "আগস্ট";
  if (s.includes("জুলাই") || s.startsWith("jul")) return "জুলাই";
  if (s.includes("জুন") || s.startsWith("jun")) return "জুন";
  if (s.includes("মে") || s === "may") return "মে";
  if (s.includes("এপ্রিল") || s.startsWith("apr")) return "এপ্রিল";
  if (s.includes("মার্চ") || s.startsWith("mar")) return "মার্চ";
  if (s.includes("ফেব্রুয়ারি") || s.startsWith("feb")) return "ফেব্রুয়ারি";
  if (s.includes("জানুয়ারি") || s.startsWith("jan")) return "জানুয়ারি";
  if (s.includes("ডিসেম্বর") || s.startsWith("dec")) return "ডিসেম্বর";
  return input;
}

function getActiveSheetLink(netType: string, month: string): string {
  if (netType === "নেট বিল") return NET_BILL_LINKS[month] || "";
  if (netType === "ডিস বিল") return DIS_BILL_LINKS[month] || "";
  if (netType === "নেট সংযোগ") return NET_CONN_LINKS[month] || "";
  if (netType === "বকেয়া বিল") return DUE_BILL_LINKS[month] || "";
  return "";
}

// ============================================================
// HELPER FUNCTIONS & BENGLI PHONETIC NORMALIZATION
// ============================================================
function toEnglishDigits(str: any): string {
  if (str === null || str === undefined) return "";
  const bnDigits = ["০", "১", "২", "৩", "৪", "৫", "৬", "৭", "৮", "৯"];
  let res = String(str);
  for (let i = 0; i < 10; i++) {
    res = res.split(bnDigits[i]).join(String(i));
  }
  return res;
}

function parseAmount(val: any): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const cleaned = toEnglishDigits(String(val)).replace(/[^0-9.-]/g, "");
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function formatPhone(s: any): string {
  if (!s || s === "---" || s === "-") return "---";
  let clean = toEnglishDigits(String(s)).replace(/\D/g, "");
  if (clean.startsWith("880") && clean.length === 13) clean = "0" + clean.slice(3);
  else if (clean.length === 11 && clean.startsWith("01")) return clean;
  else if (clean.length === 10 && !clean.startsWith("0")) clean = "0" + clean;
  return clean.length >= 10 ? clean : String(s);
}

function isLikelyPhoneNumber(val: any): boolean {
  if (!val) return false;
  const str = toEnglishDigits(String(val)).trim().replace(/[\s\-\(\)\.]/g, "");
  return (
    (str.startsWith("01") && str.length >= 10 && str.length <= 13) ||
    (str.startsWith("8801") && str.length >= 13) ||
    str.startsWith("+8801")
  );
}

function isLikelyDateString(val: any): boolean {
  if (!val) return false;
  if (typeof val === "number" && val > 30000 && val < 60000) return true;
  const str = toEnglishDigits(String(val)).trim();
  return /^(\d{1,4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,4})/.test(str);
}

function parseExcelDate(val: any): string {
  if (!val || val === "-" || val === "---") return "-";
  if (typeof val === "number" && val > 1) {
    const d = new Date((val - 25569) * 86400000);
    if (!isNaN(d.getTime())) {
      return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
    }
  }
  if (typeof val === "string") {
    const eng = toEnglishDigits(val.trim());
    const isoMatch = eng.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (isoMatch) return `${isoMatch[3].padStart(2, "0")}/${isoMatch[2].padStart(2, "0")}/${isoMatch[1]}`;
    const dmyMatch = eng.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
    if (dmyMatch) {
      const yr = dmyMatch[3].length === 2 ? `20${dmyMatch[3]}` : dmyMatch[3];
      return `${dmyMatch[1].padStart(2, "0")}/${dmyMatch[2].padStart(2, "0")}/${yr}`;
    }
    return eng;
  }
  return "-";
}

function normalizeBengali(str: any): string {
  if (!str) return "";
  return String(str)
    .replace(/[\s\u0981-\u0983\u09BC\u09BE-\u09CD\u09D7\u09E2\u09E3]/g, "")
    .replace(/[অআ]/g, "অ")
    .replace(/[ইঈএ]/g, "ই")
    .replace(/[উঊও]/g, "উ")
    .replace(/[শষস]/g, "স")
    .replace(/[ণন]/g, "ন")
    .replace(/[ড়ঢ়র]/g, "র")
    .replace(/[যজ]/g, "জ")
    .replace(/[তৎ]/g, "ত")
    .replace(/[বভ]/g, "ব")
    .toLowerCase()
    .trim();
}

// ============================================================
// WORKBOOK PARSER
// ============================================================
function parseWorkbook(wb: any, netType: string) {
  if (!wb || !wb.SheetNames || wb.SheetNames.length === 0) return [];
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows || rows.length < 2) return [];

  let headerRowIdx = -1;
  let colId = -1;
  let colName = -1;
  let colPay = -1;
  let colDue = -1;
  let colTotal = -1;
  let colDate = -1;
  let colMob = -1;
  let colAddr = -1;
  let colStatus = -1;

  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;
    const rowStr = row.map((c) => String(c || "").trim()).join(" ");
    if (
      rowStr.includes("ক্রমিক") &&
      (rowStr.includes("নাম") || rowStr.includes("ইউজার") || rowStr.includes("আইডি"))
    ) {
      headerRowIdx = r;
      row.forEach((cell, cIdx) => {
        const txt = String(cell || "").trim();
        const lower = txt.toLowerCase();
        if (txt.includes("ইউজার") || txt.includes("আইডি") || lower.includes("user") || lower === "id") {
          colId = cIdx;
        } else if (
          txt.includes("গ্রাহকের নাম") ||
          (txt.includes("নাম") &&
            !txt.includes("বিলের") &&
            !txt.includes("মাসের") &&
            !txt.includes("নাম্বার") &&
            !txt.includes("লিখুন"))
        ) {
          colName = cIdx;
        } else if (
          txt.includes("পে-বিল") ||
          txt.includes("পে বিল") ||
          txt.includes("পরিশোধ") ||
          txt.includes("আদায়") ||
          lower === "pay"
        ) {
          colPay = cIdx;
        } else if (txt.includes("বকেয়া বিল") || txt.includes("বকেয়া") || lower === "due") {
          colDue = cIdx;
        } else if (
          txt.includes("সর্বমোট") ||
          txt.includes("সর্বমোট") ||
          txt.includes("মোট ব্যালেন্স") ||
          lower === "total"
        ) {
          colTotal = cIdx;
        } else if (txt.includes("তারিখ") || lower.includes("date")) {
          colDate = cIdx;
        } else if (
          txt.includes("মোবাইল") ||
          txt.includes("মোবাইল") ||
          txt.includes("ফোন") ||
          lower.includes("phone") ||
          lower.includes("mob")
        ) {
          colMob = cIdx;
        } else if (
          txt.includes("ঠিকানা") ||
          txt.includes("এরিয়া") ||
          txt.includes("এলাকা") ||
          lower.includes("address")
        ) {
          colAddr = cIdx;
        } else if (txt.includes("স্ট্যাটাস") || txt.includes("মন্তব্য") || lower.includes("status")) {
          colStatus = cIdx;
        }
      });
      break;
    }
  }

  const isDis = netType === "ডিস বিল";
  if (colName === -1) {
    colId = isDis ? -1 : 1;
    colName = isDis ? 1 : 2;
    colPay = isDis ? 2 : 3;
    colDue = isDis ? 3 : 4;
    colTotal = isDis ? 4 : 5;
    colDate = isDis ? 5 : 6;
    colMob = isDis ? 6 : 7;
    colAddr = isDis ? 7 : 8;
    colStatus = (netType === "নেট সংযোগ" || netType === "বকেয়া বিল") ? 9 : -1;
  }

  const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 5;
  const records: any[] = [];
  const seen = new Set<string>();

  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 2) continue;

    const firstCell = row[0] ? String(row[0]).trim() : "";
    if (firstCell.includes("মোট") || firstCell === "0" || firstCell === "০") continue;

    const uid = colId !== -1 && row[colId] !== undefined ? String(row[colId]).trim() : "";
    const name = colName !== -1 && row[colName] !== undefined ? String(row[colName]).trim() : "";

    if (!name && !uid) continue;
    if (name === "0" || name === "০") continue;
    if (name.includes("মোট") || name.toLowerCase().includes("total")) continue;
    if (name.includes("গ্রাহকের নাম") || name.includes("ইউজার")) continue;

    const unique = `${uid}_${name}_${i}`;
    if (seen.has(unique)) continue;
    seen.add(unique);

    const pay = colPay !== -1 ? parseAmount(row[colPay]) : 0;
    const due = colDue !== -1 ? parseAmount(row[colDue]) : 0;
    const sheetTotal = colTotal !== -1 ? parseAmount(row[colTotal]) : 0;
    const total = sheetTotal > 0 ? sheetTotal : pay + due;

    let rawDate = colDate !== -1 ? row[colDate] : "";
    let rawMob = colMob !== -1 ? row[colMob] : "";
    const rawAddr = colAddr !== -1 && row[colAddr] !== undefined ? String(row[colAddr]).trim() : "---";
    const rawStatus = colStatus !== -1 && row[colStatus] !== undefined ? String(row[colStatus]).trim() : "";

    if (isLikelyPhoneNumber(rawDate) && !isLikelyPhoneNumber(rawMob)) {
      const temp = rawDate;
      rawDate = rawMob;
      rawMob = temp;
    } else if (isLikelyDateString(rawMob) && !isLikelyDateString(rawDate)) {
      const temp = rawDate;
      rawDate = rawMob;
      rawMob = temp;
    }

    let finalStatus = rawStatus;
    if (!finalStatus) {
      if (due === 0 && pay > 0) finalStatus = "পরিশোধিত";
      else if (due > 0 && pay > 0) finalStatus = "আংশিক";
      else if (due > 0 && pay === 0) finalStatus = "বকেয়া";
      else finalStatus = "পরিশোধিত";
    }

    records.push({
      sl: records.length + 1,
      id: colId !== -1 && uid ? uid : (isDis ? null : "-----"),
      name: name || "নাম নেই",
      pay: pay,
      due: due,
      total: total,
      date: parseExcelDate(rawDate),
      mob: formatPhone(rawMob),
      address: rawAddr || "---",
      status: finalStatus
    });
  }

  return records;
}

// In-memory cache for sheets (TTL: 60 seconds)
const sheetCache = new Map<string, { timestamp: number; records: any[] }>();
const CACHE_TTL_MS = 60 * 1000;

async function getRecordsForSheet(netType: string, month: string): Promise<any[]> {
  const url = getActiveSheetLink(netType, month);
  if (!url) return [];

  const cacheKey = `${netType}_${month}`;
  const cached = sheetCache.get(cacheKey);
  if (cached && Date.now() - cached.timestamp < CACHE_TTL_MS) {
    return cached.records;
  }

  const res = await fetch(url, {
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    },
    redirect: "follow",
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch Google Sheet: ${res.statusText}`);
  }

  const arrayBuffer = await res.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const records = parseWorkbook(wb, netType);
  sheetCache.set(cacheKey, { timestamp: Date.now(), records });
  return records;
}

// ============================================================
// EXPRESS SERVER SETUP
// ============================================================
async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  app.use(express.json({ limit: "20mb" }));

  // Global CORS Middleware - allows any external website/app to call these APIs!
  app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader(
      "Access-Control-Allow-Headers",
      "Origin, X-Requested-With, Content-Type, Accept, Authorization"
    );
    if (req.method === "OPTIONS") {
      return res.sendStatus(204);
    }
    next();
  });

  // ------------------------------------------------------------
  // API Route: Documentation & Overview
  // ------------------------------------------------------------
  app.get(["/api", "/api/docs"], (req, res) => {
    const host = req.get("host") || `localhost:${PORT}`;
    const protocol = req.protocol === "https" || req.headers["x-forwarded-proto"] === "https" ? "https" : "http";
    const baseUrl = `${protocol}://${host}`;

    res.json({
      name: "বিল ম্যানেজার প্রো - REST API (Bill Manager Pro API)",
      description: "যেকোনো ওয়েবসাইট বা মোবাইল অ্যাপ থেকে বিল ও গ্রাহকের তথ্য সহজে পাওয়ার উন্মুক্ত API।",
      baseUrl: baseUrl,
      endpoints: [
        {
          route: "/api/external/summary",
          method: "GET",
          description: "অন্য ওয়েবসাইট বা ক্যালকুলেটরের জন্য সর্বমোট পরিমাণ ও বকেয়া টাকার লাইভ সামারি API।",
          queryParams: {
            month: "মাসের নাম ('নভেম্বর', 'অক্টোবর', 'সেপ্টেম্বর' ইত্যাদি অথবা 'all')",
            netType: "সার্ভিসের নাম ('নেট বিল', 'ডিস বিল', 'নেট সংযোগ', 'বকেয়া বিল')"
          },
          example: `${baseUrl}/api/external/summary?month=নভেম্বর`
        },
        {
          route: "/api/totals",
          method: "GET",
          description: "কাস্টমার লিস্ট ছাড়া শুধুমাত্র সর্বমোট টাকা, আদায় ও বকেয়া টাকার হিসাব পেতে।",
          queryParams: {
            netType: "সার্ভিসের নাম ('নেট বিল', 'ডিস বিল', 'নেট সংযোগ', 'বকেয়া বিল' অথবা alias 'net', 'dish', 'net-conn', 'due')",
            month: "মাসের নাম ('নভেম্বর', 'অক্টোবর', 'সেপ্টেম্বর' ইত্যাদি অথবা 'nov', 'oct')",
            area: "এরিয়া ফিল্টার: 'area1', 'area2' অথবা নির্দিষ্ট এলাকা (ঐচ্ছিক)",
            status: "স্ট্যাটাস ফিল্টার (ঐচ্ছিক)"
          },
          example: `${baseUrl}/api/totals?netType=নেট বিল&month=নভেম্বর`
        },
        {
          route: "/api/bills",
          method: "GET",
          description: "নির্দিষ্ট সার্ভিস ও মাসের বিলের তালিকা এবং সামারি সংগ্রহ করুন।",
          queryParams: {
            netType: "সার্ভিসের নাম ('নেট বিল', 'ডিস বিল', 'নেট সংযোগ', 'বকেয়া বিল' অথবা alias 'net', 'dish', 'net-conn', 'due')",
            month: "মাসের নাম ('নভেম্বর', 'অক্টোবর', 'সেপ্টেম্বর' ইত্যাদি অথবা 'nov', 'oct')",
            status: "ফিল্টার: 'সকল', 'পরিশোধিত', 'বকেয়া', 'আংশিক', 'ফ্রি', 'বাকী'",
            area: "এরিয়া ফিল্টার: 'area1', 'area2' অথবা নির্দিষ্ট এলাকার নাম (যেমন: 'চর শিমুল চূড়া', 'কক্সবাজার')",
            search: "অনুসন্ধান: গ্রাহকের নাম, মোবাইল, আইডি বা ঠিকানা দিয়ে খুঁজুন (বানান ভুলেরও নির্ভুল রেজাল্ট)",
            page: "পৃষ্ঠা নম্বর (ঐচ্ছিক)",
            limit: "প্রতি পৃষ্ঠার গ্রাহক সংখ্যা (ঐচ্ছিক)"
          },
          example: `${baseUrl}/api/bills?netType=নেট বিল&month=নভেম্বর`
        },
        {
          route: "/api/search",
          method: "GET",
          description: "নাম, ফোন, আইডি বা ঠিকানা দিয়ে তাৎক্ষণিক গ্রাহক তথ্য খুঁজুন।",
          queryParams: {
            q: "সার্চ কিওয়ার্ড (যেমন: 'লাহোর' বা '017' বা 'কক্সবাজার')",
            netType: "সার্ভিস (ডিফল্ট: 'নেট বিল')",
            month: "মাস (ডিফল্ট: 'নভেম্বর')"
          },
          example: `${baseUrl}/api/search?q=লাহোর&netType=নেট বিল&month=নভেম্বর`
        },
        {
          route: "/api/sheets",
          method: "GET",
          description: "সিস্টেমে সংযুক্ত সকল গুগল শিট এবং মাসের তালিকা।",
          example: `${baseUrl}/api/sheets`
        },
        {
          route: "/api/summary",
          method: "GET",
          description: "মাসভিত্তিক মোট বিল আদায় ও গ্রাহক পরিসংখ্যান (চার্ট বা ড্যাশবোর্ডের জন্য)।",
          example: `${baseUrl}/api/summary`
        },
        {
          route: "/api/areas",
          method: "GET",
          description: "এরিয়া ১ এবং এরিয়া ২ এর আওতাভুক্ত এলাকা সমূহের তালিকা।",
          example: `${baseUrl}/api/areas`
        },
        {
          route: "/api/duplicates",
          method: "GET",
          description: "একই আইডির বা একাধিকবার এন্ট্রি হওয়া ডুপ্লিকেট গ্রাহকদের তালিকা।",
          example: `${baseUrl}/api/duplicates?netType=নেট বিল&month=নভেম্বর`
        },
        {
          route: "/api/fetch-sheet",
          method: "GET",
          description: "সরাসরি গুগল শিট XLSX ফাইল ডাউনলোড প্রক্সি।",
          queryParams: {
            url: "গুগল শিটের এক্সপোর্ট লিঙ্ক"
          }
        }
      ]
    });
  });

  // ------------------------------------------------------------
  // API Route: Health Check
  // ------------------------------------------------------------
  app.get("/api/health", (_req, res) => {
    res.json({ status: "ok", timestamp: new Date().toISOString() });
  });

  // ------------------------------------------------------------
  // API Route: Sheets Registry
  // ------------------------------------------------------------
  app.get("/api/sheets", (_req, res) => {
    res.json({
      success: true,
      services: {
        "নেট বিল": {
          availableMonths: Object.keys(NET_BILL_LINKS),
          sheets: NET_BILL_LINKS
        },
        "ডিস বিল": {
          availableMonths: Object.keys(DIS_BILL_LINKS),
          sheets: DIS_BILL_LINKS
        },
        "নেট সংযোগ": {
          availableMonths: Object.keys(NET_CONN_LINKS),
          sheets: NET_CONN_LINKS
        },
        "বকেয়া বিল": {
          availableMonths: Object.keys(DUE_BILL_LINKS),
          sheets: DUE_BILL_LINKS
        }
      }
    });
  });

  // ------------------------------------------------------------
  // API Route: Areas List
  // ------------------------------------------------------------
  app.get("/api/areas", (_req, res) => {
    res.json({
      success: true,
      areas: AREAS
    });
  });

  // ------------------------------------------------------------
  // API Route: Bills List (with filtering, fuzzy search, pagination)
  // ------------------------------------------------------------
  app.get("/api/bills", async (req, res) => {
    try {
      const netType = normalizeNetType(req.query.netType as string || req.query.service as string);
      const month = normalizeMonth(req.query.month as string);
      const sheetUrl = getActiveSheetLink(netType, month);

      if (!sheetUrl) {
        return res.status(404).json({
          success: false,
          error: `No sheet linked for service '${netType}' and month '${month}'`
        });
      }

      const allRecords = await getRecordsForSheet(netType, month);

      // Filter by Status
      let filtered = allRecords;
      const statusParam = (req.query.status as string || "").trim();
      if (statusParam && statusParam !== "সকল" && statusParam !== "all") {
        filtered = filtered.filter((r) => r.status === statusParam);
      }

      // Filter by Area
      const areaParam = (req.query.area as string || "").trim();
      if (areaParam && areaParam !== "all" && areaParam !== "সকল") {
        let areaVillages: string[] = [];
        if (areaParam === "area1" || areaParam === "এরিয়া ১") {
          areaVillages = AREAS.area1;
        } else if (areaParam === "area2" || areaParam === "এরিয়া ২") {
          areaVillages = AREAS.area2;
        } else {
          areaVillages = [areaParam];
        }

        const normalizedVillages = areaVillages.map((v) => normalizeBengali(v));
        filtered = filtered.filter((r) => {
          const normAddr = normalizeBengali(r.address);
          return normalizedVillages.some((v) => normAddr.includes(v));
        });
      }

      // Filter by Search Query
      const searchParam = (req.query.search as string || req.query.q as string || "").trim();
      if (searchParam) {
        const normQ = normalizeBengali(searchParam);
        const engCleanQ = toEnglishDigits(searchParam).toLowerCase().trim();
        filtered = filtered.filter((r) => {
          const idMatch = r.id && String(r.id).toLowerCase().includes(engCleanQ);
          const mobMatch = r.mob && toEnglishDigits(String(r.mob)).includes(engCleanQ);
          const nameMatch = r.name && normalizeBengali(r.name).includes(normQ);
          const addrMatch = r.address && normalizeBengali(r.address).includes(normQ);
          return idMatch || mobMatch || nameMatch || addrMatch;
        });
      }

      // Calculate totals for filtered dataset
      const totalPay = filtered.reduce((sum, r) => sum + (r.pay || 0), 0);
      const totalDue = filtered.reduce((sum, r) => sum + (r.due || 0), 0);
      const grandTotal = filtered.reduce((sum, r) => sum + (r.total || 0), 0);

      // Pagination
      const page = Math.max(1, parseInt(req.query.page as string || "1", 10) || 1);
      const limit = parseInt(req.query.limit as string || "0", 10);
      let paginated = filtered;
      if (limit > 0) {
        const startIndex = (page - 1) * limit;
        paginated = filtered.slice(startIndex, startIndex + limit);
      }

      const summaryOnly = req.query.summaryOnly === "true" || req.query.nodata === "true" || req.query.customers === "false";

      res.setHeader("Cache-Control", "public, max-age=60");
      if (summaryOnly) {
        return res.json({
          success: true,
          service: netType,
          month: month,
          sheetUrl: sheetUrl,
          totalCustomers: filtered.length,
          totalPay: totalPay,
          totalDue: totalDue,
          grandTotal: grandTotal
        });
      }

      res.json({
        success: true,
        service: netType,
        month: month,
        sheetUrl: sheetUrl,
        summary: {
          totalCount: filtered.length,
          totalPay: totalPay,
          totalDue: totalDue,
          grandTotal: grandTotal
        },
        page: limit > 0 ? page : 1,
        limit: limit > 0 ? limit : filtered.length,
        totalPages: limit > 0 ? Math.ceil(filtered.length / limit) : 1,
        data: paginated
      });
    } catch (err: any) {
      console.error("API error in /api/bills:", err);
      res.status(500).json({ success: false, error: err.message || "Failed to process bill records" });
    }
  });

  // ------------------------------------------------------------
  // API Route: External Summary (Requested: /api/external/summary)
  // Designed specifically for external websites & calculators
  // Returns grandTotal, totalDue, totalPay, and month data without customer list
  // ------------------------------------------------------------
  app.get(
    [
      "/api/external/summary",
      "/api/external/totals",
      "/api/totals",
      "/api/balance",
      "/api/stats",
    ],
    async (req, res) => {
      try {
        const netType = normalizeNetType(
          (req.query.netType as string) || (req.query.service as string)
        );
        const rawMonth = ((req.query.month as string) || "").trim();
        const isAllMonths =
          rawMonth === "all" || rawMonth === "সকল" || req.query.all === "true";

        const availableMonths = [
          "এপ্রিল",
          "মে",
          "জুন",
          "জুলাই",
          "আগস্ট",
          "সেপ্টেম্বর",
          "অক্টোবর",
          "নভেম্বর",
        ];

        // Preload all months summary helper
        const getAllMonthsSummary = async () => {
          const monthsResult: Record<string, any> = {};
          for (const m of availableMonths) {
            const sUrl = getActiveSheetLink(netType, m);
            if (!sUrl) continue;
            try {
              const records = await getRecordsForSheet(netType, m);
              const tPay = records.reduce((s, r) => s + (r.pay || 0), 0);
              const tDue = records.reduce((s, r) => s + (r.due || 0), 0);
              const gTotal = records.reduce((s, r) => s + (r.total || 0), 0);
              monthsResult[m] = {
                month: m,
                grandTotal: gTotal,
                total: gTotal,
                totalDue: tDue,
                due: tDue,
                totalPay: tPay,
                pay: tPay,
                netPay: tPay,
                totalCustomers: records.length,
                count: records.length,
              };
            } catch (e) {}
          }
          return monthsResult;
        };

        // If user specifically requested all months:
        if (isAllMonths) {
          const allMonthsData = await getAllMonthsSummary();
          res.setHeader("Cache-Control", "public, max-age=60");
          return res.json({
            success: true,
            service: netType,
            months: allMonthsData,
          });
        }

        // Single month calculation (default to latest: নভেম্বর)
        const month = rawMonth ? normalizeMonth(rawMonth) : "নভেম্বর";
        const sheetUrl = getActiveSheetLink(netType, month);

        if (!sheetUrl) {
          return res.status(404).json({
            success: false,
            error: `No sheet linked for service '${netType}' and month '${month}'`,
          });
        }

        const allRecords = await getRecordsForSheet(netType, month);

        // Filter by Status (optional)
        let filtered = allRecords;
        const statusParam = ((req.query.status as string) || "").trim();
        if (statusParam && statusParam !== "সকল" && statusParam !== "all") {
          filtered = filtered.filter((r) => r.status === statusParam);
        }

        // Filter by Area (optional)
        const areaParam = ((req.query.area as string) || "").trim();
        if (areaParam && areaParam !== "all" && areaParam !== "সকল") {
          let areaVillages: string[] = [];
          if (areaParam === "area1" || areaParam === "এরিয়া ১") {
            areaVillages = AREAS.area1;
          } else if (areaParam === "area2" || areaParam === "এরিয়া ২") {
            areaVillages = AREAS.area2;
          } else {
            areaVillages = [areaParam];
          }

          const normalizedVillages = areaVillages.map((v) =>
            normalizeBengali(v)
          );
          filtered = filtered.filter((r) => {
            const normAddr = normalizeBengali(r.address);
            return normalizedVillages.some((v) => normAddr.includes(v));
          });
        }

        // Calculate totals
        const totalPay = filtered.reduce((sum, r) => sum + (r.pay || 0), 0);
        const totalDue = filtered.reduce((sum, r) => sum + (r.due || 0), 0);
        const grandTotal = filtered.reduce((sum, r) => sum + (r.total || 0), 0);

        // Also get all months so external websites can read any month from the same response
        let allMonthsDict: Record<string, any> | undefined = undefined;
        if (!rawMonth) {
          try {
            allMonthsDict = await getAllMonthsSummary();
          } catch (e) {}
        }

        res.setHeader("Cache-Control", "public, max-age=60");
        res.json({
          success: true,
          service: netType,
          month: month,
          grandTotal: grandTotal,
          total: grandTotal,
          totalBill: grandTotal,
          totalDue: totalDue,
          due: totalDue,
          dueBill: totalDue,
          totalPay: totalPay,
          pay: totalPay,
          netPay: totalPay,
          totalCustomers: filtered.length,
          count: filtered.length,
          ...(allMonthsDict ? { months: allMonthsDict } : {}),
        });
      } catch (err: any) {
        console.error("API error in /api/external/summary:", err);
        res.status(500).json({
          success: false,
          error: err.message || "Failed to calculate external summary",
        });
      }
    }
  );

  // ------------------------------------------------------------
  // API Route: Quick Search
  // ------------------------------------------------------------
  app.get("/api/search", async (req, res) => {
    try {
      const q = (req.query.q as string || req.query.search as string || "").trim();
      const netType = normalizeNetType(req.query.netType as string || req.query.service as string);
      const month = normalizeMonth(req.query.month as string);

      if (!q) {
        return res.status(400).json({ success: false, error: "Search query 'q' parameter is required" });
      }

      const allRecords = await getRecordsForSheet(netType, month);
      const normQ = normalizeBengali(q);
      const engCleanQ = toEnglishDigits(q).toLowerCase().trim();

      const matched = allRecords.filter((r) => {
        const idMatch = r.id && String(r.id).toLowerCase().includes(engCleanQ);
        const mobMatch = r.mob && toEnglishDigits(String(r.mob)).includes(engCleanQ);
        const nameMatch = r.name && normalizeBengali(r.name).includes(normQ);
        const addrMatch = r.address && normalizeBengali(r.address).includes(normQ);
        return idMatch || mobMatch || nameMatch || addrMatch;
      });

      res.json({
        success: true,
        service: netType,
        month: month,
        query: q,
        foundCount: matched.length,
        data: matched
      });
    } catch (err: any) {
      console.error("API error in /api/search:", err);
      res.status(500).json({ success: false, error: err.message || "Search failed" });
    }
  });

  // ------------------------------------------------------------
  // API Route: Monthly Income Summary (for Dashboards & Charts)
  // ------------------------------------------------------------
  app.get("/api/summary", async (_req, res) => {
    try {
      const months = ["এপ্রিল", "মে", "জুন", "জুলাই", "আগস্ট", "সেপ্টেম্বর", "অক্টোবর", "নভেম্বর"];
      const summaryList = [];

      for (const m of months) {
        let netTotal = 0;
        let disTotal = 0;
        let netCount = 0;
        let disCount = 0;

        try {
          if (NET_BILL_LINKS[m]) {
            const records = await getRecordsForSheet("নেট বিল", m);
            netTotal = records.reduce((s, r) => s + (r.pay || 0), 0);
            netCount = records.length;
          }
        } catch (e) {}

        try {
          if (DIS_BILL_LINKS[m]) {
            const records = await getRecordsForSheet("ডিস বিল", m);
            disTotal = records.reduce((s, r) => s + (r.pay || 0), 0);
            disCount = records.length;
          }
        } catch (e) {}

        summaryList.push({
          month: m,
          netBillTotal: netTotal,
          netCustomers: netCount,
          dishBillTotal: disTotal,
          dishCustomers: disCount,
          combinedTotal: netTotal + disTotal
        });
      }

      res.json({
        success: true,
        months: summaryList
      });
    } catch (err: any) {
      res.status(500).json({ success: false, error: err.message });
    }
  });

  // ------------------------------------------------------------
  // API Route: Duplicate Entries Check
  // ------------------------------------------------------------
  app.get("/api/duplicates", async (req, res) => {
    try {
      const netType = normalizeNetType(req.query.netType as string || req.query.service as string);
      const month = normalizeMonth(req.query.month as string);

      const records = await getRecordsForSheet(netType, month);
      const idMap = new Map<string, any[]>();

      records.forEach((r) => {
        const idKey = r.id ? String(r.id).trim().toLowerCase() : "";
        if (idKey && idKey !== "-----") {
          if (!idMap.has(idKey)) idMap.set(idKey, []);
          idMap.get(idKey)!.push(r);
        }
      });

      const duplicateGroups = Array.from(idMap.entries())
        .filter(([_, list]) => list.length > 1)
        .map(([id, list]) => ({
          id,
          count: list.length,
          entries: list
        }));

      res.json({
        success: true,
        service: netType,
        month: month,
        totalDuplicatesFound: duplicateGroups.length,
        duplicateGroups: duplicateGroups
      });
    } catch (err: any) {
      res.status(500).json({ success: false, error: err.message });
    }
  });

  // ------------------------------------------------------------
  // API Route: Google Sheets XLSX Proxy (bypasses browser CORS restriction)
  // ------------------------------------------------------------
  app.get("/api/fetch-sheet", async (req, res) => {
    const targetUrl = req.query.url as string;
    if (!targetUrl) {
      return res.status(400).json({ error: "URL parameter is required" });
    }

    try {
      const response = await fetch(targetUrl, {
        headers: {
          "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        },
        redirect: "follow",
      });

      if (!response.ok) {
        return res
          .status(response.status)
          .json({ error: `Failed to fetch from source: ${response.statusText}` });
      }

      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
      res.setHeader("Pragma", "no-cache");
      res.setHeader("Expires", "0");
      res.send(buffer);
    } catch (err: any) {
      console.error("Error proxying sheet:", err);
      res.status(500).json({ error: err.message || "Internal server error" });
    }
  });

  // Serve public directory
  app.use(express.static(path.resolve(process.cwd(), "public")));

  // Vite middleware in dev, static files in production
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use((req, res, next) => {
      if (req.path === "/" || req.path.endsWith(".html")) {
        res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        res.setHeader("Pragma", "no-cache");
        res.setHeader("Expires", "0");
      }
      next();
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(process.cwd(), "dist");
    app.use(
      express.static(distPath, {
        setHeaders: (res, filePath) => {
          if (filePath.endsWith(".html")) {
            res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            res.setHeader("Pragma", "no-cache");
            res.setHeader("Expires", "0");
          }
        },
      })
    );
    app.get("*", (_req, res) => {
      res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
      res.setHeader("Pragma", "no-cache");
      res.setHeader("Expires", "0");
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running at http://0.0.0.0:${PORT}`);
  });
}

startServer();

