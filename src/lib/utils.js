import * as XLSX from "xlsx";

const MONTH_ABBR = {
  JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6,
  JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12,
};

export function excelSerialToDDMMYYYY(v) {
  if (v == null || v === "") return "";
  if (typeof v === "string") {
    const trimmed = v.trim();
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed)) return trimmed;
    // "08-AUG-2026" style (bank/card statement dates) -> "08/08/2026"
    const m = trimmed.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (m) {
      const mon = MONTH_ABBR[m[2].toUpperCase()];
      if (mon) return `${m[1].padStart(2, "0")}/${String(mon).padStart(2, "0")}/${m[3]}`;
    }
    return trimmed;
  }
  if (v instanceof Date) {
    const d = String(v.getDate()).padStart(2, "0");
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const y = v.getFullYear();
    return `${d}/${m}/${y}`;
  }
  if (typeof v === "number") {
    const parsed = XLSX.SSF.parse_date_code(v);
    if (parsed) {
      const d = String(parsed.d).padStart(2, "0");
      const m = String(parsed.m).padStart(2, "0");
      return `${d}/${m}/${parsed.y}`;
    }
  }
  return String(v);
}

export function toNumber(v) {
  if (v == null) return 0;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (s === "") return 0;
  const n = parseFloat(s.replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}

export function normHeader(h) {
  return String(h ?? "").trim().toLowerCase();
}

export function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename);
}
