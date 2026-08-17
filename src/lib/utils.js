import * as XLSX from "xlsx";

export function excelSerialToDDMMYYYY(v) {
  if (v == null || v === "") return "";
  if (typeof v === "string") {
    const trimmed = v.trim();
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed)) return trimmed;
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
