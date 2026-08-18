import * as XLSX from "xlsx";
import { toNumber, excelSerialToDDMMYYYY } from "./utils.js";
import {
  DIALYSIS_TOKEN,
  DIALYSIS_ALT_TOKEN,
  resolveRows,
  voucherTypeForOccurrence,
} from "./tokens.js";

// Uses the same 11-column layout as RECEIPT (Bill Type + Bill Name + Bill
// Amount are populated so Tally tracks the receivable against a bill reference).
export const DIALYSIS_HEADERS = [
  "Voucher Type",
  "Date",
  "Voucher No.",
  "Party Name",
  "Party Amount",
  "Dr/Cr",
  "Bill Type",
  "Bill Name",
  "Bill Amount",
  "Narration",
  "Cost Center",
];

const REVIEW_HEADERS = [
  "Slno",
  "Patient Name",
  "OP Number",
  "Bill Number",
  "Total Amount",
  "Pending Amount",
  "Company",
  "Reason",
];

// The Teja "Suspense Outstanding Report" places header labels across sparsely
// used columns with empty spacers. Rather than hard-code offsets, find each
// column by its label so a slight column-width change in Teja doesn't break us.
function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const row = rows[i] || [];
    const norm = row.map((v) => (v == null ? "" : String(v).trim().toLowerCase().replace(/\s+/g, " ")));
    if (norm.includes("billnumber") && norm.some((v) => v.includes("pending"))) {
      return { idx: i, cells: norm };
    }
  }
  return { idx: -1, cells: [] };
}

function findCol(cells, ...candidates) {
  for (const c of candidates) {
    const i = cells.indexOf(c);
    if (i !== -1) return i;
  }
  for (const c of candidates) {
    const i = cells.findIndex((cell) => cell.replace(/\s+/g, "").includes(c.replace(/\s+/g, "")));
    if (i !== -1) return i;
  }
  return -1;
}

export function processDialysisWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const { idx: headerRowIdx, cells } = findHeaderRow(rows);
  if (headerRowIdx === -1) {
    throw new Error(
      "Could not find the header row (looking for 'Billnumber' + 'Pending'). Check that this is the Suspense Outstanding Report from Teja."
    );
  }

  const cSlno = findCol(cells, "slno", "sl no");
  const cPatient = findCol(cells, "patient name", "patientname");
  const cOpNo = findCol(cells, "op number", "opnumber");
  const cBillNo = findCol(cells, "billnumber", "bill number");
  const cTotal = findCol(cells, "total amount", "total amnt");
  const cPending = findCol(cells, "pending amnt", "pending amount", "pending");
  const cCompany = findCol(cells, "company");
  const cEmployee = findCol(cells, "employee");

  const missing = [];
  if (cBillNo === -1) missing.push("Billnumber");
  if (cPending === -1) missing.push("Pending Amnt");
  if (cOpNo === -1) missing.push("OP Number");
  if (cPatient === -1) missing.push("Patient Name");
  if (missing.length) {
    throw new Error(
      `Could not find the following required column(s) in the Suspense Outstanding Report: ${missing.join(", ")}.`
    );
  }

  // Sections in this report look like:
  //   row: "DIALYSIS"  (single value in col ~1, rest blank)
  //   row: date-only (col ~3 has the date, other data cols blank)
  //   row: data rows
  //   row: "Total     :" footer (col ~6 has "Total", pending col has section total)
  //   next section: "PHARMACY", etc.
  //
  // Only the DIALYSIS section is receivable data — other sections in this
  // sample have all zeros in pending. Restrict to DIALYSIS.
  const salesRows = [];
  const reviewRows = [];
  const companyOccurrence = new Map();
  const companyTotals = new Map();

  let currentSection = "";
  let currentDate = "";
  let dialysisRowsScanned = 0;
  let bookedRows = 0;
  let skippedNoCompany = 0;
  let skippedZeroPending = 0;
  let totalReceivable = 0;

  const isSectionHeader = (row) => {
    const c1 = row[1];
    if (typeof c1 !== "string" || c1.trim() === "") return false;
    if (String(c1).trim().toLowerCase().startsWith("total")) return false;
    // Section header rows have text in col 1 and everything else blank
    const rest = row.filter((v, i) => i !== 1 && v != null && v !== "");
    return rest.length === 0;
  };

  const isFooterRow = (row) => {
    const c6 = row[6];
    return typeof c6 === "string" && String(c6).trim().toLowerCase().startsWith("total");
  };

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    if (isSectionHeader(row)) {
      currentSection = String(row[1]).trim();
      continue;
    }
    if (isFooterRow(row)) continue;

    // Date-only row: no bill number, but some column carries a date. Update
    // the running currentDate for subsequent data rows in this section.
    const billNoCell = row[cBillNo];
    const hasBillNo = billNoCell != null && String(billNoCell).trim() !== "";
    if (!hasBillNo) {
      for (const cell of row) {
        const d = excelSerialToDDMMYYYY(cell);
        if (d && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(d)) {
          currentDate = d;
          break;
        }
      }
      continue;
    }

    if (currentSection.toUpperCase() !== "DIALYSIS") continue;

    dialysisRowsScanned++;
    const slno = cSlno !== -1 && row[cSlno] != null ? String(row[cSlno]).trim() : "";
    const patientName = cPatient !== -1 && row[cPatient] != null ? String(row[cPatient]).trim() : "";
    const opNumber = cOpNo !== -1 && row[cOpNo] != null ? String(row[cOpNo]).trim() : "";
    const total = cTotal !== -1 ? toNumber(row[cTotal]) : 0;
    const pending = toNumber(row[cPending]);
    const company = cCompany !== -1 && row[cCompany] != null ? String(row[cCompany]).trim() : "";
    const employee = cEmployee !== -1 && row[cEmployee] != null ? String(row[cEmployee]).trim() : "";
    const billNoStr = String(billNoCell).trim();

    // The date lives in a row *before* the first data row of a section
    // (Teja renders the print date under the section title). Prefer any real
    // date the row itself carries; else use the latest we've seen; else blank.
    const rowDate = excelSerialToDDMMYYYY(row[3]);
    const dateStr = rowDate || currentDate;

    if (pending <= 0) {
      skippedZeroPending++;
      continue;
    }
    if (!company) {
      skippedNoCompany++;
      reviewRows.push([
        slno,
        patientName,
        opNumber,
        billNoStr,
        total,
        pending,
        "",
        "No company assigned in the report — cannot book as receivable. Enter the company in Teja and re-export.",
      ]);
      continue;
    }

    bookedRows++;
    totalReceivable += pending;

    const partyPatient = `${opNumber} - ${patientName}`;
    const partyCompany = `${company}-dialysis`;
    companyTotals.set(company, (companyTotals.get(company) || 0) + pending);

    // Voucher type: auto-resolve if same bill no. repeats within the sheet
    const occ = (companyOccurrence.get(billNoStr) || 0) + 1;
    companyOccurrence.set(billNoStr, occ);
    const vt = voucherTypeForOccurrence(DIALYSIS_TOKEN, occ, DIALYSIS_ALT_TOKEN);

    const narration = employee ? `Dialysis receivable — ${employee}` : "Dialysis receivable";
    // DR: company (receivable), CR: patient (settles patient's outstanding)
    // Bill Type "New Ref" with patient ledger as Bill Name lets Tally track
    // the pending balance against a named bill reference.
    salesRows.push([
      vt,
      dateStr,
      billNoStr,
      partyCompany,
      pending,
      "DR",
      "New Ref",
      partyPatient,
      pending,
      narration,
      null,
    ]);
    salesRows.push([
      vt,
      dateStr,
      billNoStr,
      partyPatient,
      pending,
      "CR",
      "New Ref",
      partyPatient,
      pending,
      narration,
      null,
    ]);

    // Track date row too — needed so date carries into subsequent data rows
    // in same section
    if (rowDate) currentDate = rowDate;
  }

  // Track dates that appear on isolated date rows (before data) so subsequent
  // data rows inherit them.
  // (Handled in-loop above — the parser reads date from col 3 of each row and
  // falls back to the last non-empty date it saw.)

  const companySummary = [...companyTotals.entries()]
    .map(([company, pending]) => ({ company, pending, count: 0 }))
    .sort((a, b) => b.pending - a.pending);
  // fill in per-company count
  for (const entry of companySummary) {
    entry.count = salesRows.filter((r) => r[5] === "DR" && r[3] === `${entry.company}-dialysis`).length;
  }

  return {
    salesRows,
    reviewRows,
    companySummary,
    stats: {
      dialysisRowsScanned,
      bookedRows,
      skippedZeroPending,
      skippedNoCompany,
      totalReceivable,
      outputRows: salesRows.length,
    },
  };
}

export function buildDialysisOutputWorkbook(result, ledgerNames) {
  const { salesRows, reviewRows } = result;
  const wb = XLSX.utils.book_new();

  const cols = [
    { wch: 18 },
    { wch: 11 },
    { wch: 14 },
    { wch: 34 },
    { wch: 14 },
    { wch: 6 },
    { wch: 10 },
    { wch: 34 },
    { wch: 14 },
    { wch: 30 },
    { wch: 12 },
  ];
  const resolved = resolveRows(salesRows, ledgerNames);
  const wsSales = XLSX.utils.aoa_to_sheet([DIALYSIS_HEADERS, ...resolved]);
  wsSales["!cols"] = cols;
  XLSX.utils.book_append_sheet(wb, wsSales, "DIALYSIS RECEIVABLES");

  if (reviewRows.length > 0) {
    const wsReview = XLSX.utils.aoa_to_sheet([REVIEW_HEADERS, ...reviewRows]);
    wsReview["!cols"] = [
      { wch: 6 },
      { wch: 26 },
      { wch: 14 },
      { wch: 14 },
      { wch: 12 },
      { wch: 14 },
      { wch: 14 },
      { wch: 60 },
    ];
    XLSX.utils.book_append_sheet(wb, wsReview, "REVIEW");
  }

  return wb;
}
