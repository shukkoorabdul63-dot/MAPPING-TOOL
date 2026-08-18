import * as XLSX from "xlsx";
import { toNumber, normHeader, excelSerialToDDMMYYYY } from "./utils.js";
import {
  OTHERS_TOKEN,
  PHARMACY_TOKEN,
  DISCHARGE_TOKEN,
  OTHERS_ALT_TOKEN,
  PHARMACY_ALT_TOKEN,
  DISCHARGE_ALT_TOKEN,
  resolveRows,
  voucherTypeForOccurrence,
} from "./tokens.js";

export const SALES_HEADERS = [
  "Voucher Type",
  "Date",
  "Voucher No.",
  "Party Name",
  "Party Amount",
  "Dr/Cr",
  "Narration",
  "Cost Center",
];

export const BILL_NEEDED_HEADERS = [
  "BILLNAME",
  "BILLNUMBER",
  "BILLAMOUNT",
  "BILLCANCELLED",
  "BILLDATE",
  "BILLDISCOUNT",
  "DOCTORCODE",
  "OPNUMBER",
  "PATIENTNAME",
  "BILLADVANCE",
  "TAXABLEAMOUNT",
  "VATAMOUNT",
  "IPCREDITRETURN",
  "IPOP",
  "GROSSAMOUNT",
];

const DUPLICATE_LOG_HEADERS = [...BILL_NEEDED_HEADERS, "Reason"];
const PHARMACY_NAMES = new Set(["pharmacy", "ip pharmacy"]);

function rawRow(c) {
  return [
    c.billName,
    c.billNo,
    c.billAmount,
    c.billCancelled,
    c.date,
    c.billDiscount,
    c.doctorCode,
    c.opNumber,
    c.patientName,
    c.billAdvance,
    c.taxableAmount,
    c.vatAmount,
    c.ipCreditReturn,
    c.ipop,
    c.grossAmount,
  ];
}

/**
 * Parse the raw workbook into a normalized candidate list. Kept separate
 * from mapping so remapping (e.g. after an IP Credit report is loaded, or
 * after ledger names change) is cheap — the parse pass only runs once.
 */
export function parseBillsWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let colMap = {};
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const idx = row.findIndex((c) => normHeader(c) === "billnumber");
    if (idx !== -1) {
      headerRowIdx = i;
      row.forEach((cell, ci) => {
        const key = normHeader(cell);
        if (key && !(key in colMap)) colMap[key] = ci;
      });
      break;
    }
  }

  if (headerRowIdx === -1) {
    throw new Error(
      "Could not find the header row (looking for a 'BILLNUMBER' column). Check that this is a bill analysis export from Teja."
    );
  }

  const col = (name) => colMap[normHeader(name)];
  const required = {
    billName: col("BILLNAME"),
    billNo: col("BILLNUMBER"),
    billAmount: col("BILLAMOUNT"),
    billCancelled: col("BILLCANCELLED"),
    billDate: col("BILLDATE"),
    billDiscount: col("BILLDISCOUNT"),
    doctorCode: col("DOCTORCODE"),
    opNumber: col("OPNUMBER"),
    patientName: col("PATIENTNAME"),
    billAdvance: col("BILLADVANCE"),
    taxableAmount: col("TAXABLEAMOUNT"),
    vatAmount: col("VATAMOUNT"),
    ipCreditReturn: col("IPCREDITRETURN"),
    ipop: col("IPOP"),
    grossAmount: col("GROSSAMOUNT"),
  };

  const missing = Object.entries(required).filter(([, v]) => v == null).map(([k]) => k);
  if (missing.length) {
    throw new Error(
      `Could not find the following required column(s) in the header row: ${missing.join(", ")}.`
    );
  }

  const candidates = [];
  let scannedDataRows = 0;
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const billNo = row[required.billNo];
    if (billNo == null || String(billNo).trim() === "") continue;
    scannedDataRows++;

    candidates.push({
      billName: row[required.billName] != null ? String(row[required.billName]).trim() : "",
      billNo: String(billNo).trim(),
      billAmount: toNumber(row[required.billAmount]),
      billCancelled: row[required.billCancelled] != null ? String(row[required.billCancelled]).trim() : "",
      date: excelSerialToDDMMYYYY(row[required.billDate]),
      billDiscount: toNumber(row[required.billDiscount]),
      doctorCode: row[required.doctorCode] != null ? String(row[required.doctorCode]).trim() : "",
      opNumber: row[required.opNumber] != null ? String(row[required.opNumber]).trim() : "",
      patientName: row[required.patientName] != null ? String(row[required.patientName]).trim() : "",
      billAdvance: toNumber(row[required.billAdvance]),
      taxableAmount: toNumber(row[required.taxableAmount]),
      vatAmount: toNumber(row[required.vatAmount]),
      ipCreditReturn: toNumber(row[required.ipCreditReturn]),
      ipop: row[required.ipop] != null ? String(row[required.ipop]).trim() : "",
      grossAmount: toNumber(row[required.grossAmount]),
    });
  }

  return { candidates, scannedDataRows };
}

/**
 * Map parsed candidates into the SALES journals for Others, Pharmacy, and
 * Discharge. Takes an optional `ipCreditMap` (settled bill no → total
 * credit) — when present, Discharge income is reduced by the mapped amount.
 */
export function mapBills(parsed, ipCreditMap = new Map()) {
  const { candidates, scannedDataRows } = parsed;

  // Step 1: drop cancelled
  let cancelledCount = 0;
  const notCancelled = [];
  for (const c of candidates) {
    if (c.billCancelled.toUpperCase() === "Y") {
      cancelledCount++;
      continue;
    }
    notCancelled.push(c);
  }

  // Step 2: Excel-style dedup on (Bill No + Bill Name). Teja emits one row per
  // PAYMENTCODE for the same underlying bill; keep the first, drop the rest.
  const seenGroups = new Set();
  const duplicateLogRows = [];
  let duplicateRowCount = 0;
  const seenGroupsForCount = new Set();
  let duplicateGroupCount = 0;
  const clean = [];
  for (const c of notCancelled) {
    const k = `${c.billNo}||${c.billName}`;
    if (seenGroups.has(k)) {
      duplicateRowCount++;
      if (!seenGroupsForCount.has(k)) {
        seenGroupsForCount.add(k);
        duplicateGroupCount++;
      }
      duplicateLogRows.push([...rawRow(c), "Duplicate Bill No. + Bill Name — kept first, dropped this"]);
      continue;
    }
    seenGroups.add(k);
    clean.push(c);
  }

  // Step 3: categorize + map to SALES rows per sheet, tracking per-sheet
  // per-billNo occurrence so we can auto-resolve duplicate voucher numbers
  // that fall inside the same sheet.
  const salesOthersRows = [];
  const salesPharmacyRows = [];
  const salesDischargeRows = [];
  const grossByBillName = new Map();
  const dischargeReviewRows = []; // negative income cases

  const othersOccurrence = new Map();
  const pharmacyOccurrence = new Map();
  const dischargeOccurrence = new Map();

  let pharmacyTaxable = 0;
  let pharmacyExempt = 0;
  let totalCgst = 0;
  let totalSgst = 0;
  let dischargeMatched = 0;
  let dischargeUnmatched = 0;
  let dischargeIpCreditApplied = 0;

  for (const c of clean) {
    if (!grossByBillName.has(c.billName)) {
      grossByBillName.set(c.billName, { billName: c.billName, grossAmount: 0, count: 0 });
    }
    const gEntry = grossByBillName.get(c.billName);
    gEntry.grossAmount += c.grossAmount;
    gEntry.count += 1;

    const nameLower = c.billName.toLowerCase();
    const party = `${c.opNumber} - ${c.patientName}`;
    const incomeHead = `${c.billName} INCOME-${c.ipop}`;

    if (nameLower === "discharge") {
      const ipCredit = ipCreditMap.get(c.billNo) || 0;
      if (ipCredit > 0) {
        dischargeMatched++;
        dischargeIpCreditApplied += ipCredit;
      } else {
        dischargeUnmatched++;
      }

      const income = c.grossAmount - ipCredit;
      const drParty = income - c.billDiscount;

      // Flag negative income (over-refund / stale IP credit) for the user
      if (income < -0.01) {
        dischargeReviewRows.push([
          ...rawRow(c),
          `Gross ₹${c.grossAmount.toFixed(2)} − IP credit ₹${ipCredit.toFixed(2)} = ₹${income.toFixed(2)} (negative). Review — likely stale or over-applied IP credit.`,
        ]);
      }

      const occ = (dischargeOccurrence.get(c.billNo) || 0) + 1;
      dischargeOccurrence.set(c.billNo, occ);
      const vt = voucherTypeForOccurrence(DISCHARGE_TOKEN, occ, DISCHARGE_ALT_TOKEN);

      salesDischargeRows.push([vt, c.date, c.billNo, party, drParty, "DR", null, null]);
      if (c.billDiscount > 0) {
        salesDischargeRows.push([vt, c.date, c.billNo, "Discount", c.billDiscount, "DR", null, null]);
      }
      salesDischargeRows.push([vt, c.date, c.billNo, incomeHead, income, "CR", null, c.doctorCode]);
      continue;
    }

    if (PHARMACY_NAMES.has(nameLower)) {
      const drAmount = c.billAmount + c.billAdvance - c.ipCreditReturn;
      const occ = (pharmacyOccurrence.get(c.billNo) || 0) + 1;
      pharmacyOccurrence.set(c.billNo, occ);
      const vt = voucherTypeForOccurrence(PHARMACY_TOKEN, occ, PHARMACY_ALT_TOKEN);

      if (c.vatAmount > 0) {
        pharmacyTaxable++;
        const taxableValue = c.grossAmount - c.vatAmount;
        const cgst = c.vatAmount / 2;
        const sgst = c.vatAmount / 2;
        totalCgst += cgst;
        totalSgst += sgst;

        salesPharmacyRows.push([vt, c.date, c.billNo, party, drAmount, "DR", null, null]);
        if (c.billDiscount > 0) {
          salesPharmacyRows.push([vt, c.date, c.billNo, "Discount", c.billDiscount, "DR", null, null]);
        }
        salesPharmacyRows.push([vt, c.date, c.billNo, incomeHead, taxableValue, "CR", null, c.doctorCode]);
        salesPharmacyRows.push([vt, c.date, c.billNo, "CGST", cgst, "CR", null, null]);
        salesPharmacyRows.push([vt, c.date, c.billNo, "SGST", sgst, "CR", null, null]);
      } else {
        pharmacyExempt++;
        salesPharmacyRows.push([vt, c.date, c.billNo, party, drAmount, "DR", null, null]);
        if (c.billDiscount > 0) {
          salesPharmacyRows.push([vt, c.date, c.billNo, "Discount", c.billDiscount, "DR", null, null]);
        }
        salesPharmacyRows.push([vt, c.date, c.billNo, incomeHead, c.grossAmount, "CR", null, c.doctorCode]);
      }
      continue;
    }

    // Others
    const drAmount = c.billAmount + c.billAdvance - c.ipCreditReturn;
    const occ = (othersOccurrence.get(c.billNo) || 0) + 1;
    othersOccurrence.set(c.billNo, occ);
    const vt = voucherTypeForOccurrence(OTHERS_TOKEN, occ, OTHERS_ALT_TOKEN);

    salesOthersRows.push([vt, c.date, c.billNo, party, drAmount, "DR", null, null]);
    salesOthersRows.push([vt, c.date, c.billNo, incomeHead, c.grossAmount, "CR", null, c.doctorCode]);
    if (c.billDiscount > 0) {
      salesOthersRows.push([vt, c.date, c.billNo, "Discount", c.billDiscount, "DR", null, null]);
    }
  }

  const billGrossByName = [...grossByBillName.values()].sort((a, b) => b.grossAmount - a.grossAmount);

  return {
    salesOthersRows,
    salesPharmacyRows,
    salesDischargeRows,
    duplicateLogRows,
    dischargeReviewRows,
    billGrossByName,
    // Also expose the parsed clean rows so the master patient list can use them.
    cleanCandidates: clean,
    stats: {
      scannedDataRows,
      cancelledCount,
      duplicateGroupCount,
      duplicateRowCount,
      cleanRows: clean.length,
      othersCount: [...othersOccurrence.values()].reduce((a, b) => a + b, 0),
      pharmacyCount: [...pharmacyOccurrence.values()].reduce((a, b) => a + b, 0),
      dischargeCount: [...dischargeOccurrence.values()].reduce((a, b) => a + b, 0),
      pharmacyTaxable,
      pharmacyExempt,
      totalCgst,
      totalSgst,
      dischargeMatched,
      dischargeUnmatched,
      dischargeIpCreditApplied,
      dischargeReviewCount: dischargeReviewRows.length,
      othersOutputRows: salesOthersRows.length,
      pharmacyOutputRows: salesPharmacyRows.length,
      dischargeOutputRows: salesDischargeRows.length,
    },
  };
}

export function buildBillsOutputWorkbook(result, ledgerNames) {
  const {
    salesOthersRows,
    salesPharmacyRows,
    salesDischargeRows,
    duplicateLogRows,
    dischargeReviewRows,
  } = result;
  const wb = XLSX.utils.book_new();

  const salesCols = [
    { wch: 14 },
    { wch: 11 },
    { wch: 14 },
    { wch: 34 },
    { wch: 14 },
    { wch: 6 },
    { wch: 16 },
    { wch: 12 },
  ];

  const addSalesSheet = (rows, name) => {
    const resolved = resolveRows(rows, ledgerNames);
    const ws = XLSX.utils.aoa_to_sheet([SALES_HEADERS, ...resolved]);
    ws["!cols"] = salesCols;
    XLSX.utils.book_append_sheet(wb, ws, name);
  };

  addSalesSheet(salesOthersRows, "SALES - OTHERS");
  addSalesSheet(salesPharmacyRows, "SALES - PHARMACY");
  addSalesSheet(salesDischargeRows, "SALES - DISCHARGE");

  const rawCols = BILL_NEEDED_HEADERS.map((h) =>
    h === "PATIENTNAME" ? { wch: 26 } : h === "BILLNAME" ? { wch: 22 } : { wch: 14 }
  );

  const wsDup = XLSX.utils.aoa_to_sheet([DUPLICATE_LOG_HEADERS, ...duplicateLogRows]);
  wsDup["!cols"] = [...rawCols, { wch: 46 }];
  XLSX.utils.book_append_sheet(wb, wsDup, "DUPLICATE LOG");

  if (dischargeReviewRows && dischargeReviewRows.length > 0) {
    const wsReview = XLSX.utils.aoa_to_sheet([[...BILL_NEEDED_HEADERS, "Note"], ...dischargeReviewRows]);
    wsReview["!cols"] = [...rawCols, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, wsReview, "DISCHARGE TO REVIEW");
  }

  return wb;
}
