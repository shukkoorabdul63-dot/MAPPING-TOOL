import * as XLSX from "xlsx";
import { toNumber, normHeader, excelSerialToDDMMYYYY } from "./utils.js";
import {
  OTHERS_TOKEN,
  PHARMACY_TOKEN,
  DISCHARGE_TOKEN,
  DISCOUNT_OTHERS_TOKEN,
  DISCOUNT_PHARMACY_TOKEN,
  DISCOUNT_DISCHARGE_TOKEN,
  CGST_TOKEN,
  SGST_TOKEN,
  resolveRows,
  nextOccurrence,
  pushToBucket,
  bucketSheetName,
} from "./tokens.js";
import { IP_CREDIT_FLAGGED_HEADERS, ipCreditFlaggedRowsToAOA } from "./ipcredit.js";

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
 * `ipCreditFlaggedRows` (IP Clear != "Y" rows that still carry a Settled
 * Bill No.) is cross-referenced against the Discharge bills actually in
 * this upload — the IP Credit report can carry stale references to bills
 * from earlier periods that aren't part of this batch at all, so only the
 * ones matching a Discharge bill no. here are surfaced as relevant.
 */
export function mapBills(parsed, ipCreditMap = new Map(), ipCreditFlaggedRows = []) {
  const { candidates, scannedDataRows } = parsed;

  const flaggedBySettled = new Map();
  for (const r of ipCreditFlaggedRows) {
    if (!r.settledBillNo) continue;
    if (!flaggedBySettled.has(r.settledBillNo)) flaggedBySettled.set(r.settledBillNo, []);
    flaggedBySettled.get(r.settledBillNo).push(r);
  }

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

  // Step 3: categorize + map to SALES rows per sheet. Occurrence is tracked
  // per output category (Others / Pharmacy / Discharge) and per bill no.
  // within that category — a repeat bill no. is routed to an additional
  // numbered sheet rather than a renamed voucher type.
  const salesOthersRows = [];
  const salesPharmacyRows = [];
  const salesDischargeRows = [];
  const othersBuckets = [];
  const pharmacyBuckets = [];
  const dischargeBuckets = [];
  const grossByBillName = new Map();
  const dischargeReviewRows = []; // negative income cases

  const othersSeen = new Map();
  const pharmacySeen = new Map();
  const dischargeSeen = new Map();

  let pharmacyTaxable = 0;
  let pharmacyExempt = 0;
  let totalCgst = 0;
  let totalSgst = 0;
  let dischargeMatched = 0;
  let dischargeUnmatched = 0;
  let dischargeIpCreditApplied = 0;
  const dischargeIpCreditFlaggedRows = [];
  const dischargeIpCreditFlaggedBills = new Set();

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

      const flaggedForThisBill = flaggedBySettled.get(c.billNo);
      if (flaggedForThisBill && !dischargeIpCreditFlaggedBills.has(c.billNo)) {
        dischargeIpCreditFlaggedBills.add(c.billNo);
        dischargeIpCreditFlaggedRows.push(...flaggedForThisBill);
      }

      const income = c.grossAmount - ipCredit;
      const drParty = income - c.billDiscount;

      if (income < -0.01) {
        dischargeReviewRows.push([
          ...rawRow(c),
          `Gross ₹${c.grossAmount.toFixed(2)} − IP credit ₹${ipCredit.toFixed(2)} = ₹${income.toFixed(2)} (negative). Review — likely stale or over-applied IP credit.`,
        ]);
      }

      const rows = [
        [DISCHARGE_TOKEN, c.date, c.billNo, party, drParty, "DR", null, null],
      ];
      if (c.billDiscount > 0) {
        rows.push([DISCHARGE_TOKEN, c.date, c.billNo, DISCOUNT_DISCHARGE_TOKEN, c.billDiscount, "DR", null, null]);
      }
      rows.push([DISCHARGE_TOKEN, c.date, c.billNo, incomeHead, income, "CR", null, c.doctorCode]);

      salesDischargeRows.push(...rows);
      const occIdx = nextOccurrence(dischargeSeen, c.billNo);
      pushToBucket(dischargeBuckets, occIdx, rows);
      continue;
    }

    if (PHARMACY_NAMES.has(nameLower)) {
      const drAmount = c.billAmount + c.billAdvance - c.ipCreditReturn;
      const rows = [];

      if (c.vatAmount > 0) {
        pharmacyTaxable++;
        const taxableValue = c.grossAmount - c.vatAmount;
        const cgst = c.vatAmount / 2;
        const sgst = c.vatAmount / 2;
        totalCgst += cgst;
        totalSgst += sgst;

        rows.push([PHARMACY_TOKEN, c.date, c.billNo, party, drAmount, "DR", null, null]);
        if (c.billDiscount > 0) {
          rows.push([PHARMACY_TOKEN, c.date, c.billNo, DISCOUNT_PHARMACY_TOKEN, c.billDiscount, "DR", null, null]);
        }
        rows.push([PHARMACY_TOKEN, c.date, c.billNo, incomeHead, taxableValue, "CR", null, c.doctorCode]);
        rows.push([PHARMACY_TOKEN, c.date, c.billNo, CGST_TOKEN, cgst, "CR", null, null]);
        rows.push([PHARMACY_TOKEN, c.date, c.billNo, SGST_TOKEN, sgst, "CR", null, null]);
      } else {
        pharmacyExempt++;
        rows.push([PHARMACY_TOKEN, c.date, c.billNo, party, drAmount, "DR", null, null]);
        if (c.billDiscount > 0) {
          rows.push([PHARMACY_TOKEN, c.date, c.billNo, DISCOUNT_PHARMACY_TOKEN, c.billDiscount, "DR", null, null]);
        }
        rows.push([PHARMACY_TOKEN, c.date, c.billNo, incomeHead, c.grossAmount, "CR", null, c.doctorCode]);
      }

      salesPharmacyRows.push(...rows);
      const occIdx = nextOccurrence(pharmacySeen, c.billNo);
      pushToBucket(pharmacyBuckets, occIdx, rows);
      continue;
    }

    // Others
    const drAmount = c.billAmount + c.billAdvance - c.ipCreditReturn;
    const rows = [
      [OTHERS_TOKEN, c.date, c.billNo, party, drAmount, "DR", null, null],
      [OTHERS_TOKEN, c.date, c.billNo, incomeHead, c.grossAmount, "CR", null, c.doctorCode],
    ];
    if (c.billDiscount > 0) {
      rows.push([OTHERS_TOKEN, c.date, c.billNo, DISCOUNT_OTHERS_TOKEN, c.billDiscount, "DR", null, null]);
    }

    salesOthersRows.push(...rows);
    const occIdx = nextOccurrence(othersSeen, c.billNo);
    pushToBucket(othersBuckets, occIdx, rows);
  }

  const billGrossByName = [...grossByBillName.values()].sort((a, b) => b.grossAmount - a.grossAmount);

  return {
    salesOthersRows,
    salesPharmacyRows,
    salesDischargeRows,
    othersBuckets,
    pharmacyBuckets,
    dischargeBuckets,
    duplicateLogRows,
    dischargeReviewRows,
    dischargeIpCreditFlaggedRows,
    billGrossByName,
    // Also expose the parsed clean rows so the master patient list can use them.
    cleanCandidates: clean,
    stats: {
      scannedDataRows,
      cancelledCount,
      duplicateGroupCount,
      duplicateRowCount,
      cleanRows: clean.length,
      othersCount: othersSeen.size ? [...othersSeen.values()].reduce((a, b) => a + b, 0) : 0,
      pharmacyCount: pharmacySeen.size ? [...pharmacySeen.values()].reduce((a, b) => a + b, 0) : 0,
      dischargeCount: dischargeSeen.size ? [...dischargeSeen.values()].reduce((a, b) => a + b, 0) : 0,
      othersSheetCount: othersBuckets.length,
      pharmacySheetCount: pharmacyBuckets.length,
      dischargeSheetCount: dischargeBuckets.length,
      pharmacyTaxable,
      pharmacyExempt,
      totalCgst,
      totalSgst,
      dischargeMatched,
      dischargeUnmatched,
      dischargeIpCreditApplied,
      dischargeReviewCount: dischargeReviewRows.length,
      dischargeIpCreditFlaggedCount: dischargeIpCreditFlaggedBills.size,
      dischargeIpCreditFlaggedAmount: dischargeIpCreditFlaggedRows.reduce((a, r) => a + r.creditAmount, 0),
      othersOutputRows: salesOthersRows.length,
      pharmacyOutputRows: salesPharmacyRows.length,
      dischargeOutputRows: salesDischargeRows.length,
    },
  };
}

export function buildBillsOutputWorkbook(result, ledgerNames) {
  const {
    othersBuckets,
    pharmacyBuckets,
    dischargeBuckets,
    duplicateLogRows,
    dischargeReviewRows,
    dischargeIpCreditFlaggedRows,
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

  const addBucketedSheets = (buckets, baseName) => {
    buckets.forEach((rows, i) => {
      if (!rows || rows.length === 0) return;
      const resolved = resolveRows(rows, ledgerNames);
      const ws = XLSX.utils.aoa_to_sheet([SALES_HEADERS, ...resolved]);
      ws["!cols"] = salesCols;
      XLSX.utils.book_append_sheet(wb, ws, bucketSheetName(baseName, i + 1));
    });
  };

  addBucketedSheets(othersBuckets, "SALES - OTHERS");
  addBucketedSheets(pharmacyBuckets, "SALES - PHARMACY");
  addBucketedSheets(dischargeBuckets, "SALES - DISCHARGE");

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

  if (dischargeIpCreditFlaggedRows && dischargeIpCreditFlaggedRows.length > 0) {
    const wsIpFlagged = XLSX.utils.aoa_to_sheet([
      IP_CREDIT_FLAGGED_HEADERS,
      ...ipCreditFlaggedRowsToAOA(dischargeIpCreditFlaggedRows),
    ]);
    wsIpFlagged["!cols"] = [
      { wch: 18 },
      { wch: 14 },
      { wch: 12 },
      { wch: 14 },
      { wch: 26 },
      { wch: 12 },
      { wch: 9 },
      { wch: 16 },
      { wch: 12 },
    ];
    XLSX.utils.book_append_sheet(wb, wsIpFlagged, "IP CREDIT FLAGGED");
  }

  return wb;
}
