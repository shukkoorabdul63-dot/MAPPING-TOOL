import * as XLSX from "xlsx";
import { toNumber, normHeader, excelSerialToDDMMYYYY } from "./utils.js";

// Matches the real SALES template layout (8 columns) — confirmed from the
// user's own upload sheet, distinct from the RECEIPT template's 11 columns.
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

// The "column needed" list from the user's own reference sheet, in order.
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

export function processBillsWorkbook(workbook) {
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

  const missing = Object.entries(required)
    .filter(([, v]) => v == null)
    .map(([k]) => k);
  if (missing.length) {
    throw new Error(
      `Could not find the following required column(s) in the header row: ${missing.join(
        ", "
      )}. Check that this is a bill analysis export from Teja.`
    );
  }

  // Pass 1: extract candidate bill rows
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

  // Pass 2: drop cancelled bills entirely
  let cancelledCount = 0;
  const notCancelled = [];
  for (const c of candidates) {
    if (c.billCancelled.toUpperCase() === "Y") {
      cancelledCount++;
      continue;
    }
    notCancelled.push(c);
  }

  // Pass 3: Excel-style dedup on (Bill Number + Bill Name). When the same
  // combination appears more than once (Teja emits identical bills once per
  // PAYMENTCODE), keep the first row and drop the extras — otherwise the
  // income would be counted multiple times.
  const groupKey = (c) => `${c.billNo}||${c.billName}`;
  const seenGroups = new Set();
  const duplicateLogRows = [];
  let duplicateRowCount = 0;
  let duplicateGroupCount = 0;
  const seenGroupsForCount = new Set();
  const clean = [];
  for (const c of notCancelled) {
    const k = groupKey(c);
    if (seenGroups.has(k)) {
      duplicateRowCount++;
      if (!seenGroupsForCount.has(k)) {
        seenGroupsForCount.add(k);
        duplicateGroupCount++;
      }
      duplicateLogRows.push([...rawRow(c), `Duplicate Bill No. + Bill Name — kept the first row, dropped this one`]);
      continue;
    }
    seenGroups.add(k);
    clean.push(c);
  }

  // Pass 4: split into Discharge / Pharmacy (both) / Others, and build a
  // per-Bill-Name gross-amount summary for cross-checking against receipts.
  const dischargeRows = [];
  const othersRawRows = [];
  const pharmacyMappedRows = []; // pharmacy rows that were mapped (audit trail)
  const salesRows = [];
  const grossByBillName = new Map();
  let pharmacyTaxable = 0;
  let pharmacyExempt = 0;
  let totalCgst = 0;
  let totalSgst = 0;

  for (const c of clean) {
    if (!grossByBillName.has(c.billName)) {
      grossByBillName.set(c.billName, { billName: c.billName, grossAmount: 0, count: 0 });
    }
    const gEntry = grossByBillName.get(c.billName);
    gEntry.grossAmount += c.grossAmount;
    gEntry.count += 1;

    const nameLower = c.billName.toLowerCase();
    if (nameLower === "discharge") {
      dischargeRows.push(rawRow(c));
      continue;
    }

    const party = `${c.opNumber} - ${c.patientName}`;
    const incomeHead = `${c.billName} INCOME-${c.ipop}`;

    if (PHARMACY_NAMES.has(nameLower)) {
      // Pharmacy: taxable when VATAMOUNT > 0, else exempted.
      // Taxable value = Gross − VAT (by design). When a bill mixes taxable and
      // exempt items, the exempt portion is absorbed into the income line —
      // this is intended, so the ledger stays balanced without a separate
      // exempt-income split.
      const drAmount = c.billAmount + c.billAdvance - c.ipCreditReturn;
      pharmacyMappedRows.push(rawRow(c));

      if (c.vatAmount > 0) {
        pharmacyTaxable++;
        const taxableValue = c.grossAmount - c.vatAmount;
        const cgst = c.vatAmount / 2;
        const sgst = c.vatAmount / 2;
        totalCgst += cgst;
        totalSgst += sgst;

        salesRows.push(["Journal", c.date, c.billNo, party, drAmount, "DR", null, null]);
        if (c.billDiscount > 0) {
          salesRows.push(["Journal", c.date, c.billNo, "Discount", c.billDiscount, "DR", null, null]);
        }
        salesRows.push(["Journal", c.date, c.billNo, incomeHead, taxableValue, "CR", null, c.doctorCode]);
        salesRows.push(["Journal", c.date, c.billNo, "CGST", cgst, "CR", null, null]);
        salesRows.push(["Journal", c.date, c.billNo, "SGST", sgst, "CR", null, null]);
      } else {
        pharmacyExempt++;
        salesRows.push(["Journal", c.date, c.billNo, party, drAmount, "DR", null, null]);
        if (c.billDiscount > 0) {
          salesRows.push(["Journal", c.date, c.billNo, "Discount", c.billDiscount, "DR", null, null]);
        }
        salesRows.push(["Journal", c.date, c.billNo, incomeHead, c.grossAmount, "CR", null, c.doctorCode]);
      }
      continue;
    }

    // Others: fully mapped SALES journal entry (unchanged).
    othersRawRows.push(rawRow(c));
    const drAmount = c.billAmount + c.billAdvance - c.ipCreditReturn;
    salesRows.push(["Journal", c.date, c.billNo, party, drAmount, "DR", null, null]);
    salesRows.push(["Journal", c.date, c.billNo, incomeHead, c.grossAmount, "CR", null, c.doctorCode]);
    if (c.billDiscount > 0) {
      salesRows.push(["Journal", c.date, c.billNo, "Discount", c.billDiscount, "DR", null, null]);
    }
  }

  const billGrossByName = [...grossByBillName.values()].sort(
    (a, b) => b.grossAmount - a.grossAmount
  );

  return {
    dischargeRows,
    pharmacyMappedRows,
    othersRawRows,
    salesRows,
    duplicateLogRows,
    billGrossByName,
    stats: {
      scannedDataRows,
      cancelledCount,
      duplicateGroupCount,
      duplicateRowCount,
      dischargeCount: dischargeRows.length,
      pharmacyCount: pharmacyMappedRows.length,
      pharmacyTaxable,
      pharmacyExempt,
      totalCgst,
      totalSgst,
      othersCount: othersRawRows.length,
      salesOutputRows: salesRows.length,
      cleanRows: clean.length,
    },
  };
}

export function buildBillsOutputWorkbook(result) {
  const { dischargeRows, pharmacyMappedRows, salesRows, duplicateLogRows } = result;
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
  const wsSales = XLSX.utils.aoa_to_sheet([SALES_HEADERS, ...salesRows]);
  wsSales["!cols"] = salesCols;
  XLSX.utils.book_append_sheet(wb, wsSales, "SALES");

  const rawCols = BILL_NEEDED_HEADERS.map((h) =>
    h === "PATIENTNAME" ? { wch: 26 } : h === "BILLNAME" ? { wch: 22 } : { wch: 14 }
  );

  const wsDischarge = XLSX.utils.aoa_to_sheet([BILL_NEEDED_HEADERS, ...dischargeRows]);
  wsDischarge["!cols"] = rawCols;
  XLSX.utils.book_append_sheet(wb, wsDischarge, "DISCHARGE (raw)");

  const wsPharmacy = XLSX.utils.aoa_to_sheet([BILL_NEEDED_HEADERS, ...pharmacyMappedRows]);
  wsPharmacy["!cols"] = rawCols;
  XLSX.utils.book_append_sheet(wb, wsPharmacy, "PHARMACY (source)");

  const wsDup = XLSX.utils.aoa_to_sheet([DUPLICATE_LOG_HEADERS, ...duplicateLogRows]);
  wsDup["!cols"] = [...rawCols, { wch: 46 }];
  XLSX.utils.book_append_sheet(wb, wsDup, "DUPLICATES REMOVED");

  return wb;
}
