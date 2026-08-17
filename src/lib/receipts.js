import * as XLSX from "xlsx";
import { toNumber, normHeader, excelSerialToDDMMYYYY } from "./utils.js";
import {
  LEDGER_TOKEN,
  DEFAULT_LEDGER_NAMES,
  resolveRows,
  voucherTypeForOccurrence,
} from "./tokens.js";

export const RECEIPT_HEADERS = [
  "Voucher Type",
  "Date",
  "Voucher No.",
  "Party Name",
  "Party Amount",
  "Dr/Cr",
  "Bill type",
  "Bill Name",
  "Bill Amount",
  "Narration",
  "Cost Center",
];

export const REVIEW_HEADERS = [
  "Bill Number",
  "Date",
  "Patient ID",
  "Patient Name",
  "Cash",
  "Card",
  "NEFT",
  "IP Credit",
  "Reason",
];

function processWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  // Locate header row by finding a row containing "Billnumber" (case-insensitive, trimmed)
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
      "Could not find the header row (looking for a 'Billnumber' column). Check that this is a Detailed Bill Register export from Teja."
    );
  }

  const col = (name) => colMap[name];
  const cBillDate = col("bill date");
  const cBillNo = col("billnumber");
  const cPatientName = col("patient name");
  // Teja exports label the patient-identifier column differently across
  // reports — sometimes "OP NO", sometimes "UHID". Accept either.
  const cOpNo = col("op no") ?? col("uhid");
  const cCash = col("cash") ?? col("cash ");
  const cCard = col("card");
  const cNeft = col("neft");
  const cIpCredit = col("ip credit");
  const cGrossAmount = col("gross amount");

  if ([cBillNo, cPatientName, cOpNo, cCash, cCard, cNeft].some((v) => v == null)) {
    const missing = [];
    if (cBillNo == null) missing.push("Billnumber");
    if (cPatientName == null) missing.push("Patient Name");
    if (cOpNo == null) missing.push("OP NO / UHID");
    if (cCash == null) missing.push("Cash");
    if (cCard == null) missing.push("Card");
    if (cNeft == null) missing.push("NEFT");
    throw new Error(
      `Could not find the following required column(s) in the header row: ${missing.join(
        ", "
      )}. Check that this is a Detailed Bill Register export from Teja.`
    );
  }

  // Pass 1: extract candidate bill rows, tracking the "Bill Name :" section
  // header each row falls under (e.g. ADVANCE, AMBULANCE, PHARMACY CASH BILL)
  const candidates = [];
  let scannedDataRows = 0;
  let currentBillName = "UNSPECIFIED";
  const billNameRegex = /^bill name\s*:?/i;

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const firstCell = row[0];
    if (typeof firstCell === "string" && billNameRegex.test(firstCell.trim())) {
      const name = firstCell.trim().replace(billNameRegex, "").trim();
      currentBillName = name || "UNSPECIFIED";
      continue;
    }

    const billNo = row[cBillNo];
    if (billNo == null || String(billNo).trim() === "") continue;
    scannedDataRows++;

    const cash = toNumber(row[cCash]);
    const card = toNumber(row[cCard]);
    const neft = toNumber(row[cNeft]);
    const ipCredit = cIpCredit != null ? toNumber(row[cIpCredit]) : 0;
    const grossAmount = cGrossAmount != null ? toNumber(row[cGrossAmount]) : 0;

    candidates.push({
      billNo: String(billNo).trim(),
      date: excelSerialToDDMMYYYY(row[cBillDate]),
      patientName: row[cPatientName] != null ? String(row[cPatientName]).trim() : "",
      patientId: row[cOpNo] != null ? String(row[cOpNo]).trim() : "",
      billName: currentBillName,
      cash,
      card,
      neft,
      ipCredit,
      grossAmount,
    });
  }

  // Bill-name summary: sum Cash / Card / NEFT per section, across every row
  // that carries a bill number (duplicates included) — this is a reconciliation
  // total against Teja's own report, independent of which rows made it into
  // the Tally-ready RECEIPT sheet.
  const summaryMap = new Map();
  for (const c of candidates) {
    const key = c.billName;
    if (!summaryMap.has(key)) {
      summaryMap.set(key, { billName: key, cash: 0, card: 0, neft: 0, grossAmount: 0, count: 0 });
    }
    const entry = summaryMap.get(key);
    entry.cash += c.cash;
    entry.card += c.card;
    entry.neft += c.neft;
    entry.grossAmount += c.grossAmount;
    entry.count += 1;
  }
  const billNameSummary = [...summaryMap.values()]
    .map((e) => ({ ...e, total: e.cash + e.card + e.neft }))
    .sort((a, b) => b.total - a.total);

  // Pass 2: find duplicate bill numbers, and track each row's occurrence
  // index within its bill-number group (in file order) so duplicates can
  // be auto-resolved with an alternate Voucher Type instead of excluded.
  const countByBillNo = new Map();
  for (const c of candidates) {
    countByBillNo.set(c.billNo, (countByBillNo.get(c.billNo) || 0) + 1);
  }
  const seenSoFar = new Map();

  const receiptRows = [];
  const paymentRows = [];
  const creditNoteRows = [];
  const reviewRows = [];
  let skippedZero = 0;
  let paymentCount = 0;
  let receiptCount = 0;
  let creditNoteCount = 0;
  let duplicateRowCount = 0;

  for (const c of candidates) {
    const totalForBillNo = countByBillNo.get(c.billNo);
    const occurrenceIndex = (seenSoFar.get(c.billNo) || 0) + 1;
    seenSoFar.set(c.billNo, occurrenceIndex);
    const isDuplicate = totalForBillNo > 1;

    const party = `${c.patientId} - ${c.patientName}`;
    const negativeCashAmount = c.cash < 0 ? Math.abs(c.cash) : 0;
    const negativeIpCreditAmount = c.ipCredit < 0 ? Math.abs(c.ipCredit) : 0;
    const receiptCash = c.cash > 0 ? c.cash : 0;
    const receiptCard = c.card > 0 ? c.card : 0;
    const receiptNeft = c.neft > 0 ? c.neft : 0;
    const receiptTotal = receiptCash + receiptCard + receiptNeft;

    if (
      negativeCashAmount === 0 &&
      negativeIpCreditAmount === 0 &&
      receiptTotal === 0
    ) {
      skippedZero++;
      continue;
    }

    if (isDuplicate) {
      duplicateRowCount++;
      reviewRows.push([
        c.billNo,
        c.date,
        c.patientId,
        c.patientName,
        c.cash,
        c.card,
        c.neft,
        c.ipCredit,
        `Duplicate #${occurrenceIndex} of ${totalForBillNo} — auto-resolved with an alternate voucher type, see voucher sheets`,
      ]);
    }

    // Negative Cash: two things happen —
    //   1) the income for this bill is reversed (Credit Note): Bill Name Dr, Party Cr
    //   2) the cash itself is refunded (Payment): Party Dr, Cash ledger Cr
    if (negativeCashAmount > 0) {
      paymentCount++;
      const vType = voucherTypeForOccurrence("Payment", occurrenceIndex);
      paymentRows.push([vType, c.date, c.billNo, party, negativeCashAmount, "DR", null, null, null, null, null]);
      paymentRows.push([
        vType,
        c.date,
        c.billNo,
        LEDGER_TOKEN.CASH,
        negativeCashAmount,
        "CR",
        null,
        null,
        null,
        null,
        null,
      ]);

      creditNoteCount++;
      const cnType = voucherTypeForOccurrence("Credit Note", occurrenceIndex);
      creditNoteRows.push([
        cnType,
        c.date,
        c.billNo,
        c.billName,
        negativeCashAmount,
        "DR",
        null,
        null,
        null,
        null,
        null,
      ]);
      creditNoteRows.push([
        cnType,
        c.date,
        c.billNo,
        party,
        negativeCashAmount,
        "CR",
        null,
        null,
        null,
        null,
        null,
      ]);
    }

    // Negative IP Credit alone: a billing/credit revision with no cash
    // movement — still an income reversal, so it also gets a Credit Note.
    if (negativeIpCreditAmount > 0) {
      creditNoteCount++;
      const cnType = voucherTypeForOccurrence("Credit Note", occurrenceIndex);
      creditNoteRows.push([
        cnType,
        c.date,
        c.billNo,
        c.billName,
        negativeIpCreditAmount,
        "DR",
        null,
        null,
        null,
        null,
        null,
      ]);
      creditNoteRows.push([
        cnType,
        c.date,
        c.billNo,
        party,
        negativeIpCreditAmount,
        "CR",
        null,
        null,
        null,
        null,
        null,
      ]);
    }

    if (receiptTotal > 0) {
      receiptCount++;
      const rType = voucherTypeForOccurrence("Receipt", occurrenceIndex);
      receiptRows.push([rType, c.date, c.billNo, party, receiptTotal, "CR", null, null, null, null, null]);
      if (receiptCash > 0) {
        receiptRows.push([
          rType,
          c.date,
          c.billNo,
          LEDGER_TOKEN.CASH,
          receiptCash,
          "DR",
          null,
          null,
          null,
          null,
          null,
        ]);
      }
      if (receiptCard > 0) {
        receiptRows.push([
          rType,
          c.date,
          c.billNo,
          LEDGER_TOKEN.CARD,
          receiptCard,
          "DR",
          null,
          null,
          null,
          null,
          null,
        ]);
      }
      if (receiptNeft > 0) {
        receiptRows.push([
          rType,
          c.date,
          c.billNo,
          LEDGER_TOKEN.NEFT,
          receiptNeft,
          "DR",
          null,
          null,
          null,
          null,
          null,
        ]);
      }
    }
  }

  return {
    receiptRows,
    paymentRows,
    creditNoteRows,
    reviewRows,
    billNameSummary,
    stats: {
      scannedDataRows,
      candidateBills: candidates.length,
      uniqueBills: candidates.length - duplicateRowCount,
      duplicateRowCount,
      duplicateBillNos: [...countByBillNo.values()].filter((v) => v > 1).length,
      skippedZero,
      paymentCount,
      receiptCount,
      creditNoteCount,
      outputRows: receiptRows.length + paymentRows.length + creditNoteRows.length,
    },
  };
}

function buildOutputWorkbook(
  receiptRows,
  paymentRows,
  creditNoteRows,
  reviewRows,
  billNameSummary,
  ledgerNames
) {
  const wb = XLSX.utils.book_new();

  const stdCols = [
    { wch: 14 },
    { wch: 11 },
    { wch: 14 },
    { wch: 32 },
    { wch: 13 },
    { wch: 6 },
    { wch: 10 },
    { wch: 10 },
    { wch: 12 },
    { wch: 14 },
    { wch: 12 },
  ];

  const addVoucherSheet = (rows, sheetName) => {
    const resolved = resolveRows(rows, ledgerNames);
    const aoa = [RECEIPT_HEADERS, ...resolved];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = stdCols;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  };

  addVoucherSheet(receiptRows, "RECEIPT");
  addVoucherSheet(paymentRows, "PAYMENT");
  addVoucherSheet(creditNoteRows, "CREDIT NOTE");

  const reviewAoa = [REVIEW_HEADERS, ...reviewRows];
  const wsReview = XLSX.utils.aoa_to_sheet(reviewAoa);
  wsReview["!cols"] = [
    { wch: 14 },
    { wch: 11 },
    { wch: 14 },
    { wch: 26 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 34 },
  ];
  XLSX.utils.book_append_sheet(wb, wsReview, "DUPLICATE LOG");

  if (billNameSummary && billNameSummary.length) {
    const totals = billNameSummary.reduce(
      (acc, r) => ({
        cash: acc.cash + r.cash,
        card: acc.card + r.card,
        neft: acc.neft + r.neft,
        total: acc.total + r.total,
        count: acc.count + r.count,
      }),
      { cash: 0, card: 0, neft: 0, total: 0, count: 0 }
    );
    const summaryAoa = [
      ["Bill Name", "Cash", "Card", "NEFT", "Total", "Bill Count"],
      ...billNameSummary.map((r) => [r.billName, r.cash, r.card, r.neft, r.total, r.count]),
      ["TOTAL", totals.cash, totals.card, totals.neft, totals.total, totals.count],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoa);
    wsSummary["!cols"] = [
      { wch: 28 },
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
      { wch: 12 },
    ];
    XLSX.utils.book_append_sheet(wb, wsSummary, "SUMMARY BY BILL TYPE");
  }

  const ledgerAoa = [
    ["Payment Mode", "Tally Ledger Name Used"],
    ["Cash", ledgerNames.cash || DEFAULT_LEDGER_NAMES.cash],
    ["Card", ledgerNames.card || DEFAULT_LEDGER_NAMES.card],
    ["NEFT", ledgerNames.neft || DEFAULT_LEDGER_NAMES.neft],
    [],
    ["Duplicate bill numbers — alternate voucher type", ledgerNames.altVoucherType || DEFAULT_LEDGER_NAMES.altVoucherType],
  ];
  const wsLedger = XLSX.utils.aoa_to_sheet(ledgerAoa);
  wsLedger["!cols"] = [{ wch: 42 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(wb, wsLedger, "LEDGERS USED");

  return wb;
}
export { processWorkbook, buildOutputWorkbook };
