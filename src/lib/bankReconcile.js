import * as XLSX from "xlsx";
import { toNumber, excelSerialToDDMMYYYY, normHeader } from "./utils.js";

// Strip a literal leading/trailing single-quote from a cell value. Some
// payment-gateway CSV exports (e.g. the UPI statement) wrap every text/ID
// cell in a literal `'` character as part of the value itself — the parsed
// cell is the string `'SUCCESS'`, not `SUCCESS`.
function stripQuote(v) {
  return typeof v === "string" ? v.replace(/^'+|'+$/g, "") : v;
}

function findHeaderRow(rows, maxScan, isHeaderRow) {
  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    const row = rows[i] || [];
    const norm = row.map((v) => normHeader(v));
    if (isHeaderRow(norm)) return { idx: i, cells: norm };
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

function sheetToRows(workbook) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

// --- Card bank/acquirer settlement statement ------------------------------
// Header row carries S.No, PAID Date, Txn Date, Txn Amt, NET, ~25 more
// columns. Data rows have a numeric S.No; footer rows (TID-SUBTOTAL,
// MID-SUBTOTAL, CREDIT TOTAL, GRAND TOTAL) and the trailing disclaimer text
// have a non-numeric S.No, so that single check separates data from
// everything else.
export function parseCardStatement(workbook) {
  const rows = sheetToRows(workbook);
  const { idx: headerRowIdx, cells } = findHeaderRow(
    rows,
    30,
    (norm) => norm.includes("s.no") && norm.includes("txn date") && norm.includes("txn amt")
  );
  if (headerRowIdx === -1) {
    throw new Error(
      "Could not find the header row of the Card statement (looking for 'S.No', 'Txn Date', 'Txn Amt'). Check that this is the Card settlement report."
    );
  }

  const cSNo = findCol(cells, "s.no");
  const cTxnDate = findCol(cells, "txn date");
  const cTxnAmt = findCol(cells, "txn amt");
  const cCardNumber = findCol(cells, "card number");

  const missing = [];
  if (cSNo === -1) missing.push("S.No");
  if (cTxnDate === -1) missing.push("Txn Date");
  if (cTxnAmt === -1) missing.push("Txn Amt");
  if (missing.length) {
    throw new Error(`Could not find the following required column(s) in the Card statement: ${missing.join(", ")}.`);
  }

  const dataRows = [];
  let skippedRows = 0;
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    if (typeof row[cSNo] !== "number") {
      skippedRows++;
      continue;
    }
    dataRows.push({
      date: excelSerialToDDMMYYYY(row[cTxnDate]),
      amount: toNumber(row[cTxnAmt]),
      // Masked card number — used to group multiple swipes by the same card
      // on the same day when a customer paid in more than one transaction
      // but Tally only booked one combined receipt for the total.
      customerRef: cCardNumber !== -1 && row[cCardNumber] != null ? String(row[cCardNumber]).trim() : "",
      raw: row,
    });
  }

  const sumAmount = dataRows.reduce((a, r) => a + r.amount, 0);
  return { rows: dataRows, stats: { dataRows: dataRows.length, skippedRows, sumAmount } };
}

// --- UPI / payment-gateway transaction statement (CSV) --------------------
export function parseUpiStatement(workbook) {
  const rows = sheetToRows(workbook);
  if (rows.length === 0) {
    throw new Error("The UPI statement appears to be empty.");
  }
  const cells = rows[0].map((v) => normHeader(stripQuote(v)));
  const cDate = findCol(cells, "transaction_date");
  const cAmount = findCol(cells, "amount");
  const cStatus = findCol(cells, "status");
  const cVpa = findCol(cells, "customer_vpa");

  const missing = [];
  if (cDate === -1) missing.push("Transaction_Date");
  if (cAmount === -1) missing.push("Amount");
  if (cStatus === -1) missing.push("Status");
  if (missing.length) {
    throw new Error(`Could not find the following required column(s) in the UPI statement: ${missing.join(", ")}.`);
  }

  const dataRows = [];
  let totalRows = 0;
  let skippedNonSuccess = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const statusCell = stripQuote(row[cStatus]);
    if (statusCell == null || String(statusCell).trim() === "") continue;
    totalRows++;
    const status = String(statusCell).trim().toUpperCase();
    if (status !== "SUCCESS") {
      skippedNonSuccess++;
      continue;
    }
    const dateCell = String(stripQuote(row[cDate]) ?? "").trim();
    const dateOnly = dateCell.slice(0, 10); // "2026-08-02 00:17:34" -> "2026-08-02"
    const isoMatch = dateOnly.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const date = isoMatch ? `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]}` : dateOnly;

    dataRows.push({
      date,
      amount: toNumber(row[cAmount]),
      // Used to group multiple payments by the same UPI handle on the same
      // day when a customer paid in more than one transaction but Tally
      // only booked one combined receipt for the total.
      customerRef: cVpa !== -1 && row[cVpa] != null ? String(stripQuote(row[cVpa])).trim() : "",
      raw: row,
    });
  }

  const sumAmount = dataRows.reduce((a, r) => a + r.amount, 0);
  return {
    rows: dataRows,
    stats: { totalRows, successRows: dataRows.length, skippedNonSuccess, sumAmount },
  };
}

// --- Card / UPI Tally ledger voucher export --------------------------------
// Same report format for both accounts — only the ledger name in the
// header text differs, which isn't needed for parsing. Data rows have a
// non-empty Vch No. AND Vch Type ("Receipt", "Receipt-Discharge", ...);
// that single test also excludes the Opening Balance row, the subtotal
// row, the Closing Balance row, and the final total row.
export function parseLedgerExport(workbook) {
  const rows = sheetToRows(workbook);
  const { idx: headerRowIdx, cells } = findHeaderRow(
    rows,
    15,
    (norm) => norm.includes("particulars") && norm.includes("vch type") && norm.includes("debit")
  );
  if (headerRowIdx === -1) {
    throw new Error(
      "Could not find the header row of the ledger export (looking for 'Particulars', 'Vch Type', 'Debit'). Check that this is a Tally ledger voucher report."
    );
  }

  const cDate = findCol(cells, "date");
  // The "Particulars" header cell is merged across two columns in the
  // source report — the label lands on the "To"/"By" direction-marker
  // column, but the actual particulars text (patient name) is one column
  // to the right (confirmed against the real export: header col holds
  // "Particulars", data rows hold "To"/"By" there and the real name in the
  // next column).
  const cParticularsLabel = findCol(cells, "particulars");
  const cParticulars = cParticularsLabel !== -1 ? cParticularsLabel + 1 : -1;
  const cVchType = findCol(cells, "vch type");
  const cVchNo = findCol(cells, "vch no.", "vch no");
  const cDebit = findCol(cells, "debit");

  const missing = [];
  if (cDate === -1) missing.push("Date");
  if (cVchType === -1) missing.push("Vch Type");
  if (cVchNo === -1) missing.push("Vch No.");
  if (cDebit === -1) missing.push("Debit");
  if (missing.length) {
    throw new Error(`Could not find the following required column(s) in the ledger export: ${missing.join(", ")}.`);
  }

  const dataRows = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const vchNo = row[cVchNo];
    const vchType = row[cVchType];
    const hasVchNo = vchNo != null && String(vchNo).trim() !== "";
    const hasVchType = vchType != null && String(vchType).trim() !== "";
    if (!hasVchNo || !hasVchType) continue;

    dataRows.push({
      date: excelSerialToDDMMYYYY(row[cDate]),
      amount: toNumber(row[cDebit]),
      vchNo: String(vchNo).trim(),
      vchType: String(vchType).trim(),
      particulars: cParticulars !== -1 && row[cParticulars] != null ? String(row[cParticulars]).trim() : "",
      raw: row,
    });
  }

  const sumAmount = dataRows.reduce((a, r) => a + r.amount, 0);
  return { rows: dataRows, stats: { dataRows: dataRows.length, sumAmount } };
}

// --- Matching ---------------------------------------------------------------
// Three passes, each only considering rows the previous pass left unmatched:
//
//   1. Exact:   same date, same amount to the paisa.
//   2. Rounded: same date, amounts equal once both sides are rounded to the
//      nearest whole rupee. Ledger exports commonly carry paise (e.g.
//      203.65) that the original bank/UPI statement never shows (it's
//      already a rounded whole-rupee figure for the same transaction), so a
//      paisa-exact comparison alone would falsely report these as
//      unmatched.
//   3. Grouped: a customer sometimes pays in two (or more) separate
//      transactions the same day, but Tally books ONE combined receipt for
//      the total. Remaining unmatched statement rows are grouped by
//      (date, customerRef) — Card Number for Card, Customer_VPA for UPI —
//      summed, and that sum is matched (exact or rounded) against a
//      remaining unmatched ledger row on the same date. All statement rows
//      in the group are recorded as matched against that one ledger row.
//
// Within a single pass, a (date, amount) key with more than one row on BOTH
// sides is a genuine duplicate (e.g. two ₹300 swipes one day) — still
// matched (pairing in original file order), just counted as "ambiguous" so
// the UI can surface one aggregate note instead of flagging individual rows.

function exactKey(date, amount) {
  return `${date}|${Math.round(amount * 100)}`;
}

function roundedKey(date, amount) {
  return `${date}|${Math.round(amount)}`;
}

// Pairs up rows from two pools that share a key (built by `keyOf`), marking
// paired rows used in both `usedS`/`usedL`. Returns { ambiguousKeyCount,
// ambiguousMatchCount } for this pass, and pushes matches (tagged
// `matchType`) onto `matched`.
function passKeyMatch(statementRows, ledgerRows, usedS, usedL, matched, keyOf, matchType) {
  const byS = new Map();
  statementRows.forEach((r, i) => {
    if (usedS[i]) return;
    const k = keyOf(r.date, r.amount);
    if (!byS.has(k)) byS.set(k, []);
    byS.get(k).push(i);
  });
  const byL = new Map();
  ledgerRows.forEach((r, i) => {
    if (usedL[i]) return;
    const k = keyOf(r.date, r.amount);
    if (!byL.has(k)) byL.set(k, []);
    byL.get(k).push(i);
  });

  let ambiguousKeyCount = 0;
  let ambiguousMatchCount = 0;
  const allKeys = new Set([...byS.keys(), ...byL.keys()]);
  for (const key of allKeys) {
    const sIdxs = byS.get(key) || [];
    const lIdxs = byL.get(key) || [];
    const n = Math.min(sIdxs.length, lIdxs.length);
    if (sIdxs.length > 1 && lIdxs.length > 1) {
      ambiguousKeyCount++;
      ambiguousMatchCount += n;
    }
    for (let i = 0; i < n; i++) {
      const si = sIdxs[i];
      const li = lIdxs[i];
      usedS[si] = true;
      usedL[li] = true;
      matched.push({
        date: statementRows[si].date,
        amount: statementRows[si].amount,
        statementRow: statementRows[si],
        ledgerRow: ledgerRows[li],
        matchType,
      });
    }
  }
  return { ambiguousKeyCount, ambiguousMatchCount };
}

function passGroupedMatch(statementRows, ledgerRows, usedS, usedL, matched) {
  const groups = new Map(); // "date||customerRef" -> statement row indices
  statementRows.forEach((r, i) => {
    if (usedS[i]) return;
    const ref = (r.customerRef || "").trim();
    if (!ref) return;
    const key = `${r.date}||${ref}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(i);
  });

  const ledgerByDate = new Map();
  ledgerRows.forEach((r, i) => {
    if (usedL[i]) return;
    if (!ledgerByDate.has(r.date)) ledgerByDate.set(r.date, []);
    ledgerByDate.get(r.date).push(i);
  });

  let groupedMatchCount = 0;
  let groupedStatementRowCount = 0;
  for (const idxs of groups.values()) {
    if (idxs.length < 2) continue; // only genuine multi-payment groups
    const date = statementRows[idxs[0]].date;
    const sum = idxs.reduce((a, i) => a + statementRows[i].amount, 0);
    const candidates = ledgerByDate.get(date) || [];
    const foundLi = candidates.find(
      (li) => !usedL[li] && (Math.abs(ledgerRows[li].amount - sum) < 0.01 || Math.round(ledgerRows[li].amount) === Math.round(sum))
    );
    if (foundLi == null) continue;

    usedL[foundLi] = true;
    groupedMatchCount++;
    groupedStatementRowCount += idxs.length;
    for (const si of idxs) {
      usedS[si] = true;
      matched.push({
        date,
        amount: statementRows[si].amount,
        statementRow: statementRows[si],
        ledgerRow: ledgerRows[foundLi],
        matchType: "grouped",
        groupSize: idxs.length,
      });
    }
  }
  return { groupedMatchCount, groupedStatementRowCount };
}

export function matchBankReconcile(statementRows, ledgerRows) {
  const usedS = new Array(statementRows.length).fill(false);
  const usedL = new Array(ledgerRows.length).fill(false);
  const matched = [];

  const exact = passKeyMatch(statementRows, ledgerRows, usedS, usedL, matched, exactKey, "exact");
  const rounded = passKeyMatch(statementRows, ledgerRows, usedS, usedL, matched, roundedKey, "rounded");
  const grouped = passGroupedMatch(statementRows, ledgerRows, usedS, usedL, matched);

  const unmatchedStatement = statementRows.filter((_, i) => !usedS[i]);
  const unmatchedLedger = ledgerRows.filter((_, i) => !usedL[i]);

  const sum = (rows, pick) => rows.reduce((a, r) => a + pick(r), 0);
  return {
    matched,
    unmatchedStatement,
    unmatchedLedger,
    stats: {
      statementCount: statementRows.length,
      statementSum: sum(statementRows, (r) => r.amount),
      ledgerCount: ledgerRows.length,
      ledgerSum: sum(ledgerRows, (r) => r.amount),
      matchedCount: matched.length,
      matchedSum: sum(matched, (r) => r.amount),
      exactMatchCount: matched.filter((m) => m.matchType === "exact").length,
      roundedMatchCount: matched.filter((m) => m.matchType === "rounded").length,
      groupedMatchCount: grouped.groupedMatchCount,
      groupedStatementRowCount: grouped.groupedStatementRowCount,
      unmatchedStatementCount: unmatchedStatement.length,
      unmatchedStatementSum: sum(unmatchedStatement, (r) => r.amount),
      unmatchedLedgerCount: unmatchedLedger.length,
      unmatchedLedgerSum: sum(unmatchedLedger, (r) => r.amount),
      ambiguousKeyCount: exact.ambiguousKeyCount + rounded.ambiguousKeyCount,
      ambiguousMatchCount: exact.ambiguousMatchCount + rounded.ambiguousMatchCount,
    },
  };
}

// --- Output workbook ---------------------------------------------------------
const MATCHED_HEADERS = [
  "Date",
  "Statement Amount",
  "Ledger Amount",
  "Match Type",
  "Group Size",
  "Customer Ref",
  "Ledger Vch No.",
  "Ledger Particulars",
];
const UNMATCHED_STATEMENT_HEADERS = ["Date", "Amount", "Customer Ref"];
const UNMATCHED_LEDGER_HEADERS = ["Date", "Amount", "Vch No.", "Vch Type", "Particulars"];

function matchedToAOA(matched) {
  return matched.map((m) => [
    m.date,
    m.statementRow.amount,
    m.ledgerRow.amount,
    m.matchType,
    m.groupSize || "",
    m.statementRow.customerRef || "",
    m.ledgerRow.vchNo,
    m.ledgerRow.particulars,
  ]);
}

function unmatchedStatementToAOA(rows) {
  return rows.map((r) => [r.date, r.amount, r.customerRef || ""]);
}

function unmatchedLedgerToAOA(rows) {
  return rows.map((r) => [r.date, r.amount, r.vchNo, r.vchType, r.particulars]);
}

const MATCHED_COLS = [
  { wch: 11 },
  { wch: 14 },
  { wch: 14 },
  { wch: 10 },
  { wch: 10 },
  { wch: 22 },
  { wch: 14 },
  { wch: 28 },
];
const UNMATCHED_STATEMENT_COLS = [{ wch: 11 }, { wch: 14 }, { wch: 22 }];
const UNMATCHED_LEDGER_COLS = [{ wch: 11 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 28 }];

function appendSection(wb, label, result) {
  if (!result) return;
  if (result.matched.length > 0) {
    const ws = XLSX.utils.aoa_to_sheet([MATCHED_HEADERS, ...matchedToAOA(result.matched)]);
    ws["!cols"] = MATCHED_COLS;
    XLSX.utils.book_append_sheet(wb, ws, `${label} MATCHED`);
  }
  if (result.unmatchedStatement.length > 0) {
    const ws = XLSX.utils.aoa_to_sheet([
      UNMATCHED_STATEMENT_HEADERS,
      ...unmatchedStatementToAOA(result.unmatchedStatement),
    ]);
    ws["!cols"] = UNMATCHED_STATEMENT_COLS;
    XLSX.utils.book_append_sheet(wb, ws, `${label} UNMATCHED - STATEMENT`);
  }
  if (result.unmatchedLedger.length > 0) {
    const ws = XLSX.utils.aoa_to_sheet([UNMATCHED_LEDGER_HEADERS, ...unmatchedLedgerToAOA(result.unmatchedLedger)]);
    ws["!cols"] = UNMATCHED_LEDGER_COLS;
    XLSX.utils.book_append_sheet(wb, ws, `${label} UNMATCHED - LEDGER`);
  }
}

export function buildBankReconcileWorkbook({ card, upi }) {
  const wb = XLSX.utils.book_new();
  appendSection(wb, "CARD", card);
  appendSection(wb, "UPI", upi);
  return wb;
}
