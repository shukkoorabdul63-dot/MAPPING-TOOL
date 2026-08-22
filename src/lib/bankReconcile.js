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
// Four passes, each only considering rows the previous pass left unmatched:
//
//   1. Exact:   same date, same amount to the paisa.
//   2. Rounded: same date, amounts equal once both sides are rounded to the
//      nearest whole rupee. Ledger exports commonly carry paise (e.g.
//      203.65) that the original bank/UPI statement never shows (it's
//      already a rounded whole-rupee figure for the same transaction), so a
//      paisa-exact comparison alone would falsely report these as
//      unmatched.
//   3. Grouped (statement side): a customer sometimes pays in two (or more)
//      separate transactions the same day, but Tally books ONE combined
//      receipt for the total. Remaining unmatched statement rows are
//      grouped by (date, customerRef) — Card Number for Card, Customer_VPA
//      for UPI — summed, and that sum is matched (exact or rounded) against
//      a remaining unmatched ledger row on the same date.
//   4. Grouped (ledger side): the mirror case — one card swipe/UPI payment
//      that Tally booked as two (or more) separate receipts for the same
//      patient, e.g. part-payments entered as distinct vouchers. Remaining
//      unmatched ledger rows are grouped by (date, particulars) — the
//      ledger has no card/VPA column, so the patient ID+name in
//      "Particulars" is the proxy for "same customer" — summed, and matched
//      against a remaining unmatched statement row on the same date.
//
// Within a single pass, a (date, amount) key with more than one row on BOTH
// sides is a genuine duplicate (e.g. two ₹300 swipes one day) — still
// matched (pairing in original file order), just counted as "ambiguous" so
// the UI can surface one aggregate note instead of flagging individual rows.

function exactKey(date, amount) {
  return `${date}|${Math.round(amount * 100)}`;
}

// Tolerance for "rounded" matches. Under a rupee is treated as noise —
// covers all common bank rounding styles (round-half-up, round-half-down,
// truncate/floor) plus 0.99-paise oddities where the statement shows a
// whole rupee and the ledger carries the fractional value.
const ROUND_TOL = 1.0;

// Widened "rounded" match: greedy file-order pairing within
// |a - b| < ROUND_TOL. Runs O(n*m) per date. The prior Map-keyed
// approach (Math.round(a) === Math.round(b)) missed cross-integer
// cases like ₹289 (stmt) vs ₹289.54 (ledger, rounds to ₹290) even
// though the diff is only 0.54.
function passRoundedMatch(statementRows, ledgerRows, usedS, usedL, matched) {
  const stmtByDate = new Map();
  statementRows.forEach((r, i) => {
    if (usedS[i]) return;
    if (!stmtByDate.has(r.date)) stmtByDate.set(r.date, []);
    stmtByDate.get(r.date).push(i);
  });
  const ledgerByDate = new Map();
  ledgerRows.forEach((r, i) => {
    if (usedL[i]) return;
    if (!ledgerByDate.has(r.date)) ledgerByDate.set(r.date, []);
    ledgerByDate.get(r.date).push(i);
  });

  for (const [date, sIdxs] of stmtByDate) {
    const lIdxs = ledgerByDate.get(date);
    if (!lIdxs || lIdxs.length === 0) continue;
    for (const si of sIdxs) {
      if (usedS[si]) continue;
      const sAmt = statementRows[si].amount;
      const foundLi = lIdxs.find(
        (li) => !usedL[li] && Math.abs(ledgerRows[li].amount - sAmt) < ROUND_TOL
      );
      if (foundLi == null) continue;
      usedS[si] = true;
      usedL[foundLi] = true;
      matched.push({
        date,
        amount: sAmt,
        statementRow: statementRows[si],
        ledgerRow: ledgerRows[foundLi],
        matchType: "rounded",
      });
    }
  }
  // Ambiguity stat doesn't map cleanly to nested-scan pairing; the exact
  // pass still contributes its (duplicated exact key) count.
  return { ambiguousKeyCount: 0, ambiguousMatchCount: 0 };
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

// Multiple statement rows (same day, same customer) sum to ONE ledger
// receipt — e.g. a customer swiped their card twice and Tally booked one
// combined receipt for the total.
function passGroupedMatchStatement(statementRows, ledgerRows, usedS, usedL, matched) {
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

// The mirror case: ONE statement row (one card swipe / one UPI payment)
// sums against MULTIPLE ledger rows for the same patient — e.g. Tally
// booked two part-payments as separate receipts for one combined swipe.
// Grouped by (date, particulars) since the ledger side has no card/VPA
// column — "particulars" (the patient ID + name) is the closest proxy for
// "same customer" on that side.
function passGroupedMatchLedger(statementRows, ledgerRows, usedS, usedL, matched) {
  const groups = new Map(); // "date||particulars" -> ledger row indices
  ledgerRows.forEach((r, i) => {
    if (usedL[i]) return;
    const particulars = (r.particulars || "").trim();
    if (!particulars) return;
    const key = `${r.date}||${particulars}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(i);
  });

  const statementByDate = new Map();
  statementRows.forEach((r, i) => {
    if (usedS[i]) return;
    if (!statementByDate.has(r.date)) statementByDate.set(r.date, []);
    statementByDate.get(r.date).push(i);
  });

  let groupedMatchCount = 0;
  let groupedLedgerRowCount = 0;
  for (const idxs of groups.values()) {
    if (idxs.length < 2) continue; // only genuine multi-receipt groups
    const date = ledgerRows[idxs[0]].date;
    const sum = idxs.reduce((a, i) => a + ledgerRows[i].amount, 0);
    const candidates = statementByDate.get(date) || [];
    const foundSi = candidates.find(
      (si) => !usedS[si] && (Math.abs(statementRows[si].amount - sum) < 0.01 || Math.round(statementRows[si].amount) === Math.round(sum))
    );
    if (foundSi == null) continue;

    usedS[foundSi] = true;
    groupedMatchCount++;
    groupedLedgerRowCount += idxs.length;
    for (const li of idxs) {
      usedL[li] = true;
      matched.push({
        date,
        amount: ledgerRows[li].amount,
        statementRow: statementRows[foundSi],
        ledgerRow: ledgerRows[li],
        matchType: "grouped",
        groupSize: idxs.length,
      });
    }
  }
  return { groupedMatchCount, groupedLedgerRowCount };
}

// --- Possible matches (no shared identity, needs manual review) -----------
// After all four passes above, some same-day amounts can still coincide
// purely by chance — e.g. two unrelated patients' receipts summing to a
// third, unrelated card swipe, with no shared card/VPA or patient name to
// justify it. These are surfaced as *suggestions* only, never folded into
// the reconciled Matched total or subtracted from Unmatched: with repeated
// round amounts (100/200/300…) common in hospital billing, a blind "any two
// rows that sum to X" search often finds more than one plausible-looking
// pairing for the same target (confirmed against real data — two separate
// unmatched ₹600 statement rows, two separate ledger pairs that both also
// sum to ₹600), and silently guessing which pairing is real risks
// misattributing a payment. So every candidate pairing is listed — capped
// at 2-row combinations, not full subset-sum, to keep the search sane — and
// a tied target (more than one candidate pairing) is left for a human to
// resolve with the Particulars/Card Ref columns, not decided here.
//
// `comboRows`/`comboUsed` is the side searched in 2- and 3-row
// combinations; `singleRows`/`singleUsed` is the side each combo's sum
// is checked against. `pairSide` (kept for backwards compat) labels
// which side that is, for the caller to render/export both directions
// (ledger combos -> one statement row, and the mirror) through the
// same shape.
//
// Enumerates PAIRS (2-row) and TRIPLETS (3-row) — capped at 3 to keep
// the search tractable and false-positive risk in check, and because a
// customer's single card/UPI payment split into 4+ separate ledger
// receipts (or the mirror) doesn't happen in practice.
function findGroupSuggestions(comboRows, comboUsed, singleRows, singleUsed, pairSide) {
  const byDate = new Map();
  comboRows.forEach((r, i) => {
    if (comboUsed[i]) return;
    if (!byDate.has(r.date)) byDate.set(r.date, []);
    byDate.get(r.date).push(i);
  });

  const singleByDate = new Map();
  singleRows.forEach((r, i) => {
    if (singleUsed[i]) return;
    if (!singleByDate.has(r.date)) singleByDate.set(r.date, []);
    singleByDate.get(r.date).push(i);
  });

  const suggestions = [];
  const classify = (target, sum) => {
    if (Math.abs(target - sum) < 0.01) return "exact";
    if (Math.abs(target - sum) < ROUND_TOL) return "rounded";
    return null;
  };
  const push = (date, idxs, si, sum, basis) => {
    suggestions.push({
      pairSide,
      date,
      comboRows: idxs.map((i) => comboRows[i]),
      comboIdxs: idxs,
      singleRow: singleRows[si],
      singleIdx: si,
      sum,
      matchBasis: basis,
    });
  };

  for (const [date, idxs] of byDate) {
    const singleIdxs = singleByDate.get(date);
    if (!singleIdxs || singleIdxs.length === 0) continue;
    // Pairs
    for (let a = 0; a < idxs.length; a++) {
      for (let b = a + 1; b < idxs.length; b++) {
        const sum2 = comboRows[idxs[a]].amount + comboRows[idxs[b]].amount;
        for (const si of singleIdxs) {
          const basis = classify(singleRows[si].amount, sum2);
          if (basis) push(date, [idxs[a], idxs[b]], si, sum2, basis);
        }
      }
    }
    // Triplets
    for (let a = 0; a < idxs.length; a++) {
      for (let b = a + 1; b < idxs.length; b++) {
        for (let c = b + 1; c < idxs.length; c++) {
          const sum3 = comboRows[idxs[a]].amount + comboRows[idxs[b]].amount + comboRows[idxs[c]].amount;
          for (const si of singleIdxs) {
            const basis = classify(singleRows[si].amount, sum3);
            if (basis) push(date, [idxs[a], idxs[b], idxs[c]], si, sum3, basis);
          }
        }
      }
    }
  }
  return suggestions;
}

export function matchBankReconcile(statementRows, ledgerRows) {
  const usedS = new Array(statementRows.length).fill(false);
  const usedL = new Array(ledgerRows.length).fill(false);
  const matched = [];

  const exact = passKeyMatch(statementRows, ledgerRows, usedS, usedL, matched, exactKey, "exact");
  const rounded = passRoundedMatch(statementRows, ledgerRows, usedS, usedL, matched);
  const groupedStatement = passGroupedMatchStatement(statementRows, ledgerRows, usedS, usedL, matched);
  const groupedLedger = passGroupedMatchLedger(statementRows, ledgerRows, usedS, usedL, matched);

  // Enumerate all identity-free pair-sum suggestions. These are the raw
  // candidates from which unambiguous ones get promoted to Matched
  // (below) and the rest stay in Possible Matches for human review.
  const rawSuggestions = [
    ...findGroupSuggestions(ledgerRows, usedL, statementRows, usedS, "ledger"),
    ...findGroupSuggestions(statementRows, usedS, ledgerRows, usedL, "statement"),
  ];

  // A suggestion is unambiguous iff:
  //   - only one pair on the "pair side" sums to this target row, AND
  //   - this pair does not also sum to any other target on the "single
  //     side" (a shared pair could belong to either target — can't say).
  // Both counts must equal 1. Everything else stays as a Possible Match
  // — no silent guessing on genuine ambiguity (e.g. the ₹600 tied case).
  //
  // Target key = the target row's object reference (rows are constructed
  // once by parseLedgerExport/parseCardStatement/parseUpiStatement and
  // never re-created, so reference identity is stable within one match
  // run — same shape the identity-based passes already rely on).
  //
  // Pair key = pairSide + a canonical index-pair on that side. Two
  // suggestions share a pair iff they name the same two rows on the
  // same side.
  const targetCount = new Map();
  const comboCount = new Map();
  const comboKey = (s) => `${s.pairSide}|${[...s.comboIdxs].sort((a, b) => a - b).join("-")}`;
  for (const s of rawSuggestions) {
    targetCount.set(s.singleRow, (targetCount.get(s.singleRow) || 0) + 1);
    const pk = comboKey(s);
    comboCount.set(pk, (comboCount.get(pk) || 0) + 1);
  }

  const possibleMatches = [];
  let pairedMatchCount = 0;
  for (const s of rawSuggestions) {
    const unambiguous = targetCount.get(s.singleRow) === 1 && comboCount.get(comboKey(s)) === 1;
    if (!unambiguous) {
      possibleMatches.push(s);
      continue;
    }
    // Flip the used bits and push one matched entry per combo row using
    // the same shape the identity-based grouping passes produce, so
    // downstream stat aggregation and rendering stays consistent. The
    // pair side determines which array holds the combo vs the single.
    // groupSize is 2 for pairs or 3 for triplets.
    const size = s.comboIdxs.length;
    if (s.pairSide === "ledger") {
      usedS[s.singleIdx] = true;
      for (const li of s.comboIdxs) usedL[li] = true;
      for (const li of s.comboIdxs) {
        matched.push({
          date: s.date,
          amount: ledgerRows[li].amount,
          statementRow: statementRows[s.singleIdx],
          ledgerRow: ledgerRows[li],
          matchType: "paired",
          groupSize: size,
        });
      }
    } else {
      usedL[s.singleIdx] = true;
      for (const si of s.comboIdxs) usedS[si] = true;
      for (const si of s.comboIdxs) {
        matched.push({
          date: s.date,
          amount: statementRows[si].amount,
          statementRow: statementRows[si],
          ledgerRow: ledgerRows[s.singleIdx],
          matchType: "paired",
          groupSize: size,
        });
      }
    }
    pairedMatchCount++;
  }

  const unmatchedStatement = statementRows.filter((_, i) => !usedS[i]);
  const unmatchedLedger = ledgerRows.filter((_, i) => !usedL[i]);

  // Ambiguity stat now runs on the surviving (non-promoted) suggestions
  // only. Every remaining tied-target entry represents a genuine "more
  // than one plausible pairing" case the human needs to break.
  const possibleMatchTargetCounts = new Map();
  for (const s of possibleMatches) {
    possibleMatchTargetCounts.set(s.singleRow, (possibleMatchTargetCounts.get(s.singleRow) || 0) + 1);
  }
  const possibleMatchTiedTargetCount = [...possibleMatchTargetCounts.values()].filter((c) => c > 1).length;

  // Group remaining unmatched rows by (date, amount). The point of the
  // whole reconciliation is to explain the statement-vs-ledger total
  // gap; per-row unmatched lists are useful for spot-checking, but the
  // gap-explaining view is "on this date, this amount had X statement
  // rows and Y ledger rows, and here are the customers on each side."
  // Date is part of the key so a run covering multiple dates doesn't
  // silently collapse rows from different dates under one amount — each
  // day's imbalance stays visible on its own row.
  const groupKey = (date, amount) => `${date}|${Math.round(amount * 100)}`;
  const unmatchedByAmountMap = new Map();
  const getEntry = (date, amount) => {
    const key = groupKey(date, amount);
    let e = unmatchedByAmountMap.get(key);
    if (!e) {
      e = { date, amount, statementRows: [], ledgerRows: [], netCount: 0, netAmount: 0 };
      unmatchedByAmountMap.set(key, e);
    }
    return e;
  };
  for (const r of unmatchedStatement) getEntry(r.date, r.amount).statementRows.push(r);
  for (const r of unmatchedLedger) getEntry(r.date, r.amount).ledgerRows.push(r);
  for (const e of unmatchedByAmountMap.values()) {
    e.netCount = e.statementRows.length - e.ledgerRows.length;
    e.netAmount = e.netCount * e.amount;
  }
  const unmatchedByAmount = [...unmatchedByAmountMap.values()].sort(
    (x, y) => Math.abs(y.netAmount) - Math.abs(x.netAmount)
  );

  const sum = (rows, pick) => rows.reduce((a, r) => a + pick(r), 0);
  return {
    matched,
    unmatchedStatement,
    unmatchedLedger,
    possibleMatches,
    unmatchedByAmount,
    stats: {
      statementCount: statementRows.length,
      statementSum: sum(statementRows, (r) => r.amount),
      ledgerCount: ledgerRows.length,
      ledgerSum: sum(ledgerRows, (r) => r.amount),
      matchedCount: matched.length,
      matchedSum: sum(matched, (r) => r.amount),
      exactMatchCount: matched.filter((m) => m.matchType === "exact").length,
      roundedMatchCount: matched.filter((m) => m.matchType === "rounded").length,
      groupedMatchCount: groupedStatement.groupedMatchCount + groupedLedger.groupedMatchCount,
      groupedStatementRowCount: groupedStatement.groupedStatementRowCount,
      groupedLedgerRowCount: groupedLedger.groupedLedgerRowCount,
      pairedMatchCount,
      unmatchedStatementCount: unmatchedStatement.length,
      unmatchedStatementSum: sum(unmatchedStatement, (r) => r.amount),
      unmatchedLedgerCount: unmatchedLedger.length,
      unmatchedLedgerSum: sum(unmatchedLedger, (r) => r.amount),
      ambiguousKeyCount: exact.ambiguousKeyCount + rounded.ambiguousKeyCount,
      ambiguousMatchCount: exact.ambiguousMatchCount + rounded.ambiguousMatchCount,
      possibleMatchCount: possibleMatches.length,
      possibleMatchTiedTargetCount,
      unmatchedByAmountCount: unmatchedByAmount.length,
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
const POSSIBLE_HEADERS = [
  "Date",
  "Statement Amount(s)",
  "Statement Ref(s)",
  "Ledger Amount(s)",
  "Ledger Vch/Particulars",
  "Sum",
  "Basis",
];
const UNMATCHED_BY_AMOUNT_HEADERS = [
  "Date",
  "Amount",
  "Statement Count",
  "Ledger Count",
  "Net Count",
  "Net Rupees",
  "Statement Customer Refs",
  "Ledger Vch / Particulars",
];

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

function unmatchedByAmountToAOA(rows) {
  return rows.map((e) => [
    e.date,
    e.amount,
    e.statementRows.length,
    e.ledgerRows.length,
    e.netCount,
    e.netAmount,
    e.statementRows.map((r) => r.customerRef || "").join("; "),
    e.ledgerRows.map((r) => `${r.vchNo}: ${r.particulars}`).join("; "),
  ]);
}

function possibleMatchesToAOA(rows) {
  return rows.map((s) => {
    const stmtRows = s.pairSide === "statement" ? s.comboRows : [s.singleRow];
    const ledgerRowsArr = s.pairSide === "ledger" ? s.comboRows : [s.singleRow];
    return [
      s.date,
      stmtRows.map((r) => r.amount).join(" + "),
      stmtRows.map((r) => r.customerRef || "").join(" / "),
      ledgerRowsArr.map((r) => r.amount).join(" + "),
      ledgerRowsArr.map((r) => `${r.vchNo}: ${r.particulars}`).join(" / "),
      s.sum,
      s.matchBasis,
    ];
  });
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
const POSSIBLE_COLS = [{ wch: 11 }, { wch: 16 }, { wch: 24 }, { wch: 16 }, { wch: 36 }, { wch: 12 }, { wch: 10 }];
const UNMATCHED_BY_AMOUNT_COLS = [{ wch: 11 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 48 }, { wch: 60 }];

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
  if (result.possibleMatches && result.possibleMatches.length > 0) {
    const ws = XLSX.utils.aoa_to_sheet([POSSIBLE_HEADERS, ...possibleMatchesToAOA(result.possibleMatches)]);
    ws["!cols"] = POSSIBLE_COLS;
    XLSX.utils.book_append_sheet(wb, ws, `${label} POSSIBLE MATCHES`);
  }
  if (result.unmatchedByAmount && result.unmatchedByAmount.length > 0) {
    const ws = XLSX.utils.aoa_to_sheet([
      UNMATCHED_BY_AMOUNT_HEADERS,
      ...unmatchedByAmountToAOA(result.unmatchedByAmount),
    ]);
    ws["!cols"] = UNMATCHED_BY_AMOUNT_COLS;
    XLSX.utils.book_append_sheet(wb, ws, `${label} UNMATCHED BY AMOUNT`);
  }
}

export function buildBankReconcileWorkbook({ card, upi }) {
  const wb = XLSX.utils.book_new();
  appendSection(wb, "CARD", card);
  appendSection(wb, "UPI", upi);
  return wb;
}
