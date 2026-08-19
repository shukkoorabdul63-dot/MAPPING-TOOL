import * as XLSX from "xlsx";
import { toNumber, excelSerialToDDMMYYYY } from "./utils.js";

// Report structure (from the "Date Wise IPCredit Bill Status Report"):
//   row 0..3: title / date / metadata
//   row 4:    header — Sl.No, Bill No, Bill Date, IP No, Patient Name,
//             Credit Amt, IP Clear, Settled BillNo, Settled Date
//             (columns are separated by empty spacer columns as of the
//             sample file, so we probe the row for known labels rather than
//             hard-code offsets)
//   row 5+:   bill-name section headers (single non-empty cell in col 0)
//             interleaved with data rows.

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const row = rows[i] || [];
    const norm = row.map((v) => (v == null ? "" : String(v).trim().toLowerCase()));
    if (norm.includes("bill no") && norm.includes("credit amt") && norm.some((v) => v.startsWith("settled"))) {
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
  // Try prefix match for e.g. "settled billno" / "settledbillno"
  for (const c of candidates) {
    const i = cells.findIndex((cell) => cell.replace(/\s+/g, "").startsWith(c.replace(/\s+/g, "")));
    if (i !== -1) return i;
  }
  return -1;
}

export function processIpCreditWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const { idx: headerRowIdx, cells } = findHeaderRow(rows);
  if (headerRowIdx === -1) {
    throw new Error(
      "Could not find the header row of the IP Credit report (looking for 'Bill No', 'Credit Amt', 'Settled BillNo'). Check the file."
    );
  }

  const cBillNo = findCol(cells, "bill no", "billno");
  const cBillDate = findCol(cells, "bill date", "billdate");
  const cIpNo = findCol(cells, "ip no", "ipno");
  const cPatient = findCol(cells, "patient name", "patientname");
  const cCredit = findCol(cells, "credit amt", "creditamt", "credit amount");
  const cIpClear = findCol(cells, "ip clear", "ipclear");
  const cSettled = findCol(cells, "settled billno", "settled bill no", "settledbillno");
  const cSettledDate = findCol(cells, "settled date", "settleddate");

  const missing = [];
  if (cBillNo === -1) missing.push("Bill No");
  if (cCredit === -1) missing.push("Credit Amt");
  if (cSettled === -1) missing.push("Settled BillNo");
  if (missing.length) {
    throw new Error(
      `Could not find the following required column(s) in the IP Credit report: ${missing.join(", ")}.`
    );
  }

  // Data rows: skip pure section-header rows (a single non-empty cell in col 0
  // that isn't a serial number). Any row with a non-empty Bill No is data.
  const creditBySettled = new Map();
  const allRows = [];
  const flaggedRows = [];
  let dataRows = 0;
  let unsettledCount = 0;
  let unsettledAmount = 0;
  let flaggedAmount = 0;
  let currentSection = "";

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const c0 = row[0];
    const billNo = row[cBillNo];

    // Section header: a single string in col 0, all other cells blank/null.
    if (typeof c0 === "string" && (billNo == null || billNo === "")) {
      const rest = row.slice(1).filter((v) => v != null && v !== "");
      if (rest.length === 0) {
        currentSection = c0.trim();
        continue;
      }
    }

    if (billNo == null || String(billNo).trim() === "") continue;
    dataRows++;

    const rowData = {
      section: currentSection,
      billNo: String(billNo).trim(),
      billDate: cBillDate !== -1 ? excelSerialToDDMMYYYY(row[cBillDate]) : "",
      ipNo: cIpNo !== -1 && row[cIpNo] != null ? String(row[cIpNo]).trim() : "",
      patientName: cPatient !== -1 && row[cPatient] != null ? String(row[cPatient]).trim() : "",
      creditAmount: toNumber(row[cCredit]),
      ipClear: cIpClear !== -1 && row[cIpClear] != null ? String(row[cIpClear]).trim().toUpperCase() : "",
      settledBillNo: cSettled !== -1 && row[cSettled] != null ? String(row[cSettled]).trim() : "",
      settledDate: cSettledDate !== -1 ? excelSerialToDDMMYYYY(row[cSettledDate]) : "",
    };

    allRows.push(rowData);

    if (rowData.settledBillNo) {
      // Teja's export sometimes carries a Settled Bill No. on a row whose
      // IP Clear flag is still "N" (the bill got referenced before the
      // credit was actually cleared). Only a "Y" row is a real settlement —
      // anything else is a data anomaly that must NOT reduce a Discharge
      // bill's income, so it's flagged for review instead of booked.
      if (rowData.ipClear === "Y") {
        creditBySettled.set(
          rowData.settledBillNo,
          (creditBySettled.get(rowData.settledBillNo) || 0) + rowData.creditAmount
        );
      } else {
        flaggedRows.push(rowData);
        flaggedAmount += rowData.creditAmount;
      }
    } else {
      unsettledCount++;
      unsettledAmount += rowData.creditAmount;
    }
  }

  const settledTotal = [...creditBySettled.values()].reduce((a, b) => a + b, 0);

  return {
    creditBySettled,
    allRows,
    flaggedRows,
    stats: {
      dataRows,
      settledCount: dataRows - unsettledCount - flaggedRows.length,
      settledAmount: settledTotal,
      unsettledCount,
      unsettledAmount,
      flaggedCount: flaggedRows.length,
      flaggedAmount,
      uniqueSettledBills: creditBySettled.size,
    },
  };
}
