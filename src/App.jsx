import React, { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

function excelSerialToDDMMYYYY(v) {
  if (v == null || v === "") return "";
  if (typeof v === "string") {
    const trimmed = v.trim();
    // Already in dd/mm/yyyy form
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

function toNumber(v) {
  if (v == null) return 0;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (s === "") return 0;
  const n = parseFloat(s.replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}

function normHeader(h) {
  return String(h ?? "").trim().toLowerCase();
}

const RECEIPT_HEADERS = [
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

const REVIEW_HEADERS = [
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
  const cOpNo = col("op no");
  const cCash = col("cash") ?? col("cash ");
  const cCard = col("card");
  const cNeft = col("neft");
  const cIpCredit = col("ip credit");

  if ([cBillNo, cPatientName, cOpNo, cCash, cCard, cNeft].some((v) => v == null)) {
    throw new Error(
      "One or more required columns (Billnumber, Patient Name, OP NO, Cash, Card, NEFT) were not found in the header row."
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
      summaryMap.set(key, { billName: key, cash: 0, card: 0, neft: 0, count: 0 });
    }
    const entry = summaryMap.get(key);
    entry.cash += c.cash;
    entry.card += c.card;
    entry.neft += c.neft;
    entry.count += 1;
  }
  const billNameSummary = [...summaryMap.values()]
    .map((e) => ({ ...e, total: e.cash + e.card + e.neft }))
    .sort((a, b) => b.total - a.total);

  // Pass 2: find duplicate bill numbers
  const countByBillNo = new Map();
  for (const c of candidates) {
    countByBillNo.set(c.billNo, (countByBillNo.get(c.billNo) || 0) + 1);
  }

  const receiptRows = [];
  const reviewRows = [];
  let skippedZero = 0;
  let paymentCount = 0;
  let receiptCount = 0;
  let duplicateRowCount = 0;

  for (const c of candidates) {
    const isDuplicate = countByBillNo.get(c.billNo) > 1;
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
        "Duplicate Bill Number across upload batch",
      ]);
      continue;
    }

    const party = `${c.patientId} - ${c.patientName}`;
    const paymentAmount = c.cash < 0 ? Math.abs(c.cash) : 0;
    const receiptCash = c.cash > 0 ? c.cash : 0;
    const receiptCard = c.card > 0 ? c.card : 0;
    const receiptNeft = c.neft > 0 ? c.neft : 0;
    const receiptTotal = receiptCash + receiptCard + receiptNeft;

    if (paymentAmount === 0 && receiptTotal === 0) {
      skippedZero++;
      continue;
    }

    if (paymentAmount > 0) {
      paymentCount++;
      // Party debited, Cash credited
      receiptRows.push([
        "Payment",
        c.date,
        c.billNo,
        party,
        paymentAmount,
        "DR",
        null,
        null,
        null,
        null,
        null,
      ]);
      receiptRows.push([
        "Payment",
        c.date,
        c.billNo,
        "CASH",
        paymentAmount,
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
      receiptRows.push([
        "Receipt",
        c.date,
        c.billNo,
        party,
        receiptTotal,
        "CR",
        null,
        null,
        null,
        null,
        null,
      ]);
      if (receiptCash > 0) {
        receiptRows.push([
          "Receipt",
          c.date,
          c.billNo,
          "CASH",
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
          "Receipt",
          c.date,
          c.billNo,
          "CARD",
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
          "Receipt",
          c.date,
          c.billNo,
          "UPI/NEFT",
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
      outputRows: receiptRows.length,
    },
  };
}

function buildOutputWorkbook(receiptRows, reviewRows, billNameSummary) {
  const wb = XLSX.utils.book_new();

  const receiptAoa = [RECEIPT_HEADERS, ...receiptRows];
  const wsReceipt = XLSX.utils.aoa_to_sheet(receiptAoa);
  wsReceipt["!cols"] = [
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
  XLSX.utils.book_append_sheet(wb, wsReceipt, "RECEIPT");

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
  XLSX.utils.book_append_sheet(wb, wsReview, "NEEDS REVIEW");

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

  return wb;
}

function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename);
}

function StatRow({ label, value, mono = true }) {
  return (
    <div className="stat-row">
      <span className="stat-label">{label}</span>
      <span className={mono ? "stat-value mono" : "stat-value"}>{value}</span>
    </div>
  );
}

function Toggle({ checked, onChange, label, description }) {
  return (
    <div className="toggle-row">
      <div className="toggle-copy">
        <span className="toggle-label">{label}</span>
        {description && <span className="toggle-desc">{description}</span>}
      </div>
      <button
        type="button"
        role="switch"
        aria-checked={checked}
        className={`switch${checked ? " on" : ""}`}
        onClick={() => onChange(!checked)}
      >
        <span className="switch-thumb" />
      </button>
    </div>
  );
}

const VOUCHER_TAG_STYLES = {
  Receipt: "tag-receipt",
  Payment: "tag-payment",
};

export default function TejaReceiptMapper() {
  const [fileName, setFileName] = useState(null);
  const [status, setStatus] = useState("idle"); // idle | processing | done | error
  const [error, setError] = useState(null);
  const [result, setResult] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const fileInputRef = useRef(null);

  const handleFile = useCallback((file) => {
    setFileName(file.name);
    setStatus("processing");
    setError(null);
    setResult(null);

    const reader = new FileReader();
    reader.onload = (e) => {
      // slight delay so the processing state actually paints before heavy sync work
      setTimeout(() => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array", cellDates: true });
          const { receiptRows, reviewRows, billNameSummary, stats } = processWorkbook(workbook);
          setResult({ receiptRows, reviewRows, billNameSummary, stats });
          setStatus("done");
        } catch (err) {
          setError(err.message || "Something went wrong while processing this file.");
          setStatus("error");
        }
      }, 30);
    };
    reader.onerror = () => {
      setError("Could not read this file.");
      setStatus("error");
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback(
    (e) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const onFileInput = useCallback(
    (e) => {
      const file = e.target.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const reset = () => {
    setFileName(null);
    setStatus("idle");
    setError(null);
    setResult(null);
    setShowPreview(false);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const downloadReceipt = () => {
    if (!result) return;
    const wb = buildOutputWorkbook(result.receiptRows, result.reviewRows, result.billNameSummary);
    const base = (fileName || "teja_export").replace(/\.[^/.]+$/, "");
    downloadWorkbook(wb, `${base}_TALLY_MAPPED.xlsx`);
  };

  const stats = result?.stats;
  const previewRows = result?.receiptRows?.slice(0, 12) ?? [];

  return (
    <div className="page">
      <style>{`
        :root {
          --navy: #16222f;
          --navy-soft: #1f3245;
          --teal: #2d8b8b;
          --teal-deep: #226e6e;
          --seafoam: #a8dadc;
          --cream: #f1faee;
          --ink: #16222f;
          --ink-muted: #55677a;
          --border: #d7e4e6;
          --card: #ffffff;
          --alert: #b5453b;
          --alert-bg: #fbeeec;
          --alert-border: #e6c3bd;
        }
        * { box-sizing: border-box; }
        .page {
          min-height: 100vh;
          background: var(--cream);
          font-family: 'Inter', sans-serif;
          color: var(--ink);
          padding: 3rem 1.25rem 4rem;
          display: flex;
          justify-content: center;
        }
        .sheet { width: 100%; max-width: 780px; }

        .masthead {
          display: flex;
          align-items: center;
          gap: 0.9rem;
          margin-bottom: 2.25rem;
        }
        .mark {
          width: 40px;
          height: 40px;
          border-radius: 9px;
          background: linear-gradient(160deg, var(--navy), var(--teal-deep));
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
        }
        .mark svg { width: 20px; height: 20px; }
        .eyebrow {
          font-family: 'IBM Plex Mono', monospace;
          font-size: 0.7rem;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          color: var(--teal-deep);
          margin: 0 0 0.3rem;
          font-weight: 500;
        }
        h1 {
          font-family: 'Manrope', sans-serif;
          font-weight: 800;
          font-size: 1.55rem;
          margin: 0;
          letter-spacing: -0.01em;
          color: var(--navy);
        }
        .sub {
          font-size: 0.93rem;
          color: var(--ink-muted);
          max-width: 56ch;
          line-height: 1.55;
          margin: 0.85rem 0 0;
        }

        .card {
          background: var(--card);
          border: 1px solid var(--border);
          border-radius: 12px;
          box-shadow: 0 1px 2px rgba(22,34,47,0.04), 0 8px 24px -18px rgba(22,34,47,0.25);
          margin-bottom: 1.5rem;
          overflow: hidden;
        }
        .card-inner { padding: 1.6rem; }
        .card-head {
          display: flex;
          align-items: center;
          gap: 0.6rem;
          margin-bottom: 1.1rem;
        }
        .step-num {
          width: 22px;
          height: 22px;
          border-radius: 6px;
          background: var(--seafoam);
          color: var(--navy);
          font-family: 'IBM Plex Mono', monospace;
          font-size: 0.72rem;
          font-weight: 600;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
        }
        .section-title {
          font-family: 'Inter', sans-serif;
          font-size: 0.86rem;
          font-weight: 700;
          color: var(--navy);
          margin: 0;
        }

        .dropzone {
          border: 1.5px dashed #b9cdd0;
          border-radius: 10px;
          padding: 2.4rem 1.5rem;
          text-align: center;
          cursor: pointer;
          transition: border-color 0.15s ease, background 0.15s ease;
          background: #fbfdfc;
        }
        .dropzone.drag, .dropzone:hover { border-color: var(--teal); background: #eef7f6; }
        .dropzone-title {
          font-family: 'Manrope', sans-serif;
          font-weight: 700;
          font-size: 1.02rem;
          margin: 0 0 0.3rem;
          color: var(--navy);
        }
        .dropzone-hint {
          font-size: 0.8rem;
          color: var(--ink-muted);
          margin: 0;
          font-family: 'IBM Plex Mono', monospace;
        }
        input[type=file] { display: none; }

        .status-row {
          display: flex;
          align-items: center;
          gap: 0.55rem;
          font-family: 'IBM Plex Mono', monospace;
          font-size: 0.83rem;
          color: var(--ink-muted);
        }
        .dot { width: 8px; height: 8px; border-radius: 50%; background: #c7d3d5; flex-shrink: 0; }
        .dot.processing { background: #c79a3f; animation: pulse 1s infinite ease-in-out; }
        .dot.done { background: var(--teal); }
        .dot.error { background: var(--alert); }
        @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.35; } }

        .stat-row {
          display: flex;
          justify-content: space-between;
          align-items: baseline;
          padding: 0.55rem 0;
          border-bottom: 1px solid #eef2ee;
          font-size: 0.9rem;
        }
        .stat-row:last-child { border-bottom: none; }
        .stat-label { color: var(--ink-muted); }
        .stat-value { font-weight: 700; color: var(--navy); }
        .stat-value.mono { font-family: 'IBM Plex Mono', monospace; }

        .stamp-wrap {
          display: flex;
          align-items: flex-start;
          gap: 0.9rem;
          padding: 0.95rem 1rem;
          background: var(--alert-bg);
          border: 1px solid var(--alert-border);
          border-radius: 9px;
          margin-top: 1.15rem;
        }
        .stamp {
          font-family: 'IBM Plex Mono', monospace;
          font-weight: 600;
          font-size: 0.66rem;
          letter-spacing: 0.06em;
          color: var(--alert);
          border: 1.5px solid var(--alert);
          border-radius: 5px;
          padding: 0.28rem 0.5rem;
          white-space: nowrap;
          flex-shrink: 0;
          text-transform: uppercase;
        }
        .stamp-text { font-size: 0.85rem; color: #7a3129; line-height: 1.5; }

        /* --- Uniform button system --- */
        .btn-row { display: flex; gap: 0.7rem; margin-top: 1.5rem; flex-wrap: wrap; }
        button.btn {
          font-family: 'Inter', sans-serif;
          font-size: 0.87rem;
          font-weight: 700;
          border-radius: 8px;
          padding: 0.68rem 1.15rem;
          cursor: pointer;
          border: 1.5px solid transparent;
          transition: background 0.15s ease, border-color 0.15s ease, color 0.15s ease;
          line-height: 1.2;
        }
        button.btn-primary { background: var(--navy); color: #fff; border-color: var(--navy); }
        button.btn-primary:hover { background: var(--navy-soft); }
        button.btn-ghost { background: transparent; color: var(--ink-muted); border-color: var(--border); }
        button.btn-ghost:hover { border-color: var(--teal); color: var(--teal-deep); }
        button.btn-sm { padding: 0.42rem 0.8rem; font-size: 0.8rem; border-radius: 7px; }

        /* --- Uniform toggle switch --- */
        .toggle-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 1rem;
          padding: 0.2rem 0 0.1rem;
        }
        .toggle-copy { display: flex; flex-direction: column; gap: 0.1rem; }
        .toggle-label { font-size: 0.88rem; font-weight: 600; color: var(--navy); }
        .toggle-desc { font-size: 0.78rem; color: var(--ink-muted); }
        button.switch {
          width: 40px;
          height: 23px;
          border-radius: 12px;
          background: #d3dedf;
          border: none;
          padding: 2px;
          cursor: pointer;
          flex-shrink: 0;
          display: flex;
          align-items: center;
          transition: background 0.18s ease;
        }
        button.switch.on { background: var(--teal); }
        .switch-thumb {
          width: 19px;
          height: 19px;
          border-radius: 50%;
          background: #fff;
          box-shadow: 0 1px 2px rgba(0,0,0,0.18);
          transition: transform 0.18s ease;
        }
        button.switch.on .switch-thumb { transform: translateX(17px); }

        .divider { height: 1px; background: #eef2ee; margin: 1.15rem 0; }

        /* --- Preview table --- */
        .preview-wrap { overflow-x: auto; border: 1px solid var(--border); border-radius: 9px; margin-top: 1rem; }
        table.preview { width: 100%; border-collapse: collapse; font-size: 0.8rem; }
        table.preview th {
          text-align: left;
          font-family: 'IBM Plex Mono', monospace;
          font-size: 0.68rem;
          letter-spacing: 0.05em;
          text-transform: uppercase;
          color: var(--ink-muted);
          background: #f6faf9;
          padding: 0.55rem 0.7rem;
          border-bottom: 1px solid var(--border);
          white-space: nowrap;
        }
        table.preview td {
          padding: 0.5rem 0.7rem;
          border-bottom: 1px solid #f0f4f2;
          white-space: nowrap;
          color: var(--ink);
        }
        table.preview td.mono { font-family: 'IBM Plex Mono', monospace; }
        table.preview tr:last-child td { border-bottom: none; }
        .tag {
          font-family: 'IBM Plex Mono', monospace;
          font-size: 0.68rem;
          font-weight: 600;
          padding: 0.15rem 0.45rem;
          border-radius: 5px;
          text-transform: uppercase;
          letter-spacing: 0.03em;
        }
        .tag-receipt { background: #e4f2f0; color: var(--teal-deep); }
        .tag-payment { background: var(--alert-bg); color: var(--alert); }
        .preview-note {
          font-size: 0.76rem;
          color: var(--ink-muted);
          margin: 0.6rem 0 0;
          font-family: 'IBM Plex Mono', monospace;
        }

        .error-box {
          background: var(--alert-bg);
          border: 1px solid var(--alert-border);
          color: #7a3129;
          padding: 1rem 1.2rem;
          border-radius: 9px;
          font-size: 0.87rem;
          line-height: 1.5;
        }

        .footnote {
          font-size: 0.78rem;
          color: #8b9ba0;
          margin-top: 2rem;
          line-height: 1.65;
          font-family: 'IBM Plex Mono', monospace;
        }
      `}</style>

      <div className="sheet">
        <div className="masthead">
          <div className="mark">
            <svg viewBox="0 0 24 24" fill="none">
              <path d="M4 12l6 6L20 6" stroke="#f1faee" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </div>
          <div>
            <p className="eyebrow">Teja → Tally · Receipt Voucher Mapper</p>
            <h1>Detailed Bill Register Converter</h1>
          </div>
        </div>
        <p className="sub" style={{ marginTop: "-1.4rem", marginBottom: "2rem" }}>
          Upload the Detailed Bill Register export from Teja. This produces a
          ready-to-import RECEIPT mapping sheet, plus a separate review sheet
          for any duplicate bill numbers that would fail on upload into Tally.
        </p>

        <div className="card">
          <div className="card-inner">
            <div className="card-head">
              <span className="step-num">1</span>
              <p className="section-title">Source file</p>
            </div>
            <div
              className={`dropzone${dragOver ? " drag" : ""}`}
              onClick={() => fileInputRef.current?.click()}
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
            >
              <p className="dropzone-title">
                {fileName ? fileName : "Drop the Teja export here"}
              </p>
              <p className="dropzone-hint">
                {fileName ? "Click to choose a different file" : ".xlsx — handles 50,000+ rows"}
              </p>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={onFileInput}
              />
            </div>

            {status !== "idle" && (
              <div style={{ marginTop: "1.1rem" }}>
                <div className="status-row">
                  <span className={`dot ${status}`} />
                  <span>
                    {status === "processing" && "Reading and mapping rows…"}
                    {status === "done" && "Mapping complete"}
                    {status === "error" && "Could not process this file"}
                  </span>
                </div>
              </div>
            )}

            {status === "error" && <div className="error-box" style={{ marginTop: "1rem" }}>{error}</div>}
          </div>
        </div>

        {status === "done" && stats && (
          <div className="card">
            <div className="card-inner">
              <div className="card-head">
                <span className="step-num">2</span>
                <p className="section-title">Summary</p>
              </div>

              <StatRow label="Bill rows scanned" value={stats.scannedDataRows.toLocaleString()} />
              <StatRow label="Bills with a bill number" value={stats.candidateBills.toLocaleString()} />
              <StatRow label="Skipped — no cash/card/NEFT movement" value={stats.skippedZero.toLocaleString()} />
              <StatRow label="Receipt vouchers generated" value={stats.receiptCount.toLocaleString()} />
              <StatRow label="Payment vouchers (negative cash / refunds)" value={stats.paymentCount.toLocaleString()} />
              <StatRow label="Total output rows (RECEIPT sheet)" value={stats.outputRows.toLocaleString()} />

              {stats.duplicateBillNos > 0 && (
                <div className="stamp-wrap">
                  <span className="stamp">Needs Review</span>
                  <span className="stamp-text">
                    {stats.duplicateBillNos} bill number{stats.duplicateBillNos === 1 ? "" : "s"} appear
                    {stats.duplicateBillNos === 1 ? "s" : ""} more than once across this export
                    ({stats.duplicateRowCount} rows total). These were excluded from the RECEIPT
                    sheet and listed on the "NEEDS REVIEW" tab — resolve manually before uploading.
                  </span>
                </div>
              )}

              <div className="divider" />

              <Toggle
                checked={showPreview}
                onChange={setShowPreview}
                label="Preview output rows"
                description="See the first 12 rows before downloading"
              />

              {showPreview && (
                <>
                  <div className="preview-wrap">
                    <table className="preview">
                      <thead>
                        <tr>
                          <th>Voucher Type</th>
                          <th>Date</th>
                          <th>Voucher No.</th>
                          <th>Party Name</th>
                          <th>Amount</th>
                          <th>Dr/Cr</th>
                        </tr>
                      </thead>
                      <tbody>
                        {previewRows.map((row, i) => (
                          <tr key={i}>
                            <td>
                              <span className={`tag ${VOUCHER_TAG_STYLES[row[0]] || ""}`}>{row[0]}</span>
                            </td>
                            <td className="mono">{row[1]}</td>
                            <td className="mono">{row[2]}</td>
                            <td>{row[3]}</td>
                            <td className="mono">{Number(row[4]).toLocaleString()}</td>
                            <td className="mono">{row[5]}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="preview-note">
                    Showing 12 of {stats.outputRows.toLocaleString()} rows
                  </p>
                </>
              )}

              <div className="btn-row">
                <button className="btn btn-primary" onClick={downloadReceipt}>
                  Download mapped workbook (.xlsx)
                </button>
                <button className="btn btn-ghost" onClick={reset}>
                  Start over
                </button>
              </div>
            </div>
          </div>
        )}

        {status === "done" && result?.billNameSummary?.length > 0 && (
          <div className="card">
            <div className="card-inner">
              <div className="card-head">
                <span className="step-num">3</span>
                <p className="section-title">Summary by bill type</p>
              </div>
              <p className="preview-note" style={{ marginTop: 0, marginBottom: "0.9rem" }}>
                Cash / Card / NEFT totals per section, across every bill row —
                a reconciliation check against Teja's own report. Included
                even for duplicate bill numbers.
              </p>
              <div className="preview-wrap">
                <table className="preview">
                  <thead>
                    <tr>
                      <th>Bill Name</th>
                      <th>Cash</th>
                      <th>Card</th>
                      <th>NEFT</th>
                      <th>Total</th>
                      <th>Bills</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.billNameSummary.map((r) => (
                      <tr key={r.billName}>
                        <td>{r.billName}</td>
                        <td className="mono">{r.cash.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                        <td className="mono">{r.card.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                        <td className="mono">{r.neft.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                        <td className="mono" style={{ fontWeight: 700 }}>
                          {r.total.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                        </td>
                        <td className="mono">{r.count.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr>
                      <td style={{ fontWeight: 700 }}>TOTAL</td>
                      <td className="mono" style={{ fontWeight: 700 }}>
                        {result.billNameSummary.reduce((a, r) => a + r.cash, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono" style={{ fontWeight: 700 }}>
                        {result.billNameSummary.reduce((a, r) => a + r.card, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono" style={{ fontWeight: 700 }}>
                        {result.billNameSummary.reduce((a, r) => a + r.neft, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono" style={{ fontWeight: 700 }}>
                        {result.billNameSummary.reduce((a, r) => a + r.total, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono" style={{ fontWeight: 700 }}>
                        {result.billNameSummary.reduce((a, r) => a + r.count, 0).toLocaleString()}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="preview-note">This sheet is also included in the downloaded workbook.</p>
            </div>
          </div>
        )}

        <p className="footnote">
          Rules applied — Receipt: unique bill no., cash/card/NEFT &gt; 0 → Party
          credited, mode(s) debited. Payment: negative cash → party debited, cash
          credited. Negative IP Credit rows are excluded here (handled on the
          SALES side). Duplicate bill numbers are excluded and listed separately.
        </p>
      </div>
    </div>
  );
}
