import React, { useCallback } from "react";
import * as XLSX from "xlsx";
import { processIpCreditWorkbook, buildIpCreditFlaggedWorkbook } from "../lib/ipcredit.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Button, StatusRow, Dropzone } from "./shared.jsx";

const STATUS_LABELS = {
  processing: "Reading the IP Credit report…",
  done: "Parsed",
  error: "Could not process this file",
};

function fmt(n) {
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function IpCreditView({ state, setState }) {
  const { fileName, status, error, result } = state;
  const patch = useCallback((updates) => setState((s) => ({ ...s, ...updates })), [setState]);

  const handleFile = useCallback(
    (file) => {
      patch({ fileName: file.name, status: "processing", error: null, result: null });
      const reader = new FileReader();
      reader.onload = (e) => {
        setTimeout(() => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const r = processIpCreditWorkbook(workbook);
            patch({ result: r, status: "done" });
          } catch (err) {
            patch({ error: err.message || "Something went wrong.", status: "error" });
          }
        }, 30);
      };
      reader.onerror = () => patch({ error: "Could not read this file.", status: "error" });
      reader.readAsArrayBuffer(file);
    },
    [patch]
  );

  const reset = () =>
    patch({ fileName: null, status: "idle", error: null, result: null });

  const downloadFlagged = () => {
    if (!result?.flaggedRows?.length) return;
    const wb = buildIpCreditFlaggedWorkbook(result);
    const base = (fileName || "teja_ipcredit").replace(/\.[^/.]+$/, "");
    downloadWorkbook(wb, `${base}_FLAGGED.xlsx`);
  };

  const stats = result?.stats;

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">IP Credit</p>
        <h1>IP Credit report</h1>
        <p className="lede">
          Upload the Date Wise IP Credit Bill Status report from Teja. It
          feeds the Bills tab — Discharge income is reduced by each patient's
          settled IP credit total so already-booked charges aren't
          double-counted.
        </p>
      </header>

      <section className="card">
        <Dropzone
          onFile={handleFile}
          fileName={fileName}
          emptyTitle="Drop the IP Credit report here"
          emptyHint=".xlsx or .xls"
        />
        <StatusRow status={status} labels={STATUS_LABELS} />
        {status === "error" && <div className="error-box">{error}</div>}
      </section>

      {status === "done" && stats && (
        <section className="card">
          <h2 className="card-title">Summary</h2>
          <StatRow label="Total credit rows scanned" value={stats.dataRows.toLocaleString()} />
          <StatRow label="Rows with a settled bill number" value={stats.settledCount.toLocaleString()} />
          <StatRow label="Unique settled bills" value={stats.uniqueSettledBills.toLocaleString()} />
          <StatRow label="Total credit settled" value={fmt(stats.settledAmount)} />
          <div className="divider" />
          <StatRow label="Rows without a settled bill number" value={stats.unsettledCount.toLocaleString()} />
          <StatRow label="Total credit still open (unsettled)" value={fmt(stats.unsettledAmount)} />

          {stats.flaggedCount > 0 && (
            <div className="note">
              <span className="pill amber">Review</span>
              <p>
                {stats.flaggedCount} row{stats.flaggedCount === 1 ? "" : "s"} (₹{fmt(stats.flaggedAmount)})
                had a Settled Bill No. even though IP Clear was not "Y" — a
                data issue in the Teja export. These are excluded from the
                Discharge credit deduction. Verify in Teja and re-export.
              </p>
            </div>
          )}

          <div className="note">
            <span className="pill teal">Feeds Bills</span>
            <p>
              This report stays loaded across tab switches. Head to the Bills
              tab — every matched Discharge bill will get its IP credit
              deducted automatically.
            </p>
          </div>

          <div className="btn-row">
            <Button variant="ghost" onClick={reset}>
              Start over
            </Button>
          </div>
        </section>
      )}

      {status === "done" && result?.flaggedRows?.length > 0 && (
        <section className="card">
          <h2 className="card-title">Flagged — IP Clear is not "Y" but a Settled Bill No. is present</h2>
          <div className="table-wrap">
            <table className="preview">
              <thead>
                <tr>
                  <th>Bill No</th>
                  <th>Bill Date</th>
                  <th>Patient Name</th>
                  <th>Credit Amt</th>
                  <th>IP Clear</th>
                  <th>Settled Bill No</th>
                  <th>Settled Date</th>
                </tr>
              </thead>
              <tbody>
                {result.flaggedRows.slice(0, 100).map((r, i) => (
                  <tr key={i}>
                    <td>{r.billNo}</td>
                    <td>{r.billDate}</td>
                    <td>{r.patientName}</td>
                    <td className="mono">{fmt(r.creditAmount)}</td>
                    <td>{r.ipClear}</td>
                    <td>{r.settledBillNo}</td>
                    <td>{r.settledDate}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {result.flaggedRows.length > 100 && (
            <p className="preview-note">
              Showing the first 100 of {result.flaggedRows.length.toLocaleString()} flagged rows.
            </p>
          )}
          <div className="btn-row">
            <Button variant="ghost" onClick={downloadFlagged}>
              Download flagged rows
            </Button>
          </div>
        </section>
      )}
    </div>
  );
}
