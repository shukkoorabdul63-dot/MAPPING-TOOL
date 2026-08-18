import React, { useCallback } from "react";
import * as XLSX from "xlsx";
import {
  processDialysisWorkbook,
  buildDialysisOutputWorkbook,
} from "../lib/dialysis.js";
import { resolveRows } from "../lib/tokens.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Button, StatusRow, Dropzone } from "./shared.jsx";

const STATUS_LABELS = {
  processing: "Reading and mapping rows…",
  done: "Mapping complete",
  error: "Could not process this file",
};

function fmt(n) {
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function DialysisView({ state, setState, ledgerNames }) {
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
            const r = processDialysisWorkbook(workbook);
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

  const download = () => {
    if (!result) return;
    const wb = buildDialysisOutputWorkbook(result, ledgerNames);
    const base = (fileName || "teja_dialysis").replace(/\.[^/.]+$/, "");
    downloadWorkbook(wb, `${base}_TALLY_MAPPED.xlsx`);
  };

  const stats = result?.stats;
  const previewRows = result ? resolveRows(result.salesRows.slice(0, 12), ledgerNames) : [];

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">Dialysis</p>
        <h1>Dialysis receivables</h1>
        <p className="lede">
          Upload the Suspense Outstanding Report from Teja. The DIALYSIS
          section becomes a Journal-Dialysis batch — Company Dr, Patient Cr —
          for booking receivables from schemes like KASP and MEDISEP.
        </p>
      </header>

      <section className="card">
        <Dropzone
          onFile={handleFile}
          fileName={fileName}
          emptyTitle="Drop the dialysis receivables report here"
          emptyHint=".xlsx or .xls"
        />
        <StatusRow status={status} labels={STATUS_LABELS} />
        {status === "error" && <div className="error-box">{error}</div>}
      </section>

      {status === "done" && stats && (
        <>
          <section className="card">
            <h2 className="card-title">Summary</h2>
            <StatRow label="Dialysis rows scanned" value={stats.dialysisRowsScanned.toLocaleString()} />
            <StatRow label="Booked as receivable" value={stats.bookedRows.toLocaleString()} />
            <StatRow label="Skipped — zero pending" value={stats.skippedZeroPending.toLocaleString()} />
            <StatRow label="Skipped — no company (needs review)" value={stats.skippedNoCompany.toLocaleString()} />
            <StatRow label="Total receivable booked" value={fmt(stats.totalReceivable)} />
            <StatRow
              label="Output rows (per voucher pair × 2)"
              value={`${stats.outputRows.toLocaleString()}${stats.sheetCount > 1 ? ` — across ${stats.sheetCount} sheets` : ""}`}
            />

            {stats.sheetCount > 1 && (
              <div className="note">
                <span className="pill teal">Auto-resolved</span>
                <p>
                  Some bill numbers repeat. Each repeat goes to its own extra sheet — "DIALYSIS
                  RECEIVABLES (2)" and so on — so upload each sheet as a separate Tally import and
                  voucher numbers never collide.
                </p>
              </div>
            )}

            {stats.skippedNoCompany > 0 && (
              <div className="note">
                <span className="pill amber">Review</span>
                <p>
                  {stats.skippedNoCompany} row{stats.skippedNoCompany === 1 ? "" : "s"} had no
                  Company assigned and could not be booked as a receivable. Full list on the
                  REVIEW sheet — fix in Teja and re-export to capture them.
                </p>
              </div>
            )}

            <div className="btn-row">
              <Button variant="primary" onClick={download}>
                Download mapped workbook
              </Button>
              <Button variant="ghost" onClick={reset}>
                Start over
              </Button>
            </div>
          </section>

          {result.companySummary.length > 0 && (
            <section className="card">
              <h2 className="card-title">By company</h2>
              <div className="table-wrap">
                <table className="preview">
                  <thead>
                    <tr>
                      <th>Company</th>
                      <th className="ta-r">Bills</th>
                      <th className="ta-r">Total Pending</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.companySummary.map((r) => (
                      <tr key={r.company}>
                        <td>{r.company}</td>
                        <td className="mono ta-r">{r.count.toLocaleString()}</td>
                        <td className="mono ta-r">{fmt(r.pending)}</td>
                      </tr>
                    ))}
                    <tr className="total-row">
                      <td>TOTAL</td>
                      <td className="mono ta-r">
                        {result.companySummary.reduce((a, r) => a + r.count, 0).toLocaleString()}
                      </td>
                      <td className="mono ta-r">
                        {fmt(result.companySummary.reduce((a, r) => a + r.pending, 0))}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {previewRows.length > 0 && (
            <section className="card">
              <h2 className="card-title">Preview</h2>
              <div className="table-wrap">
                <table className="preview">
                  <thead>
                    <tr>
                      <th>Voucher Type</th>
                      <th>Date</th>
                      <th>Voucher No.</th>
                      <th>Party</th>
                      <th className="ta-r">Amount</th>
                      <th>Dr/Cr</th>
                      <th>Bill Ref</th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.map((row, i) => (
                      <tr key={i}>
                        <td><span className="pill navy">{row[0]}</span></td>
                        <td className="mono">{row[1]}</td>
                        <td className="mono">{row[2]}</td>
                        <td>{row[3]}</td>
                        <td className="mono ta-r">{fmt(row[4])}</td>
                        <td className="mono">{row[5]}</td>
                        <td className="mono">{row[6]}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <p className="preview-note">
                Sample rows only — {stats.outputRows.toLocaleString()} total in the download.
              </p>
            </section>
          )}
        </>
      )}
    </div>
  );
}
