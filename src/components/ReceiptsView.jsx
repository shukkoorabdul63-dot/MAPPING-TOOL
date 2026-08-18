import React, { useCallback } from "react";
import * as XLSX from "xlsx";
import { processWorkbook, buildOutputWorkbook } from "../lib/receipts.js";
import { resolveRows } from "../lib/tokens.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Toggle, Button, StatusRow, Dropzone } from "./shared.jsx";

const VOUCHER_TAG_STYLES = {
  Receipt: "tag-teal",
  Payment: "tag-red",
  "Credit Note": "tag-amber",
};

const STATUS_LABELS = {
  processing: "Reading and mapping rows…",
  done: "Mapping complete",
  error: "Could not process this file",
};

export function ReceiptsView({ ledgerNames, state, setState }) {
  const { fileName, status, error, result, showPreview } = state;
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
            const r = processWorkbook(workbook);
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
    patch({ fileName: null, status: "idle", error: null, result: null, showPreview: false });

  const download = () => {
    if (!result) return;
    const wb = buildOutputWorkbook(result, ledgerNames);
    const base = (fileName || "teja_receipts").replace(/\.[^/.]+$/, "");
    downloadWorkbook(wb, `${base}_TALLY_MAPPED.xlsx`);
  };

  const stats = result?.stats;
  const previewRows = result
    ? resolveRows(
        [
          ...result.receiptRows.slice(0, 6),
          ...result.paymentRows.slice(0, 4),
          ...result.creditNoteRows.slice(0, 4),
        ],
        ledgerNames
      )
    : [];

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">Receipts</p>
        <h1>Receipt vouchers</h1>
        <p className="lede">
          Upload the Detailed Bill Register export from Teja. Produces
          Tally-ready RECEIPT, PAYMENT, and CREDIT NOTE sheets, plus a
          bill-type summary and a duplicate log.
        </p>
      </header>

      <section className="card">
        <Dropzone
          onFile={handleFile}
          fileName={fileName}
          emptyTitle="Drop the receipts export here"
          emptyHint=".xlsx or .xls — handles 50,000+ rows"
        />
        <StatusRow status={status} labels={STATUS_LABELS} />
        {status === "error" && <div className="error-box">{error}</div>}
      </section>

      {status === "done" && stats && (
        <>
          <section className="card">
            <h2 className="card-title">Summary</h2>
            <StatRow label="Bill rows scanned" value={stats.scannedDataRows.toLocaleString()} />
            <StatRow label="Bills with a bill number" value={stats.candidateBills.toLocaleString()} />
            <StatRow label="Skipped — no cash/card/NEFT movement" value={stats.skippedZero.toLocaleString()} />
            <StatRow label="Receipt vouchers" value={`${stats.receiptCount.toLocaleString()}${stats.receiptSheetCount > 1 ? ` — across ${stats.receiptSheetCount} sheets` : ""}`} />
            <StatRow label="Payment vouchers (cash refunds)" value={`${stats.paymentCount.toLocaleString()}${stats.paymentSheetCount > 1 ? ` — across ${stats.paymentSheetCount} sheets` : ""}`} />
            <StatRow label="Credit Note vouchers (income reversals)" value={`${stats.creditNoteCount.toLocaleString()}${stats.creditNoteSheetCount > 1 ? ` — across ${stats.creditNoteSheetCount} sheets` : ""}`} />
            <StatRow label="Total output rows" value={stats.outputRows.toLocaleString()} />

            {stats.duplicateRowCount > 0 && (
              <div className="note">
                <span className="pill teal">Auto-resolved</span>
                <p>
                  {stats.duplicateRowCount} voucher{stats.duplicateRowCount === 1 ? "" : "s"} share a bill
                  number with an earlier voucher of the same type. Each repeat goes to its own extra
                  sheet — "RECEIPT (2)", "PAYMENT (2)", "CREDIT NOTE (2)", and so on — so upload each
                  sheet as a separate Tally import and voucher numbers never collide. Same Voucher Type
                  throughout; only the sheet differs. Full list on the DUPLICATE LOG sheet.
                </p>
              </div>
            )}

            <div className="divider" />

            <Toggle
              checked={showPreview}
              onChange={(v) => patch({ showPreview: v })}
              label="Preview output rows"
              description="Sample rows from Receipt, Payment, and Credit Note"
            />

            {showPreview && (
              <>
                <div className="table-wrap">
                  <table className="preview">
                    <thead>
                      <tr>
                        <th>Voucher Type</th>
                        <th>Date</th>
                        <th>Voucher No.</th>
                        <th>Party Name</th>
                        <th className="ta-r">Amount</th>
                        <th>Dr/Cr</th>
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows.map((row, i) => (
                        <tr key={i}>
                          <td>
                            <span className={`pill ${VOUCHER_TAG_STYLES[row[0]] || ""}`}>{row[0]}</span>
                          </td>
                          <td className="mono">{row[1]}</td>
                          <td className="mono">{row[2]}</td>
                          <td>{row[3]}</td>
                          <td className="mono ta-r">
                            {Number(row[4]).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                          </td>
                          <td className="mono">{row[5]}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <p className="preview-note">
                  Sample rows only — {stats.outputRows.toLocaleString()} total in the download.
                </p>
              </>
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

          {result.billNameSummary?.length > 0 && (
            <section className="card">
              <h2 className="card-title">Summary by bill type</h2>
              <p className="lede small">
                Cash / Card / NEFT totals per section, across every bill row —
                a reconciliation check against Teja's own report.
              </p>
              <div className="table-wrap">
                <table className="preview">
                  <thead>
                    <tr>
                      <th>Bill Name</th>
                      <th className="ta-r">Cash</th>
                      <th className="ta-r">Card</th>
                      <th className="ta-r">NEFT</th>
                      <th className="ta-r">Total</th>
                      <th className="ta-r">Bills</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.billNameSummary.map((r) => (
                      <tr key={r.billName}>
                        <td>{r.billName}</td>
                        <td className="mono ta-r">{r.cash.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                        <td className="mono ta-r">{r.card.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                        <td className="mono ta-r">{r.neft.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                        <td className="mono ta-r bold">{r.total.toLocaleString(undefined, { minimumFractionDigits: 2 })}</td>
                        <td className="mono ta-r">{r.count.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr className="total-row">
                      <td>TOTAL</td>
                      <td className="mono ta-r">
                        {result.billNameSummary.reduce((a, r) => a + r.cash, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono ta-r">
                        {result.billNameSummary.reduce((a, r) => a + r.card, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono ta-r">
                        {result.billNameSummary.reduce((a, r) => a + r.neft, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono ta-r">
                        {result.billNameSummary.reduce((a, r) => a + r.total, 0).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                      </td>
                      <td className="mono ta-r">
                        {result.billNameSummary.reduce((a, r) => a + r.count, 0).toLocaleString()}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="preview-note">Also included as a sheet in the downloaded workbook.</p>
            </section>
          )}
        </>
      )}
    </div>
  );
}
