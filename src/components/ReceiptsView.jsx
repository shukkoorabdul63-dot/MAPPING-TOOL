import React, { useCallback, useState } from "react";
import * as XLSX from "xlsx";
import { processWorkbook, buildOutputWorkbook } from "../lib/receipts.js";
import { resolveRows, DEFAULT_LEDGER_NAMES } from "../lib/tokens.js";
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

export function ReceiptsView({ ledgerNames, onResult }) {
  const [fileName, setFileName] = useState(null);
  const [status, setStatus] = useState("idle");
  const [error, setError] = useState(null);
  const [result, setResult] = useState(null);
  const [showPreview, setShowPreview] = useState(false);

  const handleFile = useCallback(
    (file) => {
      setFileName(file.name);
      setStatus("processing");
      setError(null);
      setResult(null);
      const reader = new FileReader();
      reader.onload = (e) => {
        setTimeout(() => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const r = processWorkbook(workbook);
            setResult(r);
            setStatus("done");
            onResult?.(r);
          } catch (err) {
            setError(err.message || "Something went wrong.");
            setStatus("error");
          }
        }, 30);
      };
      reader.onerror = () => {
        setError("Could not read this file.");
        setStatus("error");
      };
      reader.readAsArrayBuffer(file);
    },
    [onResult]
  );

  const reset = () => {
    setFileName(null);
    setStatus("idle");
    setError(null);
    setResult(null);
    setShowPreview(false);
    onResult?.(null);
  };

  const download = () => {
    if (!result) return;
    const wb = buildOutputWorkbook(
      result.receiptRows,
      result.paymentRows,
      result.creditNoteRows,
      result.reviewRows,
      result.billNameSummary,
      ledgerNames
    );
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
            <StatRow label="Receipt vouchers" value={stats.receiptCount.toLocaleString()} />
            <StatRow label="Payment vouchers (cash refunds)" value={stats.paymentCount.toLocaleString()} />
            <StatRow label="Credit Note vouchers (income reversals)" value={stats.creditNoteCount.toLocaleString()} />
            <StatRow label="Total output rows" value={stats.outputRows.toLocaleString()} />

            {stats.duplicateBillNos > 0 && (
              <div className="note">
                <span className="pill teal">Auto-resolved</span>
                <p>
                  {stats.duplicateBillNos} bill number{stats.duplicateBillNos === 1 ? "" : "s"} appear
                  {stats.duplicateBillNos === 1 ? "s" : ""} more than once ({stats.duplicateRowCount} rows).
                  The repeat occurrence posts under "
                  {ledgerNames.altVoucherType || DEFAULT_LEDGER_NAMES.altVoucherType}" so voucher numbers
                  never collide. Full list on the DUPLICATE LOG sheet.
                </p>
              </div>
            )}

            <div className="divider" />

            <Toggle
              checked={showPreview}
              onChange={setShowPreview}
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
