import React, { useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { parseBillsWorkbook, mapBills, buildBillsOutputWorkbook } from "../lib/bills.js";
import { resolveRows } from "../lib/tokens.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Button, StatusRow, Dropzone } from "./shared.jsx";

const STATUS_LABELS = {
  processing: "Reading and parsing rows…",
  done: "Parsed",
  error: "Could not process this file",
};

function fmt(n) {
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function BillsView({ state, setState, ledgerNames, ipCreditResult, billMappingMode }) {
  const { fileName, status, error, parsed } = state;
  const patch = useCallback((updates) => setState((s) => ({ ...s, ...updates })), [setState]);

  const handleFile = useCallback(
    (file) => {
      patch({ fileName: file.name, status: "processing", error: null, parsed: null });
      const reader = new FileReader();
      reader.onload = (e) => {
        setTimeout(() => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const p = parseBillsWorkbook(workbook);
            patch({ parsed: p, status: "done" });
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
    patch({ fileName: null, status: "idle", error: null, parsed: null });

  // Remap whenever parsed data or IP credit changes — cheap since the raw
  // parse pass isn't re-run.
  const result = useMemo(() => {
    if (!parsed) return null;
    return mapBills(
      parsed,
      ipCreditResult?.creditBySettled || new Map(),
      ipCreditResult?.flaggedRows || [],
      billMappingMode
    );
  }, [parsed, ipCreditResult, billMappingMode]);

  const download = () => {
    if (!result) return;
    const wb = buildBillsOutputWorkbook(result, ledgerNames);
    const base = (fileName || "teja_bills").replace(/\.[^/.]+$/, "");
    downloadWorkbook(wb, `${base}_TALLY_MAPPED.xlsx`);
  };

  const stats = result?.stats;
  const isPerBillName = result?.mode === "per-bill-name";
  const previewOthers = result && !isPerBillName ? resolveRows(result.salesOthersRows.slice(0, 6), ledgerNames) : [];
  const previewPharmacy = result && !isPerBillName ? resolveRows(result.salesPharmacyRows.slice(0, 6), ledgerNames) : [];
  const previewDischarge = result && !isPerBillName ? resolveRows(result.salesDischargeRows.slice(0, 6), ledgerNames) : [];

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">Bills</p>
        <h1>Sales vouchers</h1>
        <p className="lede">
          Upload the bill analysis export from Teja.{" "}
          {billMappingMode === "per-bill-name"
            ? "Each distinct Bill Name gets its own SALES sheet — Voucher Type is the bill name itself, discount lines post to \"Discount - {Bill Name}\", and any row with VAT is split into CGST/SGST automatically."
            : "The tool splits it into three SALES journals — Others, Pharmacy (with CGST/SGST when taxable), and Discharge."}{" "}
          Discharge income is net of IP credit when the IP Credit report is
          loaded on its tab.
        </p>
      </header>

      <section className="card">
        <Dropzone
          onFile={handleFile}
          fileName={fileName}
          emptyTitle="Drop the bills export here"
          emptyHint=".xlsx or .xls"
        />
        <StatusRow status={status} labels={STATUS_LABELS} />
        {status === "error" && <div className="error-box">{error}</div>}
      </section>

      {status === "done" && stats && (
        <>
          <section className="card">
            <h2 className="card-title">Summary</h2>
            <StatRow label="Bill rows scanned" value={stats.scannedDataRows.toLocaleString()} />
            <StatRow label="Cancelled bills removed" value={stats.cancelledCount.toLocaleString()} />
            <StatRow
              label="Duplicate rows removed (kept first)"
              value={`${stats.duplicateRowCount.toLocaleString()} across ${stats.duplicateGroupCount.toLocaleString()} groups`}
            />
            <StatRow label="Clean bills carried forward" value={stats.cleanRows.toLocaleString()} />

            {!isPerBillName &&
              (stats.othersSheetCount > 1 || stats.pharmacySheetCount > 1 || stats.dischargeSheetCount > 1) && (
              <div className="note">
                <span className="pill teal">Auto-resolved</span>
                <p>
                  Some bill numbers repeat within a category. Each repeat goes to its own extra sheet
                  — "SALES - OTHERS (2)", "SALES - PHARMACY (2)", etc. — so upload each sheet as a
                  separate Tally import and voucher numbers never collide. Same Voucher Type
                  throughout; only the sheet differs.
                </p>
              </div>
            )}

            <div className="divider" />
            {isPerBillName ? (
              <>
                <StatRow
                  label="Bill-name buckets"
                  value={`${stats.perBillNameStats.length.toLocaleString()} — one SALES sheet per bill name`}
                />
                {stats.perBillNameStats.some((s) => s.sheetCount > 1) && (
                  <div className="note">
                    <span className="pill teal">Auto-resolved</span>
                    <p>
                      Some bill numbers repeat within a bill-name bucket. Each repeat goes to its
                      own numbered sheet (e.g. "SALES - PHARMACY CASH B (2)") — upload each as a
                      separate Tally import so voucher numbers never collide.
                    </p>
                  </div>
                )}
              </>
            ) : (
              <>
                <StatRow
                  label="Others — SALES journal rows"
                  value={`${stats.othersOutputRows.toLocaleString()}${stats.othersSheetCount > 1 ? ` — across ${stats.othersSheetCount} sheets` : ""}`}
                />
                <StatRow
                  label="Pharmacy — SALES journal rows"
                  value={`${stats.pharmacyOutputRows.toLocaleString()}${stats.pharmacySheetCount > 1 ? ` — across ${stats.pharmacySheetCount} sheets` : ""} (${stats.pharmacyTaxable.toLocaleString()} taxable · ${stats.pharmacyExempt.toLocaleString()} exempt bills)`}
                />
                <StatRow
                  label="Discharge — SALES journal rows"
                  value={`${stats.dischargeOutputRows.toLocaleString()}${stats.dischargeSheetCount > 1 ? ` — across ${stats.dischargeSheetCount} sheets` : ""}`}
                />
              </>
            )}
            {stats.pharmacyTaxable > 0 && (
              <>
                <div className="divider" />
                <StatRow label={isPerBillName ? "CGST collected (VAT rows)" : "Pharmacy CGST collected"} value={fmt(stats.totalCgst)} />
                <StatRow label={isPerBillName ? "SGST collected (VAT rows)" : "Pharmacy SGST collected"} value={fmt(stats.totalSgst)} />
              </>
            )}
            {stats.dischargeCount > 0 && (
              <>
                <div className="divider" />
                <StatRow
                  label="Discharge bills matched with IP Credit"
                  value={`${stats.dischargeMatched.toLocaleString()} of ${stats.dischargeCount.toLocaleString()}`}
                />
                <StatRow
                  label="Total IP credit applied to Discharge income"
                  value={fmt(stats.dischargeIpCreditApplied)}
                />
              </>
            )}

            {!ipCreditResult && stats.dischargeCount > 0 && (
              <div className="note">
                <span className="pill amber">IP Credit not loaded</span>
                <p>
                  Discharge income is currently posted as full Gross Amount.
                  Upload the IP Credit report on its own tab and the numbers
                  update here automatically — no need to re-upload bills.
                </p>
              </div>
            )}

            {stats.dischargeIpCreditFlaggedCount > 0 && (
              <div className="note">
                <span className="pill amber">Review</span>
                <p>
                  {stats.dischargeIpCreditFlaggedCount} Discharge bill
                  {stats.dischargeIpCreditFlaggedCount === 1 ? "" : "s"} in this upload{" "}
                  {stats.dischargeIpCreditFlaggedCount === 1 ? "has" : "have"} an IP Credit row with a
                  Settled Bill No. even though IP Clear wasn't "Y" (₹{fmt(stats.dischargeIpCreditFlaggedAmount)}
                  {" "}excluded from the credit applied above). Full list on the "IP CREDIT FLAGGED" sheet
                  in the download — verify in Teja.
                </p>
              </div>
            )}

            {stats.dischargeReviewCount > 0 && (
              <div className="note">
                <span className="pill red">Review</span>
                <p>
                  {stats.dischargeReviewCount} discharge bill
                  {stats.dischargeReviewCount === 1 ? "" : "s"} would post as
                  negative income (IP credit applied exceeds bill gross). The
                  entry is still generated so the ledger balances, but check
                  the "DISCHARGE TO REVIEW" sheet — usually a stale or
                  over-applied IP credit.
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

          {previewOthers.length > 0 && (
            <PreviewCard title="SALES — Others (preview)" rows={previewOthers} />
          )}
          {previewPharmacy.length > 0 && (
            <PreviewCard title="SALES — Pharmacy (preview)" rows={previewPharmacy} />
          )}
          {previewDischarge.length > 0 && (
            <PreviewCard title="SALES — Discharge (preview)" rows={previewDischarge} />
          )}
          {isPerBillName && stats.perBillNameStats.length > 0 && (
            <section className="card">
              <h2 className="card-title">By bill name</h2>
              <div className="table-wrap">
                <table className="preview">
                  <thead>
                    <tr>
                      <th>Bill Name (= Voucher Type)</th>
                      <th className="ta-r">Bills</th>
                      <th className="ta-r">Output rows</th>
                      <th className="ta-r">Sheets</th>
                    </tr>
                  </thead>
                  <tbody>
                    {stats.perBillNameStats
                      .slice()
                      .sort((a, b) => b.bills - a.bills)
                      .map((s) => (
                        <tr key={s.billName}>
                          <td>{s.billName}</td>
                          <td className="mono ta-r">{s.bills.toLocaleString()}</td>
                          <td className="mono ta-r">{s.outputRows.toLocaleString()}</td>
                          <td className="mono ta-r">{s.sheetCount}</td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
              <p className="preview-note">
                Voucher Type on each sheet is the bill name verbatim; discount
                lines use "Discount - {`{Bill Name}`}"; rows with VAT get a
                CGST/SGST split automatically.
              </p>
            </section>
          )}
        </>
      )}
    </div>
  );
}

function PreviewCard({ title, rows }) {
  return (
    <section className="card">
      <h2 className="card-title">{title}</h2>
      <div className="table-wrap">
        <table className="preview">
          <thead>
            <tr>
              <th>Voucher Type</th>
              <th>Date</th>
              <th>Voucher No.</th>
              <th>Party / Ledger</th>
              <th className="ta-r">Amount</th>
              <th>Dr/Cr</th>
              <th>Cost Center</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => (
              <tr key={i}>
                <td><span className="pill navy">{row[0]}</span></td>
                <td className="mono">{row[1]}</td>
                <td className="mono">{row[2]}</td>
                <td>{row[3]}</td>
                <td className="mono ta-r">
                  {Number(row[4]).toLocaleString(undefined, { minimumFractionDigits: 2 })}
                </td>
                <td className="mono">{row[5]}</td>
                <td className="mono">{row[7] || ""}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
