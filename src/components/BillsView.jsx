import React, { useCallback } from "react";
import * as XLSX from "xlsx";
import { processBillsWorkbook, buildBillsOutputWorkbook } from "../lib/bills.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Toggle, Button, StatusRow, Dropzone } from "./shared.jsx";

const STATUS_LABELS = {
  processing: "Reading and mapping rows…",
  done: "Mapping complete",
  error: "Could not process this file",
};

export function BillsView({ state, setState }) {
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
            const r = processBillsWorkbook(workbook);
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
    const wb = buildBillsOutputWorkbook(result);
    const base = (fileName || "teja_bills").replace(/\.[^/.]+$/, "");
    downloadWorkbook(wb, `${base}_TALLY_MAPPED.xlsx`);
  };

  const stats = result?.stats;
  const previewRows = result?.salesRows?.slice(0, 12) ?? [];

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">Bills</p>
        <h1>Sales vouchers</h1>
        <p className="lede">
          Upload the bill analysis export from Teja. Others and Pharmacy map
          directly to a SALES journal (Pharmacy with CGST/SGST split when
          taxable); Discharge is kept as a raw sheet until its working rules
          are added.
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
              label="Duplicate rows removed (kept first, dropped extras)"
              value={`${stats.duplicateRowCount.toLocaleString()} across ${stats.duplicateGroupCount.toLocaleString()} groups`}
            />
            <StatRow label="Clean bills carried forward" value={stats.cleanRows.toLocaleString()} />
            <div className="divider" />
            <StatRow label="Others — mapped to SALES journal" value={stats.othersCount.toLocaleString()} />
            <StatRow
              label="Pharmacy — mapped to SALES journal"
              value={`${stats.pharmacyCount.toLocaleString()} (${stats.pharmacyTaxable.toLocaleString()} taxable · ${stats.pharmacyExempt.toLocaleString()} exempt)`}
            />
            <StatRow label="Discharge — held as raw sheet" value={stats.dischargeCount.toLocaleString()} />
            <StatRow label="SALES output rows" value={stats.salesOutputRows.toLocaleString()} />
            {stats.pharmacyTaxable > 0 && (
              <>
                <div className="divider" />
                <StatRow
                  label="Pharmacy CGST collected"
                  value={stats.totalCgst.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                />
                <StatRow
                  label="Pharmacy SGST collected"
                  value={stats.totalSgst.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                />
              </>
            )}

            {stats.dischargeCount > 0 && (
              <div className="note">
                <span className="pill amber">Pending rules</span>
                <p>
                  Discharge has its own working rules that aren't wired in yet. Those rows go
                  to a raw sheet in the download so nothing is lost — the SALES journal covers
                  Others and Pharmacy.
                </p>
              </div>
            )}

            <div className="divider" />

            <Toggle
              checked={showPreview}
              onChange={(v) => patch({ showPreview: v })}
              label="Preview SALES journal rows"
              description="First 12 rows of the mapped Others journal"
            />

            {showPreview && previewRows.length > 0 && (
              <>
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
                      {previewRows.map((row, i) => (
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
                <p className="preview-note">
                  Sample rows only — {stats.salesOutputRows.toLocaleString()} total in the download.
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
        </>
      )}
    </div>
  );
}
