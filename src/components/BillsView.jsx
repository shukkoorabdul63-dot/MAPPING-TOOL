import React, { useCallback, useState } from "react";
import * as XLSX from "xlsx";
import { processBillsWorkbook, buildBillsOutputWorkbook } from "../lib/bills.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Toggle, Button, StatusRow, Dropzone } from "./shared.jsx";

const STATUS_LABELS = {
  processing: "Reading and mapping rows…",
  done: "Mapping complete",
  error: "Could not process this file",
};

export function BillsView({ onResult }) {
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
            const r = processBillsWorkbook(workbook);
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
          Upload the bill analysis export from Teja. Others map directly to a
          SALES journal; Discharge and Pharmacy are kept as raw sheets until
          their working rules are added.
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
              label="Duplicate (Bill No. + Bill Name) rows removed"
              value={`${stats.duplicateRowCount.toLocaleString()} across ${stats.duplicateGroupCount.toLocaleString()} groups`}
            />
            <StatRow label="Clean bills carried forward" value={stats.cleanRows.toLocaleString()} />
            <div className="divider" />
            <StatRow label="Others — mapped to SALES journal" value={stats.othersCount.toLocaleString()} />
            <StatRow label="Discharge — held as raw sheet" value={stats.dischargeCount.toLocaleString()} />
            <StatRow label="Pharmacy — held as raw sheet" value={stats.pharmacyCount.toLocaleString()} />
            <StatRow label="SALES output rows (Others)" value={stats.salesOutputRows.toLocaleString()} />

            {(stats.dischargeCount > 0 || stats.pharmacyCount > 0) && (
              <div className="note">
                <span className="pill amber">Pending rules</span>
                <p>
                  Discharge and Pharmacy have their own working rules that
                  aren't wired in yet. Those rows go to raw sheets in the
                  download so nothing is lost — the SALES journal covers
                  everything else.
                </p>
              </div>
            )}

            <div className="divider" />

            <Toggle
              checked={showPreview}
              onChange={setShowPreview}
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
