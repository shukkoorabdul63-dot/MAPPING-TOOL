import React from "react";
import { computePatientLedgerMaster, buildMasterWorkbook } from "../lib/master.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Button } from "./shared.jsx";

export function MasterView({ receiptResult, billsResult, dialysisResult }) {
  const { list, stats } = computePatientLedgerMaster(receiptResult, billsResult, dialysisResult);

  const download = () => {
    const wb = buildMasterWorkbook(list);
    downloadWorkbook(wb, `patient_ledgers_master.xlsx`);
  };

  const anyData = list.length > 0;

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">Master</p>
        <h1>Patient ledgers</h1>
        <p className="lede">
          Deduped list of Party Name ledgers (OP number − Patient name) across
          the Receipts, Bills, and Dialysis files you've loaded. Download once
          and import into Tally so every voucher's Party posts to an existing
          ledger.
        </p>
      </header>

      {!anyData ? (
        <section className="card">
          <div className="empty">
            <p className="empty-title">No party names yet</p>
            <ul className="empty-list">
              <li className={receiptResult ? "done" : ""}>
                {receiptResult ? "✓" : "•"} Receipts — {receiptResult ? "loaded" : "not loaded"}
              </li>
              <li className={billsResult ? "done" : ""}>
                {billsResult ? "✓" : "•"} Bills — {billsResult ? "loaded" : "not loaded"}
              </li>
              <li className={dialysisResult ? "done" : ""}>
                {dialysisResult ? "✓" : "•"} Dialysis — {dialysisResult ? "loaded" : "not loaded"}
              </li>
            </ul>
          </div>
        </section>
      ) : (
        <>
          <section className="card">
            <h2 className="card-title">Summary</h2>
            <StatRow label="Total unique party ledgers" value={stats.total.toLocaleString()} />
            <StatRow label="Appear in Receipts" value={stats.inReceipts.toLocaleString()} />
            <StatRow label="Appear in Bills" value={stats.inBills.toLocaleString()} />
            <StatRow label="Appear in Dialysis" value={stats.inDialysis.toLocaleString()} />
            <StatRow label="Appear in all loaded sources" value={stats.inAll.toLocaleString()} />
            <div className="btn-row">
              <Button variant="primary" onClick={download}>
                Download master workbook
              </Button>
            </div>
          </section>

          <section className="card">
            <h2 className="card-title">Preview</h2>
            <p className="lede small">First 20 rows shown — full list is in the download.</p>
            <div className="table-wrap">
              <table className="preview">
                <thead>
                  <tr>
                    <th>Party Name</th>
                    <th className="ta-r">Receipts</th>
                    <th className="ta-r">Bills</th>
                    <th className="ta-r">Dialysis</th>
                  </tr>
                </thead>
                <tbody>
                  {list.slice(0, 20).map((r) => (
                    <tr key={r.partyName}>
                      <td>{r.partyName}</td>
                      <td className="mono ta-r">{r.inReceipts ? "✓" : ""}</td>
                      <td className="mono ta-r">{r.inBills ? "✓" : ""}</td>
                      <td className="mono ta-r">{r.inDialysis ? "✓" : ""}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}
    </div>
  );
}
