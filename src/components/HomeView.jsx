import React from "react";
import { computePatientLedgerMaster } from "../lib/master.js";
import { StatRow, Button } from "./shared.jsx";
import {
  ReceiptsIcon,
  BillsIcon,
  IpCreditIcon,
  DialysisIcon,
  ReconcileIcon,
  MasterIcon,
} from "./icons.jsx";

function fmt(n) {
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function statusPill(status) {
  switch (status) {
    case "processing":
      return { cls: "amber", label: "Processing" };
    case "done":
      return { cls: "teal", label: "Loaded" };
    case "error":
      return { cls: "red", label: "Error" };
    default:
      return { cls: "navy", label: "Not started" };
  }
}

function HomeTile({ id, label, icon, pill, rows, onNavigate }) {
  return (
    <section className="card home-tile">
      <div className="home-tile-head">
        <span className="home-tile-icon">{icon}</span>
        <div className="home-tile-title">
          <h2>{label}</h2>
        </div>
        <span className={`pill ${pill.cls}`}>{pill.label}</span>
      </div>
      {rows.length > 0 ? (
        <div>
          {rows.map((r) => (
            <StatRow key={r.label} label={r.label} value={r.value} />
          ))}
        </div>
      ) : (
        <p className="home-tile-empty">No file loaded yet.</p>
      )}
      <div className="btn-row">
        <Button variant="ghost" size="sm" onClick={() => onNavigate(id)}>
          Open →
        </Button>
      </div>
    </section>
  );
}

export function HomeView({
  onNavigate,
  receiptsState,
  billsState,
  billsResult,
  ipCreditState,
  dialysisState,
}) {
  const receiptStats = receiptsState.result?.stats;
  const dialysisStats = dialysisState.result?.stats;
  const ipCreditStats = ipCreditState.result?.stats;
  const billsStats = billsResult?.stats;

  const reconcileReady = Boolean(receiptsState.result && billsResult);
  const masterStats = computePatientLedgerMaster(
    receiptsState.result,
    billsResult,
    dialysisState.result
  ).stats;

  const tiles = [
    {
      id: "receipts",
      label: "Receipts",
      icon: <ReceiptsIcon />,
      pill: statusPill(receiptsState.status),
      rows: receiptStats
        ? [
            { label: "Candidate bills", value: receiptStats.candidateBills.toLocaleString() },
            { label: "Rows exported", value: receiptStats.outputRows.toLocaleString() },
          ]
        : [],
    },
    {
      id: "bills",
      label: "Bills",
      icon: <BillsIcon />,
      pill: statusPill(billsState.status),
      rows: billsState.parsed
        ? [
            { label: "Rows scanned", value: billsState.parsed.scannedDataRows.toLocaleString() },
            { label: "Clean bills", value: (billsStats?.cleanRows ?? 0).toLocaleString() },
          ]
        : [],
    },
    {
      id: "ipcredit",
      label: "IP Credit",
      icon: <IpCreditIcon />,
      pill: statusPill(ipCreditState.status),
      rows: ipCreditStats
        ? [
            { label: "Rows scanned", value: ipCreditStats.dataRows.toLocaleString() },
            { label: "Credit settled", value: fmt(ipCreditStats.settledAmount) },
          ]
        : [],
    },
    {
      id: "dialysis",
      label: "Dialysis",
      icon: <DialysisIcon />,
      pill: statusPill(dialysisState.status),
      rows: dialysisStats
        ? [
            { label: "Booked as receivable", value: dialysisStats.bookedRows.toLocaleString() },
            { label: "Total receivable", value: fmt(dialysisStats.totalReceivable) },
          ]
        : [],
    },
    {
      id: "reconcile",
      label: "Reconcile",
      icon: <ReconcileIcon />,
      pill: reconcileReady ? { cls: "teal", label: "Ready" } : { cls: "navy", label: "Needs 2 files" },
      rows: reconcileReady
        ? [{ label: "Sources loaded", value: "Receipts + Bills" }]
        : [],
    },
    {
      id: "master",
      label: "Master",
      icon: <MasterIcon />,
      pill: masterStats.total > 0 ? { cls: "teal", label: "Ready" } : { cls: "navy", label: "Not started" },
      rows:
        masterStats.total > 0
          ? [
              { label: "Unique party ledgers", value: masterStats.total.toLocaleString() },
              { label: "In all loaded sources", value: masterStats.inAll.toLocaleString() },
            ]
          : [],
    },
  ];

  return (
    <div className="view home-view">
      <header className="view-head">
        <p className="eyebrow">Overview</p>
        <h1>Home</h1>
        <p className="lede">
          Status across every section — pick up where you left off, or jump
          straight to what needs attention.
        </p>
      </header>

      <div className="home-grid">
        {tiles.map((tile) => (
          <HomeTile key={tile.id} {...tile} onNavigate={onNavigate} />
        ))}
      </div>
    </div>
  );
}
