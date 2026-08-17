import React, { useState } from "react";
import { Sidebar } from "./components/Sidebar.jsx";
import { WorkingsPanel } from "./components/WorkingsPanel.jsx";
import { ReceiptsView } from "./components/ReceiptsView.jsx";
import { BillsView } from "./components/BillsView.jsx";
import { DEFAULT_LEDGER_NAMES } from "./lib/tokens.js";
import "./styles.css";

export default function App() {
  const [activeView, setActiveView] = useState("receipts");
  const [workingsOpen, setWorkingsOpen] = useState(false);
  const [ledgerNames, setLedgerNames] = useState(DEFAULT_LEDGER_NAMES);
  const [receiptResult, setReceiptResult] = useState(null);
  const [billsResult, setBillsResult] = useState(null);

  const counts = {
    receipts: receiptResult?.stats?.candidateBills ?? null,
    bills: billsResult?.stats?.scannedDataRows ?? null,
  };

  return (
    <div className="app">
      <Sidebar
        activeView={activeView}
        onNavigate={setActiveView}
        onOpenWorkings={() => setWorkingsOpen(true)}
        counts={counts}
      />
      <main className="workspace">
        {activeView === "receipts" && (
          <ReceiptsView ledgerNames={ledgerNames} onResult={setReceiptResult} />
        )}
        {activeView === "bills" && <BillsView onResult={setBillsResult} />}
      </main>
      <WorkingsPanel
        open={workingsOpen}
        onClose={() => setWorkingsOpen(false)}
        ledgerNames={ledgerNames}
        setLedgerNames={setLedgerNames}
      />
    </div>
  );
}
