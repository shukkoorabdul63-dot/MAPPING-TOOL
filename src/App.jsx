import React, { useState } from "react";
import { Sidebar } from "./components/Sidebar.jsx";
import { WorkingsPanel } from "./components/WorkingsPanel.jsx";
import { ReceiptsView } from "./components/ReceiptsView.jsx";
import { BillsView } from "./components/BillsView.jsx";
import { ReconcileView } from "./components/ReconcileView.jsx";
import { DEFAULT_LEDGER_NAMES } from "./lib/tokens.js";
import "./styles.css";

const EMPTY_VIEW_STATE = {
  fileName: null,
  status: "idle",
  error: null,
  result: null,
  showPreview: false,
};

export default function App() {
  const [activeView, setActiveView] = useState("receipts");
  const [workingsOpen, setWorkingsOpen] = useState(false);
  const [ledgerNames, setLedgerNames] = useState(DEFAULT_LEDGER_NAMES);
  const [receiptsState, setReceiptsState] = useState(EMPTY_VIEW_STATE);
  const [billsState, setBillsState] = useState(EMPTY_VIEW_STATE);

  const counts = {
    receipts: receiptsState.result?.stats?.candidateBills ?? null,
    bills: billsState.result?.stats?.scannedDataRows ?? null,
    reconcile: null,
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
          <ReceiptsView
            ledgerNames={ledgerNames}
            state={receiptsState}
            setState={setReceiptsState}
          />
        )}
        {activeView === "bills" && (
          <BillsView state={billsState} setState={setBillsState} />
        )}
        {activeView === "reconcile" && (
          <ReconcileView
            receiptResult={receiptsState.result}
            billsResult={billsState.result}
          />
        )}
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
