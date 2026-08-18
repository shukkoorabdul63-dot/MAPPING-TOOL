import React, { useState, useMemo } from "react";
import { Sidebar } from "./components/Sidebar.jsx";
import { WorkingsPanel } from "./components/WorkingsPanel.jsx";
import { ReceiptsView } from "./components/ReceiptsView.jsx";
import { BillsView } from "./components/BillsView.jsx";
import { IpCreditView } from "./components/IpCreditView.jsx";
import { DialysisView } from "./components/DialysisView.jsx";
import { ReconcileView } from "./components/ReconcileView.jsx";
import { MasterView } from "./components/MasterView.jsx";
import { mapBills } from "./lib/bills.js";
import { DEFAULT_LEDGER_NAMES } from "./lib/tokens.js";
import "./styles.css";

const EMPTY_VIEW_STATE = {
  fileName: null,
  status: "idle",
  error: null,
  result: null,
  showPreview: false,
};

const EMPTY_BILLS_STATE = {
  fileName: null,
  status: "idle",
  error: null,
  parsed: null, // { candidates, scannedDataRows }
};

const EMPTY_IPCREDIT_STATE = {
  fileName: null,
  status: "idle",
  error: null,
  result: null,
};

export default function App() {
  const [activeView, setActiveView] = useState("receipts");
  const [workingsOpen, setWorkingsOpen] = useState(false);
  const [ledgerNames, setLedgerNames] = useState(DEFAULT_LEDGER_NAMES);
  const [receiptsState, setReceiptsState] = useState(EMPTY_VIEW_STATE);
  const [billsState, setBillsState] = useState(EMPTY_BILLS_STATE);
  const [ipCreditState, setIpCreditState] = useState(EMPTY_IPCREDIT_STATE);
  const [dialysisState, setDialysisState] = useState(EMPTY_VIEW_STATE);

  // Bills result is derived — recomputed whenever parsed bills or IP credit
  // change. That way, uploading IP Credit after Bills doesn't need a
  // re-upload; Discharge income just refreshes.
  const billsResult = useMemo(() => {
    if (!billsState.parsed) return null;
    return mapBills(billsState.parsed, ipCreditState.result?.creditBySettled || new Map());
  }, [billsState.parsed, ipCreditState.result]);

  const counts = {
    receipts: receiptsState.result?.stats?.candidateBills ?? null,
    bills: billsState.parsed?.scannedDataRows ?? null,
    ipcredit: ipCreditState.result?.stats?.dataRows ?? null,
    dialysis: dialysisState.result?.stats?.bookedRows ?? null,
    reconcile: null,
    master: null,
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
          <ReceiptsView ledgerNames={ledgerNames} state={receiptsState} setState={setReceiptsState} />
        )}
        {activeView === "bills" && (
          <BillsView
            state={billsState}
            setState={setBillsState}
            ledgerNames={ledgerNames}
            ipCreditResult={ipCreditState.result}
          />
        )}
        {activeView === "ipcredit" && (
          <IpCreditView state={ipCreditState} setState={setIpCreditState} />
        )}
        {activeView === "dialysis" && (
          <DialysisView state={dialysisState} setState={setDialysisState} ledgerNames={ledgerNames} />
        )}
        {activeView === "reconcile" && (
          <ReconcileView receiptResult={receiptsState.result} billsResult={billsResult} />
        )}
        {activeView === "master" && (
          <MasterView
            receiptResult={receiptsState.result}
            billsResult={billsResult}
            dialysisResult={dialysisState.result}
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
