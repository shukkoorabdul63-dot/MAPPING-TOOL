import React, { useState, useMemo, useEffect } from "react";
import { Sidebar } from "./components/Sidebar.jsx";
import { WorkingsPanel } from "./components/WorkingsPanel.jsx";
import { HomeView } from "./components/HomeView.jsx";
import { ReceiptsView } from "./components/ReceiptsView.jsx";
import { BillsView } from "./components/BillsView.jsx";
import { IpCreditView } from "./components/IpCreditView.jsx";
import { DialysisView } from "./components/DialysisView.jsx";
import { ReconcileView } from "./components/ReconcileView.jsx";
import { MasterView } from "./components/MasterView.jsx";
import { mapBills } from "./lib/bills.js";
import {
  loadHospitalProfiles,
  saveHospitalProfiles,
  getActiveHospital,
  withUpdatedLedgerNames,
  withActiveHospital,
  withRemovedHospital,
  BILL_MAPPING_MODES,
} from "./lib/hospitals.js";
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
  const [activeView, setActiveView] = useState("home");
  const [workingsOpen, setWorkingsOpen] = useState(false);

  // Settings (ledger/voucher names) and bill-mapping strategy are scoped per
  // hospital and persist in localStorage — see src/lib/hospitals.js. This is
  // a config switcher, not a login: there's no backend, so nothing here
  // gates access, it just picks which naming/categorization to apply.
  const [hospitalProfiles, setHospitalProfiles] = useState(loadHospitalProfiles);

  useEffect(() => {
    saveHospitalProfiles(hospitalProfiles);
  }, [hospitalProfiles]);

  const activeHospital = getActiveHospital(hospitalProfiles);
  const ledgerNames = activeHospital?.ledgerNames ?? null;
  const billMappingMode = activeHospital?.billMappingMode ?? BILL_MAPPING_MODES.FIXED;

  const setLedgerNames = (updater) => {
    if (!activeHospital) return;
    setHospitalProfiles((profiles) => {
      const current = getActiveHospital(profiles);
      if (!current) return profiles;
      const next = typeof updater === "function" ? updater(current.ledgerNames) : updater;
      return withUpdatedLedgerNames(profiles, current.id, next);
    });
  };

  const [receiptsState, setReceiptsState] = useState(EMPTY_VIEW_STATE);
  const [billsState, setBillsState] = useState(EMPTY_BILLS_STATE);
  const [ipCreditState, setIpCreditState] = useState(EMPTY_IPCREDIT_STATE);
  const [dialysisState, setDialysisState] = useState(EMPTY_VIEW_STATE);

  // Any of these being non-null means processed-but-maybe-undownloaded data
  // is loaded — switching hospitals would silently re-resolve it under the
  // new hospital's names/categorization, so that's worth a warning.
  const hasLoadedData = Boolean(
    receiptsState.result || billsState.parsed || dialysisState.result || ipCreditState.result
  );

  const onSwitchHospital = (hospitalId) => {
    if (hospitalId === hospitalProfiles.activeHospitalId) return;
    if (
      hasLoadedData &&
      !confirm(
        "You have processed files loaded under the current hospital. Switching will re-resolve previews/downloads for that same data under the new hospital's ledger names and categorization. Re-upload fresh files after switching if this data belongs to the previous hospital."
      )
    ) {
      return;
    }
    setHospitalProfiles((profiles) => withActiveHospital(profiles, hospitalId));
  };

  const onRemoveHospital = (hospitalId) => {
    const target = hospitalProfiles.hospitals[hospitalId];
    if (!target) return;
    if (Object.keys(hospitalProfiles.hospitals).length <= 1) return;
    if (!confirm(`Remove "${target.name}"? Its saved ledger names and settings will be discarded.`)) return;
    setHospitalProfiles((profiles) => withRemovedHospital(profiles, hospitalId));
  };

  // Bills result is derived — recomputed whenever parsed bills or IP credit
  // change. That way, uploading IP Credit after Bills doesn't need a
  // re-upload; Discharge income just refreshes.
  const billsResult = useMemo(() => {
    if (!billsState.parsed) return null;
    return mapBills(
      billsState.parsed,
      ipCreditState.result?.creditBySettled || new Map(),
      ipCreditState.result?.flaggedRows || [],
      billMappingMode
    );
  }, [billsState.parsed, ipCreditState.result, billMappingMode]);

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
        hospitals={hospitalProfiles.hospitals}
        activeHospitalId={hospitalProfiles.activeHospitalId}
        onSwitchHospital={onSwitchHospital}
        onRemoveHospital={onRemoveHospital}
      />
      {!activeHospital ? (
        <main className="workspace">
          <div className="view">
            <header className="view-head">
              <h1>Select a hospital to continue</h1>
              <p className="lede">
                All hospitals were removed. Reload the page to restore the
                default Hospital A / Hospital B setup.
              </p>
            </header>
          </div>
        </main>
      ) : (
      <main className="workspace">
        {activeView === "home" && (
          <HomeView
            onNavigate={setActiveView}
            receiptsState={receiptsState}
            billsState={billsState}
            billsResult={billsResult}
            ipCreditState={ipCreditState}
            dialysisState={dialysisState}
          />
        )}
        {activeView === "receipts" && (
          <ReceiptsView ledgerNames={ledgerNames} state={receiptsState} setState={setReceiptsState} />
        )}
        {activeView === "bills" && (
          <BillsView
            state={billsState}
            setState={setBillsState}
            ledgerNames={ledgerNames}
            ipCreditResult={ipCreditState.result}
            billMappingMode={billMappingMode}
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
      )}
      {activeHospital && (
        <WorkingsPanel
          open={workingsOpen}
          onClose={() => setWorkingsOpen(false)}
          ledgerNames={ledgerNames}
          setLedgerNames={setLedgerNames}
          activeHospitalName={activeHospital.name}
        />
      )}
    </div>
  );
}
