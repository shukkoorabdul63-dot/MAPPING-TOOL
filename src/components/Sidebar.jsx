import React from "react";
import {
  HomeIcon,
  ReceiptsIcon,
  BillsIcon,
  IpCreditIcon,
  DialysisIcon,
  ReconcileIcon,
  MasterIcon,
  WorkingsIcon,
} from "./icons.jsx";

const NAV_HOME = [{ id: "home", label: "Home", hint: "Overview", icon: <HomeIcon /> }];

const NAV_PRIMARY = [
  { id: "receipts", label: "Receipts", hint: "Cash / Card / NEFT", icon: <ReceiptsIcon /> },
  { id: "bills", label: "Bills", hint: "Sales / Income", icon: <BillsIcon /> },
  { id: "ipcredit", label: "IP Credit", hint: "Feeds Discharge", icon: <IpCreditIcon /> },
  { id: "dialysis", label: "Dialysis", hint: "Scheme receivables", icon: <DialysisIcon /> },
  { id: "reconcile", label: "Reconcile", hint: "Cross-check totals", icon: <ReconcileIcon /> },
];

const NAV_SECONDARY = [
  { id: "master", label: "Master", hint: "Patient ledgers", icon: <MasterIcon /> },
];

export function Sidebar({ activeView, onNavigate, onOpenWorkings, counts }) {
  return (
    <aside className="sidebar">
      <div className="brand">
        <div className="brand-mark">
          <svg viewBox="0 0 24 24" fill="none">
            <path d="M4 12l6 6L20 6" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </div>
        <div>
          <p className="brand-name">Finova</p>
          <p className="brand-tag">Mapping tool</p>
        </div>
      </div>

      <nav className="nav">
        {NAV_HOME.map((item) => renderNavItem(item, activeView, onNavigate, counts))}
      </nav>

      <div className="nav">
        <p className="nav-group-label">Sections</p>
        {NAV_PRIMARY.map((item) => renderNavItem(item, activeView, onNavigate, counts))}
      </div>

      <div className="sidebar-foot">
        <p className="nav-group-label">Reference</p>
        <nav className="nav">
          {NAV_SECONDARY.map((item) => renderNavItem(item, activeView, onNavigate, counts))}
        </nav>
        <button type="button" className="nav-item ghost" onClick={onOpenWorkings}>
          <span className="nav-icon"><WorkingsIcon /></span>
          <span className="nav-text">
            <span className="nav-label">Workings</span>
            <span className="nav-hint">Ledger names</span>
          </span>
        </button>
      </div>
    </aside>
  );
}

function renderNavItem(item, activeView, onNavigate, counts) {
  const isActive = activeView === item.id;
  const count = counts?.[item.id];
  return (
    <button
      key={item.id}
      type="button"
      className={`nav-item${isActive ? " active" : ""}`}
      onClick={() => onNavigate(item.id)}
    >
      <span className="nav-icon">{item.icon}</span>
      <span className="nav-text">
        <span className="nav-label">{item.label}</span>
        {count == null && <span className="nav-hint">{item.hint}</span>}
      </span>
      {count != null && <span className="nav-count mono">{count.toLocaleString()}</span>}
    </button>
  );
}
