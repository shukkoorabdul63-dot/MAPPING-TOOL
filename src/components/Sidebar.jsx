import React from "react";

const NAV = [
  { id: "receipts", label: "Receipts", hint: "Cash / Card / NEFT" },
  { id: "bills", label: "Bills", hint: "Sales / Income" },
  { id: "reconcile", label: "Reconcile", hint: "Cross-check totals" },
];

export function Sidebar({ activeView, onNavigate, onOpenWorkings, counts }) {
  return (
    <aside className="sidebar">
      <div className="brand">
        <div className="brand-mark">
          <svg viewBox="0 0 24 24" fill="none">
            <path
              d="M4 12l6 6L20 6"
              stroke="currentColor"
              strokeWidth="2.4"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </div>
        <div>
          <p className="brand-name">Finova</p>
          <p className="brand-tag">Mapping tool</p>
        </div>
      </div>

      <nav className="nav">
        {NAV.map((item) => {
          const isActive = activeView === item.id;
          const count = counts?.[item.id];
          return (
            <button
              key={item.id}
              type="button"
              className={`nav-item${isActive ? " active" : ""}`}
              onClick={() => onNavigate(item.id)}
            >
              <span className="nav-label">{item.label}</span>
              {count != null ? (
                <span className="nav-count mono">{count.toLocaleString()}</span>
              ) : (
                <span className="nav-hint">{item.hint}</span>
              )}
            </button>
          );
        })}
      </nav>

      <div className="sidebar-foot">
        <button type="button" className="nav-item ghost" onClick={onOpenWorkings}>
          <span className="nav-label">Workings</span>
          <span className="nav-hint">Ledger names</span>
        </button>
      </div>
    </aside>
  );
}
