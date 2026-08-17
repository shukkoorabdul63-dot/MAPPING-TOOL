import React from "react";

const EXCLUDED_NAMES = new Set(["discharge"]);

// Discharge is excluded from reconciliation because its working rules aren't
// wired in yet. Pharmacy is now fully mapped (with GST split for taxable
// bills), so it's included — but note the Bills side reports Gross Amount
// while the Receipts side does too, so they should tie regardless of GST.
function normalize(name) {
  return String(name || "").trim().toUpperCase();
}

function isExcluded(name) {
  return EXCLUDED_NAMES.has(name.toLowerCase());
}

function fmt(n) {
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function ReconcileView({ receiptResult, billsResult }) {
  const receiptsGross = new Map();
  const billsGross = new Map();

  for (const row of receiptResult?.billNameSummary ?? []) {
    const name = normalize(row.billName);
    if (isExcluded(row.billName)) continue;
    receiptsGross.set(name, (receiptsGross.get(name) || 0) + (row.grossAmount || 0));
  }
  for (const row of billsResult?.billGrossByName ?? []) {
    const name = normalize(row.billName);
    if (isExcluded(row.billName)) continue;
    billsGross.set(name, (billsGross.get(name) || 0) + (row.grossAmount || 0));
  }

  const allNames = [...new Set([...receiptsGross.keys(), ...billsGross.keys()])];
  const matched = [];
  const unmatched = [];
  for (const name of allNames) {
    const r = receiptsGross.get(name) ?? null;
    const b = billsGross.get(name) ?? null;
    if (r != null && b != null) {
      const diff = r - b;
      matched.push({ name, receipts: r, bills: b, diff });
    } else {
      unmatched.push({ name, receipts: r, bills: b });
    }
  }
  matched.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
  unmatched.sort((a, b) => (b.receipts ?? b.bills ?? 0) - (a.receipts ?? a.bills ?? 0));

  const totals = matched.reduce(
    (acc, r) => ({
      receipts: acc.receipts + r.receipts,
      bills: acc.bills + r.bills,
      diff: acc.diff + r.diff,
    }),
    { receipts: 0, bills: 0, diff: 0 }
  );

  const mismatches = matched.filter((r) => Math.abs(r.diff) > 0.01).length;

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">Reconciliation</p>
        <h1>Cross-check</h1>
        <p className="lede">
          Compares the Gross Amount total per Bill Name from the Receipts
          export against the Bills export. Discharge is excluded since its
          working rules are still pending.
        </p>
      </header>

      {(!receiptResult || !billsResult) && (
        <section className="card">
          <div className="empty">
            <p className="empty-title">Upload both files to reconcile</p>
            <ul className="empty-list">
              <li className={receiptResult ? "done" : ""}>
                {receiptResult ? "✓" : "•"} Receipts file — {receiptResult ? "loaded" : "not loaded"}
              </li>
              <li className={billsResult ? "done" : ""}>
                {billsResult ? "✓" : "•"} Bills file — {billsResult ? "loaded" : "not loaded"}
              </li>
            </ul>
          </div>
        </section>
      )}

      {receiptResult && billsResult && (
        <>
          <section className="card">
            <h2 className="card-title">Overall</h2>
            <div className="stat-row">
              <span className="stat-label">Receipts — total Gross Amount</span>
              <span className="stat-value mono">{fmt(totals.receipts)}</span>
            </div>
            <div className="stat-row">
              <span className="stat-label">Bills — total Gross Amount</span>
              <span className="stat-value mono">{fmt(totals.bills)}</span>
            </div>
            <div className="stat-row">
              <span className="stat-label">Difference (Receipts − Bills)</span>
              <span
                className="stat-value mono"
                style={{ color: Math.abs(totals.diff) > 0.01 ? "var(--red)" : "var(--accent)" }}
              >
                {fmt(totals.diff)}
              </span>
            </div>
            {mismatches > 0 && (
              <div className="note" style={{ marginTop: "1rem" }}>
                <span className="pill red">Mismatch</span>
                <p>
                  {mismatches} bill name{mismatches === 1 ? "" : "s"} do
                  {mismatches === 1 ? "es" : ""} not tie. Check the table
                  below — a non-zero difference usually means a bill was
                  missed in one export or captured twice in the other.
                </p>
              </div>
            )}
          </section>

          <section className="card">
            <h2 className="card-title">By bill name</h2>
            <div className="table-wrap">
              <table className="preview">
                <thead>
                  <tr>
                    <th>Bill Name</th>
                    <th className="ta-r">Receipts (Gross)</th>
                    <th className="ta-r">Bills (Gross)</th>
                    <th className="ta-r">Difference</th>
                  </tr>
                </thead>
                <tbody>
                  {matched.map((r) => {
                    const isDiff = Math.abs(r.diff) > 0.01;
                    return (
                      <tr key={r.name}>
                        <td>{r.name}</td>
                        <td className="mono ta-r">{fmt(r.receipts)}</td>
                        <td className="mono ta-r">{fmt(r.bills)}</td>
                        <td
                          className="mono ta-r"
                          style={{ color: isDiff ? "var(--red)" : "var(--ink-muted)", fontWeight: isDiff ? 700 : 400 }}
                        >
                          {fmt(r.diff)}
                        </td>
                      </tr>
                    );
                  })}
                  <tr className="total-row">
                    <td>TOTAL</td>
                    <td className="mono ta-r">{fmt(totals.receipts)}</td>
                    <td className="mono ta-r">{fmt(totals.bills)}</td>
                    <td
                      className="mono ta-r"
                      style={{ color: Math.abs(totals.diff) > 0.01 ? "var(--red)" : "var(--ink-muted)" }}
                    >
                      {fmt(totals.diff)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>

          {unmatched.length > 0 && (
            <section className="card">
              <h2 className="card-title">Only in one file</h2>
              <p className="lede small">
                Bill Names that appear in only one export. Often just a naming
                mismatch (e.g. Bills says "LABORATORY BILL", Receipts says
                "LABORATORY  BILL" with two spaces). Worth checking either way.
              </p>
              <div className="table-wrap">
                <table className="preview">
                  <thead>
                    <tr>
                      <th>Bill Name</th>
                      <th className="ta-r">Receipts (Gross)</th>
                      <th className="ta-r">Bills (Gross)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {unmatched.map((r) => (
                      <tr key={r.name}>
                        <td>{r.name}</td>
                        <td className="mono ta-r">{r.receipts != null ? fmt(r.receipts) : "—"}</td>
                        <td className="mono ta-r">{r.bills != null ? fmt(r.bills) : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}
        </>
      )}
    </div>
  );
}
