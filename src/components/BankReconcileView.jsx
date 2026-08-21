import React, { useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  parseCardStatement,
  parseUpiStatement,
  parseLedgerExport,
  matchBankReconcile,
  buildBankReconcileWorkbook,
} from "../lib/bankReconcile.js";
import { downloadWorkbook } from "../lib/utils.js";
import { StatRow, Button, StatusRow, Dropzone } from "./shared.jsx";

const STATUS_LABELS = {
  processing: "Reading rows…",
  done: "Parsed",
  error: "Could not process this file",
};

function fmt(n) {
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const MATCH_TYPE_LABEL = { exact: "Exact", rounded: "Rounded", grouped: "Grouped" };

function SideDropzone({ label, hint, accept, side, onFile }) {
  const { fileName, status, error } = side;
  return (
    <div>
      <p className="dropzone-label">{label}</p>
      <Dropzone
        onFile={onFile}
        fileName={fileName}
        emptyTitle="Drop file here"
        emptyHint={hint}
        accept={accept}
      />
      <StatusRow status={status} labels={STATUS_LABELS} />
      {status === "error" && <div className="error-box">{error}</div>}
    </div>
  );
}

function ResultTables({ result }) {
  return (
    <>
      {result.matched.length > 0 && (
        <div className="table-wrap">
          <table className="preview">
            <thead>
              <tr>
                <th>Date</th>
                <th className="ta-r">Statement Amt</th>
                <th className="ta-r">Ledger Amt</th>
                <th>Type</th>
                <th>Ledger Vch No.</th>
                <th>Particulars</th>
              </tr>
            </thead>
            <tbody>
              {result.matched.map((m, i) => (
                <tr key={i}>
                  <td className="mono">{m.date}</td>
                  <td className="mono ta-r">{fmt(m.statementRow.amount)}</td>
                  <td className="mono ta-r">{fmt(m.ledgerRow.amount)}</td>
                  <td>
                    <span className="pill navy">
                      {MATCH_TYPE_LABEL[m.matchType]}
                      {m.groupSize ? ` ×${m.groupSize}` : ""}
                    </span>
                  </td>
                  <td className="mono">{m.ledgerRow.vchNo}</td>
                  <td>{m.ledgerRow.particulars}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {result.unmatchedStatement.length > 0 && (
        <>
          <p className="preview-note">Unmatched — Statement (bank shows this, no ledger entry found)</p>
          <div className="table-wrap">
            <table className="preview">
              <thead>
                <tr>
                  <th>Date</th>
                  <th className="ta-r">Amount</th>
                  <th>Customer Ref</th>
                </tr>
              </thead>
              <tbody>
                {result.unmatchedStatement.map((r, i) => (
                  <tr key={i}>
                    <td className="mono">{r.date}</td>
                    <td className="mono ta-r">{fmt(r.amount)}</td>
                    <td className="mono">{r.customerRef || ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
      {result.unmatchedLedger.length > 0 && (
        <>
          <p className="preview-note">Unmatched — Ledger (booked in Tally, no statement entry found)</p>
          <div className="table-wrap">
            <table className="preview">
              <thead>
                <tr>
                  <th>Date</th>
                  <th className="ta-r">Amount</th>
                  <th>Vch No.</th>
                  <th>Particulars</th>
                </tr>
              </thead>
              <tbody>
                {result.unmatchedLedger.map((r, i) => (
                  <tr key={i}>
                    <td className="mono">{r.date}</td>
                    <td className="mono ta-r">{fmt(r.amount)}</td>
                    <td className="mono">{r.vchNo}</td>
                    <td>{r.particulars}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </>
  );
}

function ReconcileSection({ title, hint, accept, section, sectionKey, result, patchSide }) {
  const handleFile = useCallback(
    (side, parser) => (file) => {
      patchSide(sectionKey, side, { fileName: file.name, status: "processing", error: null, parsed: null });
      const reader = new FileReader();
      reader.onload = (e) => {
        setTimeout(() => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const parsed = parser(workbook);
            patchSide(sectionKey, side, { parsed, status: "done" });
          } catch (err) {
            patchSide(sectionKey, side, { error: err.message || "Something went wrong.", status: "error" });
          }
        }, 30);
      };
      reader.onerror = () => patchSide(sectionKey, side, { error: "Could not read this file.", status: "error" });
      reader.readAsArrayBuffer(file);
    },
    [patchSide, sectionKey]
  );

  return (
    <section className="card">
      <h2 className="card-title">{title}</h2>
      <div className="dropzone-pair">
        <SideDropzone
          label="Bank / gateway statement"
          hint={hint}
          accept={accept}
          side={section.statement}
          onFile={handleFile("statement", section.statementParser)}
        />
        <SideDropzone
          label="Tally ledger export"
          hint=".xlsx or .xls"
          accept=".xlsx,.xls"
          side={section.ledger}
          onFile={handleFile("ledger", parseLedgerExport)}
        />
      </div>

      {result && (
        <>
          <div className="divider" />
          <StatRow label="Statement rows" value={`${result.stats.statementCount.toLocaleString()} (₹${fmt(result.stats.statementSum)})`} />
          <StatRow label="Ledger rows" value={`${result.stats.ledgerCount.toLocaleString()} (₹${fmt(result.stats.ledgerSum)})`} />
          <div className="divider" />
          <StatRow
            label="Matched"
            value={`${result.stats.matchedCount.toLocaleString()} (₹${fmt(result.stats.matchedSum)}) — ${result.stats.exactMatchCount} exact, ${result.stats.roundedMatchCount} rounded, ${result.stats.groupedMatchCount} grouped`}
          />
          <StatRow
            label="Unmatched — Statement"
            value={`${result.stats.unmatchedStatementCount.toLocaleString()} (₹${fmt(result.stats.unmatchedStatementSum)})`}
          />
          <StatRow
            label="Unmatched — Ledger"
            value={`${result.stats.unmatchedLedgerCount.toLocaleString()} (₹${fmt(result.stats.unmatchedLedgerSum)})`}
          />

          {result.stats.ambiguousMatchCount > 0 && (
            <div className="note">
              <span className="pill amber">Review</span>
              <p>
                {result.stats.ambiguousMatchCount} matched rows came from a date + amount combination that
                occurred more than once on both sides — paired in file order. Verify manually if exact
                correspondence matters.
              </p>
            </div>
          )}

          <ResultTables result={result} />
        </>
      )}
    </section>
  );
}

const EMPTY_SIDE = { fileName: null, status: "idle", error: null, parsed: null };
export const EMPTY_BANK_RECONCILE_STATE = {
  card: { statement: EMPTY_SIDE, ledger: EMPTY_SIDE },
  upi: { statement: EMPTY_SIDE, ledger: EMPTY_SIDE },
};

export function BankReconcileView({ state, setState }) {
  const patchSide = useCallback(
    (sectionKey, side, updates) => {
      setState((s) => ({
        ...s,
        [sectionKey]: { ...s[sectionKey], [side]: { ...s[sectionKey][side], ...updates } },
      }));
    },
    [setState]
  );

  const cardResult = useMemo(() => {
    const { statement, ledger } = state.card;
    if (!statement.parsed || !ledger.parsed) return null;
    return matchBankReconcile(statement.parsed.rows, ledger.parsed.rows);
  }, [state.card.statement.parsed, state.card.ledger.parsed]);

  const upiResult = useMemo(() => {
    const { statement, ledger } = state.upi;
    if (!statement.parsed || !ledger.parsed) return null;
    return matchBankReconcile(statement.parsed.rows, ledger.parsed.rows);
  }, [state.upi.statement.parsed, state.upi.ledger.parsed]);

  const download = () => {
    if (!cardResult && !upiResult) return;
    const wb = buildBankReconcileWorkbook({ card: cardResult, upi: upiResult });
    downloadWorkbook(wb, "bank_reconcile.xlsx");
  };

  return (
    <div className="view">
      <header className="view-head">
        <p className="eyebrow">Bank Reconcile</p>
        <h1>Card &amp; UPI / NEFT vs Ledger</h1>
        <p className="lede">
          Upload each payment mode's bank/gateway statement alongside its Tally ledger voucher export.
          Rows are matched by date and amount — first exactly, then allowing for rounding (ledger exports
          often carry paise the statement never shows), then by grouping same-day payments from the same
          customer that Tally booked as one combined receipt.
        </p>
      </header>

      <ReconcileSection
        title="Card"
        hint=".xlsx or .xls"
        accept=".xlsx,.xls"
        section={{ ...state.card, statementParser: parseCardStatement }}
        sectionKey="card"
        result={cardResult}
        patchSide={patchSide}
      />

      <ReconcileSection
        title="UPI / NEFT"
        hint=".csv, .xlsx or .xls"
        accept=".csv,.xlsx,.xls"
        section={{ ...state.upi, statementParser: parseUpiStatement }}
        sectionKey="upi"
        result={upiResult}
        patchSide={patchSide}
      />

      {(cardResult || upiResult) && (
        <section className="card">
          <div className="btn-row">
            <Button variant="primary" onClick={download}>
              Download reconciliation report
            </Button>
          </div>
        </section>
      )}
    </div>
  );
}
