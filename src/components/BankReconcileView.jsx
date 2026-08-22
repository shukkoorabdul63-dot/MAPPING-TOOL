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

const MATCH_TYPE_LABEL = { exact: "Exact", rounded: "Rounded", grouped: "Grouped", paired: "Paired" };
const PREVIEW_LIMIT = 10;

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

function PreviewCaption({ label, total }) {
  if (total <= PREVIEW_LIMIT) return <p className="preview-note">{label}</p>;
  return (
    <p className="preview-note">
      {label} — showing first {PREVIEW_LIMIT} of {total.toLocaleString()}; full list is in the downloaded report.
    </p>
  );
}

function ResultTables({ result }) {
  const matchedPreview = result.matched.slice(0, PREVIEW_LIMIT);
  const possibleMatchesPreview = result.possibleMatches.slice(0, PREVIEW_LIMIT);
  const unmatchedStatementPreview = result.unmatchedStatement.slice(0, PREVIEW_LIMIT);
  const unmatchedLedgerPreview = result.unmatchedLedger.slice(0, PREVIEW_LIMIT);
  const unmatchedByAmountPreview = result.unmatchedByAmount.slice(0, PREVIEW_LIMIT);

  return (
    <>
      {matchedPreview.length > 0 && (
        <>
          <PreviewCaption label="Matched (preview)" total={result.matched.length} />
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
                {matchedPreview.map((m, i) => (
                  <tr key={i}>
                    <td className="mono">{m.date}</td>
                    <td className="mono ta-r">{fmt(m.statementRow.amount)}</td>
                    <td className="mono ta-r">{fmt(m.ledgerRow.amount)}</td>
                    <td>
                      <span className="pill navy">
                        {MATCH_TYPE_LABEL[m.matchType]}
                        {m.matchType === "grouped" && m.groupSize ? ` ×${m.groupSize}` : ""}
                      </span>
                    </td>
                    <td className="mono">{m.ledgerRow.vchNo}</td>
                    <td>{m.ledgerRow.particulars}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
      {possibleMatchesPreview.length > 0 && (
        <>
          <div className="note">
            <span className="pill amber">Review</span>
            <p>
              {result.stats.possibleMatchCount} possible match{result.stats.possibleMatchCount === 1 ? "" : "es"}{" "}
              found by amount only — same day, no shared card/VPA or patient name to justify it. Not counted as
              matched and not removed from Unmatched below; confirm manually using the ref/particulars columns
              before treating either side as reconciled.
              {result.stats.possibleMatchTiedTargetCount > 0 &&
                ` ${result.stats.possibleMatchTiedTargetCount} of these have more than one candidate pairing for the same target — don't assume the first one shown is the right one.`}
            </p>
          </div>
          <PreviewCaption label="Possible matches (preview) — needs manual review" total={result.possibleMatches.length} />
          <div className="table-wrap">
            <table className="preview">
              <thead>
                <tr>
                  <th>Date</th>
                  <th className="ta-r">Statement Amt(s)</th>
                  <th>Statement Ref(s)</th>
                  <th className="ta-r">Ledger Amt(s)</th>
                  <th>Ledger Vch / Particulars</th>
                  <th>Basis</th>
                </tr>
              </thead>
              <tbody>
                {possibleMatchesPreview.map((s, i) => {
                  const stmtRows = s.pairSide === "statement" ? s.pairRows : [s.singleRow];
                  const ledgerRowsArr = s.pairSide === "ledger" ? s.pairRows : [s.singleRow];
                  return (
                    <tr key={i}>
                      <td className="mono">{s.date}</td>
                      <td className="mono ta-r">{stmtRows.map((r) => fmt(r.amount)).join(" + ")}</td>
                      <td className="mono">{stmtRows.map((r) => r.customerRef || "").join(" / ")}</td>
                      <td className="mono ta-r">{ledgerRowsArr.map((r) => fmt(r.amount)).join(" + ")}</td>
                      <td>{ledgerRowsArr.map((r) => `${r.vchNo}: ${r.particulars}`).join(" / ")}</td>
                      <td>
                        <span className="pill amber">{s.matchBasis === "exact" ? "Exact" : "Rounded"}</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}
      {unmatchedStatementPreview.length > 0 && (
        <>
          <PreviewCaption
            label="Unmatched — Statement (preview; bank shows this, no ledger entry found)"
            total={result.unmatchedStatement.length}
          />
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
                {unmatchedStatementPreview.map((r, i) => (
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
      {unmatchedLedgerPreview.length > 0 && (
        <>
          <PreviewCaption
            label="Unmatched — Ledger (preview; booked in Tally, no statement entry found)"
            total={result.unmatchedLedger.length}
          />
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
                {unmatchedLedgerPreview.map((r, i) => (
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
      {unmatchedByAmountPreview.length > 0 && (
        <>
          <PreviewCaption
            label="Unmatched by amount (preview) — biggest gap-causing amounts first"
            total={result.unmatchedByAmount.length}
          />
          <div className="table-wrap">
            <table className="preview">
              <thead>
                <tr>
                  <th>Date</th>
                  <th className="ta-r">Amount</th>
                  <th className="ta-r">Stmt Count</th>
                  <th className="ta-r">Ledger Count</th>
                  <th className="ta-r">Net (₹)</th>
                  <th>Statement Refs</th>
                  <th>Ledger Particulars</th>
                </tr>
              </thead>
              <tbody>
                {unmatchedByAmountPreview.map((e, i) => (
                  <tr key={i}>
                    <td className="mono">{e.date}</td>
                    <td className="mono ta-r">{fmt(e.amount)}</td>
                    <td className="mono ta-r">{e.statementRows.length}</td>
                    <td className="mono ta-r">{e.ledgerRows.length}</td>
                    <td className="mono ta-r">{fmt(e.netAmount)}</td>
                    <td className="mono">{e.statementRows.map((r) => r.customerRef || "").join(", ")}</td>
                    <td>{e.ledgerRows.map((r) => `${r.vchNo}: ${r.particulars}`).join(", ")}</td>
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

function ReconcileSection({ title, hint, accept, section, sectionKey, result, patchSide, onDownload }) {
  const handleFile = useCallback(
    (side, parser) => (file) => {
      patchSide(sectionKey, side, { fileName: file.name, status: "processing", error: null, parsed: null });
      const reader = new FileReader();
      reader.onload = (e) => {
        setTimeout(() => {
          try {
            const data = new Uint8Array(e.target.result);
            // No `cellDates: true` here (unlike the other views): SheetJS's
            // Date-object conversion for the ledger export's numeric date
            // cells is timezone-dependent and comes out a day early for
            // browsers running in Asia/Kolkata (confirmed against the real
            // ledger file — every row shifted from 2-Aug to 1-Aug). Keeping
            // dates as raw Excel serials and decoding them via
            // XLSX.SSF.parse_date_code (excelSerialToDDMMYYYY's number
            // branch) is pure day-count arithmetic with no timezone
            // involved, so it's correct in every browser.
            const workbook = XLSX.read(data, { type: "array" });
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
            value={`${result.stats.matchedCount.toLocaleString()} (₹${fmt(result.stats.matchedSum)}) — ${result.stats.exactMatchCount} exact, ${result.stats.roundedMatchCount} rounded, ${result.stats.groupedMatchCount} grouped, ${result.stats.pairedMatchCount} paired`}
          />
          <StatRow
            label="Unmatched — Statement"
            value={`${result.stats.unmatchedStatementCount.toLocaleString()} (₹${fmt(result.stats.unmatchedStatementSum)})`}
          />
          <StatRow
            label="Unmatched — Ledger"
            value={`${result.stats.unmatchedLedgerCount.toLocaleString()} (₹${fmt(result.stats.unmatchedLedgerSum)})`}
          />
          {result.stats.possibleMatchCount > 0 && (
            <StatRow
              label="Possible matches (unconfirmed)"
              value={`${result.stats.possibleMatchCount.toLocaleString()} — not counted above, see below`}
            />
          )}

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

          <div className="btn-row">
            <Button variant="primary" onClick={onDownload}>
              Download reconciliation report
            </Button>
          </div>
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

  const downloadCard = () => {
    if (!cardResult) return;
    const wb = buildBankReconcileWorkbook({ card: cardResult, upi: null });
    downloadWorkbook(wb, "bank_reconcile_card.xlsx");
  };

  const downloadUpi = () => {
    if (!upiResult) return;
    const wb = buildBankReconcileWorkbook({ card: null, upi: upiResult });
    downloadWorkbook(wb, "bank_reconcile_upi.xlsx");
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
        onDownload={downloadCard}
      />

      <ReconcileSection
        title="UPI / NEFT"
        hint=".csv, .xlsx or .xls"
        accept=".csv,.xlsx,.xls"
        section={{ ...state.upi, statementParser: parseUpiStatement }}
        sectionKey="upi"
        result={upiResult}
        patchSide={patchSide}
        onDownload={downloadUpi}
      />
    </div>
  );
}
