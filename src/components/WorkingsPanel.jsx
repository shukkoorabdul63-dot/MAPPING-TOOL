import React from "react";
import { DEFAULT_LEDGER_NAMES } from "../lib/tokens.js";
import { Button } from "./shared.jsx";

export function WorkingsPanel({ open, onClose, ledgerNames, setLedgerNames }) {
  if (!open) return null;
  const update = (key) => (e) =>
    setLedgerNames((v) => ({ ...v, [key]: e.target.value }));

  return (
    <>
      <div className="scrim" onClick={onClose} />
      <div className="drawer" role="dialog" aria-label="Workings">
        <div className="drawer-head">
          <div>
            <p className="eyebrow">Workings</p>
            <h2>Ledger names</h2>
          </div>
          <button className="icon-btn" onClick={onClose} aria-label="Close">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="none">
              <path
                d="M6 6l12 12M18 6l-6 6-6 6"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>

        <div className="drawer-body">
          <p className="lede">
            Enter the exact ledger name as it exists in Tally for each mode.
            Some accounts post NEFT to a Google Pay account ledger, or Card to
            a Swipe Control ledger — set those here and the mapped workbook
            uses them everywhere. Leave blank to use the mode's name.
          </p>

          <div className="field-stack">
            <label className="field">
              <span>Cash ledger</span>
              <input
                type="text"
                placeholder={DEFAULT_LEDGER_NAMES.cash}
                value={ledgerNames.cash}
                onChange={update("cash")}
              />
            </label>
            <label className="field">
              <span>Card ledger</span>
              <input
                type="text"
                placeholder={DEFAULT_LEDGER_NAMES.card}
                value={ledgerNames.card}
                onChange={update("card")}
              />
            </label>
            <label className="field">
              <span>NEFT ledger</span>
              <input
                type="text"
                placeholder={DEFAULT_LEDGER_NAMES.neft}
                value={ledgerNames.neft}
                onChange={update("neft")}
              />
            </label>
          </div>

          <div className="drawer-section">
            <h3>Duplicate bill numbers</h3>
            <p className="lede">
              When the same Bill Number appears more than once in a batch,
              Tally rejects it as a duplicate voucher. The repeat occurrence
              posts under this alternate Voucher Type instead, so both go
              through cleanly.
            </p>
            <div className="field-stack">
              <label className="field">
                <span>Alternate Receipt voucher type</span>
                <input
                  type="text"
                  placeholder={DEFAULT_LEDGER_NAMES.altVoucherType}
                  value={ledgerNames.altVoucherType}
                  onChange={update("altVoucherType")}
                />
              </label>
            </div>
          </div>
        </div>

        <div className="drawer-foot">
          <Button variant="primary" onClick={onClose}>
            Done
          </Button>
        </div>
      </div>
    </>
  );
}
