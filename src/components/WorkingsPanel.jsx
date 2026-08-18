import React from "react";
import { DEFAULT_LEDGER_NAMES } from "../lib/tokens.js";
import { Button } from "./shared.jsx";

function Field({ label, placeholder, value, onChange }) {
  return (
    <label className="field">
      <span>{label}</span>
      <input type="text" placeholder={placeholder} value={value} onChange={onChange} />
    </label>
  );
}

export function WorkingsPanel({ open, onClose, ledgerNames, setLedgerNames }) {
  if (!open) return null;
  const update = (key) => (e) => setLedgerNames((v) => ({ ...v, [key]: e.target.value }));

  return (
    <>
      <div className="scrim" onClick={onClose} />
      <div className="drawer" role="dialog" aria-label="Workings">
        <div className="drawer-head">
          <div>
            <p className="eyebrow">Workings</p>
            <h2>Ledger &amp; voucher names</h2>
          </div>
          <button className="icon-btn" onClick={onClose} aria-label="Close">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="none">
              <path d="M6 6l12 12M18 6l-6 6-6 6" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
            </svg>
          </button>
        </div>

        <div className="drawer-body">
          <p className="lede">
            Names apply everywhere — on-screen previews and the downloaded
            workbook — without reprocessing the source file. Leave a field
            blank to fall back to its default.
          </p>

          <div className="drawer-section noborder">
            <h3>Payment mode ledgers (Receipts)</h3>
            <div className="field-stack">
              <Field label="Cash ledger" placeholder={DEFAULT_LEDGER_NAMES.cash} value={ledgerNames.cash} onChange={update("cash")} />
              <Field label="Card ledger" placeholder={DEFAULT_LEDGER_NAMES.card} value={ledgerNames.card} onChange={update("card")} />
              <Field label="NEFT ledger" placeholder={DEFAULT_LEDGER_NAMES.neft} value={ledgerNames.neft} onChange={update("neft")} />
            </div>
          </div>

          <div className="drawer-section">
            <h3>Duplicate receipt vouchers</h3>
            <p className="lede small">
              When the same Bill No. appears more than once, the repeat
              occurrence posts under this alternate Voucher Type.
            </p>
            <div className="field-stack">
              <Field
                label="Alternate Receipt voucher type"
                placeholder={DEFAULT_LEDGER_NAMES.altVoucherType}
                value={ledgerNames.altVoucherType}
                onChange={update("altVoucherType")}
              />
            </div>
          </div>

          <div className="drawer-section">
            <h3>Sales voucher types</h3>
            <p className="lede small">
              Each Bills sheet posts under its own Voucher Type. Give the
              alternate a different name so duplicate bill numbers inside a
              single sheet don't collide on upload.
            </p>
            <div className="field-stack">
              <Field label="Others — voucher type" placeholder={DEFAULT_LEDGER_NAMES.othersVoucher} value={ledgerNames.othersVoucher} onChange={update("othersVoucher")} />
              <Field label="Others — alternate" placeholder={DEFAULT_LEDGER_NAMES.othersAltVoucher} value={ledgerNames.othersAltVoucher} onChange={update("othersAltVoucher")} />
              <Field label="Pharmacy — voucher type" placeholder={DEFAULT_LEDGER_NAMES.pharmacyVoucher} value={ledgerNames.pharmacyVoucher} onChange={update("pharmacyVoucher")} />
              <Field label="Pharmacy — alternate" placeholder={DEFAULT_LEDGER_NAMES.pharmacyAltVoucher} value={ledgerNames.pharmacyAltVoucher} onChange={update("pharmacyAltVoucher")} />
              <Field label="Discharge — voucher type" placeholder={DEFAULT_LEDGER_NAMES.dischargeVoucher} value={ledgerNames.dischargeVoucher} onChange={update("dischargeVoucher")} />
              <Field label="Discharge — alternate" placeholder={DEFAULT_LEDGER_NAMES.dischargeAltVoucher} value={ledgerNames.dischargeAltVoucher} onChange={update("dischargeAltVoucher")} />
            </div>
          </div>

          <div className="drawer-section">
            <h3>Dialysis receivables</h3>
            <p className="lede small">
              Voucher type used when booking dialysis receivables from schemes
              (KASP, MEDISEP, etc.). Company ledger is auto-composed as
              "{`{Company}-dialysis`}".
            </p>
            <div className="field-stack">
              <Field label="Dialysis — voucher type" placeholder={DEFAULT_LEDGER_NAMES.dialysisVoucher} value={ledgerNames.dialysisVoucher} onChange={update("dialysisVoucher")} />
              <Field label="Dialysis — alternate" placeholder={DEFAULT_LEDGER_NAMES.dialysisAltVoucher} value={ledgerNames.dialysisAltVoucher} onChange={update("dialysisAltVoucher")} />
            </div>
          </div>
        </div>

        <div className="drawer-foot">
          <Button variant="primary" onClick={onClose}>Done</Button>
        </div>
      </div>
    </>
  );
}
