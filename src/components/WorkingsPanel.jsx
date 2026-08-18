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
      <div className="drawer" role="dialog" aria-label="Settings">
        <div className="drawer-head">
          <div>
            <p className="eyebrow">Settings</p>
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
            blank to fall back to its default. When a bill number repeats
            within a sheet, the repeat is placed on its own additional sheet
            (e.g. "RECEIPT (2)") — the Voucher Type name stays the same, so
            there's nothing to configure for that.
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
            <h3>Receipts voucher types</h3>
            <div className="field-stack">
              <Field label="Receipt" placeholder={DEFAULT_LEDGER_NAMES.receiptVoucher} value={ledgerNames.receiptVoucher} onChange={update("receiptVoucher")} />
              <Field label="Payment" placeholder={DEFAULT_LEDGER_NAMES.paymentVoucher} value={ledgerNames.paymentVoucher} onChange={update("paymentVoucher")} />
              <Field label="Credit Note" placeholder={DEFAULT_LEDGER_NAMES.creditNoteVoucher} value={ledgerNames.creditNoteVoucher} onChange={update("creditNoteVoucher")} />
            </div>
          </div>

          <div className="drawer-section">
            <h3>Sales voucher types</h3>
            <div className="field-stack">
              <Field label="Others" placeholder={DEFAULT_LEDGER_NAMES.othersVoucher} value={ledgerNames.othersVoucher} onChange={update("othersVoucher")} />
              <Field label="Pharmacy" placeholder={DEFAULT_LEDGER_NAMES.pharmacyVoucher} value={ledgerNames.pharmacyVoucher} onChange={update("pharmacyVoucher")} />
              <Field label="Discharge" placeholder={DEFAULT_LEDGER_NAMES.dischargeVoucher} value={ledgerNames.dischargeVoucher} onChange={update("dischargeVoucher")} />
            </div>
          </div>

          <div className="drawer-section">
            <h3>Discount ledgers</h3>
            <p className="lede small">
              Ledger used for the Discount line on each sheet — can differ per
              sheet since some Tally setups book discounts separately by
              category.
            </p>
            <div className="field-stack">
              <Field label="Others — discount ledger" placeholder={DEFAULT_LEDGER_NAMES.discountOthers} value={ledgerNames.discountOthers} onChange={update("discountOthers")} />
              <Field label="Pharmacy — discount ledger" placeholder={DEFAULT_LEDGER_NAMES.discountPharmacy} value={ledgerNames.discountPharmacy} onChange={update("discountPharmacy")} />
              <Field label="Discharge — discount ledger" placeholder={DEFAULT_LEDGER_NAMES.discountDischarge} value={ledgerNames.discountDischarge} onChange={update("discountDischarge")} />
            </div>
          </div>

          <div className="drawer-section">
            <h3>Pharmacy GST ledgers</h3>
            <p className="lede small">
              Used on the CGST/SGST lines for taxable pharmacy bills.
            </p>
            <div className="field-stack">
              <Field label="CGST ledger" placeholder={DEFAULT_LEDGER_NAMES.cgst} value={ledgerNames.cgst} onChange={update("cgst")} />
              <Field label="SGST ledger" placeholder={DEFAULT_LEDGER_NAMES.sgst} value={ledgerNames.sgst} onChange={update("sgst")} />
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
              <Field label="Dialysis voucher type" placeholder={DEFAULT_LEDGER_NAMES.dialysisVoucher} value={ledgerNames.dialysisVoucher} onChange={update("dialysisVoucher")} />
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
