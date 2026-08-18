// Ledger label placeholders. Rows store these tokens instead of a literal
// ledger name so the actual Tally ledger name (set via Workings) can change
// without reprocessing the source file.
export const LEDGER_TOKEN = { CASH: "$CASH$", CARD: "$CARD$", NEFT: "$NEFT$" };

// Voucher-type tokens for occurrence 2 (auto-resolved duplicates) — one per
// output category. Occurrence 1 uses the base voucher type name, occurrence 3+
// uses the encoded suffix form "$TOKEN$||N".
export const RECEIPT_ALT_TOKEN = "$RECEIPT_ALT$";
export const OTHERS_ALT_TOKEN = "$OTHERS_ALT$";
export const PHARMACY_ALT_TOKEN = "$PHARMACY_ALT$";
export const DISCHARGE_ALT_TOKEN = "$DISCHARGE_ALT$";
export const DIALYSIS_ALT_TOKEN = "$DIALYSIS_ALT$";

// Base voucher-type tokens — occurrence-1 rows use these so their string can
// still be renamed via Workings without reprocessing.
export const OTHERS_TOKEN = "$OTHERS$";
export const PHARMACY_TOKEN = "$PHARMACY$";
export const DISCHARGE_TOKEN = "$DISCHARGE$";
export const DIALYSIS_TOKEN = "$DIALYSIS$";

export const DEFAULT_LEDGER_NAMES = {
  cash: "CASH",
  card: "CARD",
  neft: "UPI/NEFT",

  // Receipts side
  altVoucherType: "Receipt 2",

  // Bills side — base voucher type per sheet
  othersVoucher: "Sales",
  pharmacyVoucher: "Sales",
  dischargeVoucher: "Sales",

  // Bills side — alt voucher type per sheet (used when the same bill no.
  // appears more than once inside a single sheet so Tally doesn't reject the
  // upload as a duplicate voucher).
  othersAltVoucher: "Sales 2",
  pharmacyAltVoucher: "Sales 2",
  dischargeAltVoucher: "Sales 2",

  // Dialysis (receivables from insurance/scheme companies) side
  dialysisVoucher: "Journal-Dialysis",
  dialysisAltVoucher: "Journal-Dialysis 2",
};

function resolveSingle(value, ledgerNames) {
  if (value === LEDGER_TOKEN.CASH) return ledgerNames.cash || DEFAULT_LEDGER_NAMES.cash;
  if (value === LEDGER_TOKEN.CARD) return ledgerNames.card || DEFAULT_LEDGER_NAMES.card;
  if (value === LEDGER_TOKEN.NEFT) return ledgerNames.neft || DEFAULT_LEDGER_NAMES.neft;
  if (value === RECEIPT_ALT_TOKEN)
    return ledgerNames.altVoucherType || DEFAULT_LEDGER_NAMES.altVoucherType;
  if (value === OTHERS_TOKEN)
    return ledgerNames.othersVoucher || DEFAULT_LEDGER_NAMES.othersVoucher;
  if (value === PHARMACY_TOKEN)
    return ledgerNames.pharmacyVoucher || DEFAULT_LEDGER_NAMES.pharmacyVoucher;
  if (value === DISCHARGE_TOKEN)
    return ledgerNames.dischargeVoucher || DEFAULT_LEDGER_NAMES.dischargeVoucher;
  if (value === OTHERS_ALT_TOKEN)
    return ledgerNames.othersAltVoucher || DEFAULT_LEDGER_NAMES.othersAltVoucher;
  if (value === PHARMACY_ALT_TOKEN)
    return ledgerNames.pharmacyAltVoucher || DEFAULT_LEDGER_NAMES.pharmacyAltVoucher;
  if (value === DISCHARGE_ALT_TOKEN)
    return ledgerNames.dischargeAltVoucher || DEFAULT_LEDGER_NAMES.dischargeAltVoucher;
  if (value === DIALYSIS_TOKEN)
    return ledgerNames.dialysisVoucher || DEFAULT_LEDGER_NAMES.dialysisVoucher;
  if (value === DIALYSIS_ALT_TOKEN)
    return ledgerNames.dialysisAltVoucher || DEFAULT_LEDGER_NAMES.dialysisAltVoucher;
  return value;
}

export function resolveLedgerToken(value, ledgerNames) {
  if (typeof value !== "string") return value;
  // Encoded form "$TOKEN$||N" for occurrence 3+: resolve base then append N.
  const sep = value.indexOf("||");
  if (sep > 0 && value.startsWith("$")) {
    const base = value.substring(0, sep);
    const suffix = value.substring(sep + 2);
    return `${resolveSingle(base, ledgerNames)} ${suffix}`;
  }
  return resolveSingle(value, ledgerNames);
}

export function resolveRows(rows, ledgerNames) {
  return rows.map((row) => row.map((cell) => resolveLedgerToken(cell, ledgerNames)));
}

// Decides the Voucher Type string / token to store for a given occurrence
// of a bill number within its sheet. Occurrence 1 uses the base value,
// occurrence 2 the alt, and 3+ the encoded suffix form so the base can
// still be renamed via Workings without reprocessing.
export function voucherTypeForOccurrence(baseValue, occurrenceIndex, altValue = null) {
  if (occurrenceIndex <= 1) return baseValue;
  if (occurrenceIndex === 2 && altValue) return altValue;
  const isToken = typeof baseValue === "string" && baseValue.startsWith("$") && baseValue.endsWith("$");
  if (isToken) return `${baseValue}||${occurrenceIndex}`;
  return `${baseValue} ${occurrenceIndex}`;
}
