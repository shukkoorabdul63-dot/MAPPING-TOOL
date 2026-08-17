// Ledger label placeholders. Rows store these tokens instead of a literal
// ledger name, so the actual Tally ledger name (set via the "Workings"
// panel) can change without reprocessing the source file.
export const LEDGER_TOKEN = { CASH: "$CASH$", CARD: "$CARD$", NEFT: "$NEFT$" };

// Voucher-type token for a duplicated bill number's second Receipt — lets
// the alternate voucher type name (e.g. "Receipt 2") be edited later
// without reprocessing.
export const RECEIPT_ALT_TOKEN = "$RECEIPT_ALT$";

export const DEFAULT_LEDGER_NAMES = {
  cash: "CASH",
  card: "CARD",
  neft: "UPI/NEFT",
  altVoucherType: "Receipt 2",
};

export function resolveLedgerToken(value, ledgerNames) {
  if (value === LEDGER_TOKEN.CASH) return ledgerNames.cash || DEFAULT_LEDGER_NAMES.cash;
  if (value === LEDGER_TOKEN.CARD) return ledgerNames.card || DEFAULT_LEDGER_NAMES.card;
  if (value === LEDGER_TOKEN.NEFT) return ledgerNames.neft || DEFAULT_LEDGER_NAMES.neft;
  if (value === RECEIPT_ALT_TOKEN)
    return ledgerNames.altVoucherType || DEFAULT_LEDGER_NAMES.altVoucherType;
  return value;
}

export function resolveRows(rows, ledgerNames) {
  return rows.map((row) => row.map((cell) => resolveLedgerToken(cell, ledgerNames)));
}

// Decides the Voucher Type string to use for a given occurrence of a
// duplicated bill number. Occurrence 1 keeps the plain type. Occurrence 2
// of a Receipt uses the user-configurable alternate name (default
// "Receipt 2"). Anything beyond that (rare) gets an auto-numbered suffix
// so it never collides with an already-used Voucher Type + Voucher No.
// combination.
export function voucherTypeForOccurrence(baseType, occurrenceIndex) {
  if (occurrenceIndex <= 1) return baseType;
  if (occurrenceIndex === 2 && baseType === "Receipt") return RECEIPT_ALT_TOKEN;
  return `${baseType} ${occurrenceIndex}`;
}
