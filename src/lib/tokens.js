// Ledger label placeholders. Rows store these tokens instead of a literal
// ledger name so the actual Tally ledger name (set via Workings) can change
// without reprocessing the source file.
export const LEDGER_TOKEN = { CASH: "$CASH$", CARD: "$CARD$", NEFT: "$NEFT$" };

// Voucher-type tokens — one per output category. Every occurrence of a bill
// number uses the SAME base voucher type; duplicates are separated into
// additional sheets instead of a renamed voucher type (see pushOccurrence /
// bucketSheetName below), so there is no separate "alternate" token anymore.
export const RECEIPT_TOKEN = "$RECEIPT$";
export const PAYMENT_TOKEN = "$PAYMENT$";
export const CREDITNOTE_TOKEN = "$CREDITNOTE$";
export const OTHERS_TOKEN = "$OTHERS$";
export const PHARMACY_TOKEN = "$PHARMACY$";
export const DISCHARGE_TOKEN = "$DISCHARGE$";
export const DIALYSIS_TOKEN = "$DIALYSIS$";

// Non-payment-mode ledger tokens — discount and GST lines, which also vary
// by Tally setup and by sheet (Others / Pharmacy / Discharge each may use a
// different discount ledger).
export const DISCOUNT_OTHERS_TOKEN = "$DISCOUNT_OTHERS$";
export const DISCOUNT_PHARMACY_TOKEN = "$DISCOUNT_PHARMACY$";
export const DISCOUNT_DISCHARGE_TOKEN = "$DISCOUNT_DISCHARGE$";
export const CGST_TOKEN = "$CGST$";
export const SGST_TOKEN = "$SGST$";

export const DEFAULT_LEDGER_NAMES = {
  cash: "CASH",
  card: "CARD",
  neft: "UPI/NEFT",

  // Receipts side — voucher type per sheet
  receiptVoucher: "Receipt",
  paymentVoucher: "Payment",
  creditNoteVoucher: "Credit Note",

  // Bills side — voucher type per sheet
  othersVoucher: "Sales",
  pharmacyVoucher: "Sales",
  dischargeVoucher: "Sales",

  // Bills side — discount ledger per sheet (can differ across Tally setups)
  discountOthers: "Discount",
  discountPharmacy: "Discount",
  discountDischarge: "Discount",

  // Pharmacy GST ledgers
  cgst: "CGST",
  sgst: "SGST",

  // Dialysis (receivables from insurance/scheme companies) side
  dialysisVoucher: "Journal-Dialysis",
};

function resolveSingle(value, ledgerNames) {
  if (value === LEDGER_TOKEN.CASH) return ledgerNames.cash || DEFAULT_LEDGER_NAMES.cash;
  if (value === LEDGER_TOKEN.CARD) return ledgerNames.card || DEFAULT_LEDGER_NAMES.card;
  if (value === LEDGER_TOKEN.NEFT) return ledgerNames.neft || DEFAULT_LEDGER_NAMES.neft;
  if (value === RECEIPT_TOKEN) return ledgerNames.receiptVoucher || DEFAULT_LEDGER_NAMES.receiptVoucher;
  if (value === PAYMENT_TOKEN) return ledgerNames.paymentVoucher || DEFAULT_LEDGER_NAMES.paymentVoucher;
  if (value === CREDITNOTE_TOKEN)
    return ledgerNames.creditNoteVoucher || DEFAULT_LEDGER_NAMES.creditNoteVoucher;
  if (value === OTHERS_TOKEN) return ledgerNames.othersVoucher || DEFAULT_LEDGER_NAMES.othersVoucher;
  if (value === PHARMACY_TOKEN)
    return ledgerNames.pharmacyVoucher || DEFAULT_LEDGER_NAMES.pharmacyVoucher;
  if (value === DISCHARGE_TOKEN)
    return ledgerNames.dischargeVoucher || DEFAULT_LEDGER_NAMES.dischargeVoucher;
  if (value === DIALYSIS_TOKEN)
    return ledgerNames.dialysisVoucher || DEFAULT_LEDGER_NAMES.dialysisVoucher;
  if (value === DISCOUNT_OTHERS_TOKEN)
    return ledgerNames.discountOthers || DEFAULT_LEDGER_NAMES.discountOthers;
  if (value === DISCOUNT_PHARMACY_TOKEN)
    return ledgerNames.discountPharmacy || DEFAULT_LEDGER_NAMES.discountPharmacy;
  if (value === DISCOUNT_DISCHARGE_TOKEN)
    return ledgerNames.discountDischarge || DEFAULT_LEDGER_NAMES.discountDischarge;
  if (value === CGST_TOKEN) return ledgerNames.cgst || DEFAULT_LEDGER_NAMES.cgst;
  if (value === SGST_TOKEN) return ledgerNames.sgst || DEFAULT_LEDGER_NAMES.sgst;
  return value;
}

export function resolveLedgerToken(value, ledgerNames) {
  if (typeof value !== "string") return value;
  return resolveSingle(value, ledgerNames);
}

export function resolveRows(rows, ledgerNames) {
  return rows.map((row) => row.map((cell) => resolveLedgerToken(cell, ledgerNames)));
}

// --- Occurrence bucketing -------------------------------------------------
// When the same Bill No. appears more than once within one output category
// (e.g. two Receipt vouchers, or two Discharge SALES vouchers), Tally
// rejects the upload if both land in the same import batch. Rather than
// rename the voucher type for the repeat (which changes what shows up in
// Tally's voucher register), each occurrence is routed to its own bucket —
// bucket 1 holds every bill's first occurrence, bucket 2 holds only the
// bill numbers that repeated a second time, and so on. Uploading bucket 1,
// then bucket 2, etc. as separate Tally imports avoids the collision while
// every voucher keeps the same Voucher Type name.

// Returns the 1-based occurrence number for `key` and advances the map.
export function nextOccurrence(occurrenceMap, key) {
  const idx = (occurrenceMap.get(key) || 0) + 1;
  occurrenceMap.set(key, idx);
  return idx;
}

// Pushes `rows` into `buckets[occurrenceIndex - 1]`, creating the bucket
// array if this is its first row.
export function pushToBucket(buckets, occurrenceIndex, rows) {
  const i = occurrenceIndex - 1;
  if (!buckets[i]) buckets[i] = [];
  buckets[i].push(...rows);
}

// Sheet name for a given occurrence — the base name for occurrence 1,
// "Base (2)", "Base (3)", etc. for repeats.
export function bucketSheetName(base, occurrenceIndex) {
  return occurrenceIndex <= 1 ? base : `${base} (${occurrenceIndex})`;
}
