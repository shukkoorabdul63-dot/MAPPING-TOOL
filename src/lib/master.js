import * as XLSX from "xlsx";

/**
 * Collect party ledger names from both the receipts and bills results and
 * return a deduped, sorted list. Each entry keeps track of where it was
 * seen — useful for the audit/download.
 */
export function computePatientLedgerMaster(receiptResult, billsResult) {
  const map = new Map();

  const add = (party, source) => {
    if (!party || party === " - ") return;
    const key = party.trim();
    if (!key) return;
    if (!map.has(key)) {
      map.set(key, { partyName: key, inReceipts: false, inBills: false });
    }
    const entry = map.get(key);
    if (source === "receipts") entry.inReceipts = true;
    if (source === "bills") entry.inBills = true;
  };

  // Receipts: party names live as row[3] in receipt/payment/creditNote rows,
  // and only for the "party" leg (not the ledger legs like CASH/CARD/NEFT).
  // The simpler and more accurate route is to walk receiptResult's underlying
  // billNameSummary — but that only has aggregated bill names, not party
  // names. So instead we scan the row arrays and pick anything that looks like
  // a "id - name" party string.
  const scanRows = (rows, source) => {
    if (!rows) return;
    for (const row of rows) {
      const partyCell = row[3];
      if (typeof partyCell !== "string") continue;
      if (partyCell.includes(" - ")) add(partyCell, source);
    }
  };

  if (receiptResult) {
    scanRows(receiptResult.receiptRows, "receipts");
    scanRows(receiptResult.paymentRows, "receipts");
    scanRows(receiptResult.creditNoteRows, "receipts");
  }

  if (billsResult && billsResult.cleanCandidates) {
    for (const c of billsResult.cleanCandidates) {
      if (!c.opNumber && !c.patientName) continue;
      add(`${c.opNumber} - ${c.patientName}`, "bills");
    }
  }

  const list = [...map.values()].sort((a, b) => a.partyName.localeCompare(b.partyName));
  const inBoth = list.filter((r) => r.inReceipts && r.inBills).length;
  const onlyReceipts = list.filter((r) => r.inReceipts && !r.inBills).length;
  const onlyBills = list.filter((r) => !r.inReceipts && r.inBills).length;

  return {
    list,
    stats: { total: list.length, inBoth, onlyReceipts, onlyBills },
  };
}

export function buildMasterWorkbook(list) {
  const wb = XLSX.utils.book_new();
  const rows = [
    ["Party Name (Tally Ledger)", "In Receipts", "In Bills"],
    ...list.map((r) => [r.partyName, r.inReceipts ? "Y" : "", r.inBills ? "Y" : ""]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 40 }, { wch: 14 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, "PATIENT LEDGERS");
  return wb;
}
