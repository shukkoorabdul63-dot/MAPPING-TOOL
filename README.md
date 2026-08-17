# Teja → Tally Receipt Mapper

Converts a Teja "Detailed Bill Register" export into a Tally-uploadable
RECEIPT mapping sheet, and flags duplicate bill numbers separately for
manual review. Runs entirely in the browser — nothing is uploaded to a
server, so patient billing data never leaves your machine.

## Run it locally

```bash
npm install
npm run dev
```

Open the URL it prints (usually http://localhost:5173).

## Build for deployment

```bash
npm run build
```

This creates a `dist/` folder with static files you can host anywhere
(Vercel, Netlify, GitHub Pages, or just open `dist/index.html` directly).

## Deploy to GitHub + Vercel (same pattern as Finova)

1. Create a new GitHub repo and push this project:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<your-username>/teja-receipt-mapper.git
   git push -u origin main
   ```
2. Go to [vercel.com](https://vercel.com), import the repo, and deploy —
   Vercel auto-detects the Vite config, no settings needed.

## Current scope

- Handles the RECEIPT sheet only (Cash / Card / NEFT).
- Rules implemented:
  - Only source rows with a Bill Number are processed.
  - Normal bill (unique Bill No, cash/card/neft > 0): Party credited,
    payment mode(s) debited. Voucher Type = "Receipt".
  - Negative Cash (sales return, cash refunded): produces **two**
    vouchers — a Payment (Party debited, Cash ledger credited) and a
    Credit Note (Bill Name debited, Party credited — reverses the income).
  - Negative IP Credit alone (no cash movement, just a billing
    revision): produces a Credit Note only (Bill Name debited, Party
    credited).
  - Duplicate Bill Numbers across the export: no longer excluded. The
    repeat occurrence is kept in the normal voucher sheets, but posted
    under an alternate Voucher Type (default "Receipt 2", editable via
    Workings) so the Voucher Type + Voucher No. combination Tally sees
    is unique. A third+ occurrence (rare) auto-numbers further
    ("Payment 3", "Credit Note 3", etc.). Logged on the "DUPLICATE LOG"
    sheet for visibility, not action.
- Downloaded workbook has six sheets:
  - **RECEIPT** — normal receipts, Tally-ready
  - **PAYMENT** — cash refunds
  - **CREDIT NOTE** — income reversals (Bill Name Dr / Party Cr)
  - **DUPLICATE LOG** — informational list of repeat bill numbers and
    how they were resolved
  - **SUMMARY BY BILL TYPE** — Cash/Card/NEFT totals per Teja bill
    section (REGISTRATION, PHARMACY CASH BILL, ADVANCE, etc.), for
    reconciling against Teja's own report
  - **LEDGERS USED** — records which actual Tally ledger name was used
    for Cash/Card/NEFT in this export
- **Workings button** (top right) opens a panel to set the actual Tally
  ledger name used for each payment mode — e.g. Card might post to
  "Swipe Control", NEFT might post to "Google Pay Account" instead of
  the generic mode name — and the alternate Voucher Type name used for
  duplicate bill numbers. Applies across the RECEIPT and PAYMENT sheets
  and the on-screen preview without needing to reprocess the file.
- In-app preview toggle shows a sample of rows from all three voucher
  sheets before download.
- SALES sheet (billing details) mapping is not yet built — pending the
  Teja billing-details export.

## Tech

React + Vite, [SheetJS (xlsx)](https://sheetjs.com/) for parsing and
generating Excel files client-side. Tested against a real 5,000-row
export without lag; designed to comfortably handle 50,000+ rows since
all processing is a single in-memory pass.
