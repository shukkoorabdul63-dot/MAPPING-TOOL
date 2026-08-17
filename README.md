# Teja → Tally Mapping Tool

Turns Teja hospital-software exports (receipts and bills) into
Tally-ready mapping workbooks. Everything runs in the browser — patient
billing data never leaves your machine.

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

Creates a `dist/` folder you can host anywhere (Vercel, Netlify, GitHub
Pages, or just open `dist/index.html` directly).

## Deploy to GitHub + Vercel

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/teja-mapping-tool.git
git push -u origin main
```

Then import the repo on [vercel.com](https://vercel.com) — auto-detects
Vite.

## Sections

The app has a left sidebar with two sections plus Workings settings.

### Receipts — Detailed Bill Register
Upload the Detailed Bill Register export from Teja (`.xlsx` or `.xls`).
Downloaded workbook has six sheets:

- **RECEIPT** — normal receipts (Party Cr / Cash-Card-NEFT Dr)
- **PAYMENT** — cash refunds (Party Dr / Cash Cr)
- **CREDIT NOTE** — income reversals (Bill Name Dr / Party Cr)
- **DUPLICATE LOG** — repeat bill numbers and how they were resolved
- **SUMMARY BY BILL TYPE** — Cash/Card/NEFT totals per section
- **LEDGERS USED** — records which Tally ledger names were applied

Duplicate bill numbers are auto-resolved by posting the repeat
occurrence under an alternate Voucher Type (set in Workings, default
"Receipt 2") so uploads don't fail on duplicate voucher-no.

### Bills — Bill analysis
Upload the bill analysis export from Teja. Cancelled bills
(`BILLCANCELLED = Y`) are removed. Rows sharing both Bill No. and
Bill Name are treated as clear duplicates and the whole group is
removed (logged separately). Remaining bills split three ways:

- **Others (SALES journal)** — everything except Discharge and
  Pharmacy. Fully mapped:
  - DR: Party (`OPNUMBER - PATIENT NAME`), amount =
    BillAmount + BillAdvance − IPCreditReturn
  - CR: Income Head (`{BillName} INCOME-{IPOP}`),
    amount = GrossAmount, Cost Center = DoctorCode
  - Extra DR line: Discount, when BillDiscount > 0
- **Discharge (raw)** — kept as a raw sheet, rules pending
- **Pharmacy (raw)** — combined IP + OP pharmacy, rules pending

### Workings (sidebar link)
Opens a slide-over panel to set the actual Tally ledger names for Cash,
Card, and NEFT (e.g. Card might post to "Swipe Control", NEFT to a
"Google Pay Account" ledger), plus the alternate Voucher Type used for
duplicate bill numbers. Applies across the mapped workbook and the
on-screen preview without reprocessing the source file.

## Tech

React + Vite, [SheetJS (xlsx)](https://sheetjs.com/) for parsing and
generating Excel files client-side. Handles both `.xlsx` and legacy
`.xls`. Comfortably handles 50,000+ rows since processing is a single
in-memory pass.

## Project structure

```
src/
├── App.jsx                     # shell (sidebar + main + drawer)
├── main.jsx                    # entry
├── styles.css                  # all styles
├── components/
│   ├── Sidebar.jsx
│   ├── ReceiptsView.jsx
│   ├── BillsView.jsx
│   ├── WorkingsPanel.jsx
│   └── shared.jsx              # Toggle, Button, StatRow, Dropzone
└── lib/
    ├── utils.js                # numbers, dates, header parsing
    ├── tokens.js               # ledger tokens + voucher-type helpers
    ├── receipts.js             # Detailed Bill Register logic
    └── bills.js                # Bill analysis logic
```
