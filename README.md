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
  - Negative Cash (sales return): Party debited, Cash credited.
    Voucher Type = "Payment".
  - Negative IP Credit alone: excluded here — will be handled as a
    SALES-side credit note once that mapping is built.
  - Duplicate Bill Numbers across the export: excluded from the main
    output and listed on the "NEEDS REVIEW" sheet.
- Downloaded workbook has three sheets: RECEIPT (Tally-ready), NEEDS
  REVIEW (duplicate bill numbers to fix manually), and SUMMARY BY BILL
  TYPE (Cash/Card/NEFT totals per Teja bill section — e.g. REGISTRATION,
  PHARMACY CASH BILL, ADVANCE, DISCHARGE BILL — for reconciling against
  Teja's own report before you trust the output).
- In-app preview toggle shows the first 12 output rows before download.
- SALES sheet (billing details) mapping is not yet built — pending the
  Teja billing-details export.

## Tech

React + Vite, [SheetJS (xlsx)](https://sheetjs.com/) for parsing and
generating Excel files client-side. Tested against a real 5,000-row
export without lag; designed to comfortably handle 50,000+ rows since
all processing is a single in-memory pass.
