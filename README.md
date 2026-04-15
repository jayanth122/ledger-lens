# Ledger Lens

Minimal Next.js app that takes an Excel file and shows cumulative savings over time.

The app follows this logic:

- Read first worksheet from uploaded workbook.
- Use `Period`, `Amount`, and `IncomeExpense` columns.
- Exclude rows where income/expense label contains `Transfer`.
- Per month:
	- Income = sum of rows where label contains `Income`
	- Expense = sum of rows where label contains `Exp`
	- Net savings = Income - Expense
- Cumulative savings = running sum of monthly net savings.

## Tech

- Next.js (App Router)
- React + TypeScript
- Tailwind CSS
- Recharts
- SheetJS (`xlsx`)

## Local development

```bash
npm install
npm run dev
```

Build check:

```bash
npm run build
```

## Deploy on Vercel (GitHub)

1. Push this repository to GitHub.
2. In Vercel, click **Add New Project** and import the GitHub repo.
3. Keep the defaults:
	 - Framework Preset: **Next.js**
	 - Build Command: `next build`
	 - Output Directory: leave empty
4. Click **Deploy**.

No environment variables are required for this app.

## Notes

- Parsing and chart computation happen in the browser after file upload.
- Uploaded file contents are not sent to a backend by this app.
