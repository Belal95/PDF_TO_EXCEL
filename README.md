# تقارير وكلاء البيع - PDF إلى Excel

A browser-only tool for daily Arabic sales-agent reports. Select or drop one or more original PDF reports; the app processes them automatically and enables a single Excel download.

## Output

- One `data` worksheet with the existing nine Arabic columns and right-to-left layout.
- Exactly four agent rows for each successfully extracted report date: فوري، خالص، ضامن، الاهلي.
- If an agent is absent, all seven numeric columns for that agent are written as numeric zero. Existing values are preserved.
- Dates come from the labelled **من تاريخ / إلى تاريخ** report interval, including PDFs where labels and values have slightly different baselines. The header creation date and PDF metadata are never used as a fallback.
- Rows sort chronologically, then by the agent order above. Identical duplicates are skipped before missing agents are filled; conflicting duplicates produce an error.

## Supported reports

Use original text-based daily PDFs in the existing ENR layout. Images/scans, unreadable agent rows, missing report intervals, and multi-day aggregate reports produce explicit errors rather than fabricated zero rows or an assumed date. Continuation pages can inherit their report date from the preceding page.

If some PDFs fail, the download includes the successful files and the page clearly lists the failed files. Removing or adding a file rebuilds the download. There is no existing-workbook upload or append mode.

Files are processed locally in the browser with PDF.js and ExcelJS; no report data is uploaded to a backend. The interface uses an Arabic font from Google Fonts with local fallbacks.

## Development

```sh
npm ci
npm run dev
npm test
npm run lint
npm run build
```

Deploy the Vite `dist` output on the existing Vercel project. The app has no server environment variables or API credentials.

Regression tests cover report versus creation dates, Arabic numerals, label/value baseline differences, missing agents, duplicate handling, calendar sorting, malformed rows, continuation pages, and the serialized Excel workbook. Real downloaded reports are used only for local validation and are not committed to this repository.
