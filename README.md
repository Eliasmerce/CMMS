# CMMS Workbook Analyzer App

A lightweight browser app that turns your refinery maintenance workbook structure into a functional assessment report.

## What it does
- Lets you define/edit major workbook sheets and their business purpose.
- Shows an explicit data flow from form entry to KPI and chart output.
- Allows risk and improvement selection with operational focus.
- Generates a markdown report you can copy/download.

## Run locally
No build step required.

```bash
python3 -m http.server 8000
```

Then open `http://localhost:8000` in your browser.

## Files
- `index.html` – app layout.
- `styles.css` – visual styling.
- `app.js` – app behavior and report generation.
