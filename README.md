# Insurance Rule Automation Engine

**Turn a multi-sheet TW commission grid into certified, per-state insurance rule files — in one click.**

🔗 **Live app:** [insurance-rule-automation-engine.onrender.com](https://insurance-rule-automation-engine.onrender.com/)

---

## What it does

Insurance ops teams receive commission/payout grids for two-wheeler (TW) policies as sprawling multi-sheet Excel workbooks, one sheet per product structure (1+5, 1+1/SATP, SAOD), each with its own cluster, RTO, and fuel-type logic. Turning that into the flat, 50-column rule format an insurer's rating engine actually ingests is normally a manual, error-prone exercise repeated for every state.

This tool automates the whole pipeline:

1. **Upload** the source commission grid workbook (`.xlsx`)
2. **Set** the effective date range for the rules
3. **Pick** one, several, or all states found in the file
4. **Generate** a fully-formatted output workbook per state — ready to hand to the rating engine — plus a single zip of everything

## How it works

The frontend is a single-page operator console; the backend is a stateless Flask API that does the actual grid transformation.

```
┌─────────────┐        ┌──────────────────────────────┐
│   React UI   │  --->  │         Flask API             │
│  (upload,    │        │  /api/states   /api/process   │
│  date range, │  <---  │  /api/files    /api/download  │
│  state pick) │        └──────────────┬────────────────┘
└─────────────┘                        │
                                        ▼
                            ┌───────────────────────┐
                            │   processor.py         │
                            │  grid parsing engine   │
                            └───────────────────────┘
```

## Tech stack

| Layer     | Stack |
|-----------|-------|
| Frontend  | React (hooks-based, no external UI framework) |
| Backend   | Flask, Flask-CORS |
| Data      | openpyxl (Excel read/write) |
| Hosting   | Render |


*Built to replace a manual, sheet-by-sheet grid conversion process with a single upload-and-generate flow.*
