# Digit 2W Processor

Converts Digit insurance commission grid Excel files into structured rule engine upload files.

## Setup

### Backend (Flask API)
```bash
pip install -r requirements.txt
python app.py
# Runs on http://localhost:5050
```

### Frontend (React)
```bash
npm install
npm run dev
# Runs on http://localhost:3000
```

> If using the vite proxy (`/api → localhost:5050`), change `const API = ""` in App.jsx.
> If running frontend standalone without vite, keep `const API = "http://localhost:5050"`.

## Usage

1. Upload the Digit 2W Excel grid (.xlsx)
2. The app auto-detects available states from the file
3. Set the effect start/end dates (end date auto-fills to last day of month)
4. Optionally select specific states (leave all unselected = process all)
5. Click **Generate Rule Files**
6. Download individual state files or the full ZIP

## CLI (no frontend)
```bash
python run.py input.xlsx ./output 2026-02-01 2026-02-28
python run.py input.xlsx ./output 2026-02-01 2026-02-28 DELHI MAHARASHTRA
```

## File structure
```
app.py          Flask API server
processor.py    Core processing logic
run.py          CLI entrypoint
App.jsx         React frontend (place in src/ or root)
index.html      HTML entry point
vite.config.js  Vite dev server config
package.json    npm dependencies
requirements.txt Python dependencies
```
