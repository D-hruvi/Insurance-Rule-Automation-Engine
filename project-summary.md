# Project Summary: AI for Payout Grid

## 1. Project Purpose

This project converts Digit two-wheeler insurance commission grid Excel files into structured rule-engine upload files.

The user uploads a source `.xlsx` workbook containing multiple Digit payout grid sheets. The backend reads the workbook, extracts RTO clusters, state mappings, vehicle segment rules, commission values, and product-specific payout information. It then creates one formatted output Excel file per selected state and also creates a ZIP file containing all generated state files.

The main business use case is:

1. Take a broker commission grid workbook.
2. Detect all states available in the workbook.
3. Let the user select an effect date range and optional state filters.
4. Generate 50-column rule-engine Excel files for Digit two-wheeler payout rules.
5. Download individual state files or one ZIP containing all files.

## 2. Technology Stack

### Backend

- Python
- Flask
- Flask-CORS
- OpenPyXL

### Frontend

- React 18
- Vite
- Plain inline CSS styles inside `App.jsx`

### Output

- `.xlsx` files generated with OpenPyXL
- ZIP archive generated with Python `zipfile`

## 3. Project Structure

```text
AI-for-payout-grid/
+-- README.md
+-- project-summary.md
+-- backend/
|   +-- app.py
|   +-- processor.py
|   +-- requirements.txt
|   +-- run.py
+-- frontend/
    +-- index.html
    +-- package.json
    +-- package-lock.json
    +-- vite.config.js
    +-- src/
        +-- App.jsx
        +-- main.jsx
```

Generated/local folders also exist, such as:

- `frontend/node_modules/`
- `frontend/dist/`
- `backend/__pycache__/`

These are not core source files.

## 4. Backend Overview

The backend has two main files:

- `backend/app.py`: Flask API server.
- `backend/processor.py`: Core Excel processing and rule generation logic.

### 4.1 `backend/app.py`

This file exposes HTTP APIs used by the React frontend.

It creates a Flask app, enables CORS, accepts Excel uploads, calls functions from `processor.py`, stores uploaded files temporarily, writes generated output files, creates ZIP downloads, and returns JSON responses.

Important directories:

```python
UPLOAD_DIR = "/tmp/digit_uploads"
OUTPUT_DIR = "/tmp/digit_outputs"
```

On each upload or process run, the server creates a short `session_id` using UUID. That session id is used to separate generated files for each request.

### 4.2 Backend API Endpoints

#### `GET /api/health`

Checks whether the backend is running.

Response example:

```json
{
  "status": "ok",
  "timestamp": "2026-06-25T10:30:00"
}
```

#### `POST /api/states`

Uploads an Excel file and returns the detected states from the workbook.

Request:

- Form-data field: `file`
- File must end with `.xlsx`

What happens:

1. Backend saves the uploaded workbook.
2. Calls `load_input_data()`.
3. Calls `get_all_states()`.
4. Returns all state names found from RTO codes.

Response includes:

- `session_id`
- `file_path`
- `states`
- `total_states`

#### `POST /api/process`

Uploads an Excel file and generates output files.

Request form fields:

- `file`: input `.xlsx` workbook.
- `effect_start`: start date, for example `2026-02-01`.
- `effect_end`: end date, for example `2026-02-28`.
- `states`: optional JSON array of state names. If omitted, all states are processed.

What happens:

1. Backend validates the file and date fields.
2. Saves the input workbook.
3. Creates a session output folder.
4. Calls `process_all()` from `processor.py`.
5. Creates a ZIP containing all generated Excel files.
6. Returns generated file metadata and download URLs.

Response includes:

- `session_id`
- `files_generated`
- `files`
- `zip`
- `progress`

#### `GET /api/download/<session_id>/<filename>`

Downloads a generated Excel file or generated ZIP file.

#### `GET /api/files/<session_id>`

Lists all files generated for a session.

## 5. Processing Logic

All core conversion rules are in `backend/processor.py`.

The processor reads specific sheets from the input workbook:

- `2W RTO's`
- `TW 1+5`
- `TW 1+1 & SATP`
- `TW SAOD with Flexi Options`

The output file always uses the `OUTPUT_HEADERS` list, which contains 50 rule-engine columns.

### 5.1 Main Processing Functions

#### `load_input_data(path)`

Loads the source workbook in read-only mode and extracts:

- RTO to cluster mapping.
- 1+5 cluster data.
- 1+1 and SATP cluster data.
- SAOD with flexi option data.

It returns a dictionary containing normalized data used by all generator functions.

#### `get_all_states(d)`

Reads all RTO codes from loaded data and maps their first two letters to Indian state names using `RTO_STATE_NAMES`.

Example:

- `DL` -> `DELHI`
- `MH` -> `MAHARASHTRA`
- `KA` -> `KARNATAKA`

#### `process_all(input_path, output_dir, eff_s, eff_e, states=None, progress_callback=None)`

This is the main public processing function.

What it does:

1. Loads workbook data.
2. Detects all states.
3. Applies optional state filter.
4. Generates rows for each target state.
5. Writes one Excel output per state.
6. Returns a list of generated file paths.

Generated filenames look like:

```text
DELHI_2W.xlsx
MAHARASHTRA_2W.xlsx
UTTARPRADESH_2W.xlsx
```

#### `generate_for_state(d, state_name, eff_s, eff_e, counter_start=1)`

Generates all rule rows for one state.

Generation order:

1. 1+5 rows
2. SAOD rows
3. 1+1 rows

The function returns:

- generated rows
- updated rule counter

### 5.2 Rule Row Creation

The `build_row()` function creates one output row matching the 50-column rule-engine template.

Common values:

- `IC Code *`: `DIGIT`
- `Product Type *`: `TW`
- `Rule Type *`: `PAYINPAYOUT`
- `Owner Type`: `ALL`
- `Booking Mode`: `any`
- `Cover Selection Type`: `na`
- `Addon Selection Type`: `na`
- `NCB Type`: `na`
- POSP commission type: `net`
- POSP reward type: `percentage`

### 5.3 PayIn and POSP Logic

The `resolve_payin()` function converts raw commission values from the Excel grid.

Rules:

- Empty or zero value -> net 0
- `MISP` -> OD commission with `22.5`
- `D` -> net 0
- Numeric values are multiplied by 100 and stored as percentages

The `posp_amt()` function calculates POSP payout:

```text
POSP = PayIn * 0.8
```

Example:

- PayIn `20` -> POSP `16`
- PayIn `10` -> POSP `8`

### 5.4 Product Generators

#### 1+5 Generator: `gen_1p5()`

Reads data from `TW 1+5`.

It groups entries by cluster, maps clusters to RTOs, filters by state, and creates comprehensive new-business rows.

Segment handling includes:

- EV power bands
- Bike CC bands
- Scooter rows
- Mixed scooter/motorcycle rows
- All vehicle rows

Special handling:

- BAJAJ `3-7 KW` creates one payable EV power-band row and two zero-pay rows for outside bands.
- Make value `Others` becomes an exclude list:

```text
EXCLUDE: HERO MOTOCORP, BAJAJ, HONDA, ROYAL ENFIELD, TVS, SUZUKI, YAMAHA
```

#### SAOD Generator: `gen_saod()`

Reads data from `TW SAOD with Flexi Options`.

It creates SAOD rows for vehicle ages:

- Year 1
- Year 2
- Year 3
- Year 4

SAOD segment handling includes:

- `MC <155`
- `MC>155`
- `RE`
- `SC`
- `SC_EV`

Special note:

- Royal Enfield SAOD rows use vehicle type `Bike`.

#### 1+1 Generator: `gen_1p1()`

Reads data from `TW 1+1 & SATP`.

It creates:

- Comprehensive renewal/rollover rows
- TP rows
- Additional EV rows when the `SC/EV` segment exists

Ordering:

1. All RR rows for the cluster
2. All TP rows for the cluster
3. EV bonus rows at the end of the cluster block

### 5.5 Cluster and State Mapping

`RTO_STATE_NAMES` maps RTO prefixes to state names.

`CLUSTER_ABBR` maps source cluster labels to shorter rule-name prefixes.

Example rule names:

```text
DL_new_com_1_TW
MH_all_RR_10_TW
GJ_all_Od_20_TW
```

If a cluster does not exist in `CLUSTER_ABBR`, `_abbr()` creates a fallback abbreviation from the cluster text.

## 6. Output Excel Format

Each generated Excel workbook:

- Has one sheet named `Sheet1`.
- Uses the 50 headers from `OUTPUT_HEADERS`.
- Styles the header row with a dark blue fill and white bold text.
- Auto-adjusts column widths up to a maximum width of 45.

The backend also creates one ZIP file per processing session:

```text
Digit_2W_<session_id>.zip
```

## 7. Frontend Overview

The frontend is a React single-page app located in `frontend/src/App.jsx`.

It provides an operator-style UI called:

```text
Rule Engine Console
```

Main UI steps:

1. Upload Excel workbook.
2. Select effect start and end dates.
3. Optionally select states.
4. Generate rule files.
5. Download individual files or full ZIP.

### 7.1 Important Frontend State

The app tracks:

- `file`: selected `.xlsx` file.
- `effStart`: effect start date.
- `effEnd`: effect end date.
- `availableStates`: states detected from uploaded workbook.
- `selectedStates`: states selected by user.
- `status`: `idle`, `loading`, `success`, or `error`.
- `progress`: backend progress log or simulated frontend ticker.
- `result`: generated file metadata.
- `error`: UI error message.

### 7.2 Frontend Components

Main components inside `App.jsx`:

- `UploadZone`: drag/drop or click upload area.
- `StateSelector`: state filter grid.
- `ProgressRail`: progress display.
- `ResultManifest`: output file and ZIP download list.
- `Seal`: animated status indicator.
- `Step`: reusable section wrapper.
- `StatBlock`: top summary metric.

### 7.3 Frontend API Configuration

Current API setting in `frontend/src/App.jsx`:

```javascript
const API = "https://ai-for-payout-grid.onrender.com";
```

This means the frontend currently calls the deployed backend on Render.

The Vite config also contains a local API proxy:

```javascript
server: {
  port: 3000,
  proxy: {
    '/api': 'http://localhost:5050'
  }
}
```

To use the local Flask backend during development, change `API` in `App.jsx` to:

```javascript
const API = "";
```

Then requests like `/api/process` will be proxied by Vite to `http://localhost:5050`.

## 8. How The Full Flow Works

### Step-by-step runtime flow

1. User opens the React frontend.
2. User uploads a `.xlsx` workbook.
3. Frontend sends the workbook to:

```text
POST /api/states
```

4. Backend reads the workbook and returns detected states.
5. Frontend displays those states in the state selector.
6. User selects effect dates and optional states.
7. User clicks `GENERATE RULE FILES`.
8. Frontend sends form data to:

```text
POST /api/process
```

9. Backend saves the workbook in `/tmp/digit_uploads`.
10. Backend creates a session folder in `/tmp/digit_outputs`.
11. `processor.py` loads the workbook.
12. Processor generates rows for each target state.
13. Processor writes one `.xlsx` file per state.
14. Backend creates one ZIP file containing all generated state files.
15. Backend returns file metadata and download URLs.
16. Frontend shows the result manifest.
17. User downloads individual state files or the full ZIP.

## 9. Running The Project Locally

### 9.1 Backend Setup

From the project root:

```bash
cd backend
pip install -r requirements.txt
python app.py
```

Backend runs on:

```text
http://localhost:5050
```

Health check:

```text
http://localhost:5050/api/health
```

### 9.2 Frontend Setup

From the project root:

```bash
cd frontend
npm install
npm run dev
```

Frontend runs on:

```text
http://localhost:3000
```

### 9.3 Local API Setup

For full local development:

1. Start Flask backend on port `5050`.
2. Start Vite frontend on port `3000`.
3. In `frontend/src/App.jsx`, set:

```javascript
const API = "";
```

With this setup, Vite proxies frontend `/api` calls to the local backend.

### 9.4 Deployed API Setup

If using the deployed backend, keep:

```javascript
const API = "https://ai-for-payout-grid.onrender.com";
```

In this mode, the frontend sends API calls directly to Render.

## 10. CLI Usage Without Frontend

The backend also supports command-line processing through `backend/run.py`.

From the `backend` folder:

```bash
python run.py input.xlsx ./output 2026-02-01 2026-02-28
```

Process only selected states:

```bash
python run.py input.xlsx ./output 2026-02-01 2026-02-28 DELHI MAHARASHTRA
```

Arguments:

```text
python run.py <input_xlsx> <output_dir> <effect_start> <effect_end> [state1 state2 ...]
```

The CLI prints progress messages and writes generated files into the selected output directory.

## 11. Required Input Workbook Format

The workbook must be an `.xlsx` file and must contain these sheets:

```text
2W RTO's
TW 1+5
TW 1+1 & SATP
TW SAOD with Flexi Options
```

The processor expects data at specific column positions in each sheet.

### `2W RTO's`

Used for:

- RTO code
- 1+1 cluster
- 1+5 cluster
- SAOD cluster

### `TW 1+5`

Used for:

- Cluster
- Make
- Segment
- CD2 commission value

### `TW 1+1 & SATP`

Used for:

- Cluster
- Segment
- 1+1 CD2 value
- SATP CD2 value

### `TW SAOD with Flexi Options`

Used for:

- Cluster
- Segment
- Year 1 commission
- Year 2 commission
- Year 3 commission
- Year 4 commission

If any required sheet is missing or the expected columns change, the backend can fail with an error.

## 12. Important Business Rules Captured

- POSP is always `PayIn * 0.8`.
- Fuel ordering for petrol rows is:

```text
Petrol, LPG, Diesel, CNG
```

- EV rows use fuel type:

```text
Electric
```

- SAOD Royal Enfield rows use vehicle type `Bike`.
- SAOD cluster lookup is case-insensitive.
- `Others` make in 1+5 is converted to an EXCLUDE make list.
- 1+1 output ordering is RR rows first, then TP rows.
- EV bonus rows are appended at the end of each 1+1 cluster block.
- Power values are formatted as strings like `3.00` and `7.00`.
- Rule names use cluster abbreviations from `CLUSTER_ABBR`.

## 13. Error Handling

Backend validation includes:

- Missing uploaded file -> `400`
- Non-`.xlsx` file in `/api/states` -> `400`
- Missing effect dates in `/api/process` -> `400`
- Missing generated file on download -> `404`
- Processing exceptions -> `500` with error message and traceback

Frontend behavior:

- Shows an error if no file is selected.
- Shows an error if dates are missing.
- Shows backend error messages if processing fails.
- Falls back to a static India state list if state detection fails.

## 14. Known Notes And Considerations

1. Temporary upload/output paths use `/tmp/...`, which works well on Linux deployments. On Windows local development, Python usually still accepts these as root-style paths for the current drive, but using `tempfile.gettempdir()` would be more portable.

2. The frontend currently uses the deployed Render backend URL. For local backend testing, set `const API = ""`.

3. The input workbook format is tightly coupled to fixed sheet names and fixed column indexes.

4. Generated files are stored by session id, but there is no cleanup job in the current backend.

5. The backend imports `generate_for_state` and `write_output_excel` in `app.py`, but only `process_all` is directly needed by the API process endpoint.

6. `frontend/node_modules`, `frontend/dist`, and `backend/__pycache__` are generated artifacts and should normally not be edited manually.

## 15. Quick Developer Checklist

Use this checklist when working on the project:

1. Start backend:

```bash
cd backend
python app.py
```

2. Start frontend:

```bash
cd frontend
npm run dev
```

3. Confirm API mode in `frontend/src/App.jsx`:

```javascript
const API = "";
```

for local backend, or:

```javascript
const API = "https://ai-for-payout-grid.onrender.com";
```

for deployed backend.

4. Upload a valid Digit 2W `.xlsx` workbook.
5. Check state detection.
6. Select effect dates.
7. Generate output.
8. Download and inspect the Excel files.

## 16. Summary In One Line

This project is a Flask + React rule-file generator that converts Digit two-wheeler commission grid workbooks into state-wise, formatted, 50-column payout rule-engine Excel files with optional ZIP download.
