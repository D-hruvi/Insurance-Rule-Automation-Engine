# Digit 2W Converter — AI Edition

This replaces the hardcoded `processor.py` field-mapping logic with an LLM
(Groq `llama-3.1-8b-instant`, free tier) that reads raw commission grid
rows directly and outputs final rule-engine rows as JSON.

## How it works

1. **`ai_processor.py`** loads the raw source sheets (`TW 1+5`,
   `TW 1+1 & SATP`, `TW SAOD with Flexi Options`, `2W RTO's`) with no
   field-mapping logic — just raw row extraction.
2. For each target state, relevant raw rows are filtered by cluster and
   sent to Groq in batches. Batch size is set **per sheet**, not globally,
   because each sheet produces a different number of output rows per input
   row (TW 1+5 is 1:1, TW 1+1 & SATP is 1:2, SAOD is 1:4) — see
   `BATCH_SIZE_BY_SHEET` and `SHEET_OUTPUT_MULTIPLIER`. This keeps each
   call's expected JSON size predictable, which keeps `max_tokens` sized
   correctly without either truncating output or wasting token budget.
3. Every batch call includes the full 27 business rules as the **system
   prompt**, so the LLM applies them consistently call after call.
4. Calls are paced by a client-side rate limiter (`GROQ_RPM_LIMIT`, default
   25/min) to stay under Groq's free-tier cap before a 429 ever happens,
   and any 429 that does occur backs the whole pipeline off using the
   `Retry-After` Groq returns, rather than hammering the API again
   immediately.
5. The LLM returns a JSON array of fully-formed output rows — Rule Name,
   Cover/Business Type, Fuel Type, Make, PayIn/POSP amounts, everything.
6. **`validate_output.py`** then re-checks every generated row in Python
   against the hard constraints from the rules (Rule Code blank, Owner
   Type == ALL, POSP == PayIn × 0.8, valid Cover/Business Type
   combinations, etc.) and reports violations — it does not silently
   "fix" anything, so you can see exactly where the AI deviated.

## Setup

```bash
pip install -r requirements.txt
export GROQ_API_KEY="your-free-key-from-console.groq.com"
```

Optional tuning env vars (defaults shown):

```bash
export GROQ_RPM_LIMIT=25            # calls/min; stay under your account's actual Groq cap
export BATCH_SIZE_TW1P5=6           # raw rows per call for "TW 1+5"
export BATCH_SIZE_TW1P1=4           # raw rows per call for "TW 1+1 & SATP"
export BATCH_SIZE_SAOD=2            # raw rows per call for "TW SAOD with Flexi Options"
export GROQ_MAX_TOKENS_CEILING=3500 # hard ceiling on a single call's max_tokens
```

Check your account's actual rate limits at
https://console.groq.com/dashboard/limits before tuning `GROQ_RPM_LIMIT` —
published numbers vary by tier and change over time, so that dashboard is
the source of truth, not this README.

## Run directly (CLI)

```bash
python ai_processor.py source_grid.xlsx ./output 2026-02-01 2026-02-28
# optional: restrict to specific states
python ai_processor.py source_grid.xlsx ./output 2026-02-01 2026-02-28 "DELHI,MAHARASHTRA"
```

Running the CLI logs an upfront estimate of total Groq calls and rough ETA
before it starts, so you can sanity-check a full run before committing to
it.

Then validate:

```bash
python validate_output.py ./output
```

## Run as an API

```bash
python app_ai.py
```

`POST /api/process` no longer does the AI work inline and returns
immediately (HTTP 202) with `{session_id, status_url}` — a full multi-state
run can take many minutes (hundreds of sequential Groq calls), which is too
long to hold an HTTP request open for reliably. The actual processing runs
in a background thread.

Poll `GET /api/status/<session_id>` (every 3-5s is reasonable) until
`status` is `"done"` — at that point `result` has the same shape `/api/process`
used to return directly: generated files, a zip, and a `validation` block
with the violation count and details per file, so a frontend can surface AI
mistakes immediately instead of you having to spot-check output by hand.
`status` can also be `"queued"`, `"running"`, or `"error"` (with `error`
and `trace` fields in the latter case).

### Deploying (e.g. Render)

Job state lives in an in-memory dict, so the process must run as a single
worker **process** (extra threads are fine; extra Gunicorn workers are not
— they don't share memory, so a different worker could 404 on a
`session_id` it never created). Start command:

```
gunicorn app_ai:app --bind 0.0.0.0:$PORT --workers 1 --worker-class gthread --threads 4 --timeout 120
```

## Why this design

- **Per-sheet batching**: keeps each LLM call's context and expected output
  size predictable — more reliable JSON than one-size-fits-all batching,
  and far fewer total calls than batching every sheet at size 1, which is
  what was tripping Groq's rate limit.
- **Background job + polling**: a request handler that runs for 10-60+
  minutes is fragile regardless of how well-paced the Groq calls are —
  decoupling the HTTP response from the actual work fixes both the
  rate-limit pacing problem and the unrelated worker-timeout issue that
  looked like an OOM crash in earlier testing.
- **AI does everything**: no Python fallback rules for field values
  (Cover Type, Fuel Type, Make exclusions, etc.) — the LLM is the single
  source of truth for those decisions, which is why the validator exists
  as a safety net rather than a corrective layer.
- **Free tier**: Groq's `llama-3.1-8b-instant` was chosen because it's
  free, fast, and has held up well on this kind of structured-extraction
  task in earlier testing on this project.

## Known tradeoffs vs. the old rule-engine version

- Slower: many small API calls instead of pure Python loops, and now an
  async job rather than a single request/response.
- Non-deterministic: the same input grid can produce slightly different
  rule names or row counts between runs (temperature is set to 0 to
  minimize this, but it isn't a hard guarantee with hosted LLMs).
- Validation catches rule violations but doesn't catch every possible
  semantic error (e.g. wrong RTO codes for a cluster) — it's a sanity
  net, not a correctness guarantee.
- Rate-limit-bound: total runtime for a full all-states run scales with
  Groq's RPM/TPM caps on your account, not just with how fast the code
  itself runs.
