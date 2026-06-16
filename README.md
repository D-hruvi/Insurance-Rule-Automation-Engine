# Digit 2W Converter — AI Edition

This replaces the hardcoded `processor.py` field-mapping logic with an LLM
(Groq `llama-3.3-70b-versatile`, free tier) that reads raw commission grid
rows directly and outputs final rule-engine rows as JSON.

## How it works

1. **`ai_processor.py`** loads the raw source sheets (`TW 1+5`,
   `TW 1+1 & SATP`, `TW SAOD with Flexi Options`, `2W RTO's`) with no
   field-mapping logic — just raw row extraction.
2. For each target state, relevant raw rows are filtered by cluster and
   sent to Groq in small batches (`BATCH_SIZE = 8` rows per call).
3. Every batch call includes the full 27 business rules as the **system
   prompt**, so the LLM applies them consistently call after call.
4. The LLM returns a JSON array of fully-formed output rows — Rule Name,
   Cover/Business Type, Fuel Type, Make, PayIn/POSP amounts, everything.
5. **`validate_output.py`** then re-checks every generated row in Python
   against the hard constraints from the rules (Rule Code blank, Owner
   Type == ALL, POSP == PayIn × 0.8, valid Cover/Business Type
   combinations, etc.) and reports violations — it does not silently
   "fix" anything, so you can see exactly where the AI deviated.

## Setup

```bash
pip install -r requirements.txt
export GROQ_API_KEY="your-free-key-from-console.groq.com"
```

## Run directly (CLI)

```bash
python ai_processor.py source_grid.xlsx ./output 2026-02-01 2026-02-28
# optional: restrict to specific states
python ai_processor.py source_grid.xlsx ./output 2026-02-01 2026-02-28 "DELHI,MAHARASHTRA"
```

Then validate:

```bash
python validate_output.py ./output
```

## Run as an API

```bash
python app_ai.py
```

`POST /api/process` returns the generated files **and** a `validation`
block with the violation count and details per file, so a frontend can
surface AI mistakes immediately instead of you having to spot-check
output by hand.

## Why this design

- **Batches of 8 rows**: keeps each LLM call's context small, which is
  more reliable for consistent JSON than dumping an entire sheet into
  one giant prompt — at the cost of more API calls.
- **AI does everything**: no Python fallback rules for field values
  (Cover Type, Fuel Type, Make exclusions, etc.) — the LLM is the single
  source of truth for those decisions, which is why the validator exists
  as a safety net rather than a corrective layer.
- **Free tier**: Groq's `llama-3.3-70b-versatile` was chosen because it's
  free and has held up well on this kind of structured-extraction task
  in earlier testing on this project.

## Known tradeoffs vs. the old rule-engine version

- Slower: many small API calls instead of pure Python loops.
- Non-deterministic: the same input grid can produce slightly different
  rule names or row counts between runs (temperature is set to 0 to
  minimize this, but it isn't a hard guarantee with hosted LLMs).
- Validation catches rule violations but doesn't catch every possible
  semantic error (e.g. wrong RTO codes for a cluster) — it's a sanity
  net, not a correctness guarantee.
