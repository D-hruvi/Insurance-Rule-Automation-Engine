#!/usr/bin/env python3
"""
Usage:
  python3 run.py <input_xlsx> <output_dir> <effect_start> <effect_end> [state1 state2 ...]

Examples:
  python3 run.py input.xlsx ./output 2026-02-01 2026-02-28
  python3 run.py input.xlsx ./output 2026-02-01 2026-02-28 DELHI MUMBAI
"""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))
from processor import process_all

if len(sys.argv) < 5:
    print(__doc__)
    sys.exit(1)

input_path  = sys.argv[1]
output_dir  = sys.argv[2]
eff_start   = sys.argv[3]
eff_end     = sys.argv[4]
states      = sys.argv[5:] or None

def cb(msg, cur, tot):
    bar = f"[{cur:3}/{tot:3}]" if tot else ""
    print(f"  {bar} {msg}")

files = process_all(input_path, output_dir, eff_start, eff_end,
                    states=states, progress_callback=cb)

print(f"\n✓ Generated {len(files)} file(s) in '{output_dir}'")
