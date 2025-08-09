
# memoq_paste_automation.py
#
# Automate pasting Arabic translations into memoQ from Excel/CSV/TSV using Python.
#
# REQUIREMENTS (install once):
#   pip install pandas pyautogui pyperclip
#
# USAGE EXAMPLES
#   # Using the TSV I created earlier
#   python memoq_paste_automation.py --tsv "memoq_pairs.tsv"
#
#   # Using your Excel directly (first two columns = Source, Target)
#   python memoq_paste_automation.py --excel "Test2 (1).xlsx" --sheet 0 --src-col 0 --tgt-col 1
#
# WHAT IT DOES
#   For each Source→Target row:
#     1) Ctrl+F, paste Source, Enter  (search)
#     2) Tab to target cell (configurable)
#     3) Paste Target, Ctrl+Enter to confirm (configurable)
#
# TIPS
#   - Before you start, focus memoQ Editor and place the caret in the grid.
#   - Don’t touch the keyboard/mouse while it runs.
#   - Adjust delays below if your machine is slower/faster.
#
# SAFETY
#   - Dry-run mode (--dry-run) prints what would be sent without typing.
#   - You can start after a countdown (default 5s) to let you focus memoQ.
#
# LIMITATIONS
#   - This is keystroke automation; it cannot detect "not found" in UI.
#   - If you need higher reliability, share a small MQXLIFF and we’ll do ID-based merging (Option 2).
#
import argparse
import sys
import time
from pathlib import Path

import pandas as pd
import pyautogui as pag
import pyperclip

# -------- Defaults you can tweak (also exposed via CLI) --------
DEFAULT_TO_TARGET_NAV = "tab"    # how to focus target after search; often "tab"
DEFAULT_CONFIRM_SHORTCUT = ("ctrl", "enter")  # confirm/commit segment
DEFAULT_FIND_DELAY = 0.7         # seconds after pressing Enter in Find
DEFAULT_BEFORE_PASTE_DELAY = 0.2 # seconds before pasting target
DEFAULT_AFTER_CONFIRM_DELAY = 0.2# seconds after confirm
DEFAULT_COUNTDOWN = 5            # seconds before start
# ---------------------------------------------------------------

def human_hotkey(keys):
    """Send a hotkey sequence like ('ctrl','enter') using pyautogui."""
    if isinstance(keys, (list, tuple)):
        pag.hotkey(*keys)
    elif isinstance(keys, str):
        # single key name
        pag.press(keys)
    else:
        raise ValueError("Unsupported hotkey format")

def send_text(text, dry_run=False):
    """Copy text to clipboard and paste with Ctrl+V for better RTL/Unicode support."""
    if dry_run:
        print(f"[DRY] Paste: {text[:60]}{'...' if len(text)>60 else ''}")
        return
    pyperclip.copy(str(text))
    time.sleep(0.05)
    pag.hotkey("ctrl", "v")

def read_pairs(args):
    """Return list of (source, target) pairs from Excel/CSV/TSV."""
    if args.excel:
        df = pd.read_excel(args.excel, sheet_name=args.sheet)
        src = df.iloc[:, args.src_col]
        tgt = df.iloc[:, args.tgt_col]
    elif args.tsv:
        df = pd.read_csv(args.tsv, sep="\t", encoding="utf-8-sig")
        # Try common header names; else take first two columns
        if {"Source", "Target"}.issubset(df.columns):
            src = df["Source"]
            tgt = df["Target"]
        else:
            src = df.iloc[:, 0]
            tgt = df.iloc[:, 1]
    elif args.csv:
        df = pd.read_csv(args.csv, encoding="utf-8-sig")
        if {"Source", "Target"}.issubset(df.columns):
            src = df["Source"]
            tgt = df["Target"]
        else:
            src = df.iloc[:, 0]
            tgt = df.iloc[:, 1]
    else:
        raise SystemExit("Provide one of --excel/--tsv/--csv")
    def clean(x):
        if pd.isna(x):
            return None
        s = str(x).strip()
        return s if s else None
    pairs = [(clean(s), clean(t)) for s, t in zip(src, tgt)]
    # remove empties
    pairs = [(s, t) for (s, t) in pairs if s and t]
    return pairs

def main():
    ap = argparse.ArgumentParser(description="Automate pasting Arabic translations into memoQ from a file.")
    g_in = ap.add_mutually_exclusive_group(required=True)
    g_in.add_argument("--excel", help="Path to Excel file")
    g_in.add_argument("--tsv", help="Path to TSV file with columns Source\\tTarget (UTF-8)")
    g_in.add_argument("--csv", help="Path to CSV file with columns Source,Target (UTF-8)")
    ap.add_argument("--sheet", type=int, default=0, help="Excel sheet index (default 0)")
    ap.add_argument("--src-col", type=int, default=0, help="Excel source column index (default 0)")
    ap.add_argument("--tgt-col", type=int, default=1, help="Excel target column index (default 1)")
    ap.add_argument("--countdown", type=int, default=DEFAULT_COUNTDOWN, help="Seconds before start (default 5)")
    ap.add_argument("--to-target", default=DEFAULT_TO_TARGET_NAV, help="Key to focus target field (default 'tab')")
    ap.add_argument("--confirm", default="ctrl+enter", help="Confirm hotkey (default ctrl+enter)")
    ap.add_argument("--delay-find", type=float, default=DEFAULT_FIND_DELAY, help="Delay after Find (s) (default 0.7)")
    ap.add_argument("--delay-before-paste", type=float, default=DEFAULT_BEFORE_PASTE_DELAY, help="Delay before paste (s)")
    ap.add_argument("--delay-after-confirm", type=float, default=DEFAULT_AFTER_CONFIRM_DELAY, help="Delay after confirm (s)")
    ap.add_argument("--start-index", type=int, default=0, help="Start from this row index (0-based)")
    ap.add_argument("--limit", type=int, default=None, help="Process only this many rows")
    ap.add_argument("--dry-run", action="store_true", help="Print actions without sending keys")
    args = ap.parse_args()

    # Parse confirm hotkey like "ctrl+enter" or "alt+shift+s"
    confirm_keys = tuple(k.strip() for k in args.confirm.split("+")) if args.confirm else DEFAULT_CONFIRM_SHORTCUT

    pairs = read_pairs(args)
    if args.limit is not None:
        end = min(len(pairs), args.start_index + args.limit)
        pairs = pairs[args.start_index:end]
    else:
        pairs = pairs[args.start_index:]

    if not pairs:
        print("No pairs to process.")
        return

    print(f"Loaded {len(pairs)} pairs.")
    print(f"Starting in {args.countdown} seconds... Focus memoQ Editor now.")
    for i in range(args.countdown, 0, -1):
        print(i, end=" ", flush=True)
        time.sleep(1)
    print("\nGo. (Press Ctrl+C in this console to stop)")

    # Fail safe: moving mouse to a corner can abort pyautogui if fail-safe is enabled
    pag.FAILSAFE = True

    processed = 0
    try:
        for idx, (src, tgt) in enumerate(pairs, start=1):
            # 1) Find source
            if args.dry_run:
                print(f"[{idx}] Find -> {src[:80]}{'...' if len(src)>80 else ''}")
            else:
                pag.hotkey("ctrl", "f")
                time.sleep(0.1)
                pag.hotkey("ctrl", "a")
                pyperclip.copy(str(src))
                time.sleep(0.05)
                pag.hotkey("ctrl", "v")
                pag.press("enter")

            time.sleep(args.delay_find)

            # 2) Jump to target
            if args.dry_run:
                print(f"[{idx}] Navigate to target: {args.to_target}")
            else:
                if args.to_target.lower() == "tab":
                    pag.press("tab")
                else:
                    # allow sending a literal key name (e.g., "f6" or "enter")
                    pag.press(args.to_target)

            time.sleep(args.delay_before_paste)

            # 3) Paste target
            if args.dry_run:
                print(f"[{idx}] Paste target: {tgt[:80]}{'...' if len(tgt)>80 else ''}")
            else:
                pag.hotkey("ctrl", "a")
                pyperclip.copy(str(tgt))
                time.sleep(0.05)
                pag.hotkey("ctrl", "v")

            # 4) Confirm segment
            if args.dry_run:
                print(f"[{idx}] Confirm: {'+'.join(confirm_keys)}")
            else:
                human_hotkey(confirm_keys)

            time.sleep(args.delay_after_confirm)
            processed += 1
    except KeyboardInterrupt:
        print("\n[STOP] Interrupted by user (Ctrl+C).")
    except Exception as e:
        print(f"\n[ERR] {e}")
    finally:
        print(f"\nDone. Processed {processed} rows.")

if __name__ == "__main__":
    main()
