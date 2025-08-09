#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
memoq_gt_automation.py
Automate split-screen copy/paste between memoQ and Google Translate on Windows.

Requirements:
    pip install pyautogui pyperclip keyboard pillow

Usage:
    # Calibrate positions (do this once or whenever your layout changes)
    python memoq_gt_automation.py --calibrate

    # Run for N segments (example: 50)
    python memoq_gt_automation.py --segments 50 --delay 0.25

Safety:
    - Move the mouse to a screen corner to trigger PyAutoGUI fail-safe.
    - Press ESC to abort.

Notes:
    - This script uses only generic shortcuts (Ctrl+C, Ctrl+V, Ctrl+A, Tab, Enter).
    - You may customize "NEXT_SEGMENT_HOTKEYS" if you want to advance memoQ segments via a key combo.
      By default, the script just hits "Tab" then "Enter" as a conservative approach. Adjust if needed.
"""
import argparse
import json
import os
import sys
import time

import pyautogui
import pyperclip
import keyboard

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "memoq_gt_positions.json")

# Default behavior: how to move to next memoQ segment after pasting
# You can list keystrokes here (they are sent in order).
# Examples:
#   ["ctrl", "enter"]          -> hold Ctrl, press Enter (commit & next in many CAT tools)
#   ["enter"]                   -> just Enter
#   ["tab"]                     -> move focus
#   ["ctrl", "down"]            -> go down one segment
NEXT_SEGMENT_HOTKEYS = ["ctrl", "enter"]  # change if your memoQ workflow differs

# Delays to keep UI stable (seconds)
DEFAULT_DELAY_BETWEEN_STEPS = 0.25
DEFAULT_DELAY_AFTER_PASTE = 0.40
DEFAULT_DELAY_AFTER_TRANSLATE = 0.60


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    print(f"[OK] Saved positions to {CONFIG_PATH}")


def load_config():
    if not os.path.exists(CONFIG_PATH):
        print(f"[!] Config not found: {CONFIG_PATH}. Run with --calibrate first.")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def wait_for_space(prompt):
    print(prompt)
    print("  → Hover the mouse over the spot, then press SPACE to capture. (ESC to abort)")
    while True:
        if keyboard.is_pressed("esc"):
            print("[ABORT] ESC pressed.")
            sys.exit(1)
        if keyboard.is_pressed("space"):
            pos = pyautogui.position()
            # Debounce
            time.sleep(0.3)
            print(f"  Captured at {pos.x}, {pos.y}")
            return {"x": pos.x, "y": pos.y}
        time.sleep(0.05)


def calibrate():
    print("=== Calibration ===")
    print("Make sure memoQ and Google Translate are visible side-by-side.")
    time.sleep(1)

    cfg = {}
    cfg["memoq_source"] = wait_for_space("1) memoQ SOURCE area (clicking here should select the source text or the segment).")
    cfg["memoq_target"] = wait_for_space("2) memoQ TARGET area (clicking here should focus the target cell/field).")
    cfg["gt_input"] = wait_for_space("3) Google Translate INPUT box (where you paste the source).")
    cfg["gt_output"] = wait_for_space("4) Google Translate OUTPUT box (where the translation appears).")

    print("\n[Optional] Set how to advance to the next memoQ segment.")
    print(f"Current default: {NEXT_SEGMENT_HOTKEYS}")
    print("If you want to keep the default, just press ENTER. Otherwise, type a comma-separated list of keys, e.g.: ctrl,enter  or  ctrl,down")
    user_line = input("Hotkeys: ").strip()
    if user_line:
        keys = [k.strip().lower() for k in user_line.split(",") if k.strip()]
        cfg["next_segment_hotkeys"] = keys
    else:
        cfg["next_segment_hotkeys"] = NEXT_SEGMENT_HOTKEYS

    save_config(cfg)
    print("[DONE] Calibration complete.")


def click_point(pt, label="", pause=0.1):
    if label:
        print(f"[Click] {label} at ({pt['x']}, {pt['y']})")
    pyautogui.moveTo(pt["x"], pt["y"], duration=0.1)
    pyautogui.click()
    time.sleep(pause)


def send_keys(keys):
    """
    Send a list of keys in order.
    Special handling:
      - If list is like ["ctrl", "enter"] we send ctrl+enter chord.
      - If the list is longer than 2, we send them one by one with tiny pauses.
    """
    if not keys:
        return
    if len(keys) == 1:
        pyautogui.press(keys[0])
        return
    if len(keys) == 2:
        pyautogui.hotkey(keys[0], keys[1])
        return
    for k in keys:
        pyautogui.press(k)
        time.sleep(0.05)


def abort_if_requested():
    if keyboard.is_pressed("esc"):
        print("[ABORT] ESC pressed.")
        sys.exit(1)


def process_one_segment(cfg, delays):
    # 1) Focus memoQ source and copy
    click_point(cfg["memoq_source"], "memoQ Source", pause=delays["between"])
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(delays["between"])
    src_text = pyperclip.paste().strip()

    if not src_text:
        print("[WARN] Empty source text on clipboard; skipping this segment.")
        return False

    # 2) Paste into Google Translate input
    click_point(cfg["gt_input"], "Google Translate Input", pause=delays["between"])
    pyautogui.hotkey("ctrl", "a")  # clear previous
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(delays["after_paste"])

    # 3) Wait a bit for translation to appear, copy output
    click_point(cfg["gt_output"], "Google Translate Output", pause=delays["between"])
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(delays["after_translate"])
    tgt_text = pyperclip.paste().strip()

    if not tgt_text:
        print("[WARN] Empty translation from Google; will still paste (maybe UI not ready).")

    # 4) Paste into memoQ target
    click_point(cfg["memoq_target"], "memoQ Target", pause=delays["between"])
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.05)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(delays["between"])

    # 5) Advance to next segment (optional)
    if cfg.get("next_segment_hotkeys"):
        send_keys(cfg["next_segment_hotkeys"])
        time.sleep(delays["between"])

    return True


def main():
    pyautogui.FAILSAFE = True  # Move mouse to a corner to abort

    ap = argparse.ArgumentParser(description="Automate memoQ <-> Google Translate copy/paste.")
    ap.add_argument("--calibrate", action="store_true", help="Capture click positions for memoQ and Google Translate.")
    ap.add_argument("--segments", type=int, default=1, help="How many segments to process (default: 1).")
    ap.add_argument("--delay", type=float, default=DEFAULT_DELAY_BETWEEN_STEPS, help="Base delay between steps (seconds).")
    ap.add_argument("--delay-after-paste", type=float, default=DEFAULT_DELAY_AFTER_PASTE, help="Delay after pasting into Google (seconds).")
    ap.add_argument("--delay-after-translate", type=float, default=DEFAULT_DELAY_AFTER_TRANSLATE, help="Delay after copying from Google output (seconds).")
    args = ap.parse_args()

    if args.calibrate:
        calibrate()
        return

    cfg = load_config()
    delays = {
        "between": max(0.05, args.delay),
        "after_paste": max(0.1, args.delay_after_paste),
        "after_translate": max(0.1, args.delay_after_translate),
    }

    print(f"Starting automation for {args.segments} segment(s).")
    print("Press ESC to abort at any time.\n")
    time.sleep(1)

    success = 0
    for i in range(args.segments):
        abort_if_requested()
        print(f"--- Segment {i+1} ---")
        ok = process_one_segment(cfg, delays)
        if ok:
            success += 1

    print(f"\nDone. Successful segments: {success}/{args.segments}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[ABORT] KeyboardInterrupt.")
