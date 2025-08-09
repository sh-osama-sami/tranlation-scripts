
# Merge Arabic translations from Excel into a memoQ bilingual file (MQXLIFF).
#
# Usage:
#   python merge_excel_into_mqxliff.py --mqxliff IN.mqxliff --excel "Test2 (1).xlsx" --out OUT.mqxliff --src-col 0 --tgt-col 1 --target-lang ar-EG
#
# Notes:
# - Export your memoQ document as MQXLIFF (recommended over RTF).
# - Excel must have two columns: Source (English) and Target (Arabic). By default the script uses the first two columns (0,1).
# - Matching is by exact source text. If a source appears multiple times in the MQXLIFF, all matching segments will be updated.
# - The script creates or updates <target> for each trans-unit and sets the state to "translated".
# - Generates a small report: updated_count, missing_in_excel, and a CSV of MQXLIFF sources that had no match.
#
# Limitations:
# - Inline tags/markups inside <source> are ignored for matching (we match on the plain text of <source>).
# - If your project has many identical source strings, consider switching to segment-id based matching.

import argparse
import csv
import sys
import os
from xml.etree import ElementTree as ET

try:
    import pandas as pd
except Exception as e:
    print("[ERR] pandas is required. Please install it: pip install pandas")
    sys.exit(1)

def text_content(elem):
    """Return concatenated text from an XML element (including nested tags)."""
    parts = []
    if elem.text:
        parts.append(elem.text)
    for child in elem:
        parts.append(text_content(child))
        if child.tail:
            parts.append(child.tail)
    return "".join(parts)

def load_excel_pairs(excel_path, src_col=0, tgt_col=1):
    df = pd.read_excel(excel_path, sheet_name=0)
    def clean(x):
        if pd.isna(x):
            return None
        s = str(x).strip()
        return s if s else None
    src = df.iloc[:, src_col].apply(clean)
    tgt = df.iloc[:, tgt_col].apply(clean)
    pairs = {}
    for s, t in zip(src, tgt):
        if s and t:
            pairs[s] = t  # last one wins if duplicates in Excel
    return pairs

def ensure_target_child(tu, target_lang):
    # Find or create <target>
    target = tu.find("target")
    if target is None:
        target = ET.SubElement(tu, "target")
    # Set xml:lang if present in namespace; tolerate absence
    target.set("{http://www.w3.org/XML/1998/namespace}lang", target_lang)
    return target

def set_memoq_state_attrs(tu, target):
    # Try to set memoQ-like state attributes if present; otherwise, set generic ones
    # Common patterns in XLIFF 1.2:
    #   trans-unit[@approved], target[@state]
    # We'll set target state to 'translated' and approved to 'yes' if attributes exist.
    try:
        target.set("state", "translated")
    except Exception:
        pass
    try:
        tu.set("approved", "yes")
    except Exception:
        pass

def process_mqxliff(mqxliff_path, pairs, out_path, target_lang):
    tree = ET.parse(mqxliff_path)
    root = tree.getroot()

    # Iterate all trans-units regardless of namespaces
    def iter_trans_units():
        for tu in root.iter():
            if tu.tag.endswith("trans-unit"):
                yield tu

    updated = 0
    total_tus = 0
    missing_sources = []

    for tu in iter_trans_units():
        total_tus += 1
        src = None
        # find child named 'source' regardless of namespace
        for child in tu:
            if child.tag.endswith("source"):
                src = child
                break
        if src is None:
            continue
        plain_src = text_content(src).strip()
        if not plain_src:
            continue

        if plain_src in pairs:
            target = ensure_target_child(tu, target_lang)
            target.text = pairs[plain_src]
            set_memoq_state_attrs(tu, target)
            updated += 1
        else:
            missing_sources.append(plain_src)

    # Write out
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)

    # Report
    missing_csv = os.path.splitext(out_path)[0] + "_missing_sources.csv"
    with open(missing_csv, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["SourceWithoutMatch"]) 
        for s in sorted(set(missing_sources)):
            w.writerow([s])

    return {
        "total_trans_units": total_tus,
        "updated_segments": updated,
        "missing_sources_count": len(set(missing_sources)),
        "missing_sources_csv": missing_csv
    }

def main():
    ap = argparse.ArgumentParser(description="Merge Arabic translations from Excel into MQXLIFF (memoQ bilingual).")
    ap.add_argument("--mqxliff", required=True, help="Path to input MQXLIFF file exported from memoQ")
    ap.add_argument("--excel", required=True, help="Path to Excel containing Source/Target columns")
    ap.add_argument("--out", required=True, help="Path to write the merged MQXLIFF")
    ap.add_argument("--src-col", type=int, default=0, help="Zero-based index of Source column in Excel (default 0)")
    ap.add_argument("--tgt-col", type=int, default=1, help="Zero-based index of Target column in Excel (default 1)")
    ap.add_argument("--target-lang", default="ar-EG", help="xml:lang to set on <target> (default ar-EG)")
    args = ap.parse_args()

    pairs = load_excel_pairs(args.excel, args.src_col, args.tgt_col)
    if not pairs:
        print("[ERR] No valid pairs found in Excel.")
        sys.exit(2)

    # Safety: back up the original MQXLIFF if user overwrites
    if os.path.abspath(args.mqxliff) == os.path.abspath(args.out):
        backup = args.mqxliff + ".bak"
        with open(args.mqxliff, "rb") as r, open(backup, "wb") as w:
            w.write(r.read())
        print(f"[WARN] Output equals input. Backed up original to: {backup}")

    stats = process_mqxliff(args.mqxliff, pairs, args.out, args.target_lang)
    print("[DONE] Merge completed.")
    print(f"  - Total trans-units: {stats['total_trans_units']}")
    print(f"  - Updated segments:  {stats['updated_segments']}")
    print(f"  - No-match sources:  {stats['missing_sources_count']}")
    print(f"  - No-match report:   {stats['missing_sources_csv']}")

if __name__ == "__main__":
    main()
