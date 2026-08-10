"""
One-time migration: merge old per-day trade Excel files into trades_all.xlsx
============================================================================
Reads all C:\\TradeReview\\data\\trades_YYYYMMDD.xlsx files, prepends a Date
column derived from the filename, concatenates them in chronological order,
and writes a single C:\\TradeReview\\trades_all.xlsx.

Also migrates per-day notes:
  C:\\TradeReview\\data\\notes_YYYYMMDD.txt → C:\\TradeReview\\notes\\YYYY-MM-DD.txt

Setup:
  pip install pandas openpyxl

Usage:
  python migrate_trades.py             # dry-run preview
  python migrate_trades.py --apply     # actually write trades_all.xlsx and copy notes
  python migrate_trades.py --apply --force   # overwrite existing trades_all.xlsx

Safety:
  - Original files in C:\\TradeReview\\data\\ are NEVER modified or deleted.
  - If trades_all.xlsx already exists, the script aborts unless --force is given.
  - Notes are COPIED (not moved); originals stay in place.
"""
import datetime as dt
import os, glob, shutil, sys, argparse

OLD_DATA_FOLDER = r"C:\TradeReview\data"
NEW_TRADES_FILE = r"C:\TradeReview\trades_all.xlsx"
NEW_NOTES_FOLDER = r"C:\TradeReview\notes"


def parse_date_from_filename(name):
    """trades_20260408.xlsx → '2026-04-08', or None if not parseable."""
    base = os.path.basename(name).replace("trades_", "").replace(".xlsx", "")
    try:
        d = dt.datetime.strptime(base, "%Y%m%d")
        return d.strftime("%Y-%m-%d")
    except ValueError:
        return None


def parse_date_from_notes_filename(name):
    """notes_20260408.txt → '2026-04-08', or None."""
    base = os.path.basename(name).replace("notes_", "").replace(".txt", "")
    try:
        d = dt.datetime.strptime(base, "%Y%m%d")
        return d.strftime("%Y-%m-%d")
    except ValueError:
        return None


def main():
    parser = argparse.ArgumentParser(description="Migrate per-day trade files to trades_all.xlsx")
    parser.add_argument("--apply", action="store_true",
                        help="Actually write files (default is dry-run preview)")
    parser.add_argument("--force", action="store_true",
                        help="Overwrite existing trades_all.xlsx if it exists")
    parser.add_argument("--source", type=str, default=OLD_DATA_FOLDER,
                        help=f"Source folder containing old trades_*.xlsx (default {OLD_DATA_FOLDER})")
    args = parser.parse_args()

    try:
        import pandas as pd
    except ImportError:
        print("✗ pandas not installed. Run: pip install pandas openpyxl")
        sys.exit(1)

    src_folder = args.source
    print(f"\n  Trade Migration")
    print(f"  Source:        {src_folder}")
    print(f"  Output trades: {NEW_TRADES_FILE}")
    print(f"  Output notes:  {NEW_NOTES_FOLDER}")
    print(f"  Mode:          {'APPLY' if args.apply else 'dry-run preview'}\n")

    if not os.path.isdir(src_folder):
        print(f"✗ Source folder does not exist: {src_folder}")
        sys.exit(1)

    # ── Discover trade files ───────────────────────────────────────────────
    trade_files = sorted(glob.glob(os.path.join(src_folder, "trades_*.xlsx")))
    if not trade_files:
        print("  No trades_*.xlsx files found in source folder.")
    else:
        print(f"  Found {len(trade_files)} trade file(s):")
        for f in trade_files:
            d = parse_date_from_filename(f)
            print(f"    {os.path.basename(f):30s} → Date={d or '?(skipped)'}")
        print()

    # ── Read & merge ───────────────────────────────────────────────────────
    all_rows = []
    canonical_columns = None
    skipped_files = []
    for f in trade_files:
        date_str = parse_date_from_filename(f)
        if not date_str:
            skipped_files.append((f, "filename not parseable"))
            continue
        try:
            df = pd.read_excel(f, dtype=str)
            df.columns = [c.strip() for c in df.columns]
        except Exception as e:
            skipped_files.append((f, f"read error: {e}"))
            continue
        if len(df) == 0:
            skipped_files.append((f, "empty"))
            continue
        # Capture column order from the first non-empty file (without Date)
        if canonical_columns is None:
            canonical_columns = list(df.columns)
        # Insert Date as the first column
        df.insert(0, "Date", date_str)
        all_rows.append(df)

    if skipped_files:
        print("  Skipped files:")
        for f, reason in skipped_files:
            print(f"    {os.path.basename(f)}  ({reason})")
        print()

    if not all_rows:
        print("  ✗ No usable trade data to merge.")
        merged = None
    else:
        merged = pd.concat(all_rows, ignore_index=True, sort=False)
        # Reorder columns: Date first, then the canonical order from the first file
        ordered = ["Date"] + [c for c in canonical_columns if c != "Date"]
        # Append any extra columns that appeared in later files but not the first
        extras = [c for c in merged.columns if c not in ordered]
        if extras:
            print(f"  ⚠ Extra columns found in later files (kept at the end): {extras}")
        merged = merged[ordered + extras]
        print(f"  Merged: {len(merged)} rows × {len(merged.columns)} columns")
        print(f"  Columns: {list(merged.columns)}")
        print(f"  Date range: {merged['Date'].min()} → {merged['Date'].max()}\n")

    # ── Discover notes files ───────────────────────────────────────────────
    note_files = sorted(glob.glob(os.path.join(src_folder, "notes_*.txt")))
    if note_files:
        print(f"  Found {len(note_files)} notes file(s):")
        for f in note_files[:10]:
            d = parse_date_from_notes_filename(f)
            new_name = f"{d}.txt" if d else "?"
            print(f"    {os.path.basename(f):30s} → notes\\{new_name}")
        if len(note_files) > 10:
            print(f"    ... and {len(note_files) - 10} more")
        print()
    else:
        print("  No notes_*.txt files found.\n")

    # ── Apply or stop ──────────────────────────────────────────────────────
    if not args.apply:
        print("  ── Dry-run complete. Re-run with --apply to actually write files. ──\n")
        return

    # Write trades_all.xlsx
    if merged is not None:
        if os.path.exists(NEW_TRADES_FILE) and not args.force:
            print(f"  ✗ {NEW_TRADES_FILE} already exists.")
            print(f"    Re-run with --force to overwrite, or move it aside first.")
            sys.exit(1)
        os.makedirs(os.path.dirname(NEW_TRADES_FILE), exist_ok=True)
        try:
            merged.to_excel(NEW_TRADES_FILE, index=False)
            print(f"  ✓ Wrote {NEW_TRADES_FILE} ({len(merged)} rows)")
        except Exception as e:
            print(f"  ✗ Failed to write {NEW_TRADES_FILE}: {e}")
            sys.exit(1)

    # Copy notes files
    if note_files:
        os.makedirs(NEW_NOTES_FOLDER, exist_ok=True)
        copied, skipped = 0, 0
        for f in note_files:
            d = parse_date_from_notes_filename(f)
            if not d:
                skipped += 1
                continue
            dst = os.path.join(NEW_NOTES_FOLDER, f"{d}.txt")
            if os.path.exists(dst) and not args.force:
                print(f"  - notes\\{d}.txt  already exists (skipped)")
                skipped += 1
                continue
            try:
                shutil.copy2(f, dst)
                copied += 1
            except Exception as e:
                print(f"  ✗ Failed to copy {f}: {e}")
                skipped += 1
        print(f"  ✓ Notes copied: {copied}  skipped: {skipped}")

    print(f"\n  Done. Original files in {src_folder} were not modified.\n")
    print(f"  Next steps:")
    print(f"    1. Open {NEW_TRADES_FILE} in Excel and verify the Date column")
    print(f"    2. Run trade_review_app.py (Bloomberg) or trade_review_app_free.py")
    print(f"    3. After confirming everything works, you can archive the old")
    print(f"       {src_folder} folder as a backup.\n")


if __name__ == "__main__":
    main()
