"""
Flipkart Template Filler
=========================
Copies the generated listing rows into a real Flipkart template you downloaded from the
Seller portal, so you never paste by hand.

    outputs_flipkart/<model>.xlsx  ──copy rows──>  Caticker_ascbhsvcysd_268.xls

Why this exists: generate_flipkart_listings.py writes a fresh .xlsx, which cannot carry
the downloaded file's dropdowns, colour coding or the CTRL+SHIFT+S Fast Validate macro.
This script drives Excel itself, so the downloaded workbook keeps everything and only the
data rows change.

Columns are matched by attribute NAME, not position, so it still works when a Flipkart
category file orders or names its columns differently.

Run (pairs taken from the 'template_file' column of flipkart_master.xlsx):
    python fill_flipkart_template.py

Or one file at a time:
    python fill_flipkart_template.py --source outputs_flipkart/iPhone_17_pro.xlsx \
                                     --target "downloads/Caticker_ascbhsvcysd_268.xls"

Requires: pywin32 (pip install pywin32) and Microsoft Excel installed.
"""

import argparse
import os
import re
import sys
import warnings

import openpyxl
import pandas as pd

warnings.simplefilter("ignore")

# ---------- CONFIG ----------
FLIPKART_MASTER = "flipkart_master.xlsx"
SOURCE_DIR = "outputs_flipkart"
TEMPLATE_FILE_COLUMN = "template_file"   # optional column in flipkart_master.xlsx
PREFERRED_SHEET = "mobile_skin"
KEY_HEADER = "Seller SKU ID"             # used to locate the header row in either file
HEADER_BLOCK_ROWS = 4                  # names / types / example / help, then data
MAX_SCAN_COLS = 400                      # cap the header scan; UsedRange can run wide


def normalize(name):
    """Header names differ by case, padding and non-breaking spaces between templates."""
    s = str(name).replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip().casefold()


def read_source(path):
    """(headers, rows) from a generated Flipkart workbook."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[PREFERRED_SHEET] if PREFERRED_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]

    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        values = [normalize(c.value) for c in ws[r] if c.value is not None]
        if normalize(KEY_HEADER) in values:
            header_row = r
            break
    if header_row is None:
        wb.close()
        raise ValueError(f"{path}: no header row containing {KEY_HEADER!r}")

    headers = [c.value for c in ws[header_row]]
    key_col = [normalize(h) for h in headers].index(normalize(KEY_HEADER))

    # The 4-row header block is names / types / Flipkart's example / help text. The help
    # row carries prose in the SKU column and the example row a single space, so they
    # cannot be filtered by "SKU looks empty" - skip the block by position.
    rows = []
    for r in range(header_row + HEADER_BLOCK_ROWS, ws.max_row + 1):
        vals = [ws.cell(row=r, column=i + 1).value for i in range(len(headers))]
        if vals[key_col] is None or str(vals[key_col]).strip() == "":
            continue          # trailing blank rows
        rows.append(vals)

    wb.close()
    return headers, rows


def find_target_sheet(wb, source_headers):
    """The sheet whose header row best matches the source's columns.

    Width comes from the last non-empty header cell, never from UsedRange: Excel counts
    merely-formatted cells as used, and Flipkart pre-formats blank columns, which
    reported a 76-column sheet as 256 wide.
    """
    want = {normalize(h) for h in source_headers if h}
    best = None
    for sh in wb.Sheets:
        used = sh.UsedRange
        n_rows = min(used.Row + used.Rows.Count - 1, 20)
        scan_cols = min(used.Column + used.Columns.Count - 1, MAX_SCAN_COLS)
        if scan_cols < 2:
            continue
        for r in range(1, n_rows + 1):
            row = sh.Range(sh.Cells(r, 1), sh.Cells(r, scan_cols)).Value
            if not row:
                continue
            values = row[0]
            names = {normalize(v) for v in values if v is not None}
            if normalize(KEY_HEADER) not in names:
                continue
            width = max(
                (i + 1 for i, v in enumerate(values) if v is not None and str(v).strip()),
                default=0,
            )
            score = len(want & names)
            if best is None or score > best[0]:
                best = (score, sh, r, width)
    if best is None:
        raise ValueError(f"no sheet with a {KEY_HEADER!r} header row")
    _, sh, header_row, n_cols = best
    if PREFERRED_SHEET and sh.Name != PREFERRED_SHEET:
        print(f"    note: matched sheet {sh.Name!r}, not {PREFERRED_SHEET!r}")
    return sh, header_row, n_cols


def count_existing_rows(sh, data_start, key_col):
    """Data rows that actually hold a SKU. UsedRange overstates this badly: a template
    with pre-formatted empty rows reported 496 'existing data rows' when it had none."""
    used = sh.UsedRange
    last = used.Row + used.Rows.Count - 1
    if last < data_start:
        return 0
    span = sh.Range(sh.Cells(data_start, key_col), sh.Cells(last, key_col)).Value
    n = 0
    for i, row in enumerate(span or []):
        v = row[0] if isinstance(row, tuple) else row
        if v is not None and str(v).strip():
            n = i + 1
    return n


def locked_columns(sh, data_start, columns):
    """Which of these 1-based columns can't be written because the sheet is protected."""
    if not sh.ProtectContents:
        return set()
    locked = set()
    for col in columns:
        try:
            if sh.Cells(data_start, col).Locked:
                locked.add(col)
        except Exception:
            locked.add(col)
    return locked


def fill_one(xl, source_path, target_path, unprotect=False):
    headers, rows = read_source(source_path)
    if not rows:
        print(f"  [SKIPPED] {source_path}: no data rows")
        return False

    wb = xl.Workbooks.Open(os.path.abspath(target_path))
    try:
        sh, header_row, n_cols = find_target_sheet(wb, headers)
        target_row = sh.Range(sh.Cells(header_row, 1), sh.Cells(header_row, n_cols)).Value[0]
        target_by_name = {}
        for i, v in enumerate(target_row):
            if v is not None and normalize(v) not in target_by_name:
                target_by_name[normalize(v)] = i + 1   # 1-based column

        print(f"    sheet={sh.Name!r} header row={header_row} target cols={n_cols}")

        data_start = header_row + HEADER_BLOCK_ROWS

        matched, unmatched_source = [], []
        for i, h in enumerate(headers):
            if not h:
                continue
            col = target_by_name.get(normalize(h))
            if col is None:
                unmatched_source.append(str(h))
            else:
                matched.append((i, col, str(h)))

        if sh.ProtectContents:
            print(f"    sheet is PROTECTED"
                  f"{' - unprotecting as requested' if unprotect else ''}")
            if unprotect:
                sh.Unprotect()

        locked = locked_columns(sh, data_start, [c for _, c, _ in matched])
        writable = [(i, c, n) for i, c, n in matched if c not in locked]
        skipped = [(i, c, n) for i, c, n in matched if c in locked]

        if skipped:
            print(f"    SKIPPED {len(skipped)} locked column(s) - protection blocks these:")
            for _, col, name in skipped:
                print(f"        col {col}: {name}")
        if not writable:
            print("    nothing writable: every matched column is locked. Unprotect the "
                  "sheet in Excel (Review > Unprotect Sheet), or re-run with --unprotect.")
            return False

        key_col = target_by_name[normalize(KEY_HEADER)]
        existing = count_existing_rows(sh, data_start, key_col)
        if existing:
            print(f"    clearing {existing} existing data row(s) from row {data_start}")
            for _, col, _ in writable:
                sh.Range(sh.Cells(data_start, col),
                         sh.Cells(data_start + existing - 1, col)).ClearContents()

        # Protection blocks formatting changes even on unlocked cells, so only force the
        # text format when the sheet is open. A protected Flipkart template already
        # carries the right formats anyway.
        can_format = not sh.ProtectContents
        if not can_format:
            print("    leaving cell formats alone (protection blocks NumberFormat)")

        # One range assignment per column: 200x76 cell-by-cell over COM would crawl.
        for src_i, tgt_col, _ in writable:
            column = [[r[src_i] if src_i < len(r) else None] for r in rows]
            rng = sh.Range(
                sh.Cells(data_start, tgt_col),
                sh.Cells(data_start + len(rows) - 1, tgt_col),
            )
            if can_format:
                rng.NumberFormat = "@"   # keep SKUs/HSN/EAN as text, not numbers
            rng.Value = column

        if unprotect:
            sh.Protect()
        wb.Save()
        print(f"  [OK] {len(rows)} rows -> {target_path}")
        print(f"       wrote {len(writable)} of {len(matched)} matched columns "
              f"(source has {len([h for h in headers if h])})")
        if unmatched_source:
            print(f"       NOT in target ({len(unmatched_source)}): "
                  f"{', '.join(unmatched_source[:8])}")
        unmatched_target = [
            v for v in target_row
            if v is not None and normalize(v) not in {normalize(h) for h in headers if h}
        ]
        if unmatched_target:
            print(f"       target columns left untouched ({len(unmatched_target)}): "
                  f"{', '.join(str(v) for v in unmatched_target[:8])}")
        return True
    finally:
        wb.Close(SaveChanges=False)


def resolve_pairs(args):
    if args.source or args.target:
        if not (args.source and args.target):
            sys.exit("ERROR: pass both --source and --target, or neither.")
        return [(os.path.basename(args.source), args.source, args.target)]

    if not os.path.exists(FLIPKART_MASTER):
        sys.exit(f"ERROR: {FLIPKART_MASTER} not found. Use --source and --target instead.")

    df = pd.read_excel(FLIPKART_MASTER, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    if TEMPLATE_FILE_COLUMN not in df.columns:
        sys.exit(
            f"ERROR: {FLIPKART_MASTER} has no '{TEMPLATE_FILE_COLUMN}' column.\n"
            f"  Add one and put the path of the file you downloaded from Flipkart in it\n"
            f"  for each model_key, e.g. downloads/Caticker_ascbhsvcysd_268.xls\n"
            f"  Or run one at a time with --source and --target."
        )

    pairs = []
    for _, row in df.iterrows():
        model_key = str(row["model_key"]).strip()
        target = row.get(TEMPLATE_FILE_COLUMN)
        if target is None or pd.isna(target) or str(target).strip() == "":
            print(f"  [SKIPPED] {model_key}: no {TEMPLATE_FILE_COLUMN} set")
            continue
        source = os.path.join(SOURCE_DIR, f"{model_key}.xlsx")
        pairs.append((model_key, source, str(target).strip()))
    return pairs


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--source", help="generated file, e.g. outputs_flipkart/iPhone_17_pro.xlsx")
    ap.add_argument("--target", help="file downloaded from the Flipkart portal")
    ap.add_argument("--unprotect", action="store_true",
                    help="temporarily unprotect the sheet, write, then re-protect. "
                         "Only works if Flipkart set no password.")
    args = ap.parse_args()

    try:
        import win32com.client
    except ImportError:
        sys.exit(
            "ERROR: pywin32 is required to write into the downloaded Flipkart file.\n"
            "  pip install pywin32\n"
            "  (Microsoft Excel must also be installed.)"
        )

    pairs = resolve_pairs(args)
    if not pairs:
        sys.exit("Nothing to do.")

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ok = 0
    try:
        for model_key, source, target in pairs:
            print(f"\n{model_key}: {source} -> {target}")
            if not os.path.exists(source):
                print(f"  [SKIPPED] source not found: {source}")
                continue
            if not os.path.exists(target):
                print(f"  [SKIPPED] target not found: {target}")
                continue
            try:
                ok += bool(fill_one(xl, source, target, args.unprotect))
            except Exception as e:
                msg = str(e)
                if "File Block" in msg or "Trust Center" in msg:
                    print(
                        f"  [FAILED] Excel's Trust Center is blocking this file type.\n"
                        f"    Fix: Excel > File > Options > Trust Center > Trust Center\n"
                        f"    Settings > File Block Settings > find 'Excel 95-2003 Workbooks\n"
                        f"    and Templates' and set it to Open (uncheck it)."
                    )
                else:
                    print(f"  [FAILED] {type(e).__name__}: {msg[:200]}")
    finally:
        xl.Quit()

    print(f"\nFilled {ok}/{len(pairs)} file(s).")


if __name__ == "__main__":
    main()
