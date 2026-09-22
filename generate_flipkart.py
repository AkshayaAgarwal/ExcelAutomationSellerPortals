
"""
Flipkart Bulk Listing Generator
================================
Second stage of the pipeline. Takes the Amazon workbook that was already built and
checked, and turns it into a Flipkart catalog upload:

    models_master.xlsx  ─┐
                         ├─→ outputs/<model>.xlsm ─→ outputs_flipkart/<model>.xlsx
    images_master.xlsx  ─┘   (generate_amazon_listings2)   (this script)

The Amazon file is the source of truth for product data, so anything fixed there flows
through. Flipkart-only fields (Theme, Material, Tax Code, procurement, shipping fees)
have no Amazon equivalent and come from flipkart/flipkart_master.xlsx.

Run:
    python generate_flipkart_listings.py
"""

import os
import sys
import warnings
from collections import Counter

import openpyxl
import pandas as pd
import xlrd

import generate_amazon_listings2 as amazon

warnings.simplefilter("ignore")

# ---------- CONFIG ----------
FLIPKART_TEMPLATE = "flipkart_excel.xls"
FLIPKART_MASTER = "flipkart_master.xlsx"
AMAZON_OUTPUT_DIR = "outputs"
OUTPUT_DIR = "outputs_flipkart"
SHEET_NAME = "mobile_skin"
HEADER_ROW = 0          # attribute names
EXAMPLE_ROW = 2         # Flipkart's own sample values
DATA_START_ROW = 4      # "5th Row onwards: You are required to start entering product data"

# Header-row fill colours encode who fills a column and whether it is required.
# Documented on the template's own "Summary Sheet".
RGB_FLIPKART_FILLS = (192, 192, 192)   # grey  - Flipkart fills these, leave blank
RGB_MANDATORY = (141, 180, 226)        # blue  - must be filled
RGB_CONDITIONAL = (204, 153, 255)      # purple- conditionally mandatory
RGB_OPTIONAL = (148, 208, 80)          # green - good to have

# Blue, but the template's help text says "To be filled later - after QC processing".
QC_FILLED_COLUMNS = {"Product Data Status", "Disapproval Reason (if any)"}

# Flipkart attribute name -> field in the Amazon workbook.
FROM_AMAZON = {
    "Seller SKU ID":            "sku",
    "MRP (INR)":                "mrp",
    "Your selling price (INR)": "price",
    "Stock":                    "quantity",
    "HSN":                      "ext_prod_info_value",
    "Country Of Origin":        "country_of_origin",
    "Manufacturer Details":     "manufacturer_contact",
    "Packer Details":           "packer_contact",
    "Brand":                    "brand_name",
    "Brand Color":              "color",
    "Designed For":             "compatible_devices",
    "Main Image URL":           "main_image_url",
    "Other Image URL 1":        "other_image_url_1",
    "Other Image URL 2":        "other_image_url_2",
    "Other Image URL 3":        "other_image_url_3",
    "Description":              "product_description",
    "Type":                     "item_type_name",
    "Length (CM)":              "pkg_length",
    "Breadth (CM)":             "pkg_width",
    "Height (CM)":              "pkg_height",
    "Weight (KG)":              "pkg_weight",
    "Warranty Summary":         "warranty_description",
}

# Flipkart wants multi-values separated by '::' (see the template's Summary Sheet).
MULTIVALUE_SEP = "::"

SKU_MAX_LEN = 64  # "Text - limited to 64 characters (including spaces)"


def read_template(path):
    """(headers, example_row, colour_by_header) from the Flipkart .xls template."""
    bk = xlrd.open_workbook(path, formatting_info=True)
    sh = bk.sheet_by_name(SHEET_NAME)

    headers = [str(sh.cell_value(HEADER_ROW, c)).strip() for c in range(sh.ncols)]
    header_rows = [
        [sh.cell_value(r, c) for c in range(sh.ncols)]
        for r in range(DATA_START_ROW)
    ]

    colours = {}
    for c in range(sh.ncols):
        xf = bk.xf_list[sh.cell_xf_index(HEADER_ROW, c)]
        colours[c] = bk.colour_map.get(xf.background.pattern_colour_index)

    return headers, header_rows, colours


def read_amazon_children(path):
    """Every child row of an Amazon workbook as {field: value}, parents dropped."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[amazon.TEMPLATE_SHEET_NAME]
    col, _, _ = amazon.build_column_map(ws)

    children = []
    for r in range(amazon.DATA_START_ROW, ws.max_row + 1):
        row = {}
        for field, idx in col.items():
            v = ws.cell(row=r, column=idx + 1).value
            row[field] = None if v in (None, "") else str(v).strip()
        if not row.get("sku"):
            continue
        if row.get("parentage_level") == "Parent":
            continue
        children.append(row)

    wb.close()
    return children


def read_flipkart_master(path):
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    if "model_key" not in df.columns:
        raise ValueError(f"{path} is missing required column: model_key")
    df = df[df["model_key"].notna()].reset_index(drop=True)
    df["model_key"] = df["model_key"].astype(str).str.strip()
    return df


def master_value(master_row, name):
    raw = master_row.get(name)
    if raw is None or pd.isna(raw) or str(raw).strip() == "":
        return None
    return str(raw).strip()


def to_multivalue(text):
    """'a skin, a wrap, a sticker' -> 'a skin::a wrap::a sticker'"""
    if not text:
        return None
    parts = [p.strip() for p in str(text).split(",") if p.strip()]
    return MULTIVALUE_SEP.join(parts) if parts else None


def build_model_name(master_row, child, model_key):
    """Flipkart forms the product identity from Brand + Model Name + Brand Color, so
    Model Name is what groups variants together (there is no Group ID column here).

    A fixed 'Model Name' in flipkart_master puts all designs in one variant family.
    Left blank, it mirrors the Amazon parent chunks: one family per parent SKU.
    """
    override = master_value(master_row, "Model Name")
    if override:
        return override
    parent = child.get("parent_sku") or ""
    suffix = parent.rsplit("_parent_", 1)[-1] if "_parent_" in parent else "1"
    return f"{model_key.replace('_', ' ').strip()} PL-{suffix}"


def build_row(child, master_row, headers, model_key):
    """One Flipkart data row, as a list aligned to the template's columns."""
    row = [None] * len(headers)
    by_name = {}

    for name in headers:
        if not name or name in QC_FILLED_COLUMNS:
            continue
        if name in FROM_AMAZON:
            by_name[name] = child.get(FROM_AMAZON[name])
        else:
            by_name[name] = master_value(master_row, name)

    by_name["Model Name"] = build_model_name(master_row, child, model_key)
    by_name["Search Keywords"] = (
        master_value(master_row, "Search Keywords")
        or to_multivalue(child.get("generic_keyword"))
    )
    if not master_value(master_row, "Key Features"):
        bullets = [child.get(f"bullet_point_{i}") for i in (1, 2, 3, 4)]
        bullets = [b for b in bullets if b]
        by_name["Key Features"] = MULTIVALUE_SEP.join(bullets) if bullets else None

    for i, name in enumerate(headers):
        row[i] = by_name.get(name)
    return row


def validate(rows, headers, colours, model_key):
    """Flag empty mandatory columns and over-long SKUs, using the template's own colours."""
    mandatory, conditional = [], []
    for i, name in enumerate(headers):
        if not name or name in QC_FILLED_COLUMNS:
            continue
        if colours.get(i) == RGB_MANDATORY:
            mandatory.append((i, name))
        elif colours.get(i) == RGB_CONDITIONAL:
            conditional.append((i, name))

    def empty_cols(cols):
        return [n for i, n in cols if all(r[i] in (None, "") for r in rows)]

    missing = empty_cols(mandatory)
    soft = empty_cols(conditional)

    long_skus = [
        r[headers.index("Seller SKU ID")]
        for r in rows
        if r[headers.index("Seller SKU ID")]
        and len(str(r[headers.index("Seller SKU ID")])) > SKU_MAX_LEN
    ]

    if missing:
        print(f"        MANDATORY still empty ({len(missing)}): {', '.join(missing)}")
    if soft:
        print(f"        conditional empty ({len(soft)}): {', '.join(soft)}")
    if long_skus:
        print(f"        {len(long_skus)} SKU(s) exceed {SKU_MAX_LEN} chars, e.g. {long_skus[0]}")
    return missing


def write_workbook(out_path, header_rows, rows, headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    for hr in header_rows:
        ws.append(["" if v is None else v for v in hr])
    for r in rows:
        ws.append(r)
    wb.save(out_path)
    wb.close()


def main():
    for p in (FLIPKART_TEMPLATE, FLIPKART_MASTER):
        if not os.path.exists(p):
            sys.exit(f"ERROR: not found: {p}")

    headers, header_rows, colours = read_template(FLIPKART_TEMPLATE)
    tally = Counter(colours.values())
    print(f"Template   : {FLIPKART_TEMPLATE} [{SHEET_NAME}] {len(headers)} columns")
    print(
        f"  mandatory={tally.get(RGB_MANDATORY, 0)} "
        f"conditional={tally.get(RGB_CONDITIONAL, 0)} "
        f"optional={tally.get(RGB_OPTIONAL, 0)} "
        f"flipkart-filled={tally.get(RGB_FLIPKART_FILLS, 0)}"
    )
    print(f"  data starts at sheet row {DATA_START_ROW + 1}")

    master_df = read_flipkart_master(FLIPKART_MASTER)
    unknown = [
        c for c in master_df.columns
        if c != "model_key" and c not in headers
    ]
    if unknown:
        print(f"  WARNING: flipkart_master columns not in template: {', '.join(unknown)}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"Models     : {len(master_df)} from {FLIPKART_MASTER}")

    for _, master_row in master_df.iterrows():
        model_key = master_row["model_key"]
        amazon_path = os.path.join(AMAZON_OUTPUT_DIR, f"{model_key}.xlsm")
        if not os.path.exists(amazon_path):
            print(f"  [SKIPPED] {model_key}: {amazon_path} not found - run "
                  f"generate_amazon_listings2.py first")
            continue

        children = read_amazon_children(amazon_path)
        if not children:
            print(f"  [SKIPPED] {model_key}: no child rows in {amazon_path}")
            continue

        rows = [build_row(c, master_row, headers, model_key) for c in children]
        out_path = os.path.join(OUTPUT_DIR, f"{model_key}.xlsx")
        write_workbook(out_path, header_rows, rows, headers)

        families = len({r[headers.index("Model Name")] for r in rows})
        print(f"  [OK] {model_key}: {len(rows)} listings, {families} variant "
              f"family(ies) -> {out_path}")

        no_image = sum(1 for r in rows if not r[headers.index("Main Image URL")])
        if no_image:
            print(f"        {no_image}/{len(rows)} rows have NO Main Image URL "
                  f"(mandatory) - run upload_images.py for {model_key}")
        validate(rows, headers, colours, model_key)

    print("Done.")


if __name__ == "__main__":
    main()
"""
Flipkart Bulk Listing Generator
================================
Second stage of the pipeline. Takes the Amazon workbook that was already built and
checked, and turns it into a Flipkart catalog upload:

    models_master.xlsx  ─┐
                         ├─→ outputs/<model>.xlsm ─→ outputs_flipkart/<model>.xlsx
    images_master.xlsx  ─┘   (generate_amazon_listings2)   (this script)

The Amazon file is the source of truth for product data, so anything fixed there flows
through. Flipkart-only fields (Theme, Material, Tax Code, procurement, shipping fees)
have no Amazon equivalent and come from flipkart/flipkart_master.xlsx.

Run:
    python generate_flipkart_listings.py
"""

import os
import sys
import warnings
from collections import Counter

import openpyxl
import pandas as pd
import xlrd

import generate_amazon_listings2 as amazon

warnings.simplefilter("ignore")

# ---------- CONFIG ----------
FLIPKART_TEMPLATE = "flipkart_excel.xls"
FLIPKART_MASTER = "flipkart_master_fixed.xlsx"
AMAZON_OUTPUT_DIR = "outputs"
OUTPUT_DIR = "outputs_flipkart"
SHEET_NAME = "mobile_skin"
HEADER_ROW = 0          # attribute names
EXAMPLE_ROW = 2         # Flipkart's own sample values
DATA_START_ROW = 4      # "5th Row onwards: You are required to start entering product data"

# Header-row fill colours encode who fills a column and whether it is required.
# Documented on the template's own "Summary Sheet".
RGB_FLIPKART_FILLS = (192, 192, 192)   # grey  - Flipkart fills these, leave blank
RGB_MANDATORY = (141, 180, 226)        # blue  - must be filled
RGB_CONDITIONAL = (204, 153, 255)      # purple- conditionally mandatory
RGB_OPTIONAL = (148, 208, 80)          # green - good to have

# Blue, but the template's help text says "To be filled later - after QC processing".
QC_FILLED_COLUMNS = {"Product Data Status", "Disapproval Reason (if any)"}

# Flipkart attribute name -> field in the Amazon workbook.
FROM_AMAZON = {
    "Seller SKU ID":            "sku",
    "MRP (INR)":                "mrp",
    "Your selling price (INR)": "price",
    "Stock":                    "quantity",
    "HSN":                      "ext_prod_info_value",
    "Country Of Origin":        "country_of_origin",
    "Manufacturer Details":     "manufacturer_contact",
    "Packer Details":           "packer_contact",
    "Brand":                    "brand_name",
    "Brand Color":              "color",
    "Designed For":             "compatible_devices",
    "Main Image URL":           "main_image_url",
    "Other Image URL 1":        "other_image_url_1",
    "Other Image URL 2":        "other_image_url_2",
    "Other Image URL 3":        "other_image_url_3",
    "Description":              "product_description",
    "Type":                     "item_type_name",
    "Length (CM)":              "pkg_length",
    "Breadth (CM)":             "pkg_width",
    "Height (CM)":              "pkg_height",
    "Weight (KG)":              "pkg_weight",
    "Warranty Summary":         "warranty_description",
}

# Flipkart wants multi-values separated by '::' (see the template's Summary Sheet).
MULTIVALUE_SEP = "::"

SKU_MAX_LEN = 64  # "Text - limited to 64 characters (including spaces)"


def read_template(path):
    """(headers, example_row, colour_by_header) from the Flipkart .xls template."""
    bk = xlrd.open_workbook(path, formatting_info=True)
    sh = bk.sheet_by_name(SHEET_NAME)

    headers = [str(sh.cell_value(HEADER_ROW, c)).strip() for c in range(sh.ncols)]
    header_rows = [
        [sh.cell_value(r, c) for c in range(sh.ncols)]
        for r in range(DATA_START_ROW)
    ]

    colours = {}
    for c in range(sh.ncols):
        xf = bk.xf_list[sh.cell_xf_index(HEADER_ROW, c)]
        colours[c] = bk.colour_map.get(xf.background.pattern_colour_index)

    return headers, header_rows, colours


def read_amazon_children(path):
    """Every child row of an Amazon workbook as {field: value}, parents dropped."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[amazon.TEMPLATE_SHEET_NAME]
    col, _, _ = amazon.build_column_map(ws)

    children = []
    for r in range(amazon.DATA_START_ROW, ws.max_row + 1):
        row = {}
        for field, idx in col.items():
            v = ws.cell(row=r, column=idx + 1).value
            row[field] = None if v in (None, "") else str(v).strip()
        if not row.get("sku"):
            continue
        if row.get("parentage_level") == "Parent":
            continue
        children.append(row)

    wb.close()
    return children


def read_flipkart_master(path):
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    if "model_key" not in df.columns:
        raise ValueError(f"{path} is missing required column: model_key")
    df = df[df["model_key"].notna()].reset_index(drop=True)
    df["model_key"] = df["model_key"].astype(str).str.strip()
    return df


def master_value(master_row, name):
    raw = master_row.get(name)
    if raw is None or pd.isna(raw) or str(raw).strip() == "":
        return None
    return str(raw).strip()


def to_multivalue(text):
    """'a skin, a wrap, a sticker' -> 'a skin::a wrap::a sticker'"""
    if not text:
        return None
    parts = [p.strip() for p in str(text).split(",") if p.strip()]
    return MULTIVALUE_SEP.join(parts) if parts else None


def build_model_name(master_row, child, model_key):
    """Flipkart forms the product identity from Brand + Model Name + Brand Color, so
    Model Name is what groups variants together (there is no Group ID column here).

    A fixed 'Model Name' in flipkart_master puts all designs in one variant family.
    Left blank, it mirrors the Amazon parent chunks: one family per parent SKU.
    """
    override = master_value(master_row, "Model Name")
    if override:
        return override
    parent = child.get("parent_sku") or ""
    suffix = parent.rsplit("_parent_", 1)[-1] if "_parent_" in parent else "1"
    return f"{model_key.replace('_', ' ').strip()} PL-{suffix}"


def build_row(child, master_row, headers, model_key):
    """One Flipkart data row, as a list aligned to the template's columns."""
    row = [None] * len(headers)
    by_name = {}

    for name in headers:
        if not name or name in QC_FILLED_COLUMNS:
            continue
        if name in FROM_AMAZON:
            by_name[name] = child.get(FROM_AMAZON[name])
        else:
            by_name[name] = master_value(master_row, name)

    # Brand Color: always "Multicolor", not the per-design "Design-N" value
    by_name["Brand Color"] = "Multicolor"

    # Model Name: title text + ", D-<n>"
    design_n = child.get("sku", "").rsplit("_D", 1)[-1]  # "iPhone_17_pro_D1" -> "1"
    model_name_base = build_model_name(master_row, child, model_key)
    by_name["Model Name"] = f"{model_name_base}, D-{design_n}" if design_n.isdigit() else model_name_base

    by_name["Search Keywords"] = (
        master_value(master_row, "Search Keywords")
        or to_multivalue(child.get("generic_keyword"))
    )
    if not master_value(master_row, "Key Features"):
        bullets = [child.get(f"bullet_point_{i}") for i in (1, 2, 3, 4)]
        bullets = [b for b in bullets if b]
        by_name["Key Features"] = MULTIVALUE_SEP.join(bullets) if bullets else None

    for i, name in enumerate(headers):
        row[i] = by_name.get(name)
    return row

def validate(rows, headers, colours, model_key):
    """Flag empty mandatory columns and over-long SKUs, using the template's own colours."""
    mandatory, conditional = [], []
    for i, name in enumerate(headers):
        if not name or name in QC_FILLED_COLUMNS:
            continue
        if colours.get(i) == RGB_MANDATORY:
            mandatory.append((i, name))
        elif colours.get(i) == RGB_CONDITIONAL:
            conditional.append((i, name))

    def empty_cols(cols):
        return [n for i, n in cols if all(r[i] in (None, "") for r in rows)]

    missing = empty_cols(mandatory)
    soft = empty_cols(conditional)

    long_skus = [
        r[headers.index("Seller SKU ID")]
        for r in rows
        if r[headers.index("Seller SKU ID")]
        and len(str(r[headers.index("Seller SKU ID")])) > SKU_MAX_LEN
    ]

    if missing:
        print(f"        MANDATORY still empty ({len(missing)}): {', '.join(missing)}")
    if soft:
        print(f"        conditional empty ({len(soft)}): {', '.join(soft)}")
    if long_skus:
        print(f"        {len(long_skus)} SKU(s) exceed {SKU_MAX_LEN} chars, e.g. {long_skus[0]}")
    return missing


def write_workbook(out_path, header_rows, rows, headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    for hr in header_rows:
        ws.append(["" if v is None else v for v in hr])
    for r in rows:
        ws.append(r)
    wb.save(out_path)
    wb.close()


def main():
    for p in (FLIPKART_TEMPLATE, FLIPKART_MASTER):
        if not os.path.exists(p):
            sys.exit(f"ERROR: not found: {p}")

    headers, header_rows, colours = read_template(FLIPKART_TEMPLATE)
    tally = Counter(colours.values())
    print(f"Template   : {FLIPKART_TEMPLATE} [{SHEET_NAME}] {len(headers)} columns")
    print(
        f"  mandatory={tally.get(RGB_MANDATORY, 0)} "
        f"conditional={tally.get(RGB_CONDITIONAL, 0)} "
        f"optional={tally.get(RGB_OPTIONAL, 0)} "
        f"flipkart-filled={tally.get(RGB_FLIPKART_FILLS, 0)}"
    )
    print(f"  data starts at sheet row {DATA_START_ROW + 1}")

    master_df = read_flipkart_master(FLIPKART_MASTER)
    unknown = [
        c for c in master_df.columns
        if c != "model_key" and c not in headers
    ]
    if unknown:
        print(f"  WARNING: flipkart_master columns not in template: {', '.join(unknown)}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"Models     : {len(master_df)} from {FLIPKART_MASTER}")

    for _, master_row in master_df.iterrows():
        model_key = master_row["model_key"]
        amazon_path = os.path.join(AMAZON_OUTPUT_DIR, f"{model_key}.xlsm")
        if not os.path.exists(amazon_path):
            print(f"  [SKIPPED] {model_key}: {amazon_path} not found - run "
                  f"generate_amazon_listings2.py first")
            continue

        children = read_amazon_children(amazon_path)
        if not children:
            print(f"  [SKIPPED] {model_key}: no child rows in {amazon_path}")
            continue

        rows = [build_row(c, master_row, headers, model_key) for c in children]
        out_path = os.path.join(OUTPUT_DIR, f"{model_key}.xlsx")
        write_workbook(out_path, header_rows, rows, headers)

        families = len({r[headers.index("Model Name")] for r in rows})
        print(f"  [OK] {model_key}: {len(rows)} listings, {families} variant "
              f"family(ies) -> {out_path}")

        no_image = sum(1 for r in rows if not r[headers.index("Main Image URL")])
        if no_image:
            print(f"        {no_image}/{len(rows)} rows have NO Main Image URL "
                  f"(mandatory) - run upload_images.py for {model_key}")
        validate(rows, headers, colours, model_key)

    print("Done.")


if __name__ == "__main__":
    main()
