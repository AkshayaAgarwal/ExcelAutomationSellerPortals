"""
Amazon Bulk Listing Generator
==============================
Generates one Amazon-upload-ready .xlsm file per phone model, using:
  1. BASE_TEMPLATE   - your real Amazon "Template" workbook (structure/columns/validation kept intact)
  2. models_master.xlsx  - one row per model, the ~35 fields that are the SAME across all
                            designs of that model (bullets, price, keywords, etc.)
  3. images_master.xlsx  - one row per (model, design number), the main image URL for that design

Run:
    python generate_amazon_listings.py

Output:
    outputs/<model_key>.xlsm   (one file per model row in models_master.xlsx)
"""

import io
import os
import re
import sys
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd

# ---------- CONFIG ----------
BASE_TEMPLATE_PATH = "ELECTRONIC_DEVICE_SKIN.xlsm"   # if this exact name isn't found, script auto-detects any .xlsm in the folder
MODELS_MASTER_PATH = "models_master.xlsx"
IMAGES_MASTER_PATH = "images_master.xlsx"
OUTPUT_DIR = "outputs"
TEMPLATE_SHEET_NAME = "Template"
DATA_START_ROW = 7          # row 7 is the first real data row (rows 1-6 are Amazon headers)
CHUNK_SIZE_DEFAULT = 24     # designs per parent SKU, matches your existing file's pattern

# Fixed column positions in THIS exact template (0-indexed).
# Verified against ELECTRONIC_DEVICE_SKIN__3___Repaired_.xlsm -> Template sheet.
COL = {
    "sku":                     0,
    "listing_action":          1,
    "product_type":            2,
    "item_name":               3,
    "brand_name":              4,
    "exemption":                8,
    "browse_node":             9,
    "manufacturer":           16,
    "condition_type":         18,
    "tax_code":               20,
    "fulfillment_channel":    31,
    "quantity":                32,
    "inventory_always_avail": 35,
    "price":                  36,
    "mrp":                    37,
    "product_description":    46,
    "bullet_point_1":          47,
    "bullet_point_2":          48,
    "bullet_point_3":          49,
    "bullet_point_4":          50,
    "generic_keyword":         52,
    "color":                   67,
    "manufacturer_contact":    82,
    "compatible_devices":     101,
    "unit_count":             106,
    "unit_count_type":        107,
    "ext_prod_info_entity":   122,
    "ext_prod_info_value":    123,
    "parentage_level":        142,
    "child_relationship_type":143,
    "parent_sku":             144,
    "variation_theme":        145,
    "country_of_origin":      146,
    "warranty_description":   147,
    "main_image_url":         232,
    "other_image_url_1":      233,
    "other_image_url_2":      234,
    "other_image_url_3":      235,
    "pkg_length":             242,
    "pkg_length_unit":        243,
    "pkg_width":              244,
    "pkg_width_unit":         245,
    "pkg_height":             246,
    "pkg_height_unit":        247,
    "pkg_weight":             248,
    "pkg_weight_unit":        249,
}

# models_master.xlsx column -> COL key. Blank cell for a model = "keep base template's original value".
MASTER_FIELD_MAP = {
    "product_type":            "product_type",
    "brand_name":              "brand_name",
    "browse_node":              "browse_node",
    "manufacturer":            "manufacturer",
    "condition_type":          "condition_type",
    "tax_code":                "tax_code",
    "fulfillment_channel":     "fulfillment_channel",
    "quantity":                "quantity",
    "price":                   "price",
    "mrp":                     "mrp",
    "product_description":     "product_description",
    "bullet_point_1":          "bullet_point_1",
    "bullet_point_2":          "bullet_point_2",
    "bullet_point_3":          "bullet_point_3",
    "bullet_point_4":          "bullet_point_4",
    "generic_keyword":         "generic_keyword",
    "compatible_devices":      "compatible_devices",
    "unit_count":              "unit_count",
    "unit_count_type":         "unit_count_type",
    "ext_prod_info_entity":    "ext_prod_info_entity",
    "ext_prod_info_value":     "ext_prod_info_value",
    "country_of_origin":       "country_of_origin",
    "warranty_description":    "warranty_description",
    "other_image_url_1":       "other_image_url_1",
    "other_image_url_2":       "other_image_url_2",
    "other_image_url_3":       "other_image_url_3",
    "pkg_length":              "pkg_length",
    "pkg_length_unit":         "pkg_length_unit",
    "pkg_width":               "pkg_width",
    "pkg_width_unit":          "pkg_width_unit",
    "pkg_height":              "pkg_height",
    "pkg_height_unit":         "pkg_height_unit",
    "pkg_weight":              "pkg_weight",
    "pkg_weight_unit":         "pkg_weight_unit",
}

REQUIRED_MODEL_COLS = ["model_key", "sku_prefix", "title_prefix"]
# Optional column in models_master.xlsx: if a model has NO rows in images_master.xlsx,
# this many design rows get generated anyway, with main_image_url left blank for you to fill in later.
NUM_DESIGNS_FALLBACK_COL = "num_designs"


def load_base_row_template(ws):
    """Grab row 7 (first data row) as a static fallback template dict: col_index -> value."""
    row = [c.value for c in ws[DATA_START_ROW]]
    return row


def read_models_master(path):
    df = pd.read_excel(path, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for req in REQUIRED_MODEL_COLS:
        if req not in df.columns:
            raise ValueError(f"models_master.xlsx is missing required column: {req}")
    df = df[df["model_key"].notna()].reset_index(drop=True)
    # normalize: trim whitespace so " iPhone_17 " and "iPhone_17" are treated the same
    df["model_key"] = df["model_key"].astype(str).str.strip()
    return df


def read_images_master(path):
    df = pd.read_excel(path, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]
    required = {"model_key", "design_number", "main_image_url"}
    if not required.issubset(set(df.columns)):
        raise ValueError(f"images_master.xlsx must have columns: {sorted(required)}")
    df = df[df["model_key"].notna() & df["design_number"].notna()].copy()
    df["model_key"] = df["model_key"].astype(str).str.strip()
    df["design_number"] = df["design_number"].astype(int)
    return df


def check_model_key_alignment(models_df, images_df):
    """Warn loudly (with suggestions) if models_master and images_master model_keys don't line up.
    Models that deliberately supplied num_designs (no images yet, on purpose) are excluded."""
    if NUM_DESIGNS_FALLBACK_COL in models_df.columns:
        has_fallback = models_df[NUM_DESIGNS_FALLBACK_COL].notna() & (
            models_df[NUM_DESIGNS_FALLBACK_COL].astype(str).str.strip() != ""
        )
    else:
        has_fallback = pd.Series(False, index=models_df.index)
    model_keys = set(models_df.loc[~has_fallback, "model_key"])
    image_keys = set(images_df["model_key"])
    missing = model_keys - image_keys

    if not missing:
        return

    print("\n" + "=" * 70)
    print("WARNING: model_key mismatch between models_master.xlsx and images_master.xlsx")
    print("=" * 70)
    print(f"models_master.xlsx has these model_key(s) with NO matching images:")
    for mk in sorted(missing):
        print(f"    '{mk}'")
        # find close matches (case-insensitive) to help spot typos
        close = [ik for ik in image_keys if ik.lower() == mk.lower()]
        if close:
            print(f"        -> did you mean '{close[0]}' (found in images_master.xlsx, different case)?")
    print(f"\nmodel_key values actually present in images_master.xlsx:")
    for ik in sorted(image_keys):
        print(f"    '{ik}'")
    print("=" * 70 + "\n")


def build_title(title_prefix, design_number):
    """Appends the standard ' D-<n>' suffix your listings already use."""
    return f"{title_prefix.rstrip()} D-{design_number}"


def generate_model_workbook(base_bytes, model_row, images_df, chunk_size, template_row_values):
    """Returns an in-memory openpyxl workbook for one model, fully populated."""
    wb = openpyxl.load_workbook(io.BytesIO(base_bytes), keep_vba=True, read_only=False)
    ws = wb[TEMPLATE_SHEET_NAME]

    model_key = model_row["model_key"]
    sku_prefix = model_row["sku_prefix"]
    title_prefix = model_row["title_prefix"]

    model_images = images_df[images_df["model_key"] == model_key].sort_values("design_number")

    if model_images.empty:
        # No image rows for this model -> fall back to num_designs (if given), generate with blank image URLs
        num_designs_raw = model_row.get(NUM_DESIGNS_FALLBACK_COL)
        if pd.isna(num_designs_raw) or str(num_designs_raw).strip() == "":
            raise ValueError(
                f"No rows found in images_master.xlsx for model_key='{model_key}', and no "
                f"'{NUM_DESIGNS_FALLBACK_COL}' value given in models_master.xlsx either. "
                f"Add image rows, or fill in '{NUM_DESIGNS_FALLBACK_COL}' for this model to generate "
                f"it with blank image URLs."
            )
        n_designs = int(float(num_designs_raw))
        design_numbers = list(range(1, n_designs + 1))
        image_urls = [None] * n_designs
        images_missing = True
    else:
        design_numbers = model_images["design_number"].tolist()
        n_designs = len(design_numbers)
        image_urls = model_images["main_image_url"].tolist()
        images_missing = False

    n_parents = (n_designs + chunk_size - 1) // chunk_size

    # ---- Wipe existing data rows beyond header, then rewrite ----
    max_existing_row = ws.max_row
    if max_existing_row >= DATA_START_ROW:
        ws.delete_rows(DATA_START_ROW, max_existing_row - DATA_START_ROW + 1)

    current_row = DATA_START_ROW

    # ---- Write parent rows first ----
    parent_skus = [f"{sku_prefix}_parent_{k+1}" for k in range(n_parents)]
    for k, parent_sku in enumerate(parent_skus):
        row_vals = list(template_row_values)  # start from a full-width blank/base row
        row_vals = [None] * len(row_vals)
        row_vals[COL["sku"]] = parent_sku
        row_vals[COL["listing_action"]] = "Create or Replace (Full Update)"
        row_vals[COL["product_type"]] = model_row.get("product_type") or template_row_values[COL["product_type"]]
        row_vals[COL["item_name"]] = build_title(title_prefix, design_numbers[min(k * chunk_size, n_designs - 1)])
        row_vals[COL["parentage_level"]] = "Parent"
        row_vals[COL["variation_theme"]] = "COLOR"
        for c, v in enumerate(row_vals):
            ws.cell(row=current_row, column=c + 1, value=v)
        current_row += 1

    # Fields that must NEVER silently inherit the base template's value (they're inherently
    # model-specific, e.g. "Compatible Devices" would otherwise stay as whatever phone the base
    # template was originally built for). If left blank in models_master.xlsx, auto-derive from
    # model_key/title_prefix instead of falling back to the base template.
    friendly_model_name = model_key.replace("_", " ").strip()
    model_specific_defaults = {
        "compatible_devices": friendly_model_name,
        "generic_keyword": f"{friendly_model_name} skin, {friendly_model_name} wrap, {friendly_model_name} sticker",
    }

    # ---- Write child rows ----
    for idx, design_n in enumerate(design_numbers):
        parent_idx = idx // chunk_size
        parent_sku = parent_skus[parent_idx]
        image_url = image_urls[idx]

        row_vals = [None] * len(template_row_values)
        row_vals[COL["sku"]] = f"{sku_prefix}_D{design_n}"
        row_vals[COL["listing_action"]] = "Create or Replace (Full Update)"
        row_vals[COL["item_name"]] = build_title(title_prefix, design_n)
        row_vals[COL["color"]] = f"Design-{design_n}"
        row_vals[COL["main_image_url"]] = image_url
        row_vals[COL["parentage_level"]] = "Child"
        row_vals[COL["child_relationship_type"]] = "Variation"
        row_vals[COL["parent_sku"]] = parent_sku
        row_vals[COL["variation_theme"]] = "COLOR"

        # Apply every field from MASTER_FIELD_MAP: use model override if present,
        # else a smart per-model default if one exists, else the base template's value
        for master_col, col_key in MASTER_FIELD_MAP.items():
            override = model_row.get(master_col)
            if pd.notna(override) and str(override).strip() != "":
                row_vals[COL[col_key]] = override
            elif master_col in model_specific_defaults:
                row_vals[COL[col_key]] = model_specific_defaults[master_col]
            else:
                row_vals[COL[col_key]] = template_row_values[COL[col_key]]

        for c, v in enumerate(row_vals):
            ws.cell(row=current_row, column=c + 1, value=v)
        current_row += 1

    return wb, n_designs, n_parents, images_missing


def resolve_base_template_path():
    """Use BASE_TEMPLATE_PATH if it exists; otherwise auto-detect the one .xlsm in this folder."""
    if os.path.exists(BASE_TEMPLATE_PATH):
        return BASE_TEMPLATE_PATH

    candidates = [f for f in os.listdir(".") if f.lower().endswith(".xlsm")]
    if len(candidates) == 1:
        print(f"NOTE: '{BASE_TEMPLATE_PATH}' not found. Using the only .xlsm file in this folder instead: '{candidates[0]}'")
        return candidates[0]
    elif len(candidates) == 0:
        sys.exit(
            f"ERROR: base template not found. Expected '{BASE_TEMPLATE_PATH}' and no other .xlsm "
            f"file exists in this folder. Put your Amazon template .xlsm file in the same folder as this script."
        )
    else:
        sys.exit(
            f"ERROR: '{BASE_TEMPLATE_PATH}' not found, and there are multiple .xlsm files in this folder "
            f"({', '.join(candidates)}), so I can't guess which one to use.\n"
            f"Either rename the correct one to '{BASE_TEMPLATE_PATH}', or edit BASE_TEMPLATE_PATH "
            f"at the top of this script to the exact filename."
        )


def main():
    base_template_path = resolve_base_template_path()
    if not os.path.exists(MODELS_MASTER_PATH):
        sys.exit(f"ERROR: models master not found: {MODELS_MASTER_PATH}")
    if not os.path.exists(IMAGES_MASTER_PATH):
        sys.exit(f"ERROR: images master not found: {IMAGES_MASTER_PATH}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(base_template_path, "rb") as f:
        base_bytes = f.read()

    # Grab row 7 of the base template once, as fallback values for any field a model doesn't override
    probe_wb = openpyxl.load_workbook(io.BytesIO(base_bytes), keep_vba=True, read_only=False)
    probe_ws = probe_wb[TEMPLATE_SHEET_NAME]
    template_row_values = load_base_row_template(probe_ws)

    models_df = read_models_master(MODELS_MASTER_PATH)
    images_df = read_images_master(IMAGES_MASTER_PATH)

    print(f"Loaded {len(models_df)} model(s) from {MODELS_MASTER_PATH}")
    check_model_key_alignment(models_df, images_df)

    for _, model_row in models_df.iterrows():
        model_key = model_row["model_key"]
        chunk_size = int(model_row["chunk_size"]) if pd.notna(model_row.get("chunk_size")) else CHUNK_SIZE_DEFAULT

        try:
            wb, n_designs, n_parents, images_missing = generate_model_workbook(
                base_bytes, model_row, images_df, chunk_size, template_row_values
            )
        except ValueError as e:
            print(f"  [SKIPPED] {model_key}: {e}")
            continue

        out_path = os.path.join(OUTPUT_DIR, f"{model_key}.xlsm")
        wb.save(out_path)
        note = "  (NOTE: main_image_url left BLANK for all rows - fill in before uploading to Amazon)" if images_missing else ""
        print(f"  [OK] {model_key}: {n_designs} designs, {n_parents} parent SKU(s) -> {out_path}{note}")

    print("Done.")


if __name__ == "__main__":
    main()
