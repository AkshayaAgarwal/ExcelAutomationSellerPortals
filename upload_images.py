"""
Image Upload Automation for Amazon Listings
============================================
Uploads images from a local folder to imgbb and updates images_master.xlsx

Usage:
    1. Get a free API key from https://imgbb.com/api
    2. Replace API_KEY below with your key
    3. Update MODEL_KEY and IMAGES_FOLDER as needed
    4. Run: python upload_images_to_imgbb.py

Output:
    - Adds/updates rows in images_master.xlsx
    - One row per image, with (model_key, design_number, main_image_url)
"""

import os
import sys
import re
import requests
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()
# ========== CONFIG ==========
API_KEY = os.getenv("API_KEY")
MODEL_KEY = "iphone17"  # Must match the model_key in models_master.xlsx
IMAGES_FOLDER = "images/iphone17"  # Folder containing your images (1.jpg, 2.jpg, etc.)
IMAGES_MASTER_PATH = "images_master.xlsx"

# Supported image formats
SUPPORTED_FORMATS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}


def upload_image_to_imgbb(image_path, api_key):
    """Upload a single image to imgbb and return the direct URL."""
    try:
        with open(image_path, "rb") as f:
            response = requests.post(
                "https://api.imgbb.com/1/upload",
                data={"key": api_key},
                files={"image": f}
            )

        if response.status_code == 200:
            data = response.json()
            if data.get("success"):
                return data["data"]["url"]
            else:
                print(f"❌ imgbb error for {image_path}: {data.get('error', 'Unknown error')}")
                return None
        else:
            print(f"❌ HTTP {response.status_code} for {image_path}")
            return None
    except Exception as e:
        print(f"❌ Error uploading {image_path}: {e}")
        return None


def get_design_number_from_filename(filename):
    """
    Extract design number from filename.
    Handles patterns like:
    - '1.jpg' → 1
    - 'iPhone 11 (With Logo)_1_L.jpg' → 1
    - 'Samsung_S24_Design_5_XS.jpg' → 5
    """
    # Pattern: underscore, digits, underscore, letter(s)
    match = re.search(r'_(\d+)_[A-Za-z]', filename)
    if match:
        return int(match.group(1))

    # Fallback: just the filename stem as number (e.g., '1.jpg' → 1)
    try:
        name_without_ext = Path(filename).stem
        return int(name_without_ext)
    except ValueError:
        return None


def load_or_create_images_master(path):
    """Load images_master.xlsx or create a new one with proper columns."""
    if os.path.exists(path):
        df = pd.read_excel(path, dtype=str)
        df.columns = [c.strip() for c in df.columns]
        return df
    else:
        print(f"📝 Creating new {path}...")
        return pd.DataFrame(columns=["model_key", "design_number", "main_image_url"])


def upload_folder_images(api_key, images_folder, model_key, output_path):
    """
    Upload all images from a folder and update images_master.xlsx

    Returns: (upload_count, skipped_count)
    """
    if not os.path.exists(images_folder):
        print(f"❌ Folder not found: {images_folder}")
        return 0, 0

    images = sorted([
        f for f in os.listdir(images_folder)
        if Path(f).suffix.lower() in SUPPORTED_FORMATS
    ])

    if not images:
        print(f"❌ No images found in {images_folder}")
        return 0, 0

    print(f"🖼️ Found {len(images)} images. Uploading to imgbb...\n")

    # Load existing images_master or create new
    df_master = load_or_create_images_master(output_path)

    # Convert to proper types
    if "design_number" in df_master.columns:
        df_master["design_number"] = pd.to_numeric(df_master["design_number"], errors="coerce")

    upload_count = 0
    skipped_count = 0
    new_rows = []

    for filename in images:
        design_num = get_design_number_from_filename(filename)
        if design_num is None:
            print(f"⏭️  Skipping {filename} (can't extract design number)")
            skipped_count += 1
            continue

        full_path = os.path.join(images_folder, filename)
        print(f"📤 Uploading {filename} (design #{design_num})...", end=" ", flush=True)

        url = upload_image_to_imgbb(full_path, api_key)
        if url:
            print(f"✅ {url[:50]}...")
            new_rows.append({
                "model_key": model_key,
                "design_number": design_num,
                "main_image_url": url
            })
            upload_count += 1
        else:
            print(f"⚠️ Failed")
            skipped_count += 1

    # Update master: remove old rows for this model, add new ones
    if not df_master.empty:
        df_master = df_master[df_master["model_key"] != model_key]

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        df_master = pd.concat([df_master, df_new], ignore_index=True)

    # Save
    df_master.to_excel(output_path, index=False)
    print(f"\n✅ Saved to {output_path}")

    return upload_count, skipped_count


def main():
    if API_KEY == "YOUR_IMGBB_API_KEY_HERE":
        print("❌ ERROR: API_KEY not set!")
        print("   1. Get a free key from https://imgbb.com/api")
        print("   2. Replace 'YOUR_IMGBB_API_KEY_HERE' in this script")
        print("   3. Re-run this script")
        sys.exit(1)

    print("=" * 70)
    print("📤 Image Upload Automation (imgbb)")
    print("=" * 70)
    print(f"Model: {MODEL_KEY}")
    print(f"Folder: {IMAGES_FOLDER}")
    print(f"Output: {IMAGES_MASTER_PATH}\n")

    uploaded, skipped = upload_folder_images(
        API_KEY,
        IMAGES_FOLDER,
        MODEL_KEY,
        IMAGES_MASTER_PATH
    )

    print("\n" + "=" * 70)
    print(f"📊 Result: {uploaded} uploaded, {skipped} skipped")
    print("=" * 70)

    if uploaded > 0:
        print("\n✨ Next step:")
        print(f"   Run: python generate_amazon_listings2.py")
        print(f"   (Your images_master.xlsx is ready)")


if __name__ == "__main__":
    main()
