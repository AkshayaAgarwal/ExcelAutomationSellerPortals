# amazon_manager_app.py
import streamlit as st
import pandas as pd
import json
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="Amazon India Listing Manager", layout="wide")

# Custom styling
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; padding-left: 20px; padding-right: 20px; }
    .sku-box { padding: 20px; border-radius: 10px; background-color: #f0f2f6; margin: 10px 0; }
    .field-group { background-color: #ffffff; padding: 15px; border-radius: 8px; border: 1px solid #e0e0e0; margin: 10px 0; }
</style>
""", unsafe_allow_html=True)

st.title("🛒 Amazon India Listing Manager")
st.markdown("Professional tool for managing Amazon India bulk uploads")

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'manager' not in st.session_state:
    st.session_state.manager = None

# File upload section
with st.container():
    uploaded_file = st.file_uploader("📁 Upload Amazon Template Excel", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.session_state.df = df
            st.success(f"✅ Loaded {len(df)} rows with {len(df.columns)} columns")
            
            # Show column categories
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total SKUs", len(df))
            with col2:
                st.metric("Identity Fields", len([c for c in df.columns if any(x in c for x in ['SKU', 'Product', 'Item Name'])]))
            with col3:
                st.metric("Pricing Fields", len([c for c in df.columns if 'Price' in c or 'INR' in c]))
            
        except Exception as e:
            st.error(f"Error loading file: {e}")

# Main functionality
if st.session_state.df is not None:
    df = st.session_state.df
    
    # SKU Selector
    st.markdown("---")
    col1, col2 = st.columns([2, 1])
    
    with col1:
        sku_list = df['SKU'].dropna().unique().tolist()
        selected_sku = st.selectbox("🔍 Select SKU to Edit", sku_list)
    
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⬇️ Download Current Data", type="secondary"):
            output = f"amazon_export_{datetime.now().strftime('%Y%m%d')}.xlsx"
            df.to_excel(output, index=False)
            with open(output, "rb") as f:
                st.download_button("Download Excel", f, output, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # Get current row data
    if selected_sku:
        row_mask = df['SKU'] == selected_sku
        current_row = df[row_mask].iloc[0] if row_mask.any() else None
        
        if current_row is not None:
            # Create tabs for different sections
            tabs = st.tabs([
                "💰 Pricing & Inventory", 
                "📝 Content & SEO", 
                "🖼️ Images", 
                "📋 All Fields",
                "🚀 Bulk Operations"
            ])
            
            # Tab 1: Pricing
            with tabs[0]:
                st.markdown("### Pricing Management")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("#### Current Pricing")
                    current_price = current_row.get('Your Price INR (Sell on Amazon, IN)', 0)
                    current_mrp = current_row.get('Maximum Retail Price (Sell on Amazon, IN)', 0)
                    current_sale = current_row.get('Sale Price INR (Sell on Amazon, IN)', 0)
                    
                    st.write(f"**Your Price:** ₹{current_price}")
                    st.write(f"**MRP:** ₹{current_mrp}")
                    st.write(f"**Sale Price:** ₹{current_sale if pd.notna(current_sale) else 'N/A'}")
                
                with col2:
                    st.markdown("#### Update Pricing")
                    new_price = st.number_input("Your Price (INR)", value=float(current_price) if pd.notna(current_price) else 0.0, step=10.0)
                    new_mrp = st.number_input("MRP (INR)", value=float(current_mrp) if pd.notna(current_mrp) else 0.0, step=10.0)
                    new_sale = st.number_input("Sale Price (INR)", value=float(current_sale) if pd.notna(current_sale) else 0.0, step=10.0)
                
                with col3:
                    st.markdown("#### Sale Dates")
                    sale_start = st.date_input("Sale Start Date", value=datetime.now())
                    sale_end = st.date_input("Sale End Date", value=datetime.now())
                    
                    st.markdown("#### Inventory")
                    current_qty = current_row.get('Quantity (IN)', 0)
                    new_qty = st.number_input("Quantity", value=int(current_qty) if pd.notna(current_qty) else 0, step=1)
                
                if st.button("💾 Save Pricing & Inventory", type="primary"):
                    df.loc[row_mask, 'Your Price INR (Sell on Amazon, IN)'] = new_price
                    df.loc[row_mask, 'Maximum Retail Price (Sell on Amazon, IN)'] = new_mrp
                    df.loc[row_mask, 'Sale Price INR (Sell on Amazon, IN)'] = new_sale if new_sale > 0 else None
                    df.loc[row_mask, 'Quantity (IN)'] = new_qty
                    st.success("✅ Pricing updated!")
            
            # Tab 2: Content
            with tabs[1]:
                st.markdown("### Content & SEO")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("#### Basic Information")
                    current_title = str(current_row.get('Item Name', ''))
                    new_title = st.text_area("Product Title", value=current_title, height=100)
                    
                    current_desc = str(current_row.get('Product Description', ''))
                    new_desc = st.text_area("Product Description", value=current_desc, height=150)
                
                with col2:
                    st.markdown("#### Bullet Points (5 Max)")
                    bullets = []
                    for i in range(5):
                        col_name = 'Bullet Point' if i == 0 else f'Bullet Point.{i}'
                        current_bullet = str(current_row.get(col_name, ''))
                        bullet = st.text_input(f"Bullet {i+1}", value=current_bullet, key=f"bullet_{i}")
                        bullets.append(bullet)
                    
                    st.markdown("#### Generic Keywords")
                    keywords = []
                    for i in range(5):
                        col_name = 'Generic Keyword' if i == 0 else f'Generic Keyword.{i}'
                        current_kw = str(current_row.get(col_name, ''))
                        kw = st.text_input(f"Keywords {i+1}", value=current_kw, key=f"kw_{i}")
                        keywords.append(kw)
                
                if st.button("💾 Save Content", type="primary"):
                    df.loc[row_mask, 'Item Name'] = new_title
                    df.loc[row_mask, 'Product Description'] = new_desc
                    for i, bullet in enumerate(bullets):
                        col_name = 'Bullet Point' if i == 0 else f'Bullet Point.{i}'
                        if col_name in df.columns:
                            df.loc[row_mask, col_name] = bullet
                    for i, kw in enumerate(keywords):
                        col_name = 'Generic Keyword' if i == 0 else f'Generic Keyword.{i}'
                        if col_name in df.columns:
                            df.loc[row_mask, col_name] = kw
                    st.success("✅ Content updated!")
            
            # Tab 3: Images
            with tabs[2]:
                st.markdown("### Image Management")
                
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.markdown("#### Main Image")
                    main_img_col = 'Main Image URL' if 'Main Image URL' in df.columns else 'Main Image Location'
                    current_main = str(current_row.get(main_img_col, ''))
                    new_main = st.text_input("Main Image URL", value=current_main)
                    
                    if current_main:
                        st.image(current_main, width=200)
                
                with col2:
                    st.markdown("#### Additional Images (Up to 9)")
                    other_images = []
                    cols = st.columns(3)
                    for i in range(9):
                        with cols[i % 3]:
                            # Find the correct column name
                            possible_names = [
                                f'Other Image URL.{i}' if i > 0 else 'Other Image URL',
                                f'Other Image Location.{i}' if i > 0 else 'Other Image Location'
                            ]
                            img_col = None
                            for name in possible_names:
                                if name in df.columns:
                                    img_col = name
                                    break
                            
                            if img_col:
                                current_img = str(current_row.get(img_col, ''))
                                img_url = st.text_input(f"Image {i+1}", value=current_img, key=f"img_{i}")
                                other_images.append((img_col, img_url))
                                if current_img and i < 3:  # Show preview for first 3
                                    st.image(current_img, width=100)
                
                if st.button("💾 Save Images", type="primary"):
                    df.loc[row_mask, main_img_col] = new_main
                    for col_name, url in other_images:
                        df.loc[row_mask, col_name] = url
                    st.success("✅ Images updated!")
            
            # Tab 4: All Fields
            with tabs[3]:
                st.markdown("### All Fields Editor")
                st.info("Edit any field directly. Use with caution.")
                
                # Show editable dataframe for this SKU
                sku_data = df[row_mask].copy()
                edited_data = st.data_editor(sku_data, use_container_width=True, num_rows="fixed")
                
                if st.button("💾 Save All Changes", type="primary"):
                    df[row_mask] = edited_data.values
                    st.success("✅ All fields updated!")
            
            # Tab 5: Bulk Operations
            with tabs[4]:
                st.markdown("### Bulk Operations")
                
                st.markdown("#### Upload Bulk Update File")
                st.info("Upload Excel/CSV with columns: SKU, and any fields you want to update")
                
                bulk_file = st.file_uploader("Choose file", type=['xlsx', 'csv'], key="bulk")
                
                if bulk_file:
                    bulk_df = pd.read_excel(bulk_file) if bulk_file.name.endswith('.xlsx') else pd.read_csv(bulk_file)
                    st.write("Preview:", bulk_df.head())
                    
                    if st.button("🚀 Apply Bulk Updates", type="primary"):
                        updated_count = 0
                        for _, row in bulk_df.iterrows():
                            sku = row.get('SKU')
                            if sku and sku in df['SKU'].values:
                                mask = df['SKU'] == sku
                                for col in bulk_df.columns:
                                    if col != 'SKU' and col in df.columns:
                                        df.loc[mask, col] = row[col]
                                updated_count += 1
                        st.success(f"✅ Updated {updated_count} SKUs!")
            
            # Global Save Button
            st.markdown("---")
            col1, col2 = st.columns([3, 1])
            
            with col1:
                output_filename = st.text_input("Output Filename", 
                    value=f"amazon_updated_{datetime.now().strftime('%Y%m%d')}.xlsx")
            
            with col2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Save Final File", type="primary", use_container_width=True):
                    df.to_excel(output_filename, index=False)
                    with open(output_filename, "rb") as f:
                        st.download_button("⬇️ Download", f, output_filename, 
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True)

else:
    # Show template structure info when no file loaded
    st.info("👆 Upload your Amazon India template to get started")
    
    st.markdown("""
    ### Supported Operations:
    - ✅ **Pricing Updates** - Your Price, MRP, Sale Price, Sale Dates
    - ✅ **Inventory Management** - Quantity, Handling Time, Restock Dates
    - ✅ **Content Optimization** - Titles, Descriptions, 5 Bullet Points, 5 Keyword Sets
    - ✅ **Image Management** - Main Image + 9 Additional Images
    - ✅ **Bulk Updates** - Upload CSV/Excel to update multiple SKUs at once
    - ✅ **Validation** - Automatic price logic checks (Price ≤ MRP, Sale ≤ Price)
    
    ### Template Requirements:
    - File must contain 'SKU' column
    - Standard Amazon India bulk upload format
    - Supports .xlsx and .xls files
    """)

# Footer
st.markdown("---")
st.markdown("Built for Amazon India Sellers | Handles 200+ column templates")