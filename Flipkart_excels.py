
import streamlit as st
import pandas as pd
import json
from datetime import datetime
from pathlib import Path
import re
import io

# Page configuration
st.set_page_config(
    page_title="Flipkart Listing Manager Pro",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #046bd2;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .design-card {
        background-color: #f8f9fa;
        border-left: 5px solid #046bd2;
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 5px;
    }
    .variant-row {
        background-color: #ffffff;
        border: 1px solid #dee2e6;
        padding: 0.5rem;
        margin: 0.25rem 0;
        border-radius: 3px;
    }
    .update-box {
        background-color: #e7f3ff;
        border: 2px solid #046bd2;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .success-msg {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #28a745;
    }
    .warning-msg {
        background-color: #fff3cd;
        color: #856404;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #ffc107;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-left: 20px;
        padding-right: 20px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'original_df' not in st.session_state:
    st.session_state.original_df = None
if 'selected_designs' not in st.session_state:
    st.session_state.selected_designs = []
if 'update_mode' not in st.session_state:
    st.session_state.update_mode = None

def load_flipkart_file(uploaded_file):
    """Load and clean Flipkart Excel file - handles both .xls and .xlsx"""
    try:
        # Get file extension
        file_name = uploaded_file.name.lower()
        
        # Read file bytes into buffer
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)  # Reset pointer for potential reuse
        
        # Try reading 'sticker' sheet first with appropriate engine
        df = None
        
        if file_name.endswith('.xls'):
            # For .xls files, use xlrd engine only
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='sticker', engine='xlrd')
            except ValueError as e:
                # If 'sticker' sheet doesn't exist, try first sheet
                if 'Worksheet named' in str(e) or 'No sheet named' in str(e):
                    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, engine='xlrd')
                else:
                    raise e
        else:
            # For .xlsx files, use openpyxl engine
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='sticker', engine='openpyxl')
            except ValueError as e:
                # If 'sticker' sheet doesn't exist, try first sheet
                if 'Worksheet named' in str(e) or 'No sheet named' in str(e):
                    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, engine='openpyxl')
                else:
                    raise e
        
        # Skip first 3 rows (headers) - Flipkart template usually has headers in row 4
        if len(df) > 3:
            df_clean = df.iloc[3:].reset_index(drop=True)
            # Set the first row as column headers
            df_clean.columns = df_clean.iloc[0]
            df_clean = df_clean.iloc[1:].reset_index(drop=True)
        else:
            df_clean = df
            
        return df_clean
        
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        st.info("💡 Tip: Make sure you're uploading a valid Flipkart bulk upload template (.xls or .xlsx)")
        return None

def get_design_groups(df):
    """Get unique design groups"""
    if df is None or 'Group ID' not in df.columns:
        return []
    return sorted(df['Group ID'].dropna().unique().tolist())

def get_design_data(df, group_id):
    """Get all variants for a specific design"""
    return df[df['Group ID'] == group_id].copy()

def update_design_by_pattern(df, start_design_num, end_design_num, updates):
    """
    Update multiple designs by pattern
    updates: dict with keys like 'mrp', 'selling_price', 'stock', 'images', etc.
    """
    df_updated = df.copy()

    for design_num in range(start_design_num, end_design_num + 1):
        group_id = f"Flower2_DESIGN_TP_{design_num}"
        mask = df_updated['Group ID'] == group_id

        if mask.any():
            # Update pricing
            if 'mrp' in updates and 'MRP (INR)' in df_updated.columns:
                df_updated.loc[mask, 'MRP (INR)'] = updates['mrp']
            if 'selling_price' in updates and 'Your selling price (INR)' in df_updated.columns:
                df_updated.loc[mask, 'Your selling price (INR)'] = updates['selling_price']
            if 'stock' in updates and 'Stock' in df_updated.columns:
                df_updated.loc[mask, 'Stock'] = updates['stock']

            # Update images (apply to all variants in the design)
            if 'main_image' in updates and 'Main Image URL' in df_updated.columns:
                df_updated.loc[mask, 'Main Image URL'] = updates['main_image']
            if 'other_images' in updates:
                for i, img_url in enumerate(updates['other_images'][:4]):
                    col_name = f'Other Image URL {i+1}' if i > 0 else 'Other Image URL 1'
                    if col_name in df_updated.columns:
                        df_updated.loc[mask, col_name] = img_url

            # Update SKU pattern if provided
            if 'sku_pattern' in updates and 'Seller SKU ID' in df_updated.columns:
                # Update each variant's SKU
                variants = df_updated[mask]
                for idx, row in variants.iterrows():
                    try:
                        size_suffix = row['Seller SKU ID'].split('_')[-2]  # XS, S, M, L
                        design_suffix = row['Seller SKU ID'].split('_')[-1]  # 1, 2, 3, etc.
                        new_sku = updates['sku_pattern'].format(
                            size=size_suffix, 
                            design=design_suffix,
                            num=design_num
                        )
                        df_updated.loc[idx, 'Seller SKU ID'] = new_sku
                    except:
                        pass  # Skip if SKU format is different

    return df_updated

def bulk_update_by_selection(df, selected_groups, updates):
    """Update selected design groups"""
    df_updated = df.copy()

    for group_id in selected_groups:
        mask = df_updated['Group ID'] == group_id

        if 'mrp' in updates and 'MRP (INR)' in df_updated.columns:
            df_updated.loc[mask, 'MRP (INR)'] = updates['mrp']
        if 'selling_price' in updates and 'Your selling price (INR)' in df_updated.columns:
            df_updated.loc[mask, 'Your selling price (INR)'] = updates['selling_price']
        if 'stock' in updates and 'Stock' in df_updated.columns:
            df_updated.loc[mask, 'Stock'] = updates['stock']
        if 'main_image' in updates and 'Main Image URL' in df_updated.columns:
            df_updated.loc[mask, 'Main Image URL'] = updates['main_image']
        if 'other_images' in updates:
            for i, img_url in enumerate(updates['other_images'][:4]):
                col_name = f'Other Image URL {i+1}' if i > 0 else 'Other Image URL 1'
                if col_name in df_updated.columns:
                    df_updated.loc[mask, col_name] = img_url

    return df_updated

def export_file(df, original_file_name):
    """Export updated file"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"Flipkart_Updated_{timestamp}.xlsx"

    # Create Excel writer
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='sticker', index=False)
    
    output_buffer.seek(0)
    return output_buffer, output_name

# ==================== MAIN UI ====================

st.markdown('<p class="main-header">🛒 Flipkart Listing Manager Pro</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Bulk Update Tool for Flipkart Catalog Files</p>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### 📁 File Operations")

    uploaded_file = st.file_uploader(
        "Upload Flipkart Excel File", 
        type=['xls', 'xlsx'],
        help="Upload your Flipkart bulk upload template file (.xls or .xlsx)"
    )

    if uploaded_file:
        if st.session_state.df is None or st.session_state.get('last_uploaded') != uploaded_file.name:
            with st.spinner("Loading file..."):
                df = load_flipkart_file(uploaded_file)
                if df is not None:
                    st.session_state.df = df
                    st.session_state.original_df = df.copy()
                    st.session_state.last_uploaded = uploaded_file.name
                    st.success(f"✅ Loaded {len(df)} rows")
                    
                    # Show column info for debugging
                    with st.expander("📋 Loaded Columns"):
                        st.write(list(df.columns))

    if st.session_state.df is not None:
        st.markdown("---")
        st.markdown("### 📊 File Stats")
        df = st.session_state.df

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Rows", len(df))
        with col2:
            designs_count = df['Group ID'].nunique() if 'Group ID' in df.columns else 0
            st.metric("Designs", designs_count)

        if designs_count > 0:
            st.metric("Variants/Design", int(len(df) / designs_count))

        st.markdown("---")
        st.markdown("### 💾 Export")

        if st.button("📥 Prepare Download", type="primary", use_container_width=True):
            output_buffer, output_file = export_file(st.session_state.df, uploaded_file.name if uploaded_file else "data")
            st.download_button(
                label="📥 Download Updated File",
                data=output_buffer,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# Main content
if st.session_state.df is not None:
    df = st.session_state.df
    
    # Check required columns exist
    required_cols = ['Group ID', 'Seller SKU ID', 'MRP (INR)', 'Your selling price (INR)', 'Stock']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        st.warning(f"⚠️ Missing expected columns: {missing_cols}. Available columns: {list(df.columns)}")
        st.info("The app will try to work with available columns. Please check if the file format is correct.")

    # Create tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 View & Edit", 
        "🔢 Range Update", 
        "🎯 Selective Update",
        "🖼️ Image Manager"
    ])

    # ==================== TAB 1: View & Edit ====================
    with tab1:
        st.markdown("### 📋 View and Individual Edit")

        col1, col2 = st.columns([1, 2])

        with col1:
            st.markdown("#### Select Design")
            designs = get_design_groups(df)
            
            if not designs:
                st.error("No designs found. Check if 'Group ID' column exists.")
                st.stop()
                
            selected_design = st.selectbox("Choose Design Group", designs)

            if selected_design:
                design_data = get_design_data(df, selected_design)

                st.markdown("#### Design Info")
                st.write(f"**Design:** {selected_design}")
                st.write(f"**Variants:** {len(design_data)}")
                
                if 'Model Name' in design_data.columns:
                    st.write(f"**Model:** {design_data.iloc[0]['Model Name']}")

                # Quick stats
                st.markdown("#### Current Pricing")
                if 'MRP (INR)' in design_data.columns:
                    st.write(f"MRP: ₹{design_data.iloc[0]['MRP (INR)']}")
                if 'Your selling price (INR)' in design_data.columns:
                    st.write(f"Selling: ₹{design_data.iloc[0]['Your selling price (INR)']}")
                if 'Stock' in design_data.columns:
                    st.write(f"Stock: {design_data.iloc[0]['Stock']} units")

        with col2:
            if selected_design:
                st.markdown("#### Variants in this Design")

                # Display variants
                display_cols = ['Seller SKU ID']
                if 'Size in Number' in design_data.columns:
                    display_cols.append('Size in Number')
                if 'MRP (INR)' in design_data.columns:
                    display_cols.append('MRP (INR)')
                if 'Your selling price (INR)' in design_data.columns:
                    display_cols.append('Your selling price (INR)')
                if 'Stock' in design_data.columns:
                    display_cols.append('Stock')
                if 'Main Image URL' in design_data.columns:
                    display_cols.append('Main Image URL')

                # Filter to only existing columns
                display_cols = [col for col in display_cols if col in design_data.columns]

                edited_data = st.data_editor(
                    design_data[display_cols],
                    use_container_width=True,
                    num_rows="fixed",
                    key=f"editor_{selected_design}"
                )

                # Individual update section
                st.markdown("#### ✏️ Quick Update This Design")

                col_mrp, col_price, col_stock = st.columns(3)

                with col_mrp:
                    current_mrp = float(design_data.iloc[0]['MRP (INR)']) if 'MRP (INR)' in design_data.columns else 399.0
                    new_mrp = st.number_input("New MRP", 
                                             value=current_mrp,
                                             step=10.0,
                                             key=f"mrp_{selected_design}")

                with col_price:
                    current_price = float(design_data.iloc[0]['Your selling price (INR)']) if 'Your selling price (INR)' in design_data.columns else 259.0
                    new_price = st.number_input("New Selling Price", 
                                               value=current_price,
                                               step=10.0,
                                               key=f"price_{selected_design}")

                with col_stock:
                    current_stock = int(design_data.iloc[0]['Stock']) if 'Stock' in design_data.columns else 20
                    new_stock = st.number_input("New Stock", 
                                               value=current_stock,
                                               step=1,
                                               key=f"stock_{selected_design}")

                if st.button("💾 Update This Design", type="primary", key=f"update_{selected_design}"):
                    mask = df['Group ID'] == selected_design
                    if 'MRP (INR)' in df.columns:
                        st.session_state.df.loc[mask, 'MRP (INR)'] = new_mrp
                    if 'Your selling price (INR)' in df.columns:
                        st.session_state.df.loc[mask, 'Your selling price (INR)'] = new_price
                    if 'Stock' in df.columns:
                        st.session_state.df.loc[mask, 'Stock'] = new_stock
                    st.success(f"✅ Updated {selected_design}!")
                    st.rerun()

    # ==================== TAB 2: Range Update ====================
    with tab2:
        st.markdown("### 🔢 Bulk Update by Design Range")
        st.info("Update multiple designs at once by specifying the design number range")

        designs = get_design_groups(df)
        total_designs = len(designs)

        if total_designs == 0:
            st.error("No designs available to update")
        else:
            st.markdown(f"**Available Designs:** 1 to {total_designs} (Total: {total_designs})")

            col1, col2 = st.columns(2)

            with col1:
                start_num = st.number_input("From Design #", 
                                           min_value=1, 
                                           max_value=total_designs,
                                           value=1,
                                           step=1)

            with col2:
                end_num = st.number_input("To Design #", 
                                         min_value=1, 
                                         max_value=total_designs,
                                         value=min(10, total_designs),
                                         step=1)

            if start_num > end_num:
                st.error("❌ Start number must be less than or equal to end number")
            else:
                st.markdown(f"#### Selected Range: Design {start_num} to {end_num}")
                st.write(f"This will update {end_num - start_num + 1} designs")

                # Show affected designs
                affected_designs = [f"Flower2_DESIGN_TP_{i}" for i in range(start_num, end_num + 1)]
                with st.expander("View Affected Designs"):
                    st.write(affected_designs)

                st.markdown("---")
                st.markdown("#### 📝 Enter Updates")

                col_mrp, col_price, col_stock = st.columns(3)

                with col_mrp:
                    range_mrp = st.number_input("MRP (INR)", 
                                               value=399.0, 
                                               step=10.0,
                                               key="range_mrp")

                with col_price:
                    range_price = st.number_input("Selling Price (INR)", 
                                                 value=259.0, 
                                                 step=10.0,
                                                 key="range_price")

                with col_stock:
                    range_stock = st.number_input("Stock Quantity", 
                                                 value=20, 
                                                 step=1,
                                                 key="range_stock")

                # Image updates
                st.markdown("#### 🖼️ Images (Optional)")
                update_images = st.checkbox("Also update images for this range?")

                image_updates = {}
                if update_images:
                    main_img = st.text_input("Main Image URL", 
                                            placeholder="https://...",
                                            key="range_main_img")

                    col_img1, col_img2 = st.columns(2)
                    with col_img1:
                        img1 = st.text_input("Other Image 1", placeholder="https://...", key="range_img1")
                        img2 = st.text_input("Other Image 2", placeholder="https://...", key="range_img2")
                    with col_img2:
                        img3 = st.text_input("Other Image 3", placeholder="https://...", key="range_img3")
                        img4 = st.text_input("Other Image 4", placeholder="https://...", key="range_img4")

                    image_updates = {
                        'main_image': main_img if main_img else None,
                        'other_images': [img for img in [img1, img2, img3, img4] if img]
                    }

                st.markdown("---")

                if st.button("🚀 Apply Updates to Range", type="primary", use_container_width=True):
                    with st.spinner("Updating..."):
                        updates = {
                            'mrp': range_mrp,
                            'selling_price': range_price,
                            'stock': range_stock
                        }

                        if update_images and image_updates.get('main_image'):
                            updates.update(image_updates)

                        st.session_state.df = update_design_by_pattern(
                            st.session_state.df, 
                            start_num, 
                            end_num, 
                            updates
                        )

                        st.success(f"✅ Successfully updated Designs {start_num} to {end_num}!")
                        st.balloons()

    # ==================== TAB 3: Selective Update ====================
    with tab3:
        st.markdown("### 🎯 Selective Design Update")
        st.info("Hand-pick specific designs to update")

        designs = get_design_groups(df)

        if not designs:
            st.error("No designs available")
        else:
            # Multi-select designs
            selected_designs = st.multiselect(
                "Select Designs to Update",
                options=designs,
                default=[],
                help="You can select multiple designs"
            )

            if selected_designs:
                st.markdown(f"**{len(selected_designs)} designs selected**")

                with st.expander("Preview Selected Designs"):
                    preview_data = []
                    for design in selected_designs[:5]:  # Show first 5
                        design_rows = df[df['Group ID'] == design]
                        if not design_rows.empty:
                            design_row = design_rows.iloc[0]
                            preview_data.append({
                                'Design': design,
                                'Model': design_row.get('Model Name', 'N/A'),
                                'Current MRP': design_row.get('MRP (INR)', 'N/A'),
                                'Current Price': design_row.get('Your selling price (INR)', 'N/A')
                            })
                    if preview_data:
                        st.table(pd.DataFrame(preview_data))

                    if len(selected_designs) > 5:
                        st.write(f"... and {len(selected_designs) - 5} more")

                st.markdown("---")
                st.markdown("#### 📝 Enter New Values")

                col1, col2, col3 = st.columns(3)

                with col1:
                    sel_mrp = st.number_input("New MRP", value=399.0, step=10.0, key="sel_mrp")
                with col2:
                    sel_price = st.number_input("New Selling Price", value=259.0, step=10.0, key="sel_price")
                with col3:
                    sel_stock = st.number_input("New Stock", value=20, step=1, key="sel_stock")

                if st.button("🚀 Update Selected Designs", type="primary", use_container_width=True):
                    with st.spinner("Updating selected designs..."):
                        updates = {
                            'mrp': sel_mrp,
                            'selling_price': sel_price,
                            'stock': sel_stock
                        }

                        st.session_state.df = bulk_update_by_selection(
                            st.session_state.df,
                            selected_designs,
                            updates
                        )

                        st.success(f"✅ Updated {len(selected_designs)} designs successfully!")
                        st.balloons()

    # ==================== TAB 4: Image Manager ====================
    with tab4:
        st.markdown("### 🖼️ Bulk Image Manager")
        st.info("Update images for multiple designs at once")

        designs = get_design_groups(df)

        if not designs:
            st.error("No designs available")
        else:
            image_update_mode = st.radio(
                "Select Update Mode",
                ["Update by Range", "Update Selected Designs", "Update All"],
                horizontal=True
            )

            target_designs = []

            if image_update_mode == "Update by Range":
                col1, col2 = st.columns(2)
                with col1:
                    img_start = st.number_input("From Design #", 1, len(designs), 1, key="img_start")
                with col2:
                    img_end = st.number_input("To Design #", 1, len(designs), min(10, len(designs)), key="img_end")
                target_designs = [f"Flower2_DESIGN_TP_{i}" for i in range(img_start, img_end + 1)]

            elif image_update_mode == "Update Selected Designs":
                target_designs = st.multiselect("Select Designs", designs, key="img_select")

            else:  # Update All
                target_designs = designs
                st.warning(f"⚠️ This will update images for ALL {len(designs)} designs")

            if target_designs:
                st.markdown(f"**Target: {len(target_designs)} designs**")

                st.markdown("#### Enter Image URLs")

                main_image = st.text_input("Main Image URL *", 
                                          placeholder="https://example.com/image.jpg",
                                          help="This will be applied to all selected designs")

                col1, col2 = st.columns(2)
                with col1:
                    other1 = st.text_input("Other Image 1", placeholder="https://...")
                    other2 = st.text_input("Other Image 2", placeholder="https://...")
                with col2:
                    other3 = st.text_input("Other Image 3", placeholder="https://...")
                    other4 = st.text_input("Other Image 4", placeholder="https://...")

                if st.button("🖼️ Apply Images", type="primary", use_container_width=True):
                    if not main_image:
                        st.error("❌ Main Image URL is required")
                    else:
                        with st.spinner("Updating images..."):
                            updates = {
                                'main_image': main_image,
                                'other_images': [img for img in [other1, other2, other3, other4] if img]
                            }

                            st.session_state.df = bulk_update_by_selection(
                                st.session_state.df,
                                target_designs,
                                updates
                            )

                            st.success(f"✅ Updated images for {len(target_designs)} designs!")
                            st.balloons()

else:
    # No file uploaded yet
    st.markdown("""
    <div style="text-align: center; padding: 3rem; background-color: #f8f9fa; border-radius: 10px; margin-top: 2rem;">
        <h2>👋 Welcome to Flipkart Listing Manager Pro</h2>
        <p style="font-size: 1.1rem; color: #666;">
            Upload your Flipkart bulk upload file to get started.<br>
            Supports both <strong>.xls</strong> (Excel 97-2003) and <strong>.xlsx</strong> (Excel 2007+) formats.<br>
            You can update prices, stock, images, and more for multiple designs at once.
        </p>
        <br>
        <h4>✨ Features:</h4>
        <ul style="text-align: left; display: inline-block; color: #555;">
            <li>📋 View and edit individual designs</li>
            <li>🔢 Bulk update by design number range (e.g., Design 1-20)</li>
            <li>🎯 Selective update - pick specific designs</li>
            <li>🖼️ Bulk image management</li>
            <li>💾 Export updated file ready for upload</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("<p style='text-align: center; color: #666;'>Flipkart Listing Manager Pro | Built for efficient catalog management</p>", unsafe_allow_html=True)