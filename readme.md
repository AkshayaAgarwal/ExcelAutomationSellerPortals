🛒 Flipkart Listing Manager Pro
A dedicated Streamlit application for efficiently managing Flipkart bulk catalog files.
📋 Features
1. View & Edit Tab
Browse all designs individually
See all 4 size variants (XS, S, M, L) for each design
Quick edit pricing and stock for single design
Real-time preview of changes
2. Range Update Tab ⭐
Update multiple designs by number range
Example: Update Design 1 through Design 20 at once
Set MRP, Selling Price, and Stock for entire range
Optional: Update images for the range
3. Selective Update Tab
Hand-pick specific designs using multi-select
Update only the designs you want
Perfect for promotional pricing on select items
4. Image Manager Tab
Bulk update images for multiple designs
4 modes: Range, Selected, or All designs
Update Main Image + 4 Other Images
🚀 How to Use
Step 1: Install Requirements
bash
Copy
pip install streamlit pandas openpyxl xlrd
Step 2: Run the App
bash
Copy
streamlit run flipkart_listing_manager.py
Step 3: Upload Your File
Upload your Flipkart .xls or .xlsx file
The app automatically detects the 'sticker' sheet
Step 4: Make Updates
Choose your preferred method:
Method A: Range Update (Recommended for bulk)
Go to "🔢 Range Update" tab
Enter start and end design numbers (e.g., 1 to 25)
Enter new MRP, Selling Price, Stock
Click "Apply Updates to Range"
Method B: Selective Update
Go to "🎯 Selective Update" tab
Select specific designs from dropdown
Enter new values
Click "Update Selected Designs"
Method C: Individual Edit
Go to "📋 View & Edit" tab
Select a design from dropdown
Edit values directly in the table
Click "Update This Design"
Step 5: Download Updated File
Click "📥 Download Updated File" in sidebar
Upload the downloaded file to Flipkart Seller Hub
📊 Your File Structure
Based on your uploaded file:
Total Designs: 75 unique designs
Variants per Design: 4 (XS: 13.3", S: 14", M: 15.6", L: 17.3")
Total Rows: 300 (75 designs × 4 sizes)
Group ID Pattern: Flower2_DESIGN_TP_1 to Flower2_DESIGN_TP_75
SKU Pattern: Floral2_TP_XS_1, Floral2_TP_S_1, etc.
🎯 Common Use Cases
1. Update Prices for New Stock
plain
Copy
Design Range: 1-75
New MRP: 449
New Selling Price: 299
Stock: 50
2. Seasonal Sale (First 20 Designs)
plain
Copy
Design Range: 1-20
New Selling Price: 199 (reduced from 259)
3. Update Images for New Collection
plain
Copy
Select: Designs 30-45
New Main Image: https://yourdomain.com/new-collection.jpg
4. Restock Specific Designs
plain
Copy
Select: Designs 5, 12, 18, 25
New Stock: 100
⚠️ Important Notes
Backup: Always keep a backup of your original file
Validation: The app maintains Flipkart's required format
Images: When updating images, all 4 variants (XS, S, M, L) get the same images
SKU Changes: Currently preserves original SKU format
🛠️ Technical Details
Built with: Streamlit + Pandas
File Support: .xls (Excel 97-2003) and .xlsx (Excel 2007+)
Sheet Name: Automatically reads 'sticker' sheet
Data Rows: Skips first 3 header rows as per Flipkart template
📝 Flipkart Column Mapping
The app manages these key columns:
Seller SKU ID - Your product SKU
Group ID - Design grouping (Flower2_DESIGN_TP_X)
MRP (INR) - Maximum Retail Price
Your selling price (INR) - Selling price
Stock - Inventory quantity
Size in Number - 13.3, 14, 15.6, 17.3
Main Image URL - Primary product image
Other Image URL 1-4 - Additional images
🆘 Troubleshooting
File not loading?
Ensure it's a valid .xls or .xlsx file
Check if 'sticker' sheet exists
Try saving the file again from Excel
Changes not reflecting?
Click the update button after making changes
Check if you're in the correct tab
Refresh the page if needed
Export issues?
Ensure you have write permissions
Check disk space
Try a different browser
📞 Support
For issues or feature requests, please contact support.
Version: 1.0
Last Updated: 2024