import streamlit as st
import pandas as pd
import openpyxl
import os
from io import BytesIO
import tempfile

# Set page config
st.set_page_config(
    page_title="Price Editor",
    page_icon="💰",
    layout="wide"
)

def load_excel_file(file_path):
    """Load Excel file and return workbook"""
    try:
        return openpyxl.load_workbook(file_path)
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def get_current_prices(workbook):
    """Get current values from D4-D7 in Prices sheet"""
    if 'Prices' not in workbook.sheetnames:
        return None
    
    prices_sheet = workbook['Prices']
    current_prices = {
        'D4': prices_sheet['D4'].value,
        'D5': prices_sheet['D5'].value,
        'D6': prices_sheet['D6'].value,
        'D7': prices_sheet['D7'].value
    }
    return current_prices

def update_prices(workbook, new_prices):
    """Update D4-D7 cells in Prices sheet"""
    if 'Prices' not in workbook.sheetnames:
        return False
    
    prices_sheet = workbook['Prices']
    prices_sheet['D4'] = new_prices['D4']
    prices_sheet['D5'] = new_prices['D5']
    prices_sheet['D6'] = new_prices['D6']
    prices_sheet['D7'] = new_prices['D7']
    return True

def export_sheet_to_csv(workbook, sheet_name):
    """Export specified sheet to CSV bytes"""
    if sheet_name not in workbook.sheetnames:
        return None
    
    # Create a temporary file to save the workbook
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
        workbook.save(tmp_file.name)
        
        # Read the sheet with pandas
        df = pd.read_excel(tmp_file.name, sheet_name=sheet_name)
        
        # Convert to CSV
        csv_buffer = BytesIO()
        csv_string = df.to_csv(index=False)
        csv_buffer.write(csv_string.encode())
        csv_buffer.seek(0)
        
        # Clean up temp file
        os.unlink(tmp_file.name)
        
        return csv_buffer.getvalue()

def main():
    st.title("💰 Excel Price Editor")
    st.markdown("Edit prices in cells D4-D7 of the 'Prices' sheet and export the 'Export' sheet as CSV")
    
    # File handling
    excel_file_path = "your_workbook.xlsx"  # Default bundled file
    
    # Check if bundled file exists
    if os.path.exists(excel_file_path):
        st.success(f"✅ Using file: {excel_file_path}")
        use_bundled = True
    else:
        use_bundled = False
        st.warning("⚠️ Default file not found. Please upload an Excel file.")
    
    # File uploader (always show as backup option)
    uploaded_file = st.file_uploader(
        "📁 Upload Excel File (optional - overrides default)",
        type=['xlsx', 'xls'],
        help="Upload your Excel file containing 'Prices' and 'Export' sheets"
    )
    
    # Determine which file to use
    workbook = None
    temp_file_path = None
    
    if uploaded_file is not None:
        # Save uploaded file to temp location
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_file_path = tmp_file.name
        
        workbook = load_excel_file(temp_file_path)
        st.info("📤 Using uploaded file")
        
    elif use_bundled:
        workbook = load_excel_file(excel_file_path)
    
    if workbook is None:
        st.stop()
    
    # Show available sheets
    st.sidebar.subheader("📋 Available Sheets")
    for sheet in workbook.sheetnames:
        icon = "💰" if sheet == "Prices" else "📊" if sheet == "Export" else "📄"
        st.sidebar.write(f"{icon} {sheet}")
    
    # Main interface
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("✏️ Edit Prices (Cells D4-D7)")
        
        # Get current prices
        current_prices = get_current_prices(workbook)
        
        if current_prices is None:
            st.error("❌ 'Prices' sheet not found in the workbook")
            st.stop()
        
        # Create form for editing prices
        with st.form("price_editor"):
            st.write("**Current Values:**")
            
            # Input fields for each price cell
            new_d4 = st.text_input(
                "Gold:", 
                value=str(current_prices['D4']) if current_prices['D4'] is not None else "",
                help="Enter new value for cell D4"
            )
            
            new_d5 = st.text_input(
                "Silver:", 
                value=str(current_prices['D5']) if current_prices['D5'] is not None else "",
                help="Enter new value for cell D5"
            )
            
            new_d6 = st.text_input(
                "Platinum:", 
                value=str(current_prices['D6']) if current_prices['D6'] is not None else "",
                help="Enter new value for cell D6"
            )
            
            new_d7 = st.text_input(
                "Palladium:", 
                value=str(current_prices['D7']) if current_prices['D7'] is not None else "",
                help="Enter new value for cell D7"
            )
            
            # Submit button
            submitted = st.form_submit_button("💾 Update Prices", type="primary")
            
            if submitted:
                # Prepare new prices (try to convert to float if possible, otherwise keep as string)
                new_prices = {}
                for cell, value in [('D4', new_d4), ('D5', new_d5), ('D6', new_d6), ('D7', new_d7)]:
                    try:
                        # Try to convert to float if it's a number
                        new_prices[cell] = float(value) if value and value.replace('.', '').replace('-', '').isdigit() else value
                    except:
                        new_prices[cell] = value
                
                # Update the workbook
                if update_prices(workbook, new_prices):
                    # Save the file
                    file_to_save = temp_file_path if temp_file_path else excel_file_path
                    try:
                        workbook.save(file_to_save)
                        st.success("✅ Prices updated successfully!")
                        
                        # Show updated values
                        st.write("**Updated Values:**")
                        for cell, value in new_prices.items():
                            st.write(f"• {cell}: {value}")
                        
                        # Auto-rerun to refresh the interface
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ Error saving file: {e}")
                else:
                    st.error("❌ Failed to update prices")
    
    with col2:
        st.subheader("📊 Export Options")
        
        # Check if Export sheet exists
        if 'Export' in workbook.sheetnames:
            st.write("**Export Sheet Available** ✅")
            
            # Preview Export sheet
            with st.expander("👁️ Preview Export Sheet"):
                try:
                    # Create temp file to read with pandas
                    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                        workbook.save(tmp.name)
                        preview_df = pd.read_excel(tmp.name, sheet_name='Export')
                        st.dataframe(preview_df.head(10), use_container_width=True)
                        if len(preview_df) > 10:
                            st.write(f"... and {len(preview_df) - 10} more rows")
                        os.unlink(tmp.name)
                except Exception as e:
                    st.error(f"Error previewing: {e}")
            
            # Download button
            try:
                csv_data = export_sheet_to_csv(workbook, 'Export')
                if csv_data:
                    st.download_button(
                        label="📥 Download Export Sheet as CSV",
                        data=csv_data,
                        file_name="export_data.csv",
                        mime="text/csv",
                        type="primary"
                    )
                else:
                    st.error("❌ Failed to export CSV")
            except Exception as e:
                st.error(f"❌ Export error: {e}")
        else:
            st.warning("⚠️ 'Export' sheet not found")
            st.write("Available sheets:")
            for sheet in workbook.sheetnames:
                st.write(f"• {sheet}")
    
    # Display current file info
    st.markdown("---")
    with st.expander("ℹ️ File Information"):
        st.write(f"**File Path:** {excel_file_path if not temp_file_path else 'Uploaded File'}")
        st.write(f"**Available Sheets:** {', '.join(workbook.sheetnames)}")
        st.write(f"**Target Cells:** D4, D5, D6, D7 in 'Prices' sheet")
        st.write(f"**Export Sheet:** 'Export' sheet → CSV download")
    
    # Clean up
    if workbook:
        workbook.close()
    
    # Clean up temp file
    if temp_file_path and os.path.exists(temp_file_path):
        try:
            os.unlink(temp_file_path)
        except:
            pass

if __name__ == "__main__":
    main()
