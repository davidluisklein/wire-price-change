# Excel Price Editor

A Streamlit app to edit specific price cells (D4-D7) in an Excel file and export data as CSV.

## Features
- Edit cells D4, D5, D6, D7 in the 'Prices' sheet
- Download 'Export' sheet as CSV
- Web-based interface
- File upload support

## Usage
1. The app loads with a bundled Excel file
2. Edit the price values in the form
3. Click "Update Prices" to save changes
4. Download the Export sheet as CSV

## Local Development
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
