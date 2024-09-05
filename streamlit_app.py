import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import CellFormat, format_cell_range
from datetime import datetime
import warnings
import json
import os

# Suppress UserWarning related to openpyxl's default style
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Set custom background color and styles using CSS for light and dark mode
def set_bg_color():
    st.markdown(
        """
        <style>
        /* Light mode and dark mode styles */
        </style>
        """,
        unsafe_allow_html=True
    )

# Function to connect to Google Sheets
def connect_to_google_sheet(sheet_name, worksheet_name):
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_dict(
            st.secrets["GOOGLE_CREDENTIALS"], 
            scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        )
        client = gspread.authorize(creds)
        sheet = client.open(sheet_name).worksheet(worksheet_name)  # Select worksheet by name
        return sheet
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}. Please check the configuration and try again.")
        return None

def clean_column_data(column):
    """Clean column data by removing non-numeric characters and converting to float."""
    return column.replace('[\$,]', '', regex=True).astype(float)

def process_menu_sales_data(df, names_column):
    df[names_column] = df[names_column].str.strip().str.upper()  # Normalize to uppercase and strip spaces
    
    # Clean data
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    
    # Calculate counts and sums
    name_counts = df[names_column].value_counts() / 2  # Adjust counts by dividing by 2 for duplicates
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum()
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum()
    
    return name_counts, labor_gross_sums, parts_gross_sums

def update_google_sheet(sheet, name_counts, labor_gross_sums, parts_gross_sums, date):
    headers = sheet.row_values(2)  # Get headers from the sheet assuming row 3 has the dates
    if date in headers:
        date_column_index = headers.index(date) + 1
    else:
        st.error(f"Date {date} not found in the sheet.")
        return
    
    sheet_advisor_names = [name.strip().upper() for name in sheet.col_values(1)]  # Adjust if advisors are in a different column
    
    for advisor_name in name_counts.index:
        try:
            if advisor_name in sheet_advisor_names:
                row_index = sheet_advisor_names.index(advisor_name) + 2  # Adjust row index based on actual sheet layout
                # Update Menu Sales Count
                sheet.update_cell(row_index, date_column_index, int(name_counts[advisor_name]))
                
                # Update Labor Gross
                sheet.update_cell(row_index + 1, date_column_index, float(labor_gross_sums[advisor_name]))
                
                # Update Parts Gross
                sheet.update_cell(row_index + 2, date_column_index, float(parts_gross_sums[advisor_name]))
                
                # Apply black text formatting
                black_format = CellFormat(textFormat={"foregroundColor": {"red": 0, "green": 0, "blue": 0}})
                format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index, date_column_index)}", black_format)
                format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index + 1, date_column_index)}", black_format)
                format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index + 2, date_column_index)}", black_format)
            else:
                st.warning(f"{advisor_name} not found in the Google Sheet.")
        except gspread.exceptions.APIError as e:
            st.error(f"Error updating cell for {advisor_name}: {e}")
        except Exception as e:
            st.error(f"An error occurred: {e}")

def main():
    set_bg_color()

    st.title("Google Sheet Updater for Advisors")

    st.markdown(
        """
        <div class='rounded-square'>
            <p><b>Instructions:</b></p>
            <ul>
                <li>Please share the Google Sheet with the following email:</li>
                <p style='margin-left: 20px; display: flex; align-items: center;'>
                    <code style='flex: 1; white-space: nowrap;'>auto-report@auto-pop-report.iam.gserviceaccount.com</code> 
                    <button id="copy-button" class="copy-btn">Copy Email</button>
                </p>
                <li>Make sure to give the email <b>Editor</b> permissions.</li>
                <li>Ensure there are no other restrictions or permissions on the sheet.</li>
            </ul>
        </div>

        <script>
        const copyButton = document.getElementById('copy-button');
        copyButton.addEventListener('click', function() {
            navigator.clipboard.writeText('auto-report@auto-pop-report.iam.gserviceaccount.com');
            copyButton.textContent = 'Email Copied';
            setTimeout(() => { copyButton.textContent = 'Copy Email'; }, 2000);
        });
        </script>
        """,
        unsafe_allow_html=True
    )

    sheet_name = st.text_input("Enter the Google Sheet name:", "August Advisor Performance-OMAR")
    worksheet_name = st.text_input("Enter the Worksheet (tab) name:", "Menu Sales")

    # File uploads for different sections
    menu_sales_file = st.file_uploader("Upload Menu Sales Excel", type=["xlsx"])
    
    # Date input with default to today's date
    selected_date = st.date_input("Select the date:", datetime.now()).strftime('%d')
    
    if menu_sales_file is not None and sheet_name and worksheet_name:
        df = pd.read_excel(menu_sales_file)
        st.write("Menu Sales data preview:", df.head())
        
        sheet = connect_to_google_sheet(sheet_name, worksheet_name)
        if sheet is None:
            st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
            return
        
        name_counts, labor_gross_sums, parts_gross_sums = process_menu_sales_data(df, "Advisor Name")
        st.write(f"Name counts: {name_counts.to_dict()}")
        st.write(f"Labor Gross Sums: {labor_gross_sums.to_dict()}")
        st.write(f"Parts Gross Sums: {parts_gross_sums.to_dict()}")
        
        if st.button("Update Google Sheet"):
            update_google_sheet(sheet, name_counts, labor_gross_sums, parts_gross_sums, selected_date)
            st.success("Google Sheet updated successfully.")

if __name__ == "__main__":
    main()
