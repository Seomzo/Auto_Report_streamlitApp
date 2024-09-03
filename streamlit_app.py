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
        /* Light mode styles */
        @media (prefers-color-scheme: light) {
            .stApp {
                background-color: #f7f8fa;  /* Light background color */
                color: #000;  /* Light mode text color */
            }
            .rounded-square {
                background-color: white;
                border: 1px solid #ddd;  /* Light mode border color */
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);  /* Light shadow for depth */
            }
            .copy-btn {
                color: #fff;
                background-color: #1f77b4;  /* Modern button color */
                border: none;
                padding: 8px 16px;
                border-radius: 8px;
                cursor: pointer;
                font-size: 14px;
                transition: background-color 0.3s ease;
                width: auto;  /* Adjust to fit the content */
                display: inline-block;
            }
            .copy-btn:hover {
                background-color: #1a5a8a;  /* Darker shade on hover */
            }
        }

        /* Dark mode styles */
        @media (prefers-color-scheme: dark) {
            .stApp {
                background-color: #2c2c2c;  /* Dark background color */
                color: #e0e0e0;  /* Dark mode text color */
            }
            .rounded-square {
                background-color: #3a3a3a;
                border: 1px solid #555;  /* Dark mode border color */
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);  /* Dark shadow for depth */
            }
            .copy-btn {
                color: #000;
                background-color: #90caf9;  /* Modern button color for dark mode */
                border: none;
                padding: 8px 16px;
                border-radius: 8px;
                cursor: pointer;
                font-size: 14px;
                transition: background-color 0.3s ease;
                width: auto;  /* Adjust to fit the content */
                display: inline-block;
            }
            .copy-btn:hover {
                background-color: #42a5f5;  /* Darker shade on hover */
            }
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# Function to connect to Google Sheets
def connect_to_google_sheet(sheet_name, worksheet_name):
    try:
        # Load JSON key file from the same directory
        json_keyfile_path = os.path.join(os.path.dirname(__file__), 'auto-pop-report-b3525fd81b96.json')
        if not os.path.exists(json_keyfile_path):
            st.error(f"JSON key file not found at {json_keyfile_path}. Please ensure the file is in the correct location.")
            return None
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        
        # Load credentials from JSON file
        with open(json_keyfile_path) as json_file:
            json_keyfile = json.load(json_file)
        
        creds = ServiceAccountCredentials.from_json_keyfile_dict(json_keyfile, scope)
        client = gspread.authorize(creds)
        sheet = client.open(sheet_name).worksheet(worksheet_name)  # Select worksheet by name
        return sheet

    except FileNotFoundError as e:
        st.error(f"File not found error: {e}. Please make sure the JSON key file exists and is accessible.")
        return None
    except json.JSONDecodeError as e:
        st.error(f"JSON decode error: {e}. Please check the JSON key file for formatting issues.")
        return None
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"The Google Sheet named '{sheet_name}' was not found. Please verify the sheet name.")
        return None
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"The worksheet named '{worksheet_name}' was not found in the Google Sheet '{sheet_name}'. Please verify the tab name.")
        return None
    except gspread.exceptions.APIError as e:
        st.error(f"API error occurred: {e}. This may be due to insufficient permissions or API limits.")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}. Please check the configuration and try again.")
        return None

def clean_column_data(column):
    """Clean column data by removing non-numeric characters and converting to float."""
    return column.replace('[\$,]', '', regex=True).astype(float)

# Excel
def count_and_sum_gross_in_excel(excel_file_path, names_column):
    df = pd.read_excel(excel_file_path)
    df[names_column] = df[names_column].str.strip().str.upper()  # Normalize to uppercase and strip spaces
    
    # Clean data
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    
    # Calculate counts and sums
    name_counts = df[names_column].value_counts()
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum()
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum()
    
    return name_counts, labor_gross_sums, parts_gross_sums

# Function to update Google Sheet
def update_google_sheet(sheet, name_counts, labor_gross_sums, parts_gross_sums, date):
    headers = sheet.row_values(2)  # Get headers from the sheet
    if date in headers:
        date_column_index = headers.index(date) + 1
    else:
        st.error(f"Date {date} not found in the sheet.")
        return
    
    sheet_advisor_names = [name.strip().upper() for name in sheet.col_values(1)]
    
    for advisor_name in name_counts.index:
        try:
            if advisor_name in sheet_advisor_names:
                row_index = sheet_advisor_names.index(advisor_name) + 2
                # Update A-La-Cart Count
                sheet.update_cell(row_index, date_column_index, int(name_counts[advisor_name]))  # Convert to int
                
                sheet.update_cell(row_index + 1, date_column_index, float(labor_gross_sums[advisor_name]))  # Convert to float
                
                sheet.update_cell(row_index + 2, date_column_index, float(parts_gross_sums[advisor_name]))  # Convert to float
                
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
    worksheet_name = st.text_input("Enter the Worksheet (tab) name:", "A-La-Carte")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    
    selected_date = st.date_input("Select the date:", datetime.now()).strftime('%d')
    
    if uploaded_file is not None and sheet_name and worksheet_name:
        df = pd.read_excel(uploaded_file)
        st.write("Excel data preview:", df.head())
        
        sheet = connect_to_google_sheet(sheet_name, worksheet_name)
        if sheet is None:
            st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
            return
        
        name_counts, labor_gross_sums, parts_gross_sums = count_and_sum_gross_in_excel(uploaded_file, "Advisor Name")
        st.write(f"Name counts: {name_counts.to_dict()}")
        st.write(f"Labor Gross Sums: {labor_gross_sums.to_dict()}")
        st.write(f"Parts Gross Sums: {parts_gross_sums.to_dict()}")
        
        
        if st.button("Update Google Sheet"):
            update_google_sheet(sheet, name_counts, labor_gross_sums, parts_gross_sums, selected_date)
            st.success("Google Sheet updated successfully.")

if __name__ == "__main__":
    main()
