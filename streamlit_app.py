import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import CellFormat, format_cell_range
from datetime import datetime
import warnings

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

def process_alacarte_data(df, names_column):
    df[names_column] = df[names_column].str.strip().str.upper()  # Normalize to uppercase and strip spaces
    
    # Clean data
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    
    # Calculate counts and sums
    name_counts = df[names_column].value_counts()  # Regular name count
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum()
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum()
    
    return name_counts, labor_gross_sums, parts_gross_sums

def process_commodities_data(df, names_column="Primary Advisor Name"):
    # Display available columns for debugging
    st.write("Available columns in Commodities data:", df.columns.tolist())
    
    # Ensure no leading/trailing spaces in column names
    df.columns = df.columns.str.strip()
    
    # Check if the required column exists
    if names_column not in df.columns:
        st.error(f"Column '{names_column}' not found in the uploaded Commodities Excel. Please check the column names.")
        return pd.Series(dtype='int'), pd.Series(dtype='float')  # Return empty Series if column not found

    # Normalize the advisor names
    df[names_column] = df[names_column].str.strip().str.upper()  # Normalize to uppercase and strip spaces
    
    # Check for the correct 'Gross' column
    gross_column = 'Gross'
    if gross_column not in df.columns:
        st.error(f"Column '{gross_column}' not found in the uploaded Commodities Excel. Please check the column names.")
        return pd.Series(dtype='int'), pd.Series(dtype='float')  # Return empty Series if column not found

    # Clean data in the Gross column
    df[gross_column] = clean_column_data(df[gross_column])
    
    # Calculate counts and sums
    name_counts = df[names_column].value_counts()  # Regular name count
    parts_gross_sums = df.groupby(names_column)[gross_column].sum()
    
    return name_counts, parts_gross_sums

def update_google_sheet(sheet, name_counts, labor_gross_sums, parts_gross_sums, date, start_row):
    headers = sheet.row_values(2)  # Get headers from the sheet, setting date row index to 2
    if date in headers:
        date_column_index = headers.index(date) + 1
    else:
        st.error(f"Date {date} not found in the sheet.")
        return
    
    sheet_advisor_names = [name.strip().upper() for name in sheet.col_values(1)]  # Adjust if advisors are in a different column
    
    for advisor_name in name_counts.index:
        try:
            if advisor_name in sheet_advisor_names:
                row_index = sheet_advisor_names.index(advisor_name) + start_row  # Adjust row index based on actual sheet layout
                
                # Update Count
                sheet.update_cell(row_index, date_column_index, int(name_counts[advisor_name]))
                
                # Check if labor_gross_sums and parts_gross_sums are provided, if not use default values
                labor_gross = float(labor_gross_sums.get(advisor_name, 0))
                parts_gross = float(parts_gross_sums.get(advisor_name, 0))
                
                # Update Labor Gross
                sheet.update_cell(row_index + 1, date_column_index, labor_gross)
                
                # Update Parts Gross
                sheet.update_cell(row_index + 2, date_column_index, parts_gross)
                
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
    worksheet_name = st.text_input("Enter the Worksheet (tab) name:", "Input")

    # File uploads for different sections
    menu_sales_file = st.file_uploader("Upload Menu Sales Excel", type=["xlsx"])
    alacarte_file = st.file_uploader("Upload A-La-Carte Excel", type=["xlsx"])
    commodities_file = st.file_uploader("Upload Commodities Excel", type=["xlsx"])

    # Date input with default to today's date
    selected_date = st.date_input("Select the date:", datetime.now()).strftime('%d')

    # Process Menu Sales data
    if menu_sales_file is not None and sheet_name and worksheet_name:
        df_menu_sales = pd.read_excel(menu_sales_file)
        st.write("Menu Sales data preview:", df_menu_sales.head())
        
        sheet = connect_to_google_sheet(sheet_name, worksheet_name)
        if sheet is None:
            st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
            return
        
        menu_name_counts, menu_labor_gross_sums, menu_parts_gross_sums = process_menu_sales_data(df_menu_sales, "Advisor Name")
        st.write(f"Menu Name counts: {menu_name_counts.to_dict()}")
        st.write(f"Menu Labor Gross Sums: {menu_labor_gross_sums.to_dict()}")
        st.write(f"Menu Parts Gross Sums: {menu_parts_gross_sums.to_dict()}")
        
        if st.button("Update Menu Sales in Google Sheet"):
            update_google_sheet(sheet, menu_name_counts, menu_labor_gross_sums, menu_parts_gross_sums, selected_date, start_row=6)  # Adjust start_row as per your sheet layout
            st.success("Menu Sales data updated successfully.")
        
    # Process A-La-Carte data
    if alacarte_file is not None and sheet_name and worksheet_name:
        df_alacarte = pd.read_excel(alacarte_file)
        st.write("A-La-Carte data preview:", df_alacarte.head())
        
        sheet = connect_to_google_sheet(sheet_name, worksheet_name)
        if sheet is None:
            st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
            return
        
        alacarte_name_counts, alacarte_labor_gross_sums, alacarte_parts_gross_sums = process_alacarte_data(df_alacarte, "Advisor Name")
        st.write(f"A-La-Carte Name counts: {alacarte_name_counts.to_dict()}")
        st.write(f"A-La-Carte Labor Gross Sums: {alacarte_labor_gross_sums.to_dict()}")
        st.write(f"A-La-Carte Parts Gross Sums: {alacarte_parts_gross_sums.to_dict()}")
        
        if st.button("Update A-La-Carte in Google Sheet"):
            update_google_sheet(sheet, alacarte_name_counts, alacarte_labor_gross_sums, alacarte_parts_gross_sums, selected_date, start_row=9)  # Adjust start_row as per your sheet layout
            st.success("A-La-Carte data updated successfully.")

    # Process Commodities data
    if commodities_file is not None and sheet_name and worksheet_name:
        df_commodities = pd.read_excel(commodities_file)
        st.write("Commodities data preview:", df_commodities.head())
        
        sheet = connect_to_google_sheet(sheet_name, worksheet_name)
        if sheet is None:
            st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
            return
        
        # Use the correct column name for Commodities data
        commodities_name_counts, commodities_parts_gross_sums = process_commodities_data(df_commodities, "Primary Advisor Name")
        if commodities_name_counts.empty and commodities_parts_gross_sums.empty:
            return  # Return early if columns were not found

        st.write(f"Commodities Name counts: {commodities_name_counts.to_dict()}")
        st.write(f"Commodities Parts Gross Sums: {commodities_parts_gross_sums.to_dict()}")
        
        if st.button("Update Commodities in Google Sheet"):
            update_google_sheet(sheet, commodities_name_counts, pd.Series(dtype='float'), commodities_parts_gross_sums, selected_date, start_row=12)
            st.success("Commodities data updated successfully.")

if __name__ == "__main__":
    main()

