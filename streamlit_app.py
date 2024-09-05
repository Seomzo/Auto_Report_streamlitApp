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
                background-color: #f7f8fa;
                color: #000;
            }
            .rounded-square {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            .copy-btn {
                color: #fff;
                background-color: #1f77b4;
                border: none;
                padding: 8px 16px;
                border-radius: 8px;
                cursor: pointer;
                font-size: 14px;
                transition: background-color 0.3s ease;
                width: auto;
                display: inline-block;
            }
            .copy-btn:hover {
                background-color: #1a5a8a;
            }
        }

        /* Dark mode styles */
        @media (prefers-color-scheme: dark) {
            .stApp {
                background-color: #2c2c2c;
                color: #e0e0e0;
            }
            .rounded-square {
                background-color: #3a3a3a;
                border: 1px solid #555;
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
            }
            .copy-btn {
                color: #000;
                background-color: #90caf9;
                border: none;
                padding: 8px 16px;
                border-radius: 8px;
                cursor: pointer;
                font-size: 14px;
                transition: background-color 0.3s ease;
                width: auto;
                display: inline-block;
            }
            .copy-btn:hover {
                background-color: #42a5f5;
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
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        return sheet
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}. Please check the configuration and try again.")
        return None

def clean_column_data(column):
    """Clean column data by removing non-numeric characters and converting to float."""
    return column.replace('[\$,]', '', regex=True).astype(float)

def process_menu_sales_data(df, names_column):
    df[names_column] = df[names_column].str.strip().str.upper()
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    name_counts = df[names_column].value_counts() / 2
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum()
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum()
    return name_counts, labor_gross_sums, parts_gross_sums

def process_alacarte_data(df, names_column):
    df[names_column] = df[names_column].str.strip().str.upper()
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    name_counts = df[names_column].value_counts()
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum()
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum()
    return name_counts, labor_gross_sums, parts_gross_sums

def process_commodities_data(df, names_column="Primary Advisor Name"):
    st.write("Available columns in Commodities data:", df.columns.tolist())
    df.columns = df.columns.str.strip()
    if names_column not in df.columns:
        st.error(f"Column '{names_column}' not found in the uploaded Commodities Excel. Please check the column names.")
        return pd.Series(dtype='int'), pd.Series(dtype='float')
    df[names_column] = df[names_column].str.strip().str.upper()
    gross_column = 'Gross'
    if gross_column not in df.columns:
        st.error(f"Column '{gross_column}' not found in the uploaded Commodities Excel. Please check the column names.")
        return pd.Series(dtype='int'), pd.Series(dtype='float')
    df[gross_column] = clean_column_data(df[gross_column])
    name_counts = df[names_column].value_counts()
    parts_gross_sums = df.groupby(names_column)[gross_column].sum()
    return name_counts, parts_gross_sums

def process_recommendations_data(df, names_column="Name"):
    # Display available columns for debugging
    st.write("Available columns in Recommendations data:", df.columns.tolist())
    
    # Ensure no leading/trailing spaces in column names
    df.columns = df.columns.str.strip()
    
    # Check if the required column exists
    if names_column not in df.columns:
        st.error(f"Column '{names_column}' not found in the uploaded Recommendations Excel. Please check the column names.")
        return pd.Series(dtype='int'), pd.Series(dtype='int'), pd.Series(dtype='float'), pd.Series(dtype='float')  # Return empty Series if column not found

    # Normalize the advisor names
    df[names_column] = df[names_column].str.strip().str.upper()  # Normalize to uppercase and strip spaces
    
    # Ensure the necessary columns exist
    required_columns = ['Recommendations', 'Recommendations Sold', 'Recommendations $ amount', 'Recommendations Sold $ amount']
    for col in required_columns:
        if col not in df.columns:
            st.error(f"Column '{col}' not found in the uploaded Recommendations Excel. Please check the column names.")
            return pd.Series(dtype='int'), pd.Series(dtype='int'), pd.Series(dtype='float'), pd.Series(dtype='float')
    
    # Clean and process data
    rec_count = df.groupby(names_column)['Recommendations'].sum()
    rec_sold_count = df.groupby(names_column)['Recommendations Sold'].sum()
    rec_amount = clean_column_data(df.groupby(names_column)['Recommendations $ amount'].sum())
    rec_sold_amount = clean_column_data(df.groupby(names_column)['Recommendations Sold $ amount'].sum())
    
    return rec_count, rec_sold_count, rec_amount, rec_sold_amount

def update_google_sheet(sheet, name_counts, *args, date, start_row, handle_two_outputs=False, handle_four_outputs=False):
    headers = sheet.row_values(2)
    if date in headers:
        date_column_index = headers.index(date) + 1
    else:
        st.error(f"Date {date} not found in the sheet.")
        return
    
    sheet_advisor_names = [name.strip().upper() for name in sheet.col_values(1)]
    
    for advisor_name in name_counts.index:
        try:
            if advisor_name in sheet_advisor_names:
                row_index = sheet_advisor_names.index(advisor_name) + start_row
                
                sheet.update_cell(row_index, date_column_index, int(name_counts[advisor_name]))
                
                if handle_two_outputs:
                    parts_gross = float(args[0].get(advisor_name, 0))
                    sheet.update_cell(row_index + 1, date_column_index, parts_gross)
                    black_format = CellFormat(textFormat={"foregroundColor": {"red": 0, "green": 0, "blue": 0}})
                    format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index, date_column_index)}", black_format)
                    format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index + 1, date_column_index)}", black_format)
                elif handle_four_outputs:
                    rec_count = int(args[0].get(advisor_name, 0))
                    rec_sold_count = int(args[1].get(advisor_name, 0))
                    rec_amount = float(args[2].get(advisor_name, 0))
                    rec_sold_amount = float(args[3].get(advisor_name, 0))
                    
                    sheet.update_cell(row_index, date_column_index, rec_count)
                    sheet.update_cell(row_index + 1, date_column_index, rec_sold_count)
                    sheet.update_cell(row_index + 2, date_column_index, rec_amount)
                    sheet.update_cell(row_index + 3, date_column_index, rec_sold_amount)
                    
                    black_format = CellFormat(textFormat={"foregroundColor": {"red": 0, "green": 0, "blue": 0}})
                    format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index, date_column_index)}", black_format)
                    format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index + 1, date_column_index)}", black_format)
                    format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index + 2, date_column_index)}", black_format)
                    format_cell_range(sheet, f"{gspread.utils.rowcol_to_a1(row_index + 3, date_column_index)}", black_format)
                else:
            
                    labor_gross = float(args[0].get(advisor_name, 0))
                    parts_gross = float(args[1].get(advisor_name, 0))
                    
                    # Update Labor Gross
                    sheet.update_cell(row_index + 1, date_column_index, labor_gross)
                    
                    # Update Parts Gross
                    sheet.update_cell(row_index + 2, date_column_index, parts_gross)
                    
                    # Apply black text formatting to updated cells
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
    recommendations_file = st.file_uploader("Upload Recommendations Excel", type=["xlsx"])

    # Date input with default to today's date
    selected_date = st.date_input("Select the date:", datetime.now()).strftime('%d')

    # Connect to Google Sheet
    sheet = connect_to_google_sheet(sheet_name, worksheet_name)
    if sheet is None:
        st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
        return

    # Process Menu Sales data
    if menu_sales_file is not None:
        df_menu_sales = pd.read_excel(menu_sales_file)
        st.write("Menu Sales data preview:", df_menu_sales.head())
        
        menu_name_counts, menu_labor_gross_sums, menu_parts_gross_sums = process_menu_sales_data(df_menu_sales, "Advisor Name")
        st.write(f"Menu Name counts: {menu_name_counts.to_dict()}")
        st.write(f"Menu Labor Gross Sums: {menu_labor_gross_sums.to_dict()}")
        st.write(f"Menu Parts Gross Sums: {menu_parts_gross_sums.to_dict()}")
        
        if st.button("Update Menu Sales in Google Sheet"):
            update_google_sheet(sheet, menu_name_counts, menu_labor_gross_sums, menu_parts_gross_sums, date=selected_date, start_row=2)
            st.success("Menu Sales data updated successfully.")

    # Process A-La-Carte data
    if alacarte_file is not None:
        df_alacarte = pd.read_excel(alacarte_file)
        st.write("A-La-Carte data preview:", df_alacarte.head())
        
        alacarte_name_counts, alacarte_labor_gross_sums, alacarte_parts_gross_sums = process_alacarte_data(df_alacarte, "Advisor Name")
        st.write(f"A-La-Carte Name counts: {alacarte_name_counts.to_dict()}")
        st.write(f"A-La-Carte Labor Gross Sums: {alacarte_labor_gross_sums.to_dict()}")
        st.write(f"A-La-Carte Parts Gross Sums: {alacarte_parts_gross_sums.to_dict()}")
        
        if st.button("Update A-La-Carte in Google Sheet"):
            update_google_sheet(sheet, alacarte_name_counts, alacarte_labor_gross_sums, alacarte_parts_gross_sums, date=selected_date, start_row=5)
            st.success("A-La-Carte data updated successfully.")

    # Process Commodities data
    if commodities_file is not None:
        df_commodities = pd.read_excel(commodities_file)
        st.write("Commodities data preview:", df_commodities.head())
        
        commodities_name_counts, commodities_parts_gross_sums = process_commodities_data(df_commodities, "Primary Advisor Name")
        if commodities_name_counts.empty and commodities_parts_gross_sums.empty:
            return  # Return early if columns were not found

        st.write(f"Commodities Name counts: {commodities_name_counts.to_dict()}")
        st.write(f"Commodities Parts Gross Sums: {commodities_parts_gross_sums.to_dict()}")
        
        if st.button("Update Commodities in Google Sheet"):
            update_google_sheet(sheet, commodities_name_counts, commodities_parts_gross_sums, date=selected_date, start_row=8, handle_two_outputs=True)
            st.success("Commodities data updated successfully.")

    # Process Recommendations data
    if recommendations_file is not None:
        df_recommendations = pd.read_excel(recommendations_file)
        st.write("Recommendations data preview:", df_recommendations.head())
        
        rec_count, rec_sold_count, rec_amount, rec_sold_amount = process_recommendations_data(df_recommendations, "Name")
        st.write(f"Recommendations Count: {rec_count.to_dict()}")
        st.write(f"Recommendations Sold Count: {rec_sold_count.to_dict()}")
        st.write(f"Recommendations Amount: {rec_amount.to_dict()}")
        st.write(f"Recommendations Sold Amount: {rec_sold_amount.to_dict()}")
        
        if st.button("Update Recommendations in Google Sheet"):
            update_google_sheet(sheet, rec_count, rec_sold_count, rec_amount, rec_sold_amount, date=selected_date, start_row=10, handle_two_outputs=False)
            st.success("Recommendations data updated successfully.")

if __name__ == "__main__":
    main()
