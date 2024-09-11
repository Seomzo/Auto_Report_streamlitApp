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
    # Remove dollar signs, commas, and any spaces
    column = column.replace('[\$,]', '', regex=True).replace(',', '', regex=True).astype(float)
    return column


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
    #st.write("Available columns in Commodities data:", df.columns.tolist())
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
    #st.write("Available columns in Recommendations data:", df.columns.tolist())
    
    # Ensure no leading/trailing spaces in column names
    df.columns = df.columns.str.strip()
    
    # Filter out rows where the advisor name is "Total"
    df = df[df[names_column].str.strip().str.upper() != "TOTAL"]
    
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

def process_daily_data(df, names_column="Name"):
    # Ensure no leading/trailing spaces in column names
    df.columns = df.columns.str.strip()
    
    # Filter out rows where the advisor name is "Total"
    df = df[df[names_column].str.strip().str.upper() != "TOTAL"]
    
    # Filter in rows where the pay type is "ALL"
    df = df[df['Pay Type'].str.upper() == "ALL"]

    # Normalize the advisor names
    df[names_column] = df[names_column].str.strip().str.upper()

    # Check for required columns
    required_columns = ['Labor Gross', 'Parts Gross']
    for col in required_columns:
        if col not in df.columns:
            st.error(f"Column '{col}' not found in the uploaded Daily Data Excel. Please check the column names.")
            return pd.Series(dtype='float'), pd.Series(dtype='float')

    # Clean and process data
    df['Labor Gross'] = clean_column_data(df['Labor Gross'])
    df['Parts Gross'] = clean_column_data(df['Parts Gross'])

    # Group by advisor names and get the data
    labor_gross_sums = df.groupby(names_column)['Labor Gross'].sum()  # Sum to check for any duplicates
    parts_gross_sums = df.groupby(names_column)['Parts Gross'].sum()

    # # Debug: Print the processed values
    # st.write("Processed Labor Gross Sums:", labor_gross_sums)
    # st.write("Processed Parts Gross Sums:", parts_gross_sums)

    return labor_gross_sums, parts_gross_sums


def batch_update_google_sheet(sheet, updates, date, start_row):
    headers = sheet.row_values(2)  # Assuming the date is in row 2
    date = date.lstrip('0')
    
    # Check if the date exists in headers
    if date in headers:
        date_column_index = headers.index(date) + 1
    else:
        st.error(f"Date {date} not found in the sheet.")
        return
    
    sheet_advisor_names = [name.strip().upper() for name in sheet.col_values(1)]  # Assuming advisors are listed in the first column
    cells_to_update = []
    
    for update in updates:
        advisor_name = update['advisor']
        values = update['values']
        if advisor_name in sheet_advisor_names:
            row_index = sheet_advisor_names.index(advisor_name) + start_row
            
            # Convert each value to native Python types before adding to the batch update
            for i, value in enumerate(values):
                value = float(value) if isinstance(value, (int, float, np.integer, np.floating)) else value
                cells_to_update.append({
                    'range': gspread.utils.rowcol_to_a1(row_index + i, date_column_index),
                    'values': [[value]]
                })
    
    # Execute the batch update
    try:
        sheet.batch_update([{'range': cell['range'], 'values': cell['values']} for cell in cells_to_update])
        # Optionally format updated cells to have black text
        for cell in cells_to_update:
            format_cell_range(sheet, cell['range'], CellFormat(textFormat={"foregroundColor": {"red": 0, "green": 0, "blue": 0}}))
    except gspread.exceptions.APIError as e:
        st.error(f"Error updating cells: {e}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")


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
    daily_file = st.file_uploader("Upload Daily Data Excel", type=["xlsx"])

    # Date input with default to today's date
    selected_date = st.date_input("Select the date:", datetime.now()).strftime('%d')

    # Connect to Google Sheet
    sheet = connect_to_google_sheet(sheet_name, worksheet_name)
    if sheet is None:
        st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
        return

    updates = []  # List to store all updates for batch processing

    # Process each dataset and add updates
    if menu_sales_file is not None:
        df_menu_sales = pd.read_excel(menu_sales_file)
        menu_name_counts, menu_labor_gross_sums, menu_parts_gross_sums = process_menu_sales_data(df_menu_sales, "Advisor Name")
        for advisor, count in menu_name_counts.items():
            updates.append({'advisor': advisor, 'values': [int(count), float(menu_labor_gross_sums.get(advisor, 0)), float(menu_parts_gross_sums.get(advisor, 0))]})
    
    if alacarte_file is not None:
        df_alacarte = pd.read_excel(alacarte_file)
        alacarte_name_counts, alacarte_labor_gross_sums, alacarte_parts_gross_sums = process_alacarte_data(df_alacarte, "Advisor Name")
        for advisor, count in alacarte_name_counts.items():
            updates.append({'advisor': advisor, 'values': [int(count), float(alacarte_labor_gross_sums.get(advisor, 0)), float(alacarte_parts_gross_sums.get(advisor, 0))]})
    
    if commodities_file is not None:
        df_commodities = pd.read_excel(commodities_file)
        commodities_name_counts, commodities_parts_gross_sums = process_commodities_data(df_commodities, "Primary Advisor Name")
        for advisor, count in commodities_name_counts.items():
            updates.append({'advisor': advisor, 'values': [int(count), float(commodities_parts_gross_sums.get(advisor, 0))]})
    
    if recommendations_file is not None:
        df_recommendations = pd.read_excel(recommendations_file)
        rec_count, rec_sold_count, rec_amount, rec_sold_amount = process_recommendations_data(df_recommendations, "Name")
        for advisor, count in rec_count.items():
            updates.append({'advisor': advisor, 'values': [int(count), int(rec_sold_count.get(advisor, 0)), float(rec_amount.get(advisor, 0)), float(rec_sold_amount.get(advisor, 0))]})
    
    if daily_file is not None:
        df_daily = pd.read_excel(daily_file)
        daily_labor_gross, daily_parts_gross = process_daily_data(df_daily, "Name")
        for advisor, labor in daily_labor_gross.items():
            updates.append({'advisor': advisor, 'values': [float(labor), float(daily_parts_gross.get(advisor, 0))]})

    # Execute batch update if there are updates to apply
    if updates:
        batch_update_google_sheet(sheet, updates, selected_date, start_row=2)  # Adjust start_row as needed
        st.success("All data updated successfully.")

if __name__ == "__main__":
    main()

