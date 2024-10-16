import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.cell import Cell
from datetime import datetime
import warnings
import time
import numpy as np

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
    return column.replace('[\$,]', '', regex=True).replace(',', '', regex=True).astype(float)

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

def process_commodity_file(df, names_column='Primary Advisor Name', gross_column='Gross'):
    df[names_column] = df[names_column].astype(str).str.strip().str.upper()
    df[gross_column] = clean_column_data(df[gross_column])

    # Count the number of sales per advisor
    name_counts = df[names_column].value_counts()

    # Sum gross amounts per advisor
    parts_gross_sums = df.groupby(names_column)[gross_column].sum()

    # Convert to dictionaries
    name_counts = name_counts.to_dict()
    parts_gross_sums = parts_gross_sums.to_dict()

    return name_counts, parts_gross_sums

def process_tires_data(df):
    """
    Process Tires data by summing Actual Quantity and Gross per advisor.
    Handles Original and GM formats based on column headers.
    """
    # st.write("Preview of Tires Data:")
    # st.write(df.head())

    # st.write("Columns in Tires Excel:")
    # st.write(df.columns.tolist())

    # Initialize column variables
    names_column = None
    quantity_column = None
    gross_column = None

    # Identify columns based on keywords
    for col in df.columns:
        col_lower = col.lower()
        if 'advisor' in col_lower and 'name' in col_lower:
            names_column = col
        elif 'part count' in col_lower or 'actual quantity' in col_lower:
            quantity_column = col
        elif 'opcode parts gross' in col_lower or 'gross' in col_lower:
            gross_column = col

    # Check if all required columns are found
    if names_column and quantity_column and gross_column:
        if 'advisor name group' in names_column.lower():
            st.write("Detected GM Tires Format.")
        else:
            st.write("Detected Original Tires Format.")
    else:
        raise ValueError("Tires Excel does not match any known format.")

    # Normalize advisor names
    df[names_column] = df[names_column].astype(str).str.strip().str.upper()

    # Clean the 'Quantity' and 'Gross' columns
    try:
        df[quantity_column] = clean_column_data(df[quantity_column])
        df[gross_column] = clean_column_data(df[gross_column])
    except Exception as e:
        raise ValueError(f"Error cleaning columns: {e}")

    # Sum Actual Quantity and Gross per advisor
    actual_quantity_sums = df.groupby(names_column)[quantity_column].sum().to_dict()
    gross_sums = df.groupby(names_column)[gross_column].sum().to_dict()

    # Convert to native types
    actual_quantity_sums = {k: float(v) for k, v in actual_quantity_sums.items()}
    gross_sums = {k: float(v) for k, v in gross_sums.items()}

    # st.write("Tires Actual Quantity Sums:", actual_quantity_sums)
    # st.write("Tires Gross Sums:", gross_sums)

    return actual_quantity_sums, gross_sums

def process_tires_gm_format(file):
    """
    Process the GM Format Tires Excel file.
    Skips the first two rows and uses the third row as headers.
    """
    try:
        # Read the Excel file by skipping the first two rows
        df = pd.read_excel(file, skiprows=2, header=0)
        actual_quantity_sums, gross_sums = process_tires_data(df)
        return actual_quantity_sums, gross_sums
    except Exception as e:
        raise ValueError(f"Error processing GM Format Tires Excel file: {e}")

def preprocess_gm_excel(file):
    """
    Preprocess the GM Excel file to ensure correct structure.
    Removes irrelevant rows and handles merged cells.
    """
    # Read the Excel file without specifying header
    df = pd.read_excel(file, header=None)

    # Drop the first two rows which contain summary data
    df = df.drop([0, 1]).reset_index(drop=True)

    # Assign the third row as header
    new_header = df.iloc[0]
    df = df[1:]
    df.columns = new_header

    # Drop any completely empty columns
    df = df.dropna(axis=1, how='all')

    return df

def process_alignment_files(df_menus, df_alacarte, names_column='Advisor Name'):
    """
    Process Alignment data by combining menus and a-la-carte data.
    """
    # Concatenate both DataFrames
    combined_df = pd.concat([df_menus, df_alacarte], ignore_index=True)

    # Normalize advisor names
    combined_df[names_column] = combined_df[names_column].astype(str).str.strip().str.upper()

    # Clean the 'Opcode Labor Gross' and 'Opcode Parts Gross' columns
    try:
        combined_df['Opcode Labor Gross'] = clean_column_data(combined_df['Opcode Labor Gross'])
        combined_df['Opcode Parts Gross'] = clean_column_data(combined_df['Opcode Parts Gross'])
    except Exception as e:
        raise ValueError(f"Error cleaning columns in Alignment data: {e}")

    # Count the number of sales per advisor
    name_counts = combined_df[names_column].value_counts() / 2  # Assuming you need to divide by 2 as in your original function

    # Sum gross amounts per advisor
    parts_gross_sums = combined_df.groupby(names_column)['Opcode Parts Gross'].sum()
    labor_gross_sums = combined_df.groupby(names_column)['Opcode Labor Gross'].sum()

    # Convert to dictionaries
    name_counts = name_counts.to_dict()
    parts_gross_sums = parts_gross_sums.to_dict()
    labor_gross_sums = labor_gross_sums.to_dict()

    return name_counts, parts_gross_sums, labor_gross_sums

commodities_row_mapping = {
    'Air Filters': 11,
    'Cabin Filters': 12,
    'Batteries': 13,
    'Tires': 14,
    'Brakes': 15,
    'Alignments': 16,
    'Wipers': 17,
    'Belts': 18,
    'Fluids': 19,
    'Factory Chemicals': 20,
    'Labor Gross': 21, 
    'Parts Gross': 22 
}

def process_recommendations_data(df, names_column="Name"):
    """
    Process Recommendations data by summing recommendations and amounts.
    """
    # Ensure no leading/trailing spaces in column names
    df.columns = df.columns.str.strip()

    # Filter out rows where the advisor name is "Total"
    df = df[df[names_column].str.strip().str.upper() != "TOTAL"]

    # Check if the required column exists
    if names_column not in df.columns:
        raise ValueError(f"Column '{names_column}' not found in the uploaded Recommendations Excel. Please check the column names.")

    # Normalize the advisor names
    df[names_column] = df[names_column].str.strip().str.upper()  # Normalize to uppercase and strip spaces

    # Ensure the necessary columns exist
    required_columns = ['Recommendations', 'Recommendations Sold', 'Recommendations $ amount', 'Recommendations Sold $ amount']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in the uploaded Recommendations Excel. Please check the column names.")

    # Clean and process data
    rec_count = df.groupby(names_column)['Recommendations'].sum()
    rec_sold_count = df.groupby(names_column)['Recommendations Sold'].sum()
    rec_amount = clean_column_data(df.groupby(names_column)['Recommendations $ amount'].sum())
    rec_sold_amount = clean_column_data(df.groupby(names_column)['Recommendations Sold $ amount'].sum())

    return rec_count, rec_sold_count, rec_amount, rec_sold_amount

def process_daily_data(df, names_column="Name"):
    """
    Process Daily Data by summing Labor and Parts Gross per advisor.
    """
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
            raise ValueError(f"Column '{col}' not found in the uploaded Daily Data Excel. Please check the column names.")

    # Clean and process data
    df['Labor Gross'] = clean_column_data(df['Labor Gross'])
    df['Parts Gross'] = clean_column_data(df['Parts Gross'])

    # Group by advisor names and get the data
    labor_gross_sums = df.groupby(names_column)['Labor Gross'].sum()  # Sum to check for any duplicates
    parts_gross_sums = df.groupby(names_column)['Parts Gross'].sum()

    return labor_gross_sums, parts_gross_sums

def convert_to_native_type(value):
    """
    Convert value to its native Python type.
    """
    if isinstance(value, pd.Series):
        value = value.sum()
    if pd.isna(value):
        return 0
    elif isinstance(value, (np.integer, np.int64, np.int32, int)):
        return int(value)
    elif isinstance(value, (np.floating, np.float64, np.float32, float)):
        return float(value)
    elif isinstance(value, (np.bool_, bool)):
        return bool(value)
    elif isinstance(value, (np.str_, str)):
        return str(value)
    else:
        return str(value)

def update_google_sheet(sheet, data_series1, *args, date_col_index, start_row_offset, advisor_mapping):
    """
    Update Google Sheet with provided data series.
    """
    cells_to_update = []

    for advisor_name, start_row in advisor_mapping.items():
        row_index = start_row + start_row_offset

        # Get value from data_series1
        value1 = data_series1.get(advisor_name, 0)
        value1 = convert_to_native_type(value1)

        # Create cell for value1
        cell = Cell(row=row_index, col=date_col_index, value=value1)
        cells_to_update.append(cell)

        # Handle additional data series
        for i, data_series in enumerate(args):
            value = data_series.get(advisor_name, 0)
            value = convert_to_native_type(value)
            cell = Cell(row=row_index + i + 1, col=date_col_index, value=value)
            cells_to_update.append(cell)

    if cells_to_update:
        sheet.update_cells(cells_to_update)

def update_commodities_in_sheet(sheet, date_col_index, commodities_data, commodities_list, advisor_mapping, data_row_offsets):
    """
    Update Commodities data in Google Sheet.
    """
    cells_to_update = []

    # Define the row offsets for commodities within each advisor's block
    commodity_row_offsets = {
        'Air Filters': 8,
        'Cabin Filters': 9,
        'Batteries': 10,
        'Tires': 11,
        'Brakes': 12,
        'Alignments': 13,       
        'Wipers': 14,
        'Belts': 15,
        'Fluids': 16,
        'Factory Chemicals': 17,
    }

    # Offsets for Labor Gross and Parts Gross
    labor_gross_offset = data_row_offsets['Labor Gross']
    parts_gross_offset = data_row_offsets['Parts Gross'] 

    # Initialize dictionaries to accumulate gross sums per advisor
    total_parts_gross_per_advisor = {advisor: 0 for advisor in advisor_mapping.keys()}
    total_labor_gross_per_advisor = {advisor: 0 for advisor in advisor_mapping.keys()}

    for commodity in commodities_list:
        data = commodities_data.get(commodity, {})
        
        if commodity == 'Tires':
            actual_quantity_sums = data.get('actual_quantity_sums', {})
            gross_sums = data.get('gross_sums', {})
        else:
            name_counts = data.get('name_counts', {})
            parts_gross_sums = data.get('parts_gross_sums', {})
            labor_gross_sums = data.get('labor_gross_sums', {}) if commodity == 'Alignments' else {}

        for advisor_name, start_row in advisor_mapping.items():
            base_row = start_row + commodity_row_offsets[commodity] - 1  # Adjust for one-based indexing

            if commodity == 'Tires':
                # Update Actual Quantity
                actual_quantity = convert_to_native_type(actual_quantity_sums.get(advisor_name, 0))
                cell_actual_quantity = Cell(row=base_row, col=date_col_index, value=actual_quantity)
                cells_to_update.append(cell_actual_quantity)

                # Update Gross (add to Parts Gross)
                gross = convert_to_native_type(gross_sums.get(advisor_name, 0))
                total_parts_gross_per_advisor[advisor_name] += gross
            else:
                # Name Counts
                count_value = convert_to_native_type(name_counts.get(advisor_name, 0))

                # Update Name Count under the commodity
                cell_count = Cell(row=base_row, col=date_col_index, value=count_value)
                cells_to_update.append(cell_count)

                # Accumulate Parts Gross
                parts_gross_value = convert_to_native_type(parts_gross_sums.get(advisor_name, 0))
                total_parts_gross_per_advisor[advisor_name] += parts_gross_value

                # Accumulate Labor Gross (for Alignments)
                if commodity == 'Alignments':
                    labor_gross_value = convert_to_native_type(labor_gross_sums.get(advisor_name, 0))
                    total_labor_gross_per_advisor[advisor_name] += labor_gross_value

    # After processing all commodities, update the Labor Gross and Parts Gross rows
    for advisor_name, start_row in advisor_mapping.items():
        # Update Labor Gross
        labor_gross = total_labor_gross_per_advisor.get(advisor_name, 0)
        labor_gross_row = start_row + labor_gross_offset - 1
        cell_labor_gross = Cell(row=labor_gross_row, col=date_col_index, value=labor_gross)
        cells_to_update.append(cell_labor_gross)

        # Update Parts Gross
        parts_gross = total_parts_gross_per_advisor.get(advisor_name, 0)
        parts_gross_row = start_row + parts_gross_offset - 1
        cell_parts_gross = Cell(row=parts_gross_row, col=date_col_index, value=parts_gross)
        cells_to_update.append(cell_parts_gross)

    if cells_to_update:
        sheet.update_cells(cells_to_update)

def main():
    set_bg_color()
    delay_seconds = 0.01 
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

    # Pre-fill with default sheet name if desired
    sheet_name = st.text_input("Enter the Google Sheet name:", "BLANK Advisor Performance Omar")
    worksheet_name = st.text_input("Enter the Worksheet (tab) name:", "Input")

    # File uploads for different sections
    menu_sales_file = st.file_uploader("Upload Menu Sales Excel", type=["xlsx"])
    alacarte_file = st.file_uploader("Upload A-La-Carte Excel", type=["xlsx"])
    recommendations_file = st.file_uploader("Upload Recommendations Excel", type=["xlsx"])
    daily_file = st.file_uploader("Upload Daily Data Excel", type=["xlsx"])

    # Upload Commodities Files
    st.subheader("Upload Commodities Files")

    commodities_list = [
        'Air Filters', 'Cabin Filters', 'Batteries', 'Tires', 'Brakes',
        'Wipers', 'Belts', 'Fluids', 'Factory Chemicals'
    ]
    
    commodities_files = {}
    for commodity in commodities_list:
        key = f"commodity_{commodity}"
        commodities_files[commodity] = st.file_uploader(f"Upload {commodity} Excel", type=["xlsx"], key=key)

    # Alignment Files Upload
    alignment_menus_file = st.file_uploader("Upload Alignment Menus Excel", type=["xlsx"], key="alignment_menus")
    alignment_alacarte_file = st.file_uploader("Upload Alignment A-La-Carte Excel", type=["xlsx"], key="alignment_alacarte")

    # Date input
    selected_date = st.date_input("Select the date:", datetime.now()).strftime('%d').lstrip('0')

    # Connect to Google Sheet
    sheet = connect_to_google_sheet(sheet_name, worksheet_name)
    if sheet is None:
        st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
        return

    # Get date column index
    date_row = sheet.row_values(2)[2:]  # Dates start from column 3 (after columns A and B)
    date = selected_date
    if date in date_row:
        date_col_index = date_row.index(date) + 3  # +3 because columns start from 1 and we skipped two columns
    else:
        st.error(f"Date {date} not found in the sheet.")
        return

    # Get advisor names and their starting rows
    col_a_values = sheet.col_values(1)[3:]  # Get values from row 4 onwards
    advisor_names = []
    advisor_start_rows = []
    row = 4  # Starting from row 4
    idx = 0
    while idx < len(col_a_values):
        advisor_name = col_a_values[idx]
        if not advisor_name:
            break  # No more advisors
        advisor_name = advisor_name.strip().upper()
        advisor_names.append(advisor_name)
        advisor_start_rows.append(row + idx)
        idx += 26  # Assuming each advisor block has 26 rows (including empty row)
    advisor_mapping = dict(zip(advisor_names, advisor_start_rows))

    # Define the row offsets for different data types within each advisor's block
    data_row_offsets = {
        'Menu Sales': 2,
        'Menu Sales Labor Gross': 3,
        'Menu Sales Parts Gross': 4,
        'A-la-carte Count': 5,
        'A-la-carte Labor Gross': 6,
        'A-la-carte Parts Gross': 7,
        # Commodities are from row 8 to 17
        'Labor Gross': 18,
        'Parts Gross': 19,
        'Rec Count': 20,
        'Rec Sold Count': 21,
        'Rec Amount': 22,
        'Rec Sold Amount': 23,
        'Daily Labor Gross': 24,
        'Daily Parts Gross': 25,
    }

    # Creating horizontal layout for buttons
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        if menu_sales_file is not None:
            if st.button("Update Menu Sales in Google Sheet"):
                try:
                    df_menu_sales = pd.read_excel(menu_sales_file)
                    menu_name_counts, menu_labor_gross_sums, menu_parts_gross_sums = process_menu_sales_data(df_menu_sales, "Advisor Name")
                    update_google_sheet(
                        sheet,
                        menu_name_counts,
                        menu_labor_gross_sums,
                        menu_parts_gross_sums,
                        date_col_index=date_col_index,
                        start_row_offset=data_row_offsets['Menu Sales'] - 1,
                        advisor_mapping=advisor_mapping
                    )
                    st.success("Menu Sales data updated successfully.")
                except Exception as e:
                    st.error(f"Error updating Menu Sales data: {e}")
                time.sleep(delay_seconds)

    with col2:
        if alacarte_file is not None:
            if st.button("Update A-La-Carte in Google Sheet"):
                try:
                    df_alacarte = pd.read_excel(alacarte_file)
                    alacarte_name_counts, alacarte_labor_gross_sums, alacarte_parts_gross_sums = process_alacarte_data(df_alacarte, "Advisor Name")
                    update_google_sheet(
                        sheet,
                        alacarte_name_counts,
                        alacarte_labor_gross_sums,
                        alacarte_parts_gross_sums,
                        date_col_index=date_col_index,
                        start_row_offset=data_row_offsets['A-la-carte Count'] - 1,
                        advisor_mapping=advisor_mapping
                    )
                    st.success("A-La-Carte data updated successfully.")
                except Exception as e:
                    st.error(f"Error updating A-La-Carte data: {e}")
                time.sleep(delay_seconds)

    with col3:
        if any(commodities_files.values()) or alignment_menus_file or alignment_alacarte_file:
            if st.button("Update Commodities in Google Sheet"):
                # Collect commodities data
                commodities_data = {}
                
                # Process each commodity
                for commodity in commodities_list:
                    if commodities_files[commodity] is not None:
                        if commodity == 'Tires':
                            try:
                                # Attempt to read with original format (header=0)
                                df = pd.read_excel(commodities_files[commodity], header=0)
                                actual_quantity_sums, gross_sums = process_tires_data(df)
                                commodities_data['Tires'] = {
                                    'actual_quantity_sums': actual_quantity_sums,
                                    'gross_sums': gross_sums
                                }
                                st.success(f"{commodity} data (Original Format) processed successfully.")
                            except Exception as e:
                                st.warning(f"Original format not detected for {commodity}. Trying GM Format.")
                                try:
                                    # Attempt to read with GM format (skip first two rows, header=0)
                                    actual_quantity_sums, gross_sums = process_tires_gm_format(commodities_files[commodity])
                                    commodities_data['Tires'] = {
                                        'actual_quantity_sums': actual_quantity_sums,
                                        'gross_sums': gross_sums
                                    }
                                    st.success(f"{commodity} data (GM Format) processed successfully.")
                                except Exception as e2:
                                    st.error(f"Error processing {commodity} Excel file in both formats: {e2}")
                                    commodities_data['Tires'] = {
                                        'actual_quantity_sums': {},
                                        'gross_sums': {}
                                    }
                        else:
                            try:
                                df = pd.read_excel(commodities_files[commodity], header=0)
                                name_counts, parts_gross_sums = process_commodity_file(df)
                                commodities_data[commodity] = {
                                    'name_counts': name_counts,
                                    'parts_gross_sums': parts_gross_sums
                                }
                                st.success(f"{commodity} data processed successfully.")
                            except Exception as e:
                                st.error(f"Error processing {commodity} Excel file: {e}")
                                commodities_data[commodity] = {
                                    'name_counts': {},
                                    'parts_gross_sums': {}
                                }
                
                # Process Alignments separately
                if alignment_menus_file and alignment_alacarte_file:
                    try:
                        # Read both Alignment Excel files
                        df_menus = pd.read_excel(alignment_menus_file, header=0)
                        df_alacarte = pd.read_excel(alignment_alacarte_file, header=0)
                        
                        # Process combined Alignment data
                        name_counts, parts_gross_sums, labor_gross_sums = process_alignment_files(df_menus, df_alacarte)
                        
                        commodities_data['Alignments'] = {
                            'name_counts': name_counts,
                            'parts_gross_sums': parts_gross_sums,
                            'labor_gross_sums': labor_gross_sums
                        }
                        
                        st.success("Alignments data from both files processed successfully.")
                    except Exception as e:
                        st.error(f"Error processing Alignments Excel files: {e}")
                        commodities_data['Alignments'] = {
                            'name_counts': {},
                            'parts_gross_sums': {},
                            'labor_gross_sums': {}
                        }
                else:
                    if any([alignment_menus_file, alignment_alacarte_file]):
                        st.error("Please upload both Alignment Menus and Alignment A-La-Carte Excel files.")
                
                # Update Commodities in Google Sheet
                try:
                    update_commodities_in_sheet(
                        sheet,
                        date_col_index=date_col_index,
                        commodities_data=commodities_data,
                        commodities_list=commodities_list + ['Alignments'],  # Include 'Alignments' for mapping
                        advisor_mapping=advisor_mapping,
                        data_row_offsets=data_row_offsets
                    )
                    st.success("Commodities data updated successfully.")
                except Exception as e:
                    st.error(f"Error updating Commodities data: {e}")
                time.sleep(delay_seconds)

    with col4:
        if recommendations_file is not None:
            if st.button("Update Recommendations in Google Sheet"):
                try:
                    df_recommendations = pd.read_excel(recommendations_file)
                    rec_count, rec_sold_count, rec_amount, rec_sold_amount = process_recommendations_data(df_recommendations, "Name")
                    update_google_sheet(
                        sheet,
                        rec_count,
                        rec_sold_count,
                        rec_amount,
                        rec_sold_amount,
                        date_col_index=date_col_index,
                        start_row_offset=data_row_offsets['Rec Count'] - 1,
                        advisor_mapping=advisor_mapping
                    )
                    st.success("Recommendations data updated successfully.")
                except Exception as e:
                    st.error(f"Error updating Recommendations data: {e}")
                time.sleep(delay_seconds)

    with col5:
        if daily_file is not None:
            if st.button("Update Daily Data in Google Sheet"):
                try:
                    df_daily = pd.read_excel(daily_file)
                    daily_labor_gross, daily_parts_gross = process_daily_data(df_daily, "Name")
                    update_google_sheet(
                        sheet,
                        daily_labor_gross,
                        daily_parts_gross,
                        date_col_index=date_col_index,
                        start_row_offset=data_row_offsets['Daily Labor Gross'] - 1,
                        advisor_mapping=advisor_mapping
                    )
                    st.success("Daily data updated successfully.")
                except Exception as e:
                    st.error(f"Error updating Daily data: {e}")
                time.sleep(delay_seconds)

    # Handling 'Input All' button
    if st.button("Input All"):
        # Process Menu Sales
        if menu_sales_file:
            try:
                df_menu_sales = pd.read_excel(menu_sales_file)
                menu_name_counts, menu_labor_gross_sums, menu_parts_gross_sums = process_menu_sales_data(df_menu_sales, "Advisor Name")
                update_google_sheet(
                    sheet,
                    menu_name_counts,
                    menu_labor_gross_sums,
                    menu_parts_gross_sums,
                    date_col_index=date_col_index,
                    start_row_offset=data_row_offsets['Menu Sales'] - 1,
                    advisor_mapping=advisor_mapping
                )
                st.success("Menu Sales data updated successfully.")
            except Exception as e:
                st.error(f"Error updating Menu Sales data: {e}")
            time.sleep(delay_seconds)

        # Process A-La-Carte
        if alacarte_file:
            try:
                df_alacarte = pd.read_excel(alacarte_file)
                alacarte_name_counts, alacarte_labor_gross_sums, alacarte_parts_gross_sums = process_alacarte_data(df_alacarte, "Advisor Name")
                update_google_sheet(
                    sheet,
                    alacarte_name_counts,
                    alacarte_labor_gross_sums,
                    alacarte_parts_gross_sums,
                    date_col_index=date_col_index,
                    start_row_offset=data_row_offsets['A-la-carte Count'] - 1,
                    advisor_mapping=advisor_mapping
                )
                st.success("A-La-Carte data updated successfully.")
            except Exception as e:
                st.error(f"Error updating A-La-Carte data: {e}")
            time.sleep(delay_seconds)

        # Process Commodities and Alignments
        commodities_data = {}
        for commodity in commodities_list:
            if commodities_files[commodity] is not None:
                if commodity == 'Tires':
                    try:
                        # Attempt to read with original format (header=0)
                        df = pd.read_excel(commodities_files[commodity], header=0)
                        actual_quantity_sums, gross_sums = process_tires_data(df)
                        commodities_data['Tires'] = {
                            'actual_quantity_sums': actual_quantity_sums,
                            'gross_sums': gross_sums
                        }
                        st.success(f"{commodity} data (Original Format) processed successfully.")
                    except Exception as e:
                        st.warning(f"Original format not detected for {commodity}. Trying GM Format.")
                        try:
                            # Attempt to read with GM format (skip first two rows, header=0)
                            actual_quantity_sums, gross_sums = process_tires_gm_format(commodities_files[commodity])
                            commodities_data['Tires'] = {
                                'actual_quantity_sums': actual_quantity_sums,
                                'gross_sums': gross_sums
                            }
                            st.success(f"{commodity} data (GM Format) processed successfully.")
                        except Exception as e2:
                            st.error(f"Error processing {commodity} Excel file in both formats: {e2}")
                            commodities_data['Tires'] = {
                                'actual_quantity_sums': {},
                                'gross_sums': {}
                            }
                else:
                    try:
                        df = pd.read_excel(commodities_files[commodity], header=0)
                        name_counts, parts_gross_sums = process_commodity_file(df)
                        commodities_data[commodity] = {
                            'name_counts': name_counts,
                            'parts_gross_sums': parts_gross_sums
                        }
                        st.success(f"{commodity} data processed successfully.")
                    except Exception as e:
                        st.error(f"Error processing {commodity} Excel file: {e}")
                        commodities_data[commodity] = {
                            'name_counts': {},
                            'parts_gross_sums': {}
                        }

        # Process Alignments
        if alignment_menus_file and alignment_alacarte_file:
            try:
                # Read both Alignment Excel files
                df_menus = pd.read_excel(alignment_menus_file, header=0)
                df_alacarte = pd.read_excel(alignment_alacarte_file, header=0)
                
                # Process combined Alignment data
                name_counts, parts_gross_sums, labor_gross_sums = process_alignment_files(df_menus, df_alacarte)
                
                commodities_data['Alignments'] = {
                    'name_counts': name_counts,
                    'parts_gross_sums': parts_gross_sums,
                    'labor_gross_sums': labor_gross_sums
                }
                
                st.success("Alignments data from both files processed successfully.")
            except Exception as e:
                st.error(f"Error processing Alignments Excel files: {e}")
                commodities_data['Alignments'] = {
                    'name_counts': {},
                    'parts_gross_sums': {},
                    'labor_gross_sums': {}
                }
        else:
            if any([alignment_menus_file, alignment_alacarte_file]):
                st.error("Please upload both Alignment Menus and Alignment A-La-Carte Excel files.")

        # Update Commodities in Google Sheet
        try:
            update_commodities_in_sheet(
                sheet,
                date_col_index=date_col_index,
                commodities_data=commodities_data,
                commodities_list=commodities_list + ['Alignments'],  # Include 'Alignments' for mapping
                advisor_mapping=advisor_mapping,
                data_row_offsets=data_row_offsets
            )
            st.success("Commodities data updated successfully.")
        except Exception as e:
            st.error(f"Error updating Commodities data: {e}")
        time.sleep(delay_seconds)

        # Process Recommendations
        if recommendations_file:
            try:
                df_recommendations = pd.read_excel(recommendations_file)
                rec_count, rec_sold_count, rec_amount, rec_sold_amount = process_recommendations_data(df_recommendations, "Name")
                update_google_sheet(
                    sheet,
                    rec_count,
                    rec_sold_count,
                    rec_amount,
                    rec_sold_amount,
                    date_col_index=date_col_index,
                    start_row_offset=data_row_offsets['Rec Count'] - 1,
                    advisor_mapping=advisor_mapping
                )
                st.success("Recommendations data updated successfully.")
            except Exception as e:
                st.error(f"Error updating Recommendations data: {e}")
            time.sleep(delay_seconds)

        # Process Daily Data
        if daily_file:
            try:
                df_daily = pd.read_excel(daily_file)
                daily_labor_gross, daily_parts_gross = process_daily_data(df_daily, "Name")
                update_google_sheet(
                    sheet,
                    daily_labor_gross,
                    daily_parts_gross,
                    date_col_index=date_col_index,
                    start_row_offset=data_row_offsets['Daily Labor Gross'] - 1,
                    advisor_mapping=advisor_mapping
                )
                st.success("Daily data updated successfully.")
            except Exception as e:
                st.error(f"Error updating Daily data: {e}")
            time.sleep(delay_seconds)

        st.success("All data updated successfully.")

if __name__ == "__main__":
    main()