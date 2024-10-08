import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import CellFormat, format_cell_range
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

# def process_commodities_data(df, names_column="Primary Advisor Name"):
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

def process_tires_data(df, names_column='Primary Advisor Name', quantity_column='Actual Quantity', gross_column='Gross'):
  
    # Normalize advisor names
    df[names_column] = df[names_column].astype(str).str.strip().str.upper()
    
    # Define required columns
    required_columns = [names_column, quantity_column, gross_column]
    
    # Check if required columns exist
    for col in required_columns:
        if col not in df.columns:
            st.error(f"Column '{col}' not found in the uploaded Tires Excel. Available columns: {df.columns.tolist()}")
            return {}, {}
    
    # Clean the 'Actual Quantity' and 'Gross' columns
    df[quantity_column] = clean_column_data(df[quantity_column])
    df[gross_column] = clean_column_data(df[gross_column])
    
    # Sum Actual Quantity and Gross per advisor
    actual_quantity_sums = df.groupby(names_column)[quantity_column].sum().to_dict()
    gross_sums = df.groupby(names_column)[gross_column].sum().to_dict()
    
    # Convert to native types
    actual_quantity_sums = {k: float(v) for k, v in actual_quantity_sums.items()}
    gross_sums = {k: float(v) for k, v in gross_sums.items()}
    
    # Debug: Print the processed values
    # st.write("Tires Actual Quantity Sums:", actual_quantity_sums)
    # st.write("Tires Gross Sums:", gross_sums)
    
    return actual_quantity_sums, gross_sums

def process_alignment_file(df, names_column='Advisor Name'):
    df[names_column] = df[names_column].astype(str).str.strip().str.upper()
    
    
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])

    # Count the number of sales per advisor
    name_counts = df[names_column].value_counts() / 2  

    # Sum gross amounts per advisor
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum()
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum()

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
    'Parts Gross': 22 }  


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

def convert_to_native_type(value):
    """
    Convert NumPy and Pandas data types to native Python types.
    """
    if isinstance(value, pd.Series):
        value = value.sum()
    if pd.isna(value):
        return 0  # Or None, depending on how you want to handle NaNs
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

def update_commodities_in_sheet(sheet, date_col_index, commodities_data, commodities_list, advisor_mapping):
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
    labor_gross_offset = 18  
    parts_gross_offset = 19  

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
            base_row = start_row + commodity_row_offsets[commodity] - 1  

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

    # Debug: Print out the cells being updated
    #for cell in cells_to_update:
        # st.write(f"Updating cell ({cell.row}, {cell.col}): {cell.value}")

    #if cells_to_update:
        #sheet.update_cells(cells_to_update)

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

    sheet_name = st.text_input("Enter the Google Sheet name:", "GOOGLE SHEET NAME")
    worksheet_name = st.text_input("Enter the Worksheet (tab) name:", "Input")

    # File uploads for different sections
    menu_sales_file = st.file_uploader("Upload Menu Sales Excel", type=["xlsx"])
    alacarte_file = st.file_uploader("Upload A-La-Carte Excel", type=["xlsx"])
    # commodities_file = st.file_uploader("Upload Commodities Excel", type=["xlsx"])
    
    recommendations_file = st.file_uploader("Upload Recommendations Excel", type=["xlsx"])
    daily_file = st.file_uploader("Upload Daily Data Excel", type=["xlsx"])

    # In your main() function, under Commodities
    st.subheader("Upload Commodities Files")

    commodities_list = [
        'Air Filters', 'Cabin Filters', 'Batteries', 'Tires', 'Brakes',
        'Alignments', 'Wipers', 'Belts', 'Fluids', 'Factory Chemicals'
    ]
    
    commodities_files = {}
    for commodity in commodities_list:
        key = f"commodity_{commodity}"
        commodities_files[commodity] = st.file_uploader(f"Upload {commodity} Excel", type=["xlsx"], key=key)

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
        idx += 26  # Assuming each advisor block has 27 rows (including empty row)
    advisor_mapping = dict(zip(advisor_names, advisor_start_rows))

    # Define the row offsets for different data types within each advisor's block
    data_row_offsets = {
        'Menu Sales': 2,
        'Menu Sales Labor Gross': 3,
        'Menu Sales Parts Gross': 4,
        'A-la-carte Count': 5,
        'A-la-carte Labor Gross': 6,
        'A-la-carte Parts Gross': 7,
        # Commodities are from row 2 to 11
        'Labor Gross': 12,
        'Parts Gross': 13,
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
                time.sleep(delay_seconds)

    with col2:
        if alacarte_file is not None:
            if st.button("Update A-La-Carte in Google Sheet"):
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
                time.sleep(delay_seconds)

    with col3:
        if any(commodities_files.values()):
            if st.button("Update Commodities in Google Sheet"):
                # Collect commodities data
                commodities_data = {}

                for commodity in commodities_list:
                    if commodities_files[commodity] is not None:
                        if commodity == 'Alignments':
                            df = pd.read_excel(commodities_files[commodity], header=2)
                            name_counts, parts_gross_sums, labor_gross_sums = process_alignment_file(df)
                            commodities_data[commodity] = {
                                'name_counts': name_counts,
                                'parts_gross_sums': parts_gross_sums,
                                'labor_gross_sums': labor_gross_sums
                            }
                        elif commodity == 'Tires':
                            df = pd.read_excel(commodities_files[commodity], header=0)  # Headers are on row 1
                            actual_quantity_sums, gross_sums = process_tires_data(
                                df,
                                names_column='Primary Advisor Name',
                                quantity_column='Actual Quantity',
                                gross_column='Gross'
                            )
                            commodities_data[commodity] = {
                                'actual_quantity_sums': actual_quantity_sums,
                                'gross_sums': gross_sums
                            }
                        else:
                            df = pd.read_excel(commodities_files[commodity], header=0)  # Headers are on row 1
                            name_counts, parts_gross_sums = process_commodity_file(df)
                            commodities_data[commodity] = {
                                'name_counts': name_counts,
                                'parts_gross_sums': parts_gross_sums
                            }
                update_commodities_in_sheet(
                    sheet,
                    date_col_index=date_col_index,
                    commodities_data=commodities_data,
                    commodities_list=commodities_list,
                    advisor_mapping=advisor_mapping
                )
                st.success("Commodities data updated successfully.")
                time.sleep(delay_seconds)

    with col4:
        if recommendations_file is not None:
            if st.button("Update Recommendations in Google Sheet"):
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
                time.sleep(delay_seconds)

    with col5:
        if daily_file is not None:
            if st.button("Update Daily Data in Google Sheet"):
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
                time.sleep(delay_seconds)

    
    # Inside the 'Input All' button handling
    if st.button("Input All"):
        # Process Menu Sales
        if menu_sales_file:
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
            time.sleep(delay_seconds)

        # Process A-La-Carte
        if alacarte_file:
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
            time.sleep(delay_seconds)

        # Process Commodities and Alignments
        commodities_data = {}
        for commodity in commodities_list:
            if commodities_files[commodity] is not None:
                if commodity == 'Alignments':
                    df = pd.read_excel(commodities_files[commodity], header=2)
                    name_counts, parts_gross_sums, labor_gross_sums = process_alignment_file(df)
                    commodities_data[commodity] = {
                        'name_counts': name_counts,
                        'parts_gross_sums': parts_gross_sums,
                        'labor_gross_sums': labor_gross_sums
                    }
                elif commodity == 'Tires':
                    df = pd.read_excel(commodities_files[commodity], header=0)  # Headers on row 1
                    actual_quantity_sums, gross_sums = process_tires_data(
                        df,
                        names_column='Primary Advisor Name',
                        quantity_column='Actual Quantity',
                        gross_column='Gross'
                    )
                    commodities_data[commodity] = {
                        'actual_quantity_sums': actual_quantity_sums,
                        'gross_sums': gross_sums
                    }
                else:
                    df = pd.read_excel(commodities_files[commodity], header=0)
                    name_counts, parts_gross_sums = process_commodity_file(df)
                    commodities_data[commodity] = {
                        'name_counts': name_counts,
                        'parts_gross_sums': parts_gross_sums
                    }
        # Update Commodities in Google Sheet
        update_commodities_in_sheet(
            sheet,
            date_col_index=date_col_index,
            commodities_data=commodities_data,
            commodities_list=commodities_list,
            advisor_mapping=advisor_mapping
        )
        time.sleep(delay_seconds)

        # Process Recommendations
        if recommendations_file:
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
            time.sleep(delay_seconds)

        # Process Daily Data
        if daily_file:
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
            time.sleep(delay_seconds)

        st.success("All data updated successfully.")
        
if __name__ == "__main__":
    main()
