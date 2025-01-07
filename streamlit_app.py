import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.cell import Cell
from datetime import datetime
import warnings
import time
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

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
        st.error(f"An unexpected error occurred while connecting to Google Sheets: {e}. Please check the configuration and try again.")
        return None

def clean_column_data(column):
    return column.replace(r'[\$,]', '', regex=True).replace(',', '', regex=True).astype(float)

def convert_to_native_type(value):
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

def process_menu_sales_data(df, names_column='Advisor Name'):
    df[names_column] = df[names_column].str.strip().str.upper()
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    name_counts = df[names_column].value_counts() / 2
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum().to_dict()
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum().to_dict()
    return name_counts.to_dict(), labor_gross_sums, parts_gross_sums

def process_alacarte_data(df, names_column='Advisor Name'):
    df[names_column] = df[names_column].str.strip().str.upper()
    df['Opcode Labor Gross'] = clean_column_data(df['Opcode Labor Gross'])
    df['Opcode Parts Gross'] = clean_column_data(df['Opcode Parts Gross'])
    name_counts = df[names_column].value_counts().to_dict()
    labor_gross_sums = df.groupby(names_column)['Opcode Labor Gross'].sum().to_dict()
    parts_gross_sums = df.groupby(names_column)['Opcode Parts Gross'].sum().to_dict()
    return name_counts, labor_gross_sums, parts_gross_sums

def process_commodity_file(df, names_column='Primary Advisor Name', gross_column='Gross'):
    df[names_column] = df[names_column].astype(str).str.strip().str.upper()
    df[gross_column] = clean_column_data(df[gross_column])
    name_counts = df[names_column].value_counts()
    parts_gross_sums = df.groupby(names_column)[gross_column].sum()
    name_counts = name_counts.to_dict()
    parts_gross_sums = parts_gross_sums.to_dict()
    return name_counts, parts_gross_sums

# ----------------------- OLD TIRES LOGIC -----------------------
def process_tires_data(df):
    names_column = None
    quantity_column = None
    gross_column = None

    for col in df.columns:
        col_lower = col.lower()
        if 'advisor' in col_lower and 'name' in col_lower:
            names_column = col
        elif 'part count' in col_lower or 'actual quantity' in col_lower:
            quantity_column = col
        elif 'opcode parts gross' in col_lower or 'gross' in col_lower:
            gross_column = col

    if names_column and quantity_column and gross_column:
        if 'advisor name group' in names_column.lower():
            st.write("Detected GM Tires Format.")
        else:
            st.write("Detected Original Tires Format.")
    else:
        raise ValueError("Tires Excel does not match any known format.")

    df[names_column] = df[names_column].astype(str).str.strip().str.upper()

    try:
        df[quantity_column] = clean_column_data(df[quantity_column])
        df[gross_column] = clean_column_data(df[gross_column])
    except Exception as e:
        raise ValueError(f"Error cleaning columns: {e}")

    actual_quantity_sums = df.groupby(names_column)[quantity_column].sum().to_dict()
    gross_sums = df.groupby(names_column)[gross_column].sum().to_dict()

    actual_quantity_sums = {k: float(v) for k, v in actual_quantity_sums.items()}
    gross_sums = {k: float(v) for k, v in gross_sums.items()}
    return actual_quantity_sums, gross_sums

def process_tires_gm_format(file):
    try:
        df = pd.read_excel(file, skiprows=2, header=0)
        actual_quantity_sums, gross_sums = process_tires_data(df)
        return actual_quantity_sums, gross_sums
    except Exception as e:
        raise ValueError(f"Error processing GM Format Tires Excel file: {e}")

# ---------------------- OLD ALIGNMENT LOGIC --------------------
def process_alignment_files(df_menus, df_alacarte, names_column='Advisor Name'):
    combined_df = pd.concat([df_menus, df_alacarte], ignore_index=True)
    combined_df[names_column] = combined_df[names_column].astype(str).str.strip().str.upper()
    try:
        combined_df['Opcode Labor Gross'] = clean_column_data(combined_df['Opcode Labor Gross'])
        combined_df['Opcode Parts Gross'] = clean_column_data(combined_df['Opcode Parts Gross'])
    except Exception as e:
        raise ValueError(f"Error cleaning columns in Alignment data: {e}")

    # Old logic: dividing name_counts by 2 to avoid double counting
    name_counts = (combined_df[names_column].value_counts() / 2).to_dict()
    parts_gross_sums = combined_df.groupby(names_column)['Opcode Parts Gross'].sum().to_dict()
    labor_gross_sums = combined_df.groupby(names_column)['Opcode Labor Gross'].sum().to_dict()
    return name_counts, parts_gross_sums, labor_gross_sums

# --------------------- NEW MENUS LOGIC: WHEEL ALIGNMENT --------
def process_alignment_menus_new_format(df, advisor_col='Advisor Name', story_col='Operation Tech Story'):
    """
    This function parses the new 'Alignment Menus' Excel, 
    looking for the phrase 'wheel alignment' in the Operation Tech Story.
    """
    df[advisor_col] = df[advisor_col].astype(str).str.strip().str.upper()
    alignment_counts = {}
    for _, row in df.iterrows():
        advisor = row[advisor_col]
        # Lowercase so we can search for "wheel alignment" case-insensitively
        story_text = str(row.get(story_col, "")).lower()
        if "wheel alignment" in story_text:
            alignment_counts[advisor] = alignment_counts.get(advisor, 0) + 1

    # We'll only have name_counts. No labor/parts in this new approach.
    return alignment_counts, {}, {}

def process_recommendations_data(df, names_column="Name"):
    df.columns = df.columns.str.strip()
    df = df[df[names_column].str.strip().str.upper() != "TOTAL"]
    required_columns = ['Recommendations', 'Recommendations Sold', 'Recommendations $ amount', 'Recommendations Sold $ amount']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in the uploaded Recommendations Excel. Please check the column names.")
    df[names_column] = df[names_column].str.strip().str.upper()
    rec_count = df.groupby(names_column)['Recommendations'].sum().to_dict()
    rec_sold_count = df.groupby(names_column)['Recommendations Sold'].sum().to_dict()
    rec_amount = clean_column_data(df.groupby(names_column)['Recommendations $ amount'].sum()).to_dict()
    rec_sold_amount = clean_column_data(df.groupby(names_column)['Recommendations Sold $ amount'].sum()).to_dict()
    return rec_count, rec_sold_count, rec_amount, rec_sold_amount

def process_daily_data(df, names_column="Name"):
    df.columns = df.columns.str.strip()
    df = df[df[names_column].str.strip().str.upper() != "TOTAL"]
    df = df[df['Pay Type'].str.upper() == "ALL"]
    df[names_column] = df[names_column].str.strip().str.upper()
    required_columns = ['Labor Gross', 'Parts Gross']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in the uploaded Daily Data Excel. Please check the column names.")
    df['Labor Gross'] = clean_column_data(df['Labor Gross'])
    df['Parts Gross'] = clean_column_data(df['Parts Gross'])
    labor_gross_sums = df.groupby(names_column)['Labor Gross'].sum().to_dict()
    parts_gross_sums = df.groupby(names_column)['Parts Gross'].sum().to_dict()
    return labor_gross_sums, parts_gross_sums

def process_ro_count_data(df, advisor_column='Advisor Name', ro_number_column='RO Number'):
    df.columns = df.columns.str.strip()
    if advisor_column not in df.columns or ro_number_column not in df.columns:
        raise ValueError(f"Columns '{advisor_column}' or '{ro_number_column}' not found in the uploaded RO Count Excel.")
    df[advisor_column] = df[advisor_column].str.strip().str.upper()
    df = df.dropna(subset=[ro_number_column])
    df[ro_number_column] = df[ro_number_column].astype(str).str.strip()
    unique_ro = df.drop_duplicates(subset=[advisor_column, ro_number_column])
    ro_counts = unique_ro.groupby(advisor_column)[ro_number_column].nunique().to_dict()
    return ro_counts

def update_google_sheet(sheet, data_series1, *args, date_col_index, start_row_offset, advisor_mapping):
    cells_to_update = []
    for advisor_name, start_row in advisor_mapping.items():
        row_index = start_row + start_row_offset
        value1 = data_series1.get(advisor_name, 0)
        value1 = convert_to_native_type(value1)
        cell = Cell(row=row_index, col=date_col_index, value=value1)
        cells_to_update.append(cell)
        for i, data_series in enumerate(args):
            value = data_series.get(advisor_name, 0)
            value = convert_to_native_type(value)
            cell = Cell(row=row_index + i + 1, col=date_col_index, value=value)
            cells_to_update.append(cell)
    if cells_to_update:
        try:
            sheet.update_cells(cells_to_update)
        except Exception as e:
            st.error(f"Failed to update Google Sheet cells: {e}")

def update_commodities_in_sheet(sheet, date_col_index, commodities_data, commodities_list, advisor_mapping, data_row_offsets):
    cells_to_update = {}
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
    labor_gross_offset = data_row_offsets['Labor Gross']
    parts_gross_offset = data_row_offsets['Parts Gross'] 
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
                actual_quantity = convert_to_native_type(actual_quantity_sums.get(advisor_name, 0))
                cell_actual_quantity = Cell(row=base_row, col=date_col_index, value=actual_quantity)
                cells_to_update.setdefault(advisor_name, []).append(cell_actual_quantity)
                gross = convert_to_native_type(gross_sums.get(advisor_name, 0))
                total_parts_gross_per_advisor[advisor_name] += gross
            else:
                count_value = convert_to_native_type(name_counts.get(advisor_name, 0))
                cell_count = Cell(row=base_row, col=date_col_index, value=count_value)
                cells_to_update.setdefault(advisor_name, []).append(cell_count)
                parts_gross_value = convert_to_native_type(parts_gross_sums.get(advisor_name, 0))
                total_parts_gross_per_advisor[advisor_name] += parts_gross_value

                # If alignments, we may have labor sums from old logic:
                if commodity == 'Alignments':
                    labor_gross_value = convert_to_native_type(labor_gross_sums.get(advisor_name, 0))
                    total_labor_gross_per_advisor[advisor_name] += labor_gross_value

    # Now fill in the "Labor Gross" and "Parts Gross" rows for each advisor
    for advisor_name, start_row in advisor_mapping.items():
        labor_gross = total_labor_gross_per_advisor.get(advisor_name, 0)
        labor_gross_row = start_row + labor_gross_offset - 1
        cell_labor_gross = Cell(row=labor_gross_row, col=date_col_index, value=labor_gross)

        parts_gross = total_parts_gross_per_advisor.get(advisor_name, 0)
        parts_gross_row = start_row + parts_gross_offset - 1
        cell_parts_gross = Cell(row=parts_gross_row, col=date_col_index, value=parts_gross)

        cells_to_update.setdefault(advisor_name, []).extend([cell_labor_gross, cell_parts_gross])

    all_cells = []
    for advisor_cells in cells_to_update.values():
        all_cells.extend(advisor_cells)

    if all_cells:
        try:
            sheet.update_cells(all_cells)
        except Exception as e:
            st.error(f"Failed to update Commodities in Google Sheet: {e}")

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

    sheet_name = st.text_input("Enter the Google Sheet name:", "SHEET NAME HERE")
    worksheet_name = st.text_input("Enter the Worksheet (tab) name:", "Input")

    st.subheader("Upload Excel Files")

    st.markdown("#### **Upload RO Count Excel**")
    ro_count_file = st.file_uploader("Select RO Count Excel file", type=["xlsx"], key="ro_count", label_visibility="hidden")

    st.markdown("#### **Upload Menu Sales Excel**")
    menu_sales_file = st.file_uploader("Upload Menu Sales Excel", type=["xlsx"], key="menu_sales_file", label_visibility="hidden")

    st.markdown("#### **Upload A-La-Carte Excel**")
    alacarte_file = st.file_uploader("Upload A-La-Carte Excel", type=["xlsx"], key="alacarte_file", label_visibility="hidden")

    st.markdown("#### **Upload Recommendations Excel**")
    recommendations_file = st.file_uploader("Upload Recommendations Excel", type=["xlsx"], key="recommendations_file", label_visibility="hidden")

    st.markdown("#### **Upload Daily Data Excel**")
    daily_file = st.file_uploader("Upload Daily Data Excel", type=["xlsx"], key="daily_file", label_visibility="hidden")

    # -------------- Commodities Files --------------
    st.markdown("### **Upload Commodities Files**")

    commodities_list = [
        'Air Filters', 'Cabin Filters', 'Batteries', 'Tires', 'Brakes',
        'Wipers', 'Belts', 'Fluids', 'Factory Chemicals'
    ]
    
    commodities_files = {}
    for commodity in commodities_list:
        key = f"commodity_{commodity.replace(' ', '_').lower()}"
        commodities_files[commodity] = st.file_uploader(f"Upload {commodity} Excel", type=["xlsx"], key=key)

    # -------------- Alignment Files --------------
    st.markdown("### **Upload Alignment Files**")
    alignment_menus_file = st.file_uploader("Upload Alignment Menus Excel", type=["xlsx"], key="alignment_menus")
    alignment_alacarte_file = st.file_uploader("Upload Alignment A-La-Carte Excel", type=["xlsx"], key="alignment_alacarte")

    # -------------- Date Input --------------
    selected_date = st.date_input("Select the date:", datetime.now(), key="selected_date").strftime('%d').lstrip('0')

    # -------------- Connect to Google Sheet --------------
    sheet = connect_to_google_sheet(sheet_name, worksheet_name)
    if sheet is None:
        st.error("Failed to connect to the Google Sheet. Please check the inputs and try again.")
        return

    date_row = sheet.row_values(2)[2:]
    date = selected_date
    if date in date_row:
        date_col_index = date_row.index(date) + 3
    else:
        st.error(f"Date {date} not found in the sheet.")
        return

    # -------------- Get Advisors --------------
    col_a_values = sheet.col_values(1)[3:]
    advisor_names = []
    advisor_start_rows = []
    row = 4
    idx = 0
    while idx < len(col_a_values):
        advisor_name = col_a_values[idx]
        if not advisor_name:
            break
        advisor_name = advisor_name.strip().upper()
        advisor_names.append(advisor_name)
        advisor_start_rows.append(row + idx)
        idx += 26
    advisor_mapping = dict(zip(advisor_names, advisor_start_rows))

    data_row_offsets = {
        'RO Count': 1,
        'Menu Sales': 2,
        'Menu Sales Labor Gross': 3,
        'Menu Sales Parts Gross': 4,
        'A-la-carte Count': 5,
        'A-la-carte Labor Gross': 6,
        'A-la-carte Parts Gross': 7,
        'Labor Gross': 18,
        'Parts Gross': 19,
        'Rec Count': 20,
        'Rec Sold Count': 21,
        'Rec Amount': 22,
        'Rec Sold Amount': 23,
        'Daily Labor Gross': 24,
        'Daily Parts Gross': 25,
    }

    # -------------- Buttons Layout --------------
    col1, col2, col3, col4, col5, col6 = st.columns(6)

    # ----- Update RO Count -----
    with col1:
        if ro_count_file is not None:
            if st.button("Update RO Count in Google Sheet", key="update_ro_count"):
                try:
                    df_ro_count = pd.read_excel(ro_count_file)
                    ro_counts = process_ro_count_data(df_ro_count, advisor_column='Advisor Name', ro_number_column='RO Number')
                    update_google_sheet(
                        sheet,
                        ro_counts,
                        date_col_index=date_col_index,
                        start_row_offset=data_row_offsets['RO Count'] - 1,
                        advisor_mapping=advisor_mapping
                    )
                    st.success("RO Count data updated successfully.")
                except Exception as e:
                    st.error(f"Error updating RO Count data: {e}")
                time.sleep(delay_seconds)

    # ----- Update Menu Sales -----
    with col2:
        if menu_sales_file is not None:
            if st.button("Update Menu Sales in Google Sheet", key="update_menu_sales"):
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

    # ----- Update A-La-Carte -----
    with col3:
        if alacarte_file is not None:
            if st.button("Update A-La-Carte in Google Sheet", key="update_alacarte"):
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

    # ----- Update Commodities -----
    with col4:
        if any(commodities_files.values()) or alignment_menus_file or alignment_alacarte_file:
            if st.button("Update Commodities in Google Sheet", key="update_commodities"):
                commodities_data = {}

                # ~~~~~ Process Commodities ~~~~~
                for commodity in commodities_list:
                    if commodities_files[commodity] is not None:
                        if commodity == 'Tires':
                            try:
                                df = pd.read_excel(commodities_files[commodity], header=0)
                                actual_quantity_sums, gross_sums = process_tires_data(df)
                                commodities_data['Tires'] = {
                                    'actual_quantity_sums': actual_quantity_sums,
                                    'gross_sums': gross_sums
                                }
                                st.success(f"{commodity} data (Original Format) processed successfully.")
                            except Exception:
                                try:
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

                # ~~~~~ Process Alignments ~~~~~
                # We'll combine the new "Alignment Menus" approach with old "A-La-Carte" approach
                alignment_menus_counts = {}
                alignment_alacarte_counts = {}
                alignment_alacarte_parts = {}
                alignment_alacarte_labor = {}

                # 1) If alignment menus file is uploaded => NEW logic
                if alignment_menus_file is not None:
                    try:
                        df_menus_new = pd.read_excel(alignment_menus_file, header=0)
                        # EXAMPLE: 'Advisor Name' / 'Operation Tech Story' => Adjust if different
                        alignment_menus_counts, _, _ = process_alignment_menus_new_format(
                            df_menus_new,
                            advisor_col="Advisor Name",
                            story_col="Operation Tech Story"
                        )
                        st.success("Alignment Menus (New Format) processed successfully.")
                    except Exception as e:
                        st.error(f"Error processing new-format Alignment Menus: {e}")

                # 2) If alignment a-la-carte file is uploaded => OLD logic
                if alignment_alacarte_file is not None:
                    try:
                        df_alacarte_align = pd.read_excel(alignment_alacarte_file, header=0)
                        # We pretend there's no "menus" DF for old logic => pass empty
                        df_menus_empty = pd.DataFrame()
                        alignment_alacarte_counts, alignment_alacarte_parts, alignment_alacarte_labor = process_alignment_files(
                            df_menus_empty,
                            df_alacarte_align,
                            names_column="Advisor Name"
                        )
                        st.success("Alignment A-La-Carte (Old Logic) processed successfully.")
                    except Exception as e:
                        st.error(f"Error processing old-format Alignment A-La-Carte: {e}")

                # 3) Merge the two sets into final
                final_align_counts = {}
                for adv, c in alignment_menus_counts.items():
                    final_align_counts[adv] = final_align_counts.get(adv, 0) + c
                for adv, c in alignment_alacarte_counts.items():
                    final_align_counts[adv] = final_align_counts.get(adv, 0) + c

                final_align_parts = alignment_alacarte_parts
                final_align_labor = alignment_alacarte_labor

                # 4) Put them into commodities_data['Alignments']
                commodities_data['Alignments'] = {
                    'name_counts': final_align_counts,
                    'parts_gross_sums': final_align_parts,
                    'labor_gross_sums': final_align_labor
                }

                # ~~~~~ Update in Google Sheet ~~~~~
                try:
                    update_commodities_in_sheet(
                        sheet,
                        date_col_index=date_col_index,
                        commodities_data=commodities_data,
                        commodities_list=commodities_list + ['Alignments'],
                        advisor_mapping=advisor_mapping,
                        data_row_offsets=data_row_offsets
                    )
                    st.success("Commodities data updated successfully.")
                except Exception as e:
                    st.error(f"Error updating Commodities data: {e}")
                time.sleep(delay_seconds)

    # ----- Update Recommendations -----
    with col5:
        if recommendations_file is not None:
            if st.button("Update Recommendations in Google Sheet", key="update_recommendations"):
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

    # ----- Update Daily Data -----
    with col6:
        if daily_file is not None:
            if st.button("Update Daily Data in Google Sheet", key="update_daily_data"):
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

    # -------------- Input All Button --------------
    if st.button("Input All", key="input_all"):
        updated_sections = []

        # ---------- A) RO Count ----------
        if ro_count_file:
            try:
                df_ro_count = pd.read_excel(ro_count_file)
                ro_counts = process_ro_count_data(df_ro_count, advisor_column='Advisor Name', ro_number_column='RO Number')
                update_google_sheet(
                    sheet,
                    ro_counts,
                    date_col_index=date_col_index,
                    start_row_offset=data_row_offsets['RO Count'] - 1,
                    advisor_mapping=advisor_mapping
                )
                updated_sections.append("RO Count")
                st.success("RO Count data updated successfully.")
            except Exception as e:
                st.error(f"Error updating RO Count data: {e}")
            time.sleep(delay_seconds)

        # ---------- B) Menu Sales ----------
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
                updated_sections.append("Menu Sales")
                st.success("Menu Sales data updated successfully.")
            except Exception as e:
                st.error(f"Error updating Menu Sales data: {e}")
            time.sleep(delay_seconds)

        # ---------- C) A-La-Carte ----------
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
                updated_sections.append("A-La-Carte")
                st.success("A-La-Carte data updated successfully.")
            except Exception as e:
                st.error(f"Error updating A-La-Carte data: {e}")
            time.sleep(delay_seconds)

        # ---------- D) Commodities + Alignments ----------
        commodities_data = {}

        # 1) Normal Commodities
        for commodity in commodities_list:
            if commodities_files[commodity] is not None:
                if commodity == 'Tires':
                    try:
                        df = pd.read_excel(commodities_files[commodity], header=0)
                        actual_quantity_sums, gross_sums = process_tires_data(df)
                        commodities_data['Tires'] = {
                            'actual_quantity_sums': actual_quantity_sums,
                            'gross_sums': gross_sums
                        }
                        updated_sections.append("Tires")
                        st.success(f"{commodity} data (Original Format) processed successfully.")
                    except Exception:
                        try:
                            actual_quantity_sums, gross_sums = process_tires_gm_format(commodities_files[commodity])
                            commodities_data['Tires'] = {
                                'actual_quantity_sums': actual_quantity_sums,
                                'gross_sums': gross_sums
                            }
                            updated_sections.append("Tires (GM Format)")
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
                        updated_sections.append(commodity)
                        st.success(f"{commodity} data processed successfully.")
                    except Exception as e:
                        st.error(f"Error processing {commodity} Excel file: {e}")
                        commodities_data[commodity] = {
                            'name_counts': {},
                            'parts_gross_sums': {}
                        }

        # 2) Alignments
        alignment_menus_counts = {}
        alignment_alacarte_counts = {}
        alignment_alacarte_parts = {}
        alignment_alacarte_labor = {}

        # If alignment menus file => new logic
        if alignment_menus_file:
            try:
                df_menus_new = pd.read_excel(alignment_menus_file, header=0)
                alignment_menus_counts, _, _ = process_alignment_menus_new_format(
                    df_menus_new,
                    advisor_col="Advisor Name",
                    story_col="Operation Tech Story"
                )
                st.success("Alignment Menus (New Format) processed successfully.")
            except Exception as e:
                st.error(f"Error processing new-format Alignment Menus: {e}")

        # If alignment a-la-carte file => old logic
        if alignment_alacarte_file:
            try:
                df_alacarte_align = pd.read_excel(alignment_alacarte_file, header=0)
                df_menus_empty = pd.DataFrame()
                alignment_alacarte_counts, alignment_alacarte_parts, alignment_alacarte_labor = process_alignment_files(
                    df_menus_empty,
                    df_alacarte_align,
                    names_column="Advisor Name"
                )
                st.success("Alignment A-La-Carte (Old Logic) processed successfully.")
            except Exception as e:
                st.error(f"Error processing old-format Alignment A-La-Carte: {e}")

        # Merge new menus + old a-la-carte
        final_align_counts = {}
        for adv, c in alignment_menus_counts.items():
            final_align_counts[adv] = final_align_counts.get(adv, 0) + c
        for adv, c in alignment_alacarte_counts.items():
            final_align_counts[adv] = final_align_counts.get(adv, 0) + c

        final_align_parts = alignment_alacarte_parts
        final_align_labor = alignment_alacarte_labor

        # Put them into 'Alignments'
        commodities_data['Alignments'] = {
            'name_counts': final_align_counts,
            'parts_gross_sums': final_align_parts,
            'labor_gross_sums': final_align_labor
        }

        # Now update Commodities in Google Sheet
        if any(commodities_data.values()):
            try:
                update_commodities_in_sheet(
                    sheet,
                    date_col_index=date_col_index,
                    commodities_data=commodities_data,
                    commodities_list=commodities_list + ['Alignments'],
                    advisor_mapping=advisor_mapping,
                    data_row_offsets=data_row_offsets
                )
                updated_sections.append("Commodities")
                st.success("Commodities data updated successfully.")
            except Exception as e:
                st.error(f"Error updating Commodities data: {e}")
            time.sleep(delay_seconds)

        # ---------- E) Recommendations ----------
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
                updated_sections.append("Recommendations")
                st.success("Recommendations data updated successfully.")
            except Exception as e:
                st.error(f"Error updating Recommendations data: {e}")
            time.sleep(delay_seconds)

        # ---------- F) Daily Data ----------
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
                updated_sections.append("Daily Data")
                st.success("Daily data updated successfully.")
            except Exception as e:
                st.error(f"Error updating Daily data: {e}")
            time.sleep(delay_seconds)

        # ---------- Final Success or Warning ----------
        if updated_sections:
            st.success(f"Updated the following sections successfully: {', '.join(updated_sections)}")
        else:
            st.warning("No data sections were updated. Please ensure you've uploaded the necessary Excel files.")

if __name__ == "__main__":
    main()