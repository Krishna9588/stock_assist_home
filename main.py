# import json
# import time
# import pandas as pd
# import os
# from extract import *
# import xlsxwriter
# from hist_data import hist_daily, hist_weekly, hist_monthly, hist_quarterly
#
#
# def create_final_report(json_file_path, hist_dfs, excel_file_path):
#     """
#     Creates a single Excel report with all data sections placed side-by-side,
#     each with its own header, separated by empty columns, and with proper formatting.
#     """
#     try:
#         with open(json_file_path, 'r', encoding='utf-8') as f:
#             data = json.load(f)
#
#         with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
#             workbook = writer.book
#             worksheet = workbook.add_worksheet('Combined Report')
#
#             # --- Define Cell Formats ---
#             primary_header_format = workbook.add_format({
#                 'bold': True,
#                 'font_color': 'white',
#                 'align': 'center',
#                 'valign': 'vcenter',
#                 'fg_color': '#0b5394',
#                 'border': 1
#             })
#             secondary_header_format = workbook.add_format({
#                 'bold': True,
#                 'font_color': '#2F5597',  # Blue
#                 'fg_color': '#F2F2F2',  # Light Gray
#                 'border': 1
#             })
#             # Get currency from the 'Finance' section, default to USD
#             currency_code = data.get("Finance", {}).get("Currency", "USD")
#             currency_symbols = {"USD": "$", "GBP": "£", "EUR": "€", "JPY": "¥", "INR": "₹"}
#             currency_symbol = currency_symbols.get(currency_code, "")
#
#             currency_format = workbook.add_format({'num_format': f'{currency_symbol}#,##0.00', 'border': 1})
#             percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
#             default_cell_format = workbook.add_format({'border': 1})
#             date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#
#             # --- Prepare Data Sections ---
#             section_data = []
#             section_order = ["About", "Finance", "Location", "Growth", "Buyer Group", "Mutual Fund Holders",
#                              "Ownership Structure"]
#             for key in section_order:
#                 if key not in data:
#                     continue
#                 value = data[key]
#                 df = None
#                 if isinstance(value, dict):
#                     df = pd.DataFrame.from_dict(value, orient='index', columns=['Value'])
#                     df.index.name = key
#                     df = df.reset_index()
#                 elif isinstance(value, list) and value:
#                     df = pd.DataFrame(value)
#                 elif isinstance(value, (str, int, float)):
#                     df = pd.DataFrame({key: [value]})
#                 if df is not None and not df.empty:
#                     section_data.append((key, df))
#
#             # Add the multiple historical dataframes
#             for title, df in hist_dfs:
#                 if df is not None and not df.empty:  # The index is already reset in hist_data.py
#                     section_data.append((title, df))
#
#             # --- Write to Excel Horizontally ---
#             current_col = 0
#             for section_title, df in section_data:
#                 if df.empty:
#                     continue
#
#                 # --- Write Headers ---
#                 num_cols = df.shape[1]
#                 if num_cols > 1:
#                     worksheet.merge_range(0, current_col, 0, current_col + num_cols - 1, section_title,
#                                           primary_header_format)
#                 else:
#                     worksheet.write(0, current_col, section_title, primary_header_format)
#
#                 worksheet.write_row(1, current_col, df.columns, secondary_header_format)
#
#                 # --- Write Data and Format Cells ---
#                 for r_idx, row in enumerate(df.itertuples(index=False), start=2):
#                     for c_idx, value in enumerate(row):
#                         col_name = df.columns[c_idx]
#                         cell_format = default_cell_format
#
#                         # --- THIS IS THE NEW, MORE ROBUST FORMATTING LOGIC ---
#
#                         # Condition 1: Column name itself contains 'date' (e.g., 'Date', 'End Date').
#                         is_date_col = 'date' in col_name.lower()
#
#                         # Condition 2: For key-value tables (like 'Finance'), check if the key in the first column contains 'date'.
#                         is_date_key_row = (
#                             df.shape[1] == 2 and  # It's a key-value table
#                             'date' in str(row[0]).lower() and  # The key for this row contains 'date'
#                             c_idx == 1  # We are currently on the 'Value' cell for that key
#                         )
#
#                         if is_date_col or is_date_key_row:
#                             cell_format = date_format
#                         elif col_name in ['Current Price', 'Market Cap', 'Total Revenue', 'open', 'high', 'low',
#                                           'close', 'adjclose', 'Shares']:
#                             cell_format = currency_format
#                         elif col_name in ['Revenue Growth', 'Profit Growth', 'Dividend Yield', 'Growth', 'Holding']:
#                             cell_format = percent_format
#
#                         if pd.isna(value):
#                             value = None  # xlsxwriter handles None as a blank cell
#
#                         worksheet.write(r_idx, current_col + c_idx, value, cell_format)
#
#                 # --- Auto-fit Columns for this section ---
#                 for i, col_name in enumerate(df.columns):
#                     col_idx = current_col + i
#                     # Ensure data_width calculation handles non-string data gracefully
#                     try:
#                         data_width = df[col_name].astype(str).str.len().max()
#                     except:
#                         data_width = 10  # Fallback width
#                     header_width = len(str(col_name))
#                     width = max(data_width if not pd.isna(data_width) else 0, header_width)
#                     worksheet.set_column(col_idx, col_idx, width + 2)  # Add padding
#
#                 # Update start column for next section
#                 current_col += df.shape[1] + 1  # +1 for separator column
#
#             print(f"Step 3. Successfully created combined report: '{excel_file_path}'")
#
#     except FileNotFoundError:
#         print(f"Error: The file '{json_file_path}' was not found.")
#     except json.JSONDecodeError:
#         print(f"Error: Could not decode JSON from '{json_file_path}'. Check the file format.")
#     except Exception as e:
#         print(f"An unexpected error occurred during Excel report creation: {e}")
#
#
# # --- Read tickers from input.csv ---
# try:
#     input_df = pd.read_csv('input.csv')
#     # Select the 'company_id' column and convert it to a list
#     tickers = input_df['company_id'].tolist()
#     print(f"Successfully loaded {len(tickers)} tickers from input.csv")
# except FileNotFoundError:
#     print("Error: 'input.csv' not found. Please create it with a 'company_id' column.")
#     tickers = []  # Use an empty list to avoid crashing the script
# except KeyError:
#     print("Error: 'input.csv' must have a column named 'company_id'.")
#     tickers = []  # Use an empty list to avoid crashing the script
#
# for ticker_id in tickers:  # Loop through the tickers loaded from the CSV
#     print(f"\n----- Processing Ticker: {ticker_id} -----")
#
#     # Create a dedicated output folder for the ticker under a 'result' directory
#     output_dir = os.path.join('result', ticker_id)
#     os.makedirs(output_dir, exist_ok=True)
#
#     # Define file paths within the new directory
#     ext_json_path = os.path.join(output_dir, f"{ticker_id}_ext.json")
#     final_excel_path = os.path.join(output_dir, f"{ticker_id}_Final.xlsx")
#
#     # Step 1: Fetch and save company profile JSON
#     extract(ticker_id, ext_json_path)
#
#     # Step 2: Fetch all historical data as DataFrames
#     print("\n--- Fetching Historical Data ---")
#     hist_dfs = [
#         ("Daily History", hist_daily(ticker_id)),
#         ("Weekly History", hist_weekly(ticker_id)),
#         ("Monthly History", hist_monthly(ticker_id)),
#         ("Quarterly History", hist_quarterly(ticker_id))
#     ]
#
#     # Step 3: Combine JSON data and historical DataFrame into one final Excel file
#     create_final_report(ext_json_path, hist_dfs, final_excel_path)
#     print(f"----- Finished Processing {ticker_id} -----")

# import json
# import time
# import pandas as pd
# import os
# from extract import *
# import xlsxwriter
# from hist_data import hist_daily, hist_weekly, hist_monthly, hist_quarterly
#
#
# def create_final_report(json_file_path, hist_dfs, excel_file_path):
#     """
#     Creates a single Excel report with all data sections placed side-by-side,
#     each with its own header, separated by empty columns, and with proper formatting.
#     """
#     try:
#         with open(json_file_path, 'r', encoding='utf-8') as f:
#             data = json.load(f)
#
#         # This option tells xlsxwriter to automatically convert timezone-aware
#         # datetimes to timezone-naive datetimes, resolving the error.
#         writer_options = {'options': {'remove_timezone': True}}
#         with pd.ExcelWriter(excel_file_path, engine='xlsxwriter', engine_kwargs=writer_options) as writer:
#             workbook = writer.book
#             worksheet = workbook.add_worksheet('Combined Report')
#
#             # --- Define Cell Formats ---
#             primary_header_format = workbook.add_format({
#                 'bold': True,
#                 'font_color': 'white',
#                 'align': 'center',
#                 'valign': 'vcenter',
#                 'fg_color': '#0b5394',
#                 'border': 1
#             })
#             secondary_header_format = workbook.add_format({
#                 'bold': True,
#                 'font_color': '#2F5597',  # Blue
#                 'fg_color': '#F2F2F2',  # Light Gray
#                 'border': 1
#             })
#             # Get currency from the 'Finance' section, default to USD
#             currency_code = data.get("Finance", {}).get("Currency", "USD")
#             currency_symbols = {"USD": "$", "GBP": "£", "EUR": "€", "JPY": "¥", "INR": "₹"}
#             currency_symbol = currency_symbols.get(currency_code, "")
#
#             currency_format = workbook.add_format({'num_format': f'{currency_symbol}#,##0.00', 'border': 1})
#             percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
#             default_cell_format = workbook.add_format({'border': 1})
#             date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#
#             # --- Prepare Data Sections ---
#             section_data = []
#             section_order = ["About", "Finance", "Location", "Growth", "Buyer Group", "Mutual Fund Holders",
#                              "Ownership Structure"]
#             for key in section_order:
#                 if key not in data:
#                     continue
#                 value = data[key]
#                 df = None
#                 if isinstance(value, dict):
#                     df = pd.DataFrame.from_dict(value, orient='index', columns=['Value'])
#                     df.index.name = key
#                     df = df.reset_index()
#                 elif isinstance(value, list) and value:
#                     df = pd.DataFrame(value)
#                 elif isinstance(value, (str, int, float)):
#                     df = pd.DataFrame({key: [value]})
#                 if df is not None and not df.empty:
#                     section_data.append((key, df))
#
#             # Add the multiple historical dataframes
#             for title, df in hist_dfs:
#                 if df is not None and not df.empty:  # The index is already reset in hist_data.py
#                     section_data.append((title, df))
#
#             # --- Write to Excel Horizontally ---
#             current_col = 0
#             for section_title, df in section_data:
#                 if df.empty:
#                     continue
#
#                 # --- Write Headers ---
#                 num_cols = df.shape[1]
#                 if num_cols > 1:
#                     worksheet.merge_range(0, current_col, 0, current_col + num_cols - 1, section_title,
#                                           primary_header_format)
#                 else:
#                     worksheet.write(0, current_col, section_title, primary_header_format)
#
#                 worksheet.write_row(1, current_col, df.columns, secondary_header_format)
#
#                 # --- Write Data and Format Cells ---
#                 for r_idx, row in enumerate(df.itertuples(index=False), start=2):
#                     for c_idx, value in enumerate(row):
#                         col_name = df.columns[c_idx]
#                         cell_format = default_cell_format
#
#                         # --- THIS IS THE NEW, MORE ROBUST FORMATTING LOGIC ---
#
#                         # Condition 1: Column name itself contains 'date' (e.g., 'Date', 'End Date').
#                         is_date_col = 'date' in col_name.lower()
#
#                         # Condition 2: For key-value tables (like 'Finance'), check if the key in the first column contains 'date'.
#                         is_date_key_row = (
#                             df.shape[1] == 2 and  # It's a key-value table
#                             'date' in str(row[0]).lower() and  # The key for this row contains 'date'
#                             c_idx == 1  # We are currently on the 'Value' cell for that key
#                         )
#
#                         if is_date_col or is_date_key_row:
#                             cell_format = date_format
#                         elif col_name in ['Current Price', 'Market Cap', 'Total Revenue', 'open', 'high', 'low',
#                                           'close', 'adjclose', 'Shares']:
#                             cell_format = currency_format
#                         elif col_name in ['Revenue Growth', 'Profit Growth', 'Dividend Yield', 'Growth', 'Holding']:
#                             cell_format = percent_format
#
#                         if pd.isna(value):
#                             value = None  # xlsxwriter handles None as a blank cell
#
#                         worksheet.write(r_idx, current_col + c_idx, value, cell_format)
#
#                 # --- Auto-fit Columns for this section ---
#                 for i, col_name in enumerate(df.columns):
#                     col_idx = current_col + i
#                     # Ensure data_width calculation handles non-string data gracefully
#                     try:
#                         data_width = df[col_name].astype(str).str.len().max()
#                     except:
#                         data_width = 10  # Fallback width
#                     header_width = len(str(col_name))
#                     width = max(data_width if not pd.isna(data_width) else 0, header_width)
#                     worksheet.set_column(col_idx, col_idx, width + 2)  # Add padding
#
#                 # Update start column for next section
#                 current_col += df.shape[1] + 1  # +1 for separator column
#
#             print(f"Step 3. Successfully created combined report: '{excel_file_path}'")
#
#     except FileNotFoundError:
#         print(f"Error: The file '{json_file_path}' was not found.")
#     except json.JSONDecodeError:
#         print(f"Error: Could not decode JSON from '{json_file_path}'. Check the file format.")
#     except Exception as e:
#         print(f"An unexpected error occurred during Excel report creation: {e}")
#
#
# # --- Read tickers from input.csv ---
# try:
#     input_df = pd.read_csv('input.csv')
#     # Select the 'company_id' column and convert it to a list
#     tickers = input_df['company_id'].tolist()
#     print(f"Successfully loaded {len(tickers)} tickers from input.csv")
# except FileNotFoundError:
#     print("Error: 'input.csv' not found. Please create it with a 'company_id' column.")
#     tickers = []  # Use an empty list to avoid crashing the script
# except KeyError:
#     print("Error: 'input.csv' must have a column named 'company_id'.")
#     tickers = []  # Use an empty list to avoid crashing the script
#
# for ticker_id in tickers:  # Loop through the tickers loaded from the CSV
#     print(f"\n----- Processing Ticker: {ticker_id} -----")
#
#     # Create a dedicated output folder for the ticker under a 'result' directory
#     output_dir = os.path.join('result', ticker_id)
#     os.makedirs(output_dir, exist_ok=True)
#
#     # Define file paths within the new directory
#     ext_json_path = os.path.join(output_dir, f"{ticker_id}_ext.json")
#     final_excel_path = os.path.join(output_dir, f"{ticker_id}_Final.xlsx")
#
#     # Step 1: Fetch and save company profile JSON
#     extract(ticker_id, ext_json_path)
#
#     # Step 2: Fetch all historical data as DataFrames
#     print("\n--- Fetching Historical Data ---")
#     hist_dfs = [
#         ("Daily History", hist_daily(ticker_id)),
#         ("Weekly History", hist_weekly(ticker_id)),
#         ("Monthly History", hist_monthly(ticker_id)),
#         ("Quarterly History", hist_quarterly(ticker_id))
#     ]
#
#     # Step 3: Combine JSON data and historical DataFrame into one final Excel file
#     create_final_report(ext_json_path, hist_dfs, final_excel_path)
#     print(f"----- Finished Processing {ticker_id} -----")

import json
import time
import pandas as pd
import os
import datetime  # <-- Added this import
from extract import *
import xlsxwriter
from hist_data import hist_daily, hist_weekly, hist_monthly, hist_quarterly


def create_final_report(json_file_path, hist_dfs, excel_file_path):
    """
    Creates a single Excel report with all data sections placed side-by-side,
    each with its own header, separated by empty columns, and with proper formatting.
    """
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # This option tells xlsxwriter to automatically convert timezone-aware
        # datetimes to timezone-naive datetimes, resolving the error.
        writer_options = {'options': {'remove_timezone': True}}
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter', engine_kwargs=writer_options) as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Combined Report')

            # --- Define Cell Formats ---
            primary_header_format = workbook.add_format({
                'bold': True,
                'font_color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': '#0b5394',
                'border': 1
            })
            secondary_header_format = workbook.add_format({
                'bold': True,
                'font_color': '#2F5597',  # Blue
                'fg_color': '#F2F2F2',  # Light Gray
                'border': 1
            })
            # Get currency from the 'Finance' section, default to USD
            currency_code = data.get("Finance", {}).get("Currency", "USD")
            currency_symbols = {"USD": "$", "GBP": "£", "EUR": "€", "JPY": "¥", "INR": "₹"}
            currency_symbol = currency_symbols.get(currency_code, "")

            currency_format = workbook.add_format({'num_format': f'{currency_symbol}#,##0.00', 'border': 1})
            percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
            default_cell_format = workbook.add_format({'border': 1})
            date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})

            # --- Prepare Data Sections ---
            section_data = []
            section_order = ["About", "Finance", "Location", "Growth", "Buyer Group", "Mutual Fund Holders",
                             "Ownership Structure"]
            for key in section_order:
                if key not in data:
                    continue
                value = data[key]
                df = None
                if isinstance(value, dict):
                    df = pd.DataFrame.from_dict(value, orient='index', columns=['Value'])
                    df.index.name = key
                    df = df.reset_index()
                elif isinstance(value, list) and value:
                    df = pd.DataFrame(value)
                elif isinstance(value, (str, int, float)):
                    df = pd.DataFrame({key: [value]})
                if df is not None and not df.empty:
                    section_data.append((key, df))

            # Add the multiple historical dataframes
            for title, df in hist_dfs:
                if df is not None and not df.empty:  # The index is already reset in hist_data.py
                    section_data.append((title, df))

            # --- Write to Excel Horizontally ---
            current_col = 0
            for section_title, df in section_data:
                if df.empty:
                    continue

                # --- Write Headers ---
                num_cols = df.shape[1]
                if num_cols > 1:
                    worksheet.merge_range(0, current_col, 0, current_col + num_cols - 1, section_title,
                                          primary_header_format)
                else:
                    worksheet.write(0, current_col, section_title, primary_header_format)

                worksheet.write_row(1, current_col, df.columns, secondary_header_format)

                # --- Write Data and Format Cells ---
                for r_idx, row in enumerate(df.itertuples(index=False), start=2):
                    for c_idx, value in enumerate(row):
                        col_name = df.columns[c_idx]

                        # --- NEW, MORE ROBUST FORMATTING LOGIC ---

                        # 1. Check if the field is likely a date based on its name.
                        is_date_field = 'date' in col_name.lower() or (
                            df.shape[1] == 2 and 'date' in str(row[0]).lower() and c_idx == 1
                        )

                        # 2. If it's a date field and the value is a number, convert it from Excel serial format.
                        if is_date_field and isinstance(value, (int, float)):
                            try:
                                # Excel's epoch starts on 1899-12-30. Day 1 is Jan 1 1900.
                                value = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)
                            except (ValueError, TypeError, OverflowError):
                                # If conversion fails (e.g., it's just a regular number), leave it as is.
                                pass

                        # 3. Apply the correct format based on the final type of the value.
                        cell_format = default_cell_format
                        if isinstance(value, (datetime.datetime, datetime.date)):
                            cell_format = date_format
                        elif col_name in ['Current Price', 'Market Cap', 'Total Revenue', 'open', 'high', 'low',
                                          'close', 'adjclose', 'Shares']:
                            cell_format = currency_format
                        elif col_name in ['Revenue Growth', 'Profit Growth', 'Dividend Yield', 'Growth', 'Holding']:
                            cell_format = percent_format

                        if pd.isna(value):
                            value = None  # xlsxwriter handles None as a blank cell

                        worksheet.write(r_idx, current_col + c_idx, value, cell_format)

                # --- Auto-fit Columns for this section ---
                for i, col_name in enumerate(df.columns):
                    col_idx = current_col + i
                    # Ensure data_width calculation handles non-string data gracefully
                    try:
                        data_width = df[col_name].astype(str).str.len().max()
                    except:
                        data_width = 10  # Fallback width
                    header_width = len(str(col_name))
                    width = max(data_width if not pd.isna(data_width) else 0, header_width)
                    worksheet.set_column(col_idx, col_idx, width + 2)  # Add padding

                # Update start column for next section
                current_col += df.shape[1] + 1  # +1 for separator column

            print(f"Step 3. Successfully created combined report: '{excel_file_path}'")

    except FileNotFoundError:
        print(f"Error: The file '{json_file_path}' was not found.")
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{json_file_path}'. Check the file format.")
    except Exception as e:
        print(f"An unexpected error occurred during Excel report creation: {e}")


# --- Read tickers from input.csv ---
try:
    input_df = pd.read_csv('input.csv')
    # Select the 'company_id' column and convert it to a list
    tickers = input_df['company_id'].tolist()
    print(f"Successfully loaded {len(tickers)} tickers from input.csv")
except FileNotFoundError:
    print("Error: 'input.csv' not found. Please create it with a 'company_id' column.")
    tickers = []  # Use an empty list to avoid crashing the script
except KeyError:
    print("Error: 'input.csv' must have a column named 'company_id'.")
    tickers = []  # Use an empty list to avoid crashing the script

for ticker_id in tickers:  # Loop through the tickers loaded from the CSV
    print(f"\n----- Processing Ticker: {ticker_id} -----")

    # Create a dedicated output folder for the ticker under a 'result' directory
    output_dir = os.path.join('result', ticker_id)
    os.makedirs(output_dir, exist_ok=True)

    # Define file paths within the new directory
    ext_json_path = os.path.join(output_dir, f"{ticker_id}_ext.json")
    final_excel_path = os.path.join(output_dir, f"{ticker_id}_Final.xlsx")

    # Step 1: Fetch and save company profile JSON
    extract(ticker_id, ext_json_path)

    # Step 2: Fetch all historical data as DataFrames
    print("\n--- Fetching Historical Data ---")
    hist_dfs = [
        ("Daily History", hist_daily(ticker_id)),
        ("Weekly History", hist_weekly(ticker_id)),
        ("Monthly History", hist_monthly(ticker_id)),
        ("Quarterly History", hist_quarterly(ticker_id))
    ]

    # Step 3: Combine JSON data and historical DataFrame into one final Excel file
    create_final_report(ext_json_path, hist_dfs, final_excel_path)
    print(f"----- Finished Processing {ticker_id} -----")