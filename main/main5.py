import json
import time
import pandas as pd
import os
from extract import *
from historical_data import *


def create_final_report(json_file_path, hist_df, excel_file_path):
    """
    Creates a single Excel report with all data sections combined onto one sheet,
    separated by empty columns.
    """
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # A list to hold all the individual DataFrames for each section
        section_dfs = []

        # Define the order of sections for a consistent layout
        section_order = ["About", "Finance", "Location", "Growth", "Buyer Group", "Mutual Fund Holders",
                         "Ownership Structure"]

        for key in section_order:
            if key not in data:
                continue

            value = data[key]
            df = None

            if isinstance(value, dict):
                # Convert dictionary to a DataFrame, naming columns after the section
                df = pd.DataFrame.from_dict(value, orient='index', columns=[f'{key} Value'])
                df.index.name = key
                df = df.reset_index()
            elif isinstance(value, list) and value:
                # Convert list of dictionaries to a DataFrame
                df = pd.DataFrame(value)
            elif isinstance(value, (str, int, float)):
                # For simple key-value pairs, create a small DataFrame
                df = pd.DataFrame({key: [value]})

            if df is not None and not df.empty:
                section_dfs.append(df)

        # Add the historical data DataFrame
        if hist_df is not None and not hist_df.empty:
            # Reset index to make 'date' and 'symbol' columns
            hist_df_processed = hist_df.reset_index()
            section_dfs.append(hist_df_processed)

        # Combine all DataFrames horizontally with an empty column as a separator
        final_df_parts = []
        for i, df in enumerate(section_dfs):
            df_reset = df.reset_index(drop=True)
            final_df_parts.append(df_reset)
            if i < len(section_dfs) - 1:
                final_df_parts.append(pd.DataFrame({' ': ['']}))  # Empty separator column

        combined_df = pd.concat(final_df_parts, axis=1)
        combined_df.to_excel(excel_file_path, sheet_name='Combined Report', index=False)

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

    # Step 2: Fetch historical data as a DataFrame
    hist_df = hist(ticker_id)

    # Step 3: Combine JSON data and historical DataFrame into one final Excel file
    create_final_report(ext_json_path, hist_df, final_excel_path)
    print(f"----- Finished Processing {ticker_id} -----")