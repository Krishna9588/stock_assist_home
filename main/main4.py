import json
import time
import pandas as pd
import os
from extract import *
from historical_data import *

def create_final_report(json_file_path, hist_df, excel_file_path):
    """
    Creates a single, multi-sheet Excel report by combining company profile data
    from a JSON file and historical price data from a DataFrame.
    """
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            for key, value in data.items():
                # Excel sheet names must be <= 31 chars
                sheet_name = key[:31]

                if isinstance(value, dict):
                    # Convert dictionary to a DataFrame (Attribute | Value)
                    df = pd.DataFrame.from_dict(value, orient='index', columns=['Value'])
                    df.index.name = 'Attribute'
                    df.to_excel(writer, sheet_name=sheet_name)
                elif isinstance(value, list):
                    # Convert list of dictionaries to a DataFrame
                    if value:  # Ensure list is not empty
                        df = pd.DataFrame(value)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                elif isinstance(value, (str, int, float)):
                    # For simple key-value pairs, create a small sheet
                    df = pd.DataFrame({'Attribute': [key], 'Value': [value]})
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Part 2: Write the historical data DataFrame to its own sheet
            if hist_df is not None and not hist_df.empty:
                hist_df.to_excel(writer, sheet_name='Historical Data')

        print(f"Step 3. Successfully created final report: '{excel_file_path}'")

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