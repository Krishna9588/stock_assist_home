import json
import time
import pandas as pd
import os
from extract import *
from historical_data import *

def convert_json_to_multi_sheet_excel(json_file_path, excel_file_path):

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

        print(f"Step 3. Successfully converted '{json_file_path}' to a multi-sheet Excel file.")

    except FileNotFoundError:
        print(f"Error: The file '{json_file_path}' was not found.")
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{json_file_path}'. Check the file format.")
    except Exception as e:
        print(f"An unexpected error occurred during Excel conversion: {e}")


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
    hist_excel_path = os.path.join(output_dir, f"{ticker_id}_hist.xlsx")
    ext_excel_path = os.path.join(output_dir, f"{ticker_id}_ext.xlsx")

    # Fetch and save data
    extract(ticker_id, ext_json_path)
    hist(ticker_id, hist_excel_path)

    time.sleep(2)
    convert_json_to_multi_sheet_excel(ext_json_path, ext_excel_path)
    print(f"----- Finished Processing {ticker_id} -----")