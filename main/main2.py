import json
import time
import pandas as pd
from extract import *
from historical_data import *
import pandas as pd
import json
def convert_json_to_excel(json_file_path, excel_file_path):
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)


        if isinstance(data, dict):
            df = pd.DataFrame([data])
        else:
            df = pd.DataFrame(data)

        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        print(f"Successfully converted '{json_file_path}' to '{excel_file_path}'.")

    except FileNotFoundError:
        print(f"Error: The file '{json_file_path}' was not found.")
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{json_file_path}'. Check the file format.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


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

for i in tickers: # Loop through the tickers loaded from the CSV
    ext_path = f"{i}_ext.json"
    his_path = f"{i}_hist.xlsx"
    extract(i,ext_path)
    hist(i,his_path)
    # ---
    time.sleep(2)
    # ---
    ext_output_path_excel = f"{i}_ext.xlsx"
    # his_output_path_excel = f"{i}_hist.xlsx"
    print("Converting to excel")
    convert_json_to_excel(ext_path, ext_output_path_excel)
    # convert_json_to_excel(his_path, his_output_path_excel)
    print(f"\nStep 3. Excel Sheet created successfully: {ext_output_path_excel}")




""" old working
import json
import time
import pandas as pd
from extract import *
from historical_data import *
import pandas as pd
import json
def convert_json_to_excel(json_file_path, excel_file_path):
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)


        if isinstance(data, dict):
            df = pd.DataFrame([data])
        else:
            df = pd.DataFrame(data)

        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        print(f"Successfully converted '{json_file_path}' to '{excel_file_path}'.")

    except FileNotFoundError:
        print(f"Error: The file '{json_file_path}' was not found.")
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{json_file_path}'. Check the file format.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")



tickers = ["ORCL","BCS","MSFT","TATAMOTORS.NS","8058.T"]

for i in tickers:
    ext_path = f"{i}_ext.json"
    his_path = f"{i}_hist.xlsx"
    extract(i,ext_path)
    hist(i,his_path)
    # ---
    time.sleep(2)
    # ---
    ext_output_path_excel = f"{i}_ext.xlsx"
    # his_output_path_excel = f"{i}_hist.xlsx"
    print("Converting to excel")
    convert_json_to_excel(ext_path, ext_output_path_excel)
    # convert_json_to_excel(his_path, his_output_path_excel)
    print(f"\nStep 3. Excel Sheet created successfully: {ext_output_path_excel}")

"""