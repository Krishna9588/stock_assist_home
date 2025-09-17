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


tickers = ["ORCL","BCS"]
for i in tickers:
    ext_path = f"{i}_ext.json"
    his_path = f"{i}_hist.json"
    extract(i,ext_path)
    hist(i,his_path)
    # ---
    # time.sleep(2)
    # ---
    # ext_output_path_excel = f"{i}_ext.xlsx"
    # his_output_path_excel = f"{i}_hist.xlsx"
    # print("Converting to excel")
    # convert_json_to_excel(ext_path, ext_output_path_excel)
    # convert_json_to_excel(his_path, his_output_path_excel)
    # print(f"\nStep 3. Excel Sheet created successfully: {ext_output_path_excel} and {his_output_path_excel}")
    # # combine_and_convert(ext_path, his_path, output_path_json, output_path_excel)
