import json
from yahooquery import Ticker
import pandas as pd

# --------------------- put  ticker_id here -------------------------


def filter_company_info(ticker_id):
    tk = Ticker(ticker_id)
    full_info_json = tk.all_modules
    filtered_data = {}

    # ---------------------------------------------------------
    for ticker, data in full_info_json.items():
        # Safely get all necessary data modules at the beginning
        asset_profile = data.get('assetProfile', {})
        quote_type = data.get('quoteType', {})
        price = data.get('price', {})
        grow = data.get('earningsTrend',{})
        mutual_fund = data.get('fundOwnership', {})
        summary = data.get('summaryDetail', {})
        tot_rev = data.get('financialData',{})
        inside_tran = data.get('insiderTransactions', {})

        # ---------------------------------------------------------
        # Call the dedicated functions to build the output dictionary
        filtered_data = {
            "About": about(asset_profile, quote_type, price),
            "Finance": finance(price,tot_rev,summary),
            "Location": location(asset_profile),
            "Ownership Structure": "Public",
            "Growth": growth(grow),
            "Buyer Group": insider_Tran(inside_tran),
            "Mutual Fund Holders": mutual_fund_holders(mutual_fund)
        }

    return filtered_data
def about(asset_profile, quote_type, price):
    """
    Extracts and formats data for the 'About' section.
    """
    return {
        "ID": price.get("symbol", ""),
        "Name": price.get("longName", ""),
        "Unique ID": quote_type.get("uuid", ""),
        "Industry": asset_profile.get("industry", ""),
        "Sector": asset_profile.get("sector", ""),
        "Full Time employees": str(asset_profile.get("fullTimeEmployees", 0)),
        "Website": asset_profile.get("website", ""),
        "Investor Website": asset_profile.get("irWebsite", ""),
        "Exchange": quote_type.get("exchange", "")
    }

def finance(price,tot_rev,summary):
    """
    Extracts and formats data for the 'Finance' section.
    """
    return {
        "Date & Time": price.get("regularMarketTime", ""),
        "Current Price": f"{str(price.get("regularMarketPrice", 0)) + " " + price.get("currencySymbol", "")}",
        "Market Cap": f"{str(price.get("marketCap", 0))+" "+price.get("currencySymbol", "")}",
        "Total Revenue": f"{str(tot_rev.get("totalRevenue",0))+" "+price.get("currencySymbol", "")}",
        "Revenue Growth": f"{tot_rev.get("revenueGrowth",0) * 100:.2f} %",
        "Profit Growth": f"{tot_rev.get("earningsGrowth",0) * 100:.2f} %",
        "Dividend Rate": summary.get("dividendRate", 0),
        "Dividend Yield": summary.get("dividendYield", 0),
        "Date of Last Dividend": summary.get("exDividendDate", ""),
        "Five Years Average Dividend Yield": summary.get("fiveYearAvgDividendYield", ""),
        "Currency": price.get("currency", "")
    }

"""
def get_yearly_data(data):

    return {}
"""
def location(asset_profile):
    """
    Extracts and formats data for the 'Location' section.
    """
    return {
        "Address": asset_profile.get('address1', ''),
        "City": asset_profile.get("city", ""),
        "State": asset_profile.get("state", ""),
        "Country": asset_profile.get("country", ""),
        "Contact": asset_profile.get("phone",0)
    }

def growth(earnings_trend_data):
    # The list is under the 'trend' key in the 'earningsTrend' dictionary
    trend_list = earnings_trend_data.get("trend", [])
    filtered_growth_list = []
    for trend_item in trend_list:
        filtered_item = {
            "Period": trend_item.get("period", ""),
            "End Date": trend_item.get("endDate", ""),
            # "Growth": f"{trend_item.get("growth", 0)*100:.2f} %"  # Use None for numbers if missing
            "Growth": trend_item.get("growth", 0) # Use None for numbers if missing
        }
        filtered_growth_list.append(filtered_item)

    return filtered_growth_list

def revenue_growth(earnings_trend_data):
    trend_list = earnings_trend_data.get("trend", [])
    filtered_growth_list = []



"""
def revenue(revenue_data):
       # Get the raw list from the "yearly" key.
    yearly_list_raw = revenue_data.get("yearly", [])

    filtered_yearly_list = []
    # Loop through each item in the raw list.
    for r_item in yearly_list_raw:
        # Create a new dictionary with your desired structure.
        filtered_item = {
            "Year": r_item.get("date", ""),
            "Revenue": r_item.get("revenue", 0),
            "Profitability": r_item.get("earnings", 0)
        }
        filtered_yearly_list.append(filtered_item)

    return filtered_yearly_list
"""
def insider_Tran(inside_tran):
    in_ppl = inside_tran.get("transactions",[])
    inside_list = []

    for lt in in_ppl:
        filtered_item = {
            "Name": lt.get("filerName",""),
            "Relation" : lt.get("filerRelation",""),
            "Shares": str(lt.get("shares", 0)),
            "Description": lt.get("transactionText",""),
            "Date": lt.get("startDate","")
        }
        inside_list.append(filtered_item)
    return inside_list[:5]

def mutual_fund_holders(ownership_data):
    # if "fundOwnership" not in ownership_data:
    #     return []

    ownership_list = ownership_data.get("ownershipList", [])
    formatted_list = []

    for holder in ownership_list:
        formatted_holder = {
            "Date": holder.get("reportDate", ""),
            "Name": holder.get("organization", ""),
            "Holding": holder.get("pctHeld", 0.0),
            "Shares": holder.get("position", 0)
        }
        formatted_list.append(formatted_holder)

    return formatted_list

# ---------------------------------------------------------------------------

# Example Usage
# if __name__ == "__main__":
#     list = ["ORCL"]
#     for i in list:
#         ticker_symbol = i
#         company_data = filter_company_info(ticker_symbol)
#         output = f"{ticker_symbol}_output_file_2_new.json"
#         if company_data:
#             print(json.dumps(company_data, indent=4))
#             with open(output, 'w', encoding='utf-8') as jsonfile:
#                 json.dump(company_data, jsonfile, indent=4)
#             print(f"\n Data collection complete. Results saved to '{output}'.")

# list = ["ORCL"]
# for i in list:
#     extract(ticker_symbol)

# ---------------------------------------------------------------------------

def extract(ticker_symbol,output):
    company_data = filter_company_info(ticker_symbol)
    out_path = output
    if company_data:
        # print(json.dumps(company_data, indent=4))
        with open(out_path, 'w', encoding='utf-8') as jsonfile:
            json.dump(company_data, jsonfile, indent=4)
        print(f"\nStep 1. Data collection complete of {ticker_symbol}. Results saved to '{output}'")
