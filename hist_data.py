# from yahooquery import Ticker
# import pandas as pd
#
# def hist_daily(ticker_id):
#
#     try:
#         tickers = Ticker(ticker_id, asynchronous=True)
#         df = tickers.history(period='1mo')
#         print(f"\nStep 2. Daily historical data collection complete for {ticker_id}.")
#         if not df.empty:
#             # This is the key change: move 'symbol' and 'date' from the index to columns
#             df = df.reset_index()
#             # df = df.tail(11).iloc[:-1][::-1]
#             df = df.tail(11).iloc[:-1]
#             # This line is now enabled to add the 'Frequency' column
#             df.insert(0, 'Frequency', 'Daily')
#         return df
#     except Exception as e:
#         print(f"An error occurred while fetching daily historical data for {ticker_id}: {e}")
#         return pd.DataFrame()
#
# # The test code below can be removed for the final script
# # tk_id = "orcl"
# # a = hist_daily(tk_id)
# # print(a)
#
# def hist_weekly(ticker_id):
#
#     try:
#         tickers = Ticker(ticker_id, asynchronous=True)
#         df = tickers.history(period = '3mo',interval='1wk')
#         print(f"Step 2. Weekly historical data collection complete for {ticker_id}.")
#         if not df.empty:
#             # This is the key change: move 'symbol' and 'date' from the index to columns
#             df = df.reset_index()
#             # df = df.tail(11).iloc[:-1][::-1]
#             df = df.tail(11).iloc[:-1]
#             # This line is now enabled to add the 'Frequency' column
#             df.insert(0, 'Frequency', 'Weekly')
#         return df
#     except Exception as e:
#         print(f"An error occurred while fetching weekly historical data for {ticker_id}: {e}")
#         return pd.DataFrame()
#
# def hist_monthly(ticker_id):
#
#     try:
#         tickers = Ticker(ticker_id, asynchronous=True)
#         df = tickers.history(period = '1y',interval='1mo')
#         print(f"Step 2. Monthly historical data collection complete for {ticker_id}.")
#         if not df.empty:
#             # This is the key change: move 'symbol' and 'date' from the index to columns
#             df = df.reset_index()
#             # df = df.tail(11).iloc[:-1][::-1]
#             df = df.tail(11).iloc[:-1]
#             # This line is now enabled to add the 'Frequency' column
#             df.insert(0, 'Frequency', 'Monthly')
#         return df
#     except Exception as e:
#         print(f"An error occurred while fetching monthly historical data for {ticker_id}: {e}")
#         return pd.DataFrame()
#
# def hist_quarterly(ticker_id):
#
#     try:
#         tickers = Ticker(ticker_id, asynchronous=True)
#         df = tickers.history(period = '5y',interval='3mo')
#         print(f"Step 2. Quarterly historical data collection complete for {ticker_id}.")
#         if not df.empty:
#             # This is the key change: move 'symbol' and 'date' from the index to columns
#             df = df.reset_index()
#             df = df.tail(11).iloc[:-1][::-1]
#             df = df.tail(11).iloc[:-1]
#             # This line is now enabled to add the 'Frequency' column
#             df.insert(0, 'Frequency', 'Quarterly')
#         return df
#     except Exception as e:
#         print(f"An error occurred while fetching quarterly historical data for {ticker_id}: {e}")
#         return pd.DataFrame()

from yahooquery import Ticker
import pandas as pd

def hist_daily(ticker_id):
    """Fetches the last 10 days of historical data."""
    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        df = tickers.history(period='1mo')
        print(f"\nStep 2. Daily historical data collection complete for {ticker_id}.")
        if not df.empty:
            # Move 'symbol' and 'date' from the index to columns
            df = df.reset_index()
            # Get the last 10 records and sort them with the newest date first
            df = df.tail(10).sort_index(ascending=False)
            # The 'Frequency' column is no longer added
            # df.insert(0, 'Frequency', 'Daily')
        return df
    except Exception as e:
        print(f"An error occurred while fetching daily historical data for {ticker_id}: {e}")
        return pd.DataFrame()

def hist_weekly(ticker_id):
    """Fetches the last 10 weeks of historical data."""
    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        df = tickers.history(period='3mo', interval='1wk')
        print(f"Step 2. Weekly historical data collection complete for {ticker_id}.")
        if not df.empty:
            # Move 'symbol' and 'date' from the index to columns
            df = df.reset_index()
            # Get the last 10 records and sort them with the newest date first
            df = df.tail(10).sort_index(ascending=False)
            # The 'Frequency' column is no longer added
            # df.insert(0, 'Frequency', 'Weekly')
        return df
    except Exception as e:
        print(f"An error occurred while fetching weekly historical data for {ticker_id}: {e}")
        return pd.DataFrame()

def hist_monthly(ticker_id):
    """Fetches the last 10 months of historical data."""
    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        df = tickers.history(period='1y', interval='1mo')
        print(f"Step 2. Monthly historical data collection complete for {ticker_id}.")
        if not df.empty:
            # Move 'symbol' and 'date' from the index to columns
            df = df.reset_index()
            # Get the last 10 records and sort them with the newest date first
            df = df.tail(10).sort_index(ascending=False)
            # The 'Frequency' column is no longer added
            # df.insert(0, 'Frequency', 'Monthly')
        return df
    except Exception as e:
        print(f"An error occurred while fetching monthly historical data for {ticker_id}: {e}")
        return pd.DataFrame()

def hist_quarterly(ticker_id):
    """Fetches the last 10 quarters of historical data."""
    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        df = tickers.history(period='5y', interval='3mo')
        print(f"Step 2. Quarterly historical data collection complete for {ticker_id}.")
        if not df.empty:
            # Move 'symbol' and 'date' from the index to columns
            df = df.reset_index()
            # Get the last 10 records and sort them with the newest date first
            df = df.tail(10).sort_index(ascending=False)
            # The 'Frequency' column is no longer added
            # df.insert(0, 'Frequency', 'Quarterly')
        return df
    except Exception as e:
        print(f"An error occurred while fetching quarterly historical data for {ticker_id}: {e}")
        return pd.DataFrame()