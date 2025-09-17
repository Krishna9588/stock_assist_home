from yahooquery import Ticker
import pandas as pd

tk_id = "orcl"

def hist_daily(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period='1mo')
        print(f"\nStep 2. Historical data collection complete for {ticker_id}.")

        return df.tail(10).iloc[::-1]
    except Exception as e:
        print(f"An error occurred while fetching historical data for {ticker_id}: {e}")
        return pd.DataFrame()

a = hist_daily(tk_id)
print(a.head(2))

def hist_weekly(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period = '3mo',interval='1wk')
        print(f"\nStep 2. Historical data collection complete for {ticker_id}.")
        return df.tail(10).iloc[::-1]
    except Exception as e:
        print(f"An error occurred while fetching historical data for {ticker_id}: {e}")
        return pd.DataFrame()

def hist_monthly(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period = '1y',interval='1mo')
        print(f"\nStep 2. Historical data collection complete for {ticker_id}.")
        return df.tail(10).iloc[::-1]
    except Exception as e:
        print(f"An error occurred while fetching historical data for {ticker_id}: {e}")
        return pd.DataFrame()

def hist_quarterly(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period = '5y',interval='3mo')
        print(f"\nStep 2. Historical data collection complete for {ticker_id}.")
        return df.tail(10).iloc[::-1]
    except Exception as e:
        print(f"An error occurred while fetching historical data for {ticker_id}: {e}")
        return pd.DataFrame()

# backup - 2
"""
from yahooquery import Ticker
import pandas as pd

def hist_daily(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period='1mo')
        print(f"\nStep 2. Daily historical data collection complete for {ticker_id}.")
        if not df.empty:
            df = df.tail(11).iloc[:-1][::-1]
            df.insert(0, 'Frequency', 'Daily')
        return df
    except Exception as e:
        print(f"An error occurred while fetching daily historical data for {ticker_id}: {e}")
        return pd.DataFrame()

# tk_id = "orcl"
# a = hist_daily(tk_id)
# print(a)
def hist_weekly(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period = '3mo',interval='1wk')
        print(f"Step 2. Weekly historical data collection complete for {ticker_id}.")
        if not df.empty:
            df = df.tail(11).iloc[:-1][::-1]
            df.insert(0, 'Frequency', 'Weekly')
        return df
    except Exception as e:
        print(f"An error occurred while fetching weekly historical data for {ticker_id}: {e}")
        return pd.DataFrame()

def hist_monthly(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period = '1y',interval='1mo')
        print(f"Step 2. Monthly historical data collection complete for {ticker_id}.")
        if not df.empty:
            df = df.tail(11).iloc[:-1][::-1]
            df.insert(0, 'Frequency', 'Monthly')
        return df
    except Exception as e:
        print(f"An error occurred while fetching monthly historical data for {ticker_id}: {e}")
        return pd.DataFrame()

def hist_quarterly(ticker_id):

    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        df = tickers.history(period = '5y',interval='3mo')
        print(f"Step 2. Quarterly historical data collection complete for {ticker_id}.")
        if not df.empty:
            df = df.tail(11).iloc[:-1][::-1]
            df.insert(0, 'Frequency', 'Quarterly')
        return df
    except Exception as e:
        print(f"An error occurred while fetching quarterly historical data for {ticker_id}: {e}")
        return pd.DataFrame()

"""

