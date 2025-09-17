from yahooquery import Ticker
import pandas as pd

def hist(ticker_id):
    """
    Fetches historical data for a given ticker and returns it as a pandas DataFrame.
    """
    try:
        tickers = Ticker(ticker_id, asynchronous=True)
        # Fetch weekly data for the specified date range
        df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')
        print(f"\nStep 2. Historical data collection complete for {ticker_id}.")
        return df
    except Exception as e:
        print(f"An error occurred while fetching historical data for {ticker_id}: {e}")
        # Return an empty DataFrame to prevent the main script from crashing
        return pd.DataFrame()

""" working
def hist(ticker_id, output):

    tickers = Ticker(ticker_id, asynchronous=True)
    df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')

    # Save the DataFrame directly to an Excel file
    try:
        df.to_excel(output, engine='xlsxwriter')
        print(f"\nStep 2. Data collection complete for {ticker_id}. Results saved to '{output}'.")
    except Exception as e:
        print(f"An error occurred while saving the Excel file: {e}")

"""