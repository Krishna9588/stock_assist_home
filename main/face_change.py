import pandas as pd
from yahooquery import Ticker

id = "orcl"
tickers = Ticker(id, asynchronous=True)
df = tickers.history(interval='1wk', start='2025-01-01', end='2025-08-01')

print(df.to_string())


df.to_csv(f'{id}_history.csv')

df.to_json(f'{id}_history.json', indent=4)

df.to_excel(f'{id}_history.xlsx')

print(f"\nDataFrame for {id} has been saved to:")
print(f"- {id}_history.csv")
print(f"- {id}_history.json")
print(f"- {id}_history.xlsx")