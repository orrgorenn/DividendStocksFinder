import math
from io import StringIO
from typing import Optional

import pandas as pd
import datetime

import requests
from bs4 import BeautifulSoup

# Set you min and max values for the filters
MIN_DIV_YIELD = 1.53
MIN_DGR_1Y = 5
MIN_DGR_3Y = 5
MIN_DGR_5Y = 5
MIN_DGR_10Y = 5
MAX_DEBT_CAPITAL = 0.85
MAX_PE = 29
MAX_PAYOUT_RATIO = 65
MIN_PAST_5_YEARS = 0
MIN_NEXT_5_YEARS = 0
MIN_MARKET_CAP = 1  # billion
MAX_MONEY_TO_INVEST = 1000  # dollars

file_path = "stocks.xlsx"
df = pd.read_excel(file_path, sheet_name="Champions", header=2)

# Get the current year
current_year = datetime.datetime.now().year


def fetch_and_extract_table(symbol: str) -> Optional[str]:
    url = f"https://finviz.com/quote.ashx?t={symbol}"
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        table = soup.find("table", class_="snapshot-table2")
        return table
    else:
        print(f"Error with fetch data for {url}")
        return None


# Filter the DataFrame based on your criteria
filtered_df = df[
    (df["No Years"] > (current_year - 2000))
    & (df["Div Yield"] > MIN_DIV_YIELD)
    & (df["DGR 1Y"] > MIN_DGR_1Y)
    & (df["DGR 3Y"] > MIN_DGR_3Y)
    & (df["DGR 5Y"] > MIN_DGR_5Y)
    & (df["DGR 10Y"] > MIN_DGR_10Y)
    & (df["Debt/Capital"] < MAX_DEBT_CAPITAL)
    & (df["P/E"] < MAX_PE)
]

# Sort the filtered DataFrame by column E in descending order
filtered_df = filtered_df.sort_values(by="No Years", ascending=False)

# Convert the filtered DataFrame to a list of dictionaries
filtered_data = filtered_df.to_dict("records")

end_rows = []

# Print the resulting array of rows
for row in filtered_data:
    table = fetch_and_extract_table(row.get("Symbol").replace(".", ""))

    if table is not None:
        table_str = str(table)
        table_io = StringIO(table_str)

        df_table = pd.read_html(table_io)[0]
        try:
            payout = float(df_table[7][10].rstrip("%"))
            past_5_years = float(df_table[5][6].rstrip("%"))
            next_5_years = float(df_table[5][5].rstrip("%"))
            market_cap = float(df_table[1][1].rstrip("B"))
            if (
                payout < MAX_PAYOUT_RATIO
                and past_5_years > MIN_PAST_5_YEARS
                and next_5_years > MIN_NEXT_5_YEARS
                and market_cap > MIN_MARKET_CAP
            ):
                end_rows.append(row)
        except ValueError:
            continue

for row in end_rows:
    shares_to_buy = math.floor(MAX_MONEY_TO_INVEST / row.get("Price"))
    print(
        f"{row.get('Symbol')}: {row.get('Company')} - Buy {shares_to_buy} shares for a total of ${round(shares_to_buy* row.get('Price'), 2)} (${row.get('Price')} per share)"
    )
