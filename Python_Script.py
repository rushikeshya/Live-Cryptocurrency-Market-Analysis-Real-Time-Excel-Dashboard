import requests
import pandas as pd
import time
from datetime import datetime

# API URL for fetching top 50 cryptocurrencies from CoinGecko
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False,
}

# Function to fetch live cryptocurrency data
def fetch_crypto_data():
    print(f"Fetching data at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    response = requests.get(API_URL, params=PARAMS)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data:", response.status_code)
        return None

# Function to analyze cryptocurrency data
def analyze_crypto_data(data):
    df = pd.DataFrame(data, columns=["name", "symbol", "current_price_USD", "market_cap", "total_volume", "price_change_percentage_24h"])
    
    # Identifying top 5 cryptocurrencies by market capitalization
    top_5_cryptos = df.nlargest(5, 'market_cap')[["name", "market_cap"]]
    
    # Calculating the average price of the top 50 cryptocurrencies
    avg_price = df["current_price_USD"].mean()
    
    # Finding the highest and lowest 24-hour price change percentages
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax(), ["name", "price_change_percentage_24h"]]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin(), ["name", "price_change_percentage_24h"]]

    return top_5_cryptos, avg_price, highest_change, lowest_change, df

# Function to update the Excel sheet
def update_excel(dataframe):
    with pd.ExcelWriter("Crypto_Live_Data.xlsx", engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, sheet_name="Live Crypto Data", index=False)
        print("Excel sheet updated successfully!")

# Main loop to fetch, analyze, and update every 5 minutes
while True:
    crypto_data = fetch_crypto_data()
    if crypto_data:
        top_5, avg_price, high_change, low_change, df = analyze_crypto_data(crypto_data)

        # Print analysis results
        print("\nTop 5 Cryptos by Market Cap:\n", top_5)
        print("\nAverage Price of Top 50 Cryptos: $", round(avg_price, 2))
        print("\nHighest 24h Change:", high_change["name"], "(", round(high_change["price_change_percentage_24h"], 2), "% )")
        print("\nLowest 24h Change:", low_change["name"], "(", round(low_change["price_change_percentage_24h"], 2), "% )")

        # Update Excel
        update_excel(df)

    # Wait for 5 minutes before updating again
    time.sleep(300)
