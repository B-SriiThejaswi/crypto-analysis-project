import requests
import pandas as pd
from datetime import datetime
import time

# Constants
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
CURRENCY = "usd"
LIMIT = 50  # Top 50 cryptocurrencies
EXCEL_FILE = "Live_Crypto_Data.xlsx"
SHEET_NAME = "Crypto Data"

def fetch_data():
    """Fetches live cryptocurrency data from CoinGecko."""
    params = {
        "vs_currency": CURRENCY,
        "order": "market_cap_desc",  # Sort by market cap
        "per_page": LIMIT,
        "page": 1
    }
    try:
        response = requests.get(API_URL, params=params)
        response.raise_for_status()  # Raise an exception for invalid responses
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return []

def process_data(data):
    """Processes and extracts the required data from the API response."""
    processed_data = []
    for item in data:
        processed_data.append({
            "Name": item["name"],
            "Symbol": item["symbol"],
            "Current Price (USD)": item["current_price"],
            "Market Cap (USD)": item["market_cap"],
            "24h Trading Volume (USD)": item["total_volume"],
            "24h Price Change (%)": item["price_change_percentage_24h"]
        })
    return pd.DataFrame(processed_data)

def analyze_data(df):
    """Performs analysis on the cryptocurrency data."""
    # Top 5 Cryptocurrencies by Market Cap
    top_5 = df[["Name", "Market Cap (USD)"]].sort_values(by="Market Cap (USD)", ascending=False).head(5)
    
    # Average price of the top 50 cryptocurrencies
    avg_price = df["Current Price (USD)"].mean()
    
    # Highest and Lowest 24-hour Price Change
    highest_change = df.loc[df["24h Price Change (%)"].idxmax()]
    lowest_change = df.loc[df["24h Price Change (%)"].idxmin()]

    # Print analysis
    print("\nTop 5 Cryptocurrencies by Market Cap:")
    print(top_5)
    print(f"\nAverage price of the top 50 cryptocurrencies: ${avg_price:.2f}")
    print(f"\nHighest 24h Change: {highest_change['Name']} with {highest_change['24h Price Change (%)']:.2f}%")
    print(f"Lowest 24h Change: {lowest_change['Name']} with {lowest_change['24h Price Change (%)']:.2f}%")
    
    return top_5, avg_price, highest_change, lowest_change

def update_excel(df):
    """Updates the Excel file with the live data."""
    try:
        # Use openpyxl to handle Excel file updates
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            # Check if the sheet exists, and if it does, remove it before writing new data
            try:
                # Open the Excel file to check if the sheet exists
                existing_data = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
                if existing_data is not None:
                    print(f"Sheet '{SHEET_NAME}' exists, replacing it with new data.")
            except ValueError:
                print(f"Sheet '{SHEET_NAME}' does not exist, creating a new one.")
            
            # Write the new data to the Excel sheet
            df.to_excel(writer, index=False, sheet_name=SHEET_NAME)
            print(f"Updated Excel file: {EXCEL_FILE}")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    """Main function to fetch, process, analyze, and update live data."""
    while True:
        print(f"Fetching live data at {datetime.now()}")
        data = fetch_data()
        if data:
            df = process_data(data)
            top_5, avg_price, highest_change, lowest_change = analyze_data(df)
            update_excel(df)
        else:
            print("No data to update.")
        
        # Wait for 5 minutes before fetching new data
        time.sleep(300)

if __name__ == "__main__":
    main()
