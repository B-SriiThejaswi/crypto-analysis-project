Live Crypto Analysis Project
This project provides a real-time analysis of the top 50 cryptocurrencies by market capitalization. It fetches live data from a public API, performs basic analysis, and generates a live-updating Excel sheet. The analysis includes insights such as the top 5 cryptocurrencies by market cap, average price, and the highest and lowest 24-hour price changes.

Features:
i) Fetches real-time cryptocurrency data using a public API.
ii) Analyzes key metrics, including:
    a)Top 5 cryptocurrencies by market capitalization.
    b)Average price of the top 50 cryptocurrencies.
    c)Highest and lowest 24-hour percentage price changes.
iii) Outputs results to a live-updating Excel file.
iv) Includes a detailed analysis report.

Data Fields
The project fetches and processes the following data fields:
a)Cryptocurrency Name
b)Symbol
c)Current Price (USD)
d)Market Capitalization
e)24-Hour Trading Volume
f)24-Hour Price Change (Percentage)

Files in the Repository
crypto_data_fetcher.py: The Python script for fetching live data, performing analysis, and updating the Excel sheet.
Live_Crypto_Data.xlsx: The live-updating Excel file showing the latest cryptocurrency data and key metrics.
Crypto_Analysis_Report.docx: A comprehensive analysis report summarizing the project insights with charts and graphs.

Install Python
Windows:

Download the Python installer from the official Python website:
https://www.python.org/downloads/
Run the installer and make sure to check the box for "Add Python to PATH" before clicking "Install Now".
python -m pip install --upgrade pip
pip install pandas
pip install requests
pip install openpyxl
