# Live Cryptocurrency Market Analysis & Real-Time Excel Dashboard

## 🚀 Overview
This data science project fetches live cryptocurrency data using the CoinGecko API, performs real-time analysis, and updates an Excel sheet every 5 minutes. The project aims to identify key market trends, including:

- Top 5 cryptocurrencies by market capitalization
- Average price of the top 50 cryptocurrencies
- Highest and lowest price change percentage in the last 24 hours

## 📊 Features
✅ **Live Data Fetching** – Retrieves top 50 cryptocurrencies by market cap in real-time  
✅ **Automated Analysis** – Identifies key insights such as price movements and market trends  
✅ **Excel Integration** – Updates an Excel file every 5 minutes with fresh data  
✅ **Logging** – Prints timestamps of each fetch operation for tracking  

## 🛠️ Tech Stack
- **Python** (requests, pandas, time, datetime, xlsxwriter)
- **API Source:** CoinGecko
- **Data Storage:** Excel Spreadsheet

## 📥 Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/rushikeshya/Live-Cryptocurrency-Market-Analysis-Real-Time-Excel-Dashboard.git
   cd Live-Cryptocurrency-Market-Analysis-Real-Time-Excel-Dashboard
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the script:**
   ```bash
   python Python_Script.py
   ```

## 📌 Usage
Once executed, the script will:
1. Fetch live cryptocurrency data every 5 minutes
2. Analyze and display market insights in the terminal
3. Save the latest data into an Excel sheet (`Crypto_Live_Data.xlsx`)

## 📊 Example Output
```
Fetching data at 2025-03-03 12:00:00
Top 5 Cryptos by Market Cap:
1. Bitcoin
2. Ethereum
3. XRP
4. Tether
5. BNB

Highest 24h Change: Cardano (+43.01%)
Lowest 24h Change: Litecoin (-4.77%)
Excel sheet updated successfully!
```

## 🔍 Future Improvements
- Add data visualization (graphs & charts)
- Develop a Flask API to serve real-time data
- Implement machine learning models for crypto trend prediction

## 🤝 Contributing
Feel free to fork this project, submit issues, and create pull requests!  

## 📜 License
This project is licensed under the MIT License.

---
**⭐ If you found this project useful, consider giving it a star on GitHub! ⭐**
