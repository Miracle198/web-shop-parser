# 📊 Automated Books Shop Parser

A professional, synchronous Python script designed to scrape and parse data from the "Books to Scrape" bookstore. It automatically extracts books' info and generates a beautifully styled Excel spreadsheet with advanced typography.

## 🛠️ Key Features
- **Accurate Data Extraction:** Uses `BeautifulSoup4` and `requests` to accurately pull book titles, prices, and stock availability from the web catalog.
- **Advanced Excel Formatting:** Uses `pandas` and `openpyxl` to build structured tables with custom headers (styled in corporate blue), auto-fitted column widths, frozen top panes for comfortable scrolling, and built-in interactive Excel data filters.
- **Encoding Safety:** Forces UTF-8 encoding to prevent broken currency symbols (preserving the exact "£" symbol without system glitches).

## 📦 Tech Stack
- Python 3.10+
- Requests
- BeautifulSoup4 (HTML Parser)
- Pandas & OpenPyXL (Excel Engines)

## 🔧 How to Run
1. Install all required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the main script from your terminal:
   ```bash
   python parser.py
   ```
3. Find your custom-formatted spreadsheet inside the newly generated `products.xlsx` file in the root directory!
