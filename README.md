# scraping-online-store
A Python-based web scraper that automates the collection of product names and prices from a Pharmacy website. It navigates through multiple pages, cleans the pricing data, and appends the results to an existing Excel workbook.

## 🚀 Features
Automated Pagination: Iterates through the first 5 pages of the specified collection.

Data Cleaning: Strips unnecessary UI text (e.g., "Regular price", "Sale price") to extract pure numerical values.

Data Structuring: Splits price strings into separate value (float) and currency columns.

Excel Integration: Uses pandas to append new data to products.xlsx without overwriting existing entries.

## 🛠️ Prerequisites
Before running the script, ensure you have the following installed:

Python 3.x

Chrome Browser (The script uses webdriver.Chrome())

ChromeDriver: Ensure it matches your Chrome version (or use webdriver-manager).

### Required Libraries
Install the necessary Python packages via pip:

### Bash
pip install selenium pandas openpyxl
## 📖 How It Works
Initialization: Opens a Chrome instance via Selenium.

Navigation: Loops through pages 1 to 5 of the Bloom Pharmacy "x-1" collection.

Extraction:

Locates products by class product-item__title.

Locates prices by class price.

Transformation: * Converts the price string (e.g., "EGP 150.00") into a float.

Stores data in a list of dictionaries.

Output: Converts the list to a DataFrame and appends it to the bottom of Sheet1 in products.xlsx.

## ⚠️ Important Notes
Excel File: The script is set to mode='a' (append). Ensure a file named products.xlsx already exists with a "Sheet1" before the first run, or modify the script to create a new one.

Wait Times: A time.sleep(1) is included to allow elements to load. Depending on your internet speed, you might need to increase this or use WebDriverWait for better reliability.

Selectivity: The script only scrapes pages 1 through 5. You can adjust the range(1, 6) to scrape more or fewer pages.
