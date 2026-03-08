import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd

#create a webdriver instance
driver = webdriver.Chrome()

rows=[]

#Looping through the pages of the website 
for i in range(1,6):
    driver.get(f"https://www.bloompharmacy.com/collections/x-1?page={i}")
    time.sleep(1)
    #get the product name and the price of all products in each page using Class name
    products = driver.find_elements(By.CLASS_NAME,"product-item__title")
    prices = driver.find_elements(By.CLASS_NAME,"price")


    #looping on both the product name and prices
    for product, price in zip(products,prices):

        #remove any hidden text from the price to only get the number and the currency
        price = price.text.replace("Sale price","").replace("FROM","").replace("Regular price","").strip()
        
        #check if the name of the product exists to avoid having empty rows
        if product.text: 
            #divide the price to value number and currency
            currency, value = price.split()
            #change the value of price from string to float
            value = float(value.replace(",",""))
            #put the values in dictionary 
            row = {"product":product.text, "price":value,"currency":currency}
            #append the dictionary into list of dictionaries
            rows.append(row)

#Create dataframe to add to be ready for the excel sheet
df_rows= pd.DataFrame(rows)


with pd.ExcelWriter("products.xlsx", mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
    # Read existing data to find the last row
    start_row = writer.book['Sheet1'].max_row
    
    # Write new data starting from the next available row
    # header=False prevents re-writing the column names
    df_rows.to_excel(writer, index=False, header=False, startrow=start_row)

#close the instance
driver.quit()
