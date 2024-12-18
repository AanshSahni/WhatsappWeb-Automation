from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import pandas as pd
import time


# Load the Chrome driver
driver = webdriver.Chrome()

# Open WhatsApp URL in Chrome browser
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 20)
time.sleep(20)

# Read data from Excel
excel_data = pd.read_excel('Recipients data.xlsx', sheet_name='Recipients')


# Initialize counter
count = 0

# Iterate through Excel rows
for index, row in excel_data.iterrows():
    contact = row['Contact']
    name = row['Name']
    message = row['Message']
    
    # Locate search box using XPath
    search_box_class = '#side > div._ak9t > div > div._ai04 > div._ai05 > div > div.x1hx0egp.x6ikm8r.x1odjw0f.x6prxxf.x1k6rcq7.x1whj5v > p'
    search_box = wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, search_box_class))
    time.sleep(10)
    # Clear search box if any contact number is written in it
    search_box.clear()

    # Send contact number to search box
    search_box.send_keys(contact)

    # Wait for 3 seconds to allow the contact to be found
    time.sleep(10)

    
    # Format the message with the customer's name
    

    # Send the Enter key to select the contact
    search_box.send_keys(Keys.ENTER)
    contact_name_add = '#main > header > div._amie > div > div > div > span'
    contact_name = wait.until(lambda driver: driver.find_element(By.CSS_SELECTOR, contact_name_add))


    time.sleep(10)
    
    # Send the message using ActionChains
    actions = ActionChains(driver)
    actions.send_keys(str(message))
    actions.send_keys(Keys.ENTER)
    actions.perform()
    print(f"Message sent to {name}")
    time.sleep(10)

time.sleep(15)
# Close Chrome browser
driver.quit()
