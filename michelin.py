#Import Important Libraries

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (NoSuchElementException, StaleElementReferenceException, TimeoutException)
import pandas as pd
import json
from time import sleep
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.action_chains import ActionChains
import gspread
from google.oauth2.service_account import Credentials

# Open the Chrome browser to Automate the Data Extraction process

options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(options=options)
url = 'https://guide.michelin.com/en/hotels-stays/italy?'
driver.get(url)
driver.maximize_window()

# Create Empty Lists to store data

hotel_links = []
name = []
country = []
city = []
hotel_url = []
description = []
photo_1 = []
photo_2 = []
photo_3 = []


while True:
    sleep(5)
    xpath = '//a[@class="link"]'
    hLinks = driver.find_elements(By.XPATH, xpath)
    for hLink in hLinks:
        hotel_links.append(hLink.get_attribute('href'))
    
    try:
        next_page = driver.find_element(By.XPATH, '(//i[@class="icon fal fa-angle-right"])[1]')
        driver.execute_script("arguments[0].click();", next_page)
    except NoSuchElementException:
        break

print('Total Number of Link:',len(hotel_links))
for index, url in enumerate(hotel_links):
    print(f"Processing Links No : {index}")
    driver.get(url)
    sleep(3)

    xpath = '//script[@type="application/ld+json"]'
    script_tag = driver.find_element(By.XPATH, xpath)
    json_data = script_tag.get_attribute("textContent")
    result = json.loads(json_data)

    try:
        name_text = result['name']
        name_split = name_text.split(' in ')[0]
        name.append(name_split)
    except (KeyError,IndexError):
        name.append("N/A")

    try:
        country_text = result['address']['addressCountry']
        if country_text == 'ITA':
            text = "Italy"
            country.append(text)
        else:
            country.append(result['address']['addressCountry'])
    except KeyError:
        country.append('N/A')

    try:
        city.append(result['address']['addressLocality'])
    except KeyError:
        city.append("N/A")

    try:
        hotel_url.append(result['url'])
    except KeyError:
        hotel_url.append("N/A")

    try:
        photo_1.append(driver.find_element(By.XPATH,'(//div[@class="masthead__gallery-image-item lazy entered loaded adjusted"])[1]').get_attribute('data-bg'))
    except NoSuchElementException:
        photo_1.append('N/A')

    try:
        photo_2.append(driver.find_element(By.XPATH,'(//div[@class="masthead__gallery-image-item lazy entered loaded adjusted"])[2]').get_attribute('data-bg'))
    except NoSuchElementException:
        photo_2.append('N/A')

    try:
        photo_3.append(driver.find_element(By.XPATH,'(//div[@class="masthead__gallery-image-item lazy entered loaded adjusted"])[3]').get_attribute('data-bg'))
    except NoSuchElementException:
        photo_3.append('N/A')

    try:
        description.append(driver.find_element(By.XPATH, '//div[@class="hotelpage__block--description js-hotel__content-translate"]').text)
    except NoSuchElementException:
        description.append('N/A')

# Create the Data Frame to Save the results

df = pd.DataFrame(zip(name,country,city, hotel_url, description,photo_1,photo_2,photo_3), columns=['Name of Hotel','Country', 'City', 'Hotel URL', 'Description Text', 'Photo 1', 'Photo 2', 'Photo 3'])
file_path = os.path.join(r"C:\Users\Mauz Khan\Desktop\Michelin.xlsx")
df.to_excel(file_path, index=False)

# Export the data to Google Sheets

creds = Credentials.from_service_account_file(r"C:\Users\Mauz Khan\Desktop\vscode\GS API.json",
                                              scopes=["https://spreadsheets.google.com/feeds",
                                                      "https://www.googleapis.com/auth/drive"])

try:
    client = gspread.authorize(creds)

    spreadsheet_id = '1-Yi0BNz6XXE_aqVZ_0v-8EQvFCsvLYWDeA'
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.get_worksheet(0)
    df = df.astype(str)
    worksheet.append_rows(df.values.tolist(), value_input_option='USER_ENTERED')
    
except ValueError as ve:
    print("ValueError:", ve)
except gspread.exceptions.APIError as api_error:
    print("API Error:", api_error)
except Exception as e:
    print("An unexpected error occurred:", e)
