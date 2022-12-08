#importing libraries
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import pandas as pd
import json
import os
from bs4 import BeautifulSoup

# Fetch the keywords from the googlesheet
creds = ServiceAccountCredentials.from_json_keyfile_name("D:\PycharmProjects\Scrape-Google-Search-Results\data\secret_key.json")
file = gspread.authorize(creds)
workbook = file.open("google_search")
sheet = workbook.sheet1
keywords = []
for cell in sheet.range('B2:B5'):
    keywords.append(cell.value)

for keyword in keywords:
    print(keyword)

# setting the webdriver for chrome
service = Service(executable_path="C:\Development\chromedriver.exe")
options = Options()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=service, options=options)
driver.get("https://www.google.com/")
result = []
final_result = []
final_data = {}
for i in range(len(keywords)):
    search_bar = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input")
    search_bar.send_keys(keywords[i])
    search_bar.send_keys(Keys.ENTER)
    sleep(3)
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'html.parser')

    header_contents = soup.find_all('div', class_="yuRUbf")
    # Scraping process
    for content in header_contents:
        title = content.find('h3').text
        link = content.find('a')['href']
        data = {
            "title": title,
             "link": link,
        }
        result.append(data)
    final_data[keywords[i]] = result
    final_result.append(final_data)
    result = []
    final_data = {}
    driver.get("https://www.google.com/")

print(len(final_result))
print(final_result)
try:
    os.mkdir('json_result')
except FileExistsError:
    pass
with open('json_result/final_data.json', "w+") as json_data:
    json.dump(final_result, json_data)
print('json created')

# create csv
df = pd.DataFrame(final_result)
df.to_csv('googlesearch_data.csv', index=False)
df.to_excel('googlesearch_data.xlsx', index=False)
print("Data created success")
print("Total rows", len(final_result))
