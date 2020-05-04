from typing import TextIO
from selenium import webdriver
from selenium.webdriver import Chrome as Chrome
import pandas as pd
import os
import shutil
import time
import datetime

# Data storage folders
xls_data_folder_path = r"C:\Users\sleep\OneDrive\Documents\Projects\Data Analysis\tbilisi_air_pollution\xls_data"

# Webdriver points to the Selenium chromedriver file; driver initiates it
webdriver = r"C:\bin\selenium-drivers\chromedriver.exe"
driver = Chrome(webdriver)

os.chdir(xls_data_folder_path)
date_list = []

#Loops over year 2016-current year, months 1-12
for year in range(2016, int(datetime.datetime.now().year) + 1):
    for month in range(1, 13):
        # Georgian government air website; different reports accessible by altering dates in the URL
        url = "http://air.gov.ge/en/reports_page?station=AGMS%2CKZBG%2CVRKT%2CTSRT%2CTBL01&report_type=monthly&date_from=" \
              + str(year) + "-" + str(month)

        # Open URL, find the XLS download button, and click it using JS
        driver.get(url)
        driver.find_element_by_id('xls').click()

        # Wait for the file to download so we can rename and move it
        while "export.xlsx" not in os.listdir(r"C:\Users\sleep\Downloads"):
            time.sleep(0.2)

        # Move the file to the project folder
        shutil.move(r"C:\Users\sleep\Downloads" + "\\" + "export.xlsx",
                    xls_data_folder_path + "\\" + "export.xlsx")

        # Rename the file with the correct date (because date info isn't in the file)
        os.rename(xls_data_folder_path + "\\export.xlsx",
                  xls_data_folder_path + "\\" + str(year) + "-" + str(month) + ".xlsx")
        date_list.append((year, month))

#Get daily data

xls_day_data_folder_path = r"C:\Users\sleep\OneDrive\Documents\Projects\Data Analysis\tbilisi_air_pollution\xls_day_data"

# Webdriver points to the Selenium chromedriver file; driver initiates it
webdriver = r"C:\bin\selenium-drivers\chromedriver.exe"
driver = Chrome(webdriver)

os.chdir(xls_day_data_folder_path)
day_date_list = []

#Loops over year 2016-current year, months 1-12, days in month
now = datetime.datetime.today()
start_date = datetime.date(2016, 10, 25)
end_date = datetime.date(int(now.year), int(now.month), int(now.day))
delta = datetime.timedelta(days=1)

while start_date <= end_date:
    url = "http://air.gov.ge/en/reports_page?station=BTUM%2CKZBG%2CVRKT%2CTBL01%2CTSRT%2CAGMS%2CRST18%2CKUTS&report_type=daily&date_from="\
            + str(start_date) + "&export_type=xlsx"

    day_date_list.append((start_date))

    # Open URL, find the XLS download button, and click it using JS
    driver.get(url)

    # if driver.find_element_by_id('report_page'):

    # Wait for the file to download so we can rename and move it
    time.sleep(0.2)
    if "export.xlsx" not in os.listdir(r"C:\Users\sleep\Downloads"):
        ms = 0
        while ms < 21:
            if "export.xlsx" not in os.listdir(r"C:\Users\sleep\Downloads"):
                time.sleep(0.1)
                ms += 1
                break

    if "export.xlsx" in os.listdir(r"C:\Users\sleep\Downloads"):
        # Move the file to the project folder
        shutil.move(r"C:\Users\sleep\Downloads" + "\\" + "export.xlsx",
                    xls_day_data_folder_path + "\\" + "export.xlsx")

        # Rename the file with the correct date (because date info isn't in the file)
        os.rename(xls_day_data_folder_path + "\\export.xlsx",
                  xls_day_data_folder_path + "\\" + str(start_date) + ".xlsx")

    start_date += delta
