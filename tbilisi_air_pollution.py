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
xls_data_daily_folder_path = r"C:\Users\sleep\OneDrive\Documents\Projects\Data Analysis\tbilisi_air_pollution\xls_data_daily"


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
