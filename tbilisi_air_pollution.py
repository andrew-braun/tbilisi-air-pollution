from typing import TextIO
from selenium import webdriver
from selenium.webdriver import Chrome as Chrome
import pandas as pd
import os
import shutil
import time
import datetime

# Webdriver points to the Selenium chromedriver file; driver initiates it
webdriver = r"C:\Users\sleep\Documents\Projects\Data Analysis\tbilisi_air_pollution\chromedriver.exe"
driver = Chrome(webdriver)

os.chdir(r"C:\Users\sleep\Documents\Projects\Data Analysis\tbilisi_air_pollution\xls_data")
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
            time.sleep(0.1)

        # Move the file to the project folder
        shutil.move(r"C:\Users\sleep\Downloads" + "\\" + "export.xlsx",
                    r"C:\Users\sleep\Documents\Projects\Data Analysis\tbilisi_air_pollution\xls_data" + "\\" + "export.xlsx")

        # Rename the file with the correct date (because date info isn't in the file)
        os.rename(r"C:\Users\sleep\Documents\Projects\Data Analysis\tbilisi_air_pollution\xls_data\export.xlsx",
                  r"C:\Users\sleep\Documents\Projects\Data Analysis\tbilisi_air_pollution\xls_data"
                  + "\\" + str(year) + "-" + str(month) + ".xlsx")
        date_list.append((year, month))








from selenium.webdriver import Firefox
from selenium import webdriver
import pandas as pd
import os
import datetime


os.chdir(r"C:\Users\sleep\Documents\Projects\Data Analysis\tbilisi_air_pollution")



webdriver = r"C:\Users\sleep\Documents\Projects\Data Analysis\tbilisi_air_pollution"

driver = Firefox(webdriver)

browser = webdriver.firefox(profile)

profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2) # custom location
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', '/tmp')
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')

for year in range(2016, 2019):
    for month in range(1, 13):
        url = "http://air.gov.ge/en/reports_page?station=AGMS%2CKZBG%2CVRKT%2CTSRT%2CTBL01&report_type=monthly&date_from=" + str(year) + "-" + str(month)

        browser.get(url)
        browser.find_element_by_id('xls').click()
