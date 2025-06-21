import csv, time, os
import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions

import warnings
warnings.filterwarnings('ignore')

agent = 'Firefox'
headless = True
retries = 2

title_filter = "-"

# open the excel file
wb = openpyxl.load_workbook('xxx.xlsx')
ws = wb.get_sheet_by_name('Future Calls')

# load URLs from the excel file
urls = []
for cell in ws['A']:
    if isinstance(cell.value, str) and title_filter in cell.value:
        try:    urls.append([cell.hyperlink.target, cell.value, ''])
        except: urls.append(['', cell.value, ''])

# iterate over the URLs and update the list with search announcements
for item in urls:
    url = item[0]
    item[2] = '-'

    if url != '':
        for attempt in range(retries):

            if agent == 'Firefox':
                options = FirefoxOptions()
                if headless:    options.add_argument("--headless")
                driver = webdriver.Firefox(options=options)
            if agent == 'Chrome':
                options = ChromeOptions()
                if headless:    options.add_argument("--headless")
                driver = webdriver.Chrome(executable_path="E:\\Desktop\\Dropbox\\Playground\\Python\\ParseEC\\chromedriver.exe")
            driver.get(url)
            pageSource = driver.page_source
            print(pageSource)
            pageSource = driver.execute_script("return document.documentElement.outerHTML;")
            print(pageSource)
            # wait = WebDriverWait(driver, 10)
            # wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, 'Contact')))
            time.sleep(8)

            # driver.execute_script('alert("test");')
            # driver.execute_script('var k = 5;')
            driver.execute_script(''
                                  'var k=document.querySelector("body > sedia-opportunities-v10 > ng-component > eui-page > eui-page-content > eui-page-columns > eui-page-column:nth-child(2) > div > eui-page-column-body > div:nth-child(5) > eui-card > mat-card > eui-card-content > mat-card-content > div > div.font-weight-bold.pr-4.eui-u-font-size-4xl");'
                                  'alert(k);'
                                  '')

            time.sleep(8)

            # search = driver.find_element_by_xpath('/html/body/div[1]/p/a[4]/span')
            # search = driver.find_element_by_css_selector('body > div:nth-child(5) > p:nth-child(2) > a:nth-child(2)')
            # actions = ActionChains(driver)
            # actions.move_to_element(search).perform()

            if item[2] == '-':
                try:
                    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/sedia-opportunities-v10/ng-component/eui-page/eui-page-content/eui-page-columns/eui-page-column[2]/div/eui-page-column-body/div[5]/eui-card/mat-card/eui-card-content/mat-card-content/div/div[1]')))
                    item[2] = element.text
                except:
                    #
                    item[2] = '-'

            if item[2] == '-':
                try:
                    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/sedia-opportunities-v10/ng-component/eui-page/eui-page-content/eui-page-columns/eui-page-column[2]/div/eui-page-column-body/div[5]/eui-card/mat-card/eui-card-content/mat-card-content/div/div[1]')))
                    item[2] = element.text
                except:
                    #
                    item[2] = '-'

            driver.quit()
            # driver.close()

            if item[2] != '-': break

    print(item[1], '|', item[2])

# print the output in a way that can be pasted to the relevant column of the sheet
print(' ')
for item in urls:
    #
    print(item[2])

