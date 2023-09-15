import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import *
from openpyxl import Workbook
from time import sleep
import numpy as np
import xlsxwriter

def read_input(filepath):
    df = pd.read_excel(filepath)
    print(df.columns[0])
    url_array = []
    for i in range(0, len(df.to_numpy())):
        for j in range(len(df.to_numpy()[i])):
            url_array.append(df.to_numpy()[i][j])
    return url_array

def scrapping_data(url_array):
    result_array = []
    workbook = xlsxwriter.Workbook("Output.xlsx")
    worksheet = workbook.add_worksheet()
    hyperlink_format = workbook.add_format({'color': 'blue', 'underline': 1})
    for i in range(len(url_array)):
        driver = webdriver.Chrome()
        driver.maximize_window()
        driver.get(url_array[i])
        result_array.append([url_array[i], "", "", "", ""])
        table_element = driver.find_element(By.XPATH, "//*[@id='newsList']")
        rows = table_element.find_elements(By.TAG_NAME, "tr")
        for j in range(1, len(rows)):
            columns = rows[j].find_elements(By.TAG_NAME, "td")
            title = columns[3].text
            symbol = columns[4].text
            company = columns[5].text
            headline_element = columns[3].find_elements(By.TAG_NAME, "a")
            # headline_link = headline_element[1].get_attribute("href")
            headline_link = headline_element[0].get_attribute("href") if headline_element else ""
            symbol_element = columns[4].find_elements(By.TAG_NAME, "a")
            # symbol_link = symbol_element[1].get_attribute("href")
            symbol_link = symbol_element[0].get_attribute("href") if symbol_element else ""
            company_element = columns[5].find_elements(By.TAG_NAME, "a")
            # company_link = company_element[1].get_attribute("href")
            company_link = company_element[0].get_attribute("href") if company_element else ""
            if any(keyword in title for keyword in ["beneficial owner", "Holding(s) in Company", "Purchase of Shares", "Director/PDMR Shareholding", "Holding in Company", "Beneficial Owner"]):
                result_array.append(["", columns[0].text, [title, headline_link], [symbol, symbol_link], [company, company_link]])
                print(title)
            else:
                continue
        driver.quit()
    for i in range(len(result_array)):
        if result_array[i][0] != "":
            worksheet.write_url("A"+str(i+1), result_array[i][0], hyperlink_format, result_array[i][0])
            worksheet.write("B"+str(i+1), "")
            worksheet.write("C"+str(i+1), "")
            worksheet.write("D"+str(i+1), "")
            worksheet.write("E"+str(i+1), "")
        else:
            worksheet.write("A"+str(i+1), "")
            worksheet.write("B"+str(i+1), result_array[i][1])
            worksheet.write_url("C"+str(i+1), result_array[i][2][1], hyperlink_format, result_array[i][2][0])
            worksheet.write_url("D"+str(i+1), result_array[i][3][1], hyperlink_format, result_array[i][3][0])
            worksheet.write_url("E"+str(i+1), result_array[i][4][1], hyperlink_format, result_array[i][4][0])
    workbook.close()
       
        


if __name__ == "__main__":
    filepath = read_input("ExampleInput.xlsx")
    scrapping_data(filepath)