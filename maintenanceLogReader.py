# InstantMaximo.py
# 
# This script intends to check if a daily maintenance log status matches the Maximo status
#
# Author: Seiya Nozawa-Temchenko
##########################################################################################

# Load all relevant packages that were downloaded using pip
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyPDF2 import PdfReader
from docx import Document
import datetime as time
import xlwings as Excel

# Ensure browser version and web driver version match
browser = webdriver.Edge()
actions = ActionChains(browser) 

# Define wait
wait = WebDriverWait(browser, 20)

# Navigate to Maximo login page
browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   

# Maximize window
browser.maximize_window()

# Enter login information
UserElem = wait.until(EC.element_to_be_clickable((By.ID, "username")))
UserElem.send_keys('NOZASEIY')

Cont1Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
Cont1Elem.click()

passElem = wait.until(EC.element_to_be_clickable((By.ID, "password")))
passElem.send_keys('Roofing1SN!') 

Cont2Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
Cont2Elem.click()

# Navigate to Quick Reporting
iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[7]/iframe")))
browser.switch_to.frame(iframe)

# Define function to read Word documents
crew_details = []

def extract_data_from_docx(file_path):
    # Load docx file
    docx = Document(file_path)

    # Extract the date and format it
    date_cell = docx.cell(1, 1).text
    date_string = date_cell.split()
    date_object = time.strptime(date_string, '%b %d, %Y')
    format_date = date_object.strptime('%m/%d/%Y')

    # Extract the crew list
    for row in docx.rows[9:]:
        name = row.cells[0].text
        position = row.cells[1].text
        aid = row.cells[3].text

        if row.cells[2].text != "A":
            crew_details.append((name, position, aid))

# Define function to fetch Maximo status using Selenium
# Define function to write to Excel
# Main function or script execution
# Close Selenium driver and save Excel file
