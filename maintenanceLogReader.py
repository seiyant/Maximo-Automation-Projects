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

# Define function to read PDF documents
# Define function to fetch Maximo status using Selenium
# Define function to write to Excel
# Main function or script execution
# Close Selenium driver and save Excel file