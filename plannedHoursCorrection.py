# plannedHoursCorrection
# 
# This script intends to check work order planned hours and compare them to actual hours
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
from docx import Document
import re
import time
import datetime
import xlwings as xw

# Main function 
def main():
    #run maximo_navigation
    browser = webdriver.Edge()
    print("Web browser initiated...\n")
    maximo_navigation(browser)
    
    #error search - separate function?
    #run excel_analysis
    
def maximo_navigation(browser):
    actions = ActionChains(browser) 
    wait = WebDriverWait(browser, 20)
    browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
    browser.maximize_window()

    # Extract login information from text file
    credentials = {}
    with open('config.txt', 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            credentials[key] = value

    # Enter login information
    UserElem = wait.until(EC.element_to_be_clickable((By.ID, "username")))
    UserElem.send_keys(credentials['username'])

    Cont1Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
    Cont1Elem.click()

    passElem = wait.until(EC.element_to_be_clickable((By.ID, "password")))
    passElem.send_keys(credentials['password'])

    Cont2Elem = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/div[2]/div[2]/form/div[3]/button")
    Cont2Elem.click()

    # Navigate to Work Order Tracking
    iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[7]/iframe")))
    browser.switch_to.frame(iframe)
    wotrackElem = wait.until(EC.element_to_be_clickable((By.ID, "FavoriteApp_WOTRACK")))
    actions.move_to_element_with_offset(wotrackElem, 5, 5).click().perform()

    # Specify information under More Search Fields
    search_frame = wait.until(EC.presence_of_element_located((By.XPATH,
    #history, target start, target finish
    #copy and paste job plan
    #error encoding

def excel_analysis():

main()
#Go to WO Tracking
#Go to More Search Fields
#Set History to "Y"
#Set Target Start to First of a month
#Set Target Finish to Last day of a month
# Target dates can be manually entered
#Normal distribution of difference in work hours
#Job plan copy and search again
#Plans contain hours
#Actual contains real work hours done

#CONSIDERATIONS
#Manage repeats, if information is identical
