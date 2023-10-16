# InstantMaximo.py
# 
# This script intends to shorten paper log entry time on Maximo
#
# Author: Seiya Nozawa-Temchenko
#################################################################

# Load all relevant packages that were downloaded using pip
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime, time
import xlwings as xw 

workOrders = []

# Get user input for ROLLBACK command
def inputRollback(prompt):
    value = input(prompt)

    if value == 'xx':
        raise ValueError("Rollback last work order")
    
    return value

# Improvement: add entry method for WPLAN and WSCHED in addition to RELEASED
# Handles work order and SKIP command from user input
def getOrderInfo():
    workOrder = {}
    labors = []
    materials = []
    workOrder_number = inputRollback("Enter work order string (DONE 'd', ROLLBACK 'xx'):")

    if workOrder_number == 'd':
        return 'x'

    workOrder["number"] = workOrder_number

    # Improvement: add the option to create new line in the entry
    workOrder_log = inputRollback("Enter work log details (SKIP 'x', ROLLBACK 'xx'):")
    if workOrder_log != 'x':
        workOrder["log"] = workOrder_log

    # Improvement: enter any name and using suggestion array, choose best option
    while True:
        workOrder_labor = inputRollback("Enter labor details (Name, Date, Hours) separated by commas (DONE 'd', SKIP 'x', ROLLBACK 'xx'):")
        if workOrder_labor == 'd':
            break
        elif workOrder_labor != 'x':
            name, date, hours = workOrder_labor.split(',')
            labor = {
                "name": name.strip(),
                "date": date.strip(),
                "hours": hours.strip()
            }
            labors.append(labor)
    workOrder["labors"] = labors

# Item is numbered, transaction is either ISSUE or RETURN
    while True:
        workOrder_material = inputRollback("Enter material details (Item, Transaction, Quantity) separated by commas (DONE 'd', SKIP 'x', ROLLBACK 'xx'):")
        if workOrder_material == 'd':
            break
        elif workOrder_material != 'x':
            item, transaction, quantity = workOrder_material.split(',')
            material = {
                "item": item.strip(),
                "transaction": transaction.strip(),
                "quantity": quantity.strip()
            }
            materials.append(material)
    workOrder["materials"] = materials

    workOrder_progress = inputRollback("Enter progress (DONE 'd' or IN-PROGRESS 'x', ROLLBACK 'xx'):")
    if workOrder_progress != 'x':
        workOrder["progress"] = 'DONE'

    print(workOrder) 
    return workOrder

# Loop creating a list of work orders from singular work orders
while True:
    try:
        print("\nStarting new work order input...\n")
        workOrder = getOrderInfo()

        if workOrder == 'x':
            break

    # Append each filled workOrder to list of workOrders
        workOrders.append(workOrder)

    # Trigger for ROLLBACK
    except ValueError as e:
        print("\nRolling back the last work order entry...\n")
        continue

# Ensure your browser version and web driver version match
browser = webdriver.Edge()
actions = ActionChains(browser) 

# Navigate to Maximo login page
browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo') 
wait = WebDriverWait(browser, 20)

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

qrepElem = wait.until(EC.element_to_be_clickable((By.ID, "FavoriteApp_QUICKREP")))
actions.move_to_element_with_offset(qrepElem, 5, 5).click().perform()

# Type and search for Work Orders
# Improvement: add entry method for WPLAN and WSCHED in addition to RELEASED
for workOrder in workOrders:
    searchWO_number = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tfrow_[C:1]_txt-tb")))
    searchWO_number.send_keys(workOrder["number"])
    searchWO_number.send_keys(Keys.ENTER)

    # Improvement: create failsafes no elements, go to Work Order Tracking
    first_workOrder = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]")))
    first_workOrder.click()

    if "log" in workOrder:
        addLog = wait.until(EC.element_to_be_clickable((By.ID, "mfda81699-hb_addrow")))
        addLog.click()
 
        addSummary = wait.until(EC.element_to_be_clickable((By.ID, "mb97068ca-tb")))
        addSummary.send_keys('Paper Log')

        addDetails = wait.until(EC.element_to_be_clickable((By.ID, "mce77585c-rte_iframe")))
        addDetails.send_keys(workOrder['log'])
        print(f"Work Order {workOrder['number']} Work Log Details entered...\n")

    # Navigate to Quick Reporting tab
    move2QR = wait.until(EC.element_to_be_clickable((By.ID, "m8ddd952b-tab")))
    move2QR.click()

    for labor in workOrder.get("labors",[]):
        go2Labor = wait.until(EC.element_to_be_clickable((By.ID, "mb6f8aa93_bg_button_addrow-pb_addrow_a")))
        go2Labor.click()

        addLabor = wait.until(EC.element_to_be_clickable((By.ID, "m5f97af0-tb")))
        addLabor.send_keys(labor['name'])
        garbage_value = wait.until(EC.element_to_be_clickable((By.ID, "m5f97af0-tb2")))
        garbage_value.click()
        print(f"Work Order {workOrder['number']} Laborer entered...")
        
        addDate = wait.until(EC.element_to_be_clickable((By.ID, "m336d0567-tb")))
        addDate.click()
        addDate.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
        addDate.send_keys(labor['date'])
        garbage_value = wait.until(EC.element_to_be_clickable((By.ID, "m5f97af0-tb2")))
        garbage_value.click()
        print(f"Work Order {workOrder['number']} Start Date entered...")
        
        addHours = wait.until(EC.element_to_be_clickable((By.ID, "ma3d218f6-tb")))
        addHours.click()
        addHours.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
        addHours.send_keys(labor['hours'])
        garbage_value = wait.until(EC.element_to_be_clickable((By.ID, "m5f97af0-tb2")))
        garbage_value.click()
        print(f"Work Order {workOrder['number']} Regular Hours entered...")
        print(f"Work Order {workOrder['number']} Labor entered...")
    
    # Improvement: create failsafe for when there is a mismatch in quantity
    for material in workOrder.get("materials",[]):
        go2Material = wait.until(EC.element_to_be_clickable((By.ID, "m5cf77a07_bg_button_addrow-pb_addrow_a")))
        go2Material.click()

        addItem = wait.until(EC.element_to_be_clickable((By.ID, "ma9a49433-tb")))
        addItem.send_keys(material['item'])

        addTransaction = wait.until(EC.element_to_be_clickable((By.ID, "m30adc589-tb")))
        addTransaction.send_keys(material['transaction'])

        addQuantity = wait.until(EC.element_to_be_clickable((By.ID, "ma012d818-tb")))
        addQuantity.send_keys(material['quantity'])

    if "progress" in workOrder:
        go2Progress = wait.until(EC.element_to_be_clickable((By.ID, "m8bb73832-pb")))
        go2Progress.click()
'''        
    #saveButton = wait.until(EC.element_to_be_clickable(By.ID, "toolactions_SAVE-tbb")))
    #saveButton.click()
    print(f"Work order {workOrder['number']} completed... Moving on")

    #backButton = wait.until(EC.element_to_be_clickable((By.ID, "toolactions_CLEAR-tbb")))
    #backButton.click()
 
browser.quit() ''' 
print("\n\nSeiya, out")
'''
    try: #contains(@id, 'NAME') #By.NAME, 
        second_workOrder = wait.until(EC.visibility_of_element_located((By.ID, "m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:1]")))
        print(f"\nMultiple work orders found for {workOrder['number']}. Skipping...\n")
        continue 
            
    except TimeoutException:
        try:
            first_workOrder = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]")))
            first_workOrder.click()
        
        except TimeoutException:
            print(f"\nNo work orders found for {workOrder['number']}. Skipping...\n")
            continue '''