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
from docx import Document
import re
import time
import datetime
import xlwings as xw

# Main function 
def main():
    # Assumes maintenance log is in the same folder
    word_file_path = 'Nov 25, 2023 Maintenance Daily Log.docx'
    date, work_details = extract_data_from_doc(word_file_path)
    print(f"Total work details extracted: {len(work_details)}\n")
    excel_file_path = 'Maintenance Daily Log Checker.xlsx'
    excel_file_sheet = 'November 23'
    write_to_excel(date, work_details, excel_file_path, excel_file_sheet) # Load excel from here not inside
    browser = webdriver.Edge()
    print("Web browser initiated...\n")
    extract_maximo_status(browser, excel_file_path, excel_file_sheet)
    browser.quit()
    print("Process complete...\n")

# Extract the work details from the Word document
def extract_data_from_doc(file_path):
    # Load docx file
    doc = Document(file_path)
    print('Loaded Word document...\n')
    
    # Initialize an empty list to store the work details
    work_details = []

    # Find cell locations
    row_index = 0
    for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if (''.join((cell.text).split()).lower() == 'date') and (cell.text != last_cell_text):
                        date_row = row_index # Date is in the same row
                        print('Date:', date_row)
                    if (''.join((cell.text).split()).lower() == 'ertorlevel3firstaid') and (cell.text != last_cell_text):
                        crew_row = row_index + 1
                        print('Crew:', crew_row)
                    if (''.join((cell.text).split()).lower() == 'status') and (cell.text != last_cell_text):
                        jobs_row = row_index + 1
                        print('Jobs:', jobs_row)
                    last_cell_text = cell.text
                row_index += 1

    # Extract the date and format it
    for cell in doc.tables[0].rows[date_row].cells:
        if (''.join((cell.text).split()).lower() != 'date') and (cell.text != last_cell_text):
            date = cell.text
            break
        last_cell_text = cell.text
    print('DATE: ', date, '\n')
    
    # Extract crew attendance details
    crew_names = {}
    crew_positions = {}
    row_index = 0 # Keep track of index
    end_of_loop = False
    
    for cell in doc.tables[0].rows[crew_row - 1].cells:
        if ''.join((cell.text).split()).lower() == 'crew':
            crew_count = row_index
        elif ''.join((cell.text).split()).lower() == 'position':
            position_count = row_index
        elif ''.join((cell.text).split()).lower() == 'present(p)/absent(a)':
            attendance_count = row_index
        elif ''.join((cell.text).split()).lower() == 'ertorlevel3firstaid':
            certification_count = row_index # Not really important
        row_index += 1
    print('Crew cell size:', crew_count)
    print('Position cell size:', position_count)
    print('Attendance cell size:', attendance_count)
    print('Certification cell size:', certification_count, '\n')

    for row in doc.tables[0].rows[crew_row:]:
        if row.cells[attendance_count].text == 'P':
            first_name, last_name = row.cells[crew_count].text.split()
            position = row.cells[position_count].text
            
            real_initials = first_name[0] + last_name[0]
            crew_names[real_initials] = row.cells[crew_count].text
            crew_positions[real_initials] = position
            
            print(first_name, last_name, position)
            
        for cell in row.cells:
            if ''.join((cell.text).split()).lower() == 'jobassignedto':
                end_of_loop = True
                break  
        if end_of_loop == True:
            break
        
    print('\n')

    # Consider common letter swapping nicknames
    nickname_initials = {
        'Bill': 'William',
        'Bob': 'Robert',
        'Dick': 'Richard',
        'Ted': 'Edward'
    }

    # Extracts work order details
    row_index = 0
    for cell in doc.tables[0].rows[jobs_row - 1].cells:
        if ''.join((cell.text).split()).lower() == 'jobassignedto':
            assignment_count = row_index
        elif ''.join((cell.text).split()).lower() == 'description':
            description_count = row_index
        elif ''.join((cell.text).split()).lower() == 'status':
            status_count = row_index
        row_index += 1
    print('Assignment cell size:', assignment_count)
    print('Description cell size:', description_count)
    print('Status cell size:', status_count, '\n')
    
    for row in doc.tables[0].rows[jobs_row:]:
        if row.cells[assignment_count].text.lower() in ['everyone', 'all']:
            name = 'Everyone'
        elif row.cells[assignment_count].text.lower() != '':
            assigned_names = []
            for initials in row.cells[assignment_count].text.split('/'):
                if initials in crew_names:
                    assigned_names.append(crew_names[initials])
                    print('Match found:', crew_names[initials], initials)
                else:
                    initials_found = False
                    for nickname, root_name in nickname_initials.items():
                        possible_initials = root_name[0] + initials[1]
                        print(f'Assumed nickname: {nickname}, swapping {initials} for {possible_initials}...')
                        if possible_initials in crew_names:
                            assigned_names.append(crew_names[possible_initials])
                            initials_found = True
                            print('Match found:', crew_names[possible_initials], initials)
                            break
                        else:
                            print(f'Cannot find using {nickname}\n')
                    if not initials_found:
                        print(f'Searching Excel for initial: {initials}...\n')
                        crew_names[initials] = search_in_excel(initials)
                        assigned_names.append(crew_names[initials])
                        print('Match found:', crew_names[initials], initials)

            name = ', '.join(assigned_names)

        description = row.cells[description_count].text
        status = row.cells[status_count].text
        if status not in ['DONE', 'IP']:
            break
        
        work_details.append((name, description, status))
        print(name, description, status, '\n')
        
    return (date, work_details)

# Function to extract data from laborAssignments as a fail case
def search_in_excel(initial):
    # Load Excel file
    workbook = xw.Book('laborAssignment.xlsx')
    sheet = workbook.sheets['List of Records']

    # Find the last row with data in column H
    last_row = sheet.range('H' + str(sheet.cells.last_cell.row)).end('up').row

    matched_names = []
    for i in range(1,last_row + 1):
        names = sheet.range(f'B{i}').value
        initials = sheet.range(f'H{i}').value
        if initials == initial:
            matched_names.append(names)
            print(f"\nMatch found for {initials} in Excel: {names}")

    if len(matched_names) > 1:
        print(f"\nMatches found for initials {initial}:")
        for index, name_option in enumerate(matched_names):
            print(f"{index + 1}. {name_option}")
        choice = int(input("\nChoose the correct match (enter the number): ")) - 1
        name = matched_names[choice]
    elif len(matched_names) == 0:
        print('\nMatch not found: DNE in Maximo')
        name = 'DNE in Maximo'
    else:
        name = matched_names[0]
                                
    return name

# Define function to write to Excel
def write_to_excel(date, work_details, excel_file_path, excel_file_sheet):
    print("Writing extracted data to Excel...\n")
    # Connect to the workbook
    wb = xw.Book(excel_file_path)
    
    sheet = wb.sheets[excel_file_sheet] 

    # Find the last used row in the Excel sheet
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    print(last_row) # how to find the last filled table row

    # Start entering data from the next available row
    for detail in work_details:
        name, description, status = detail
        print(f'{name} worked on {description} and it is {status}')
        # Split the description to extract 9-digit ID
        work_order_id = None
        for word in description.split():
            if word.startswith("W23") and len(word) == 9:
                work_order_id = word
                break
            if work_order_id:
                description = description.replace(work_order_id, '').strip()

        # Writing data to Excel
        sheet.range(f"A{last_row}").value = date
        sheet.range(f"B{last_row}").value = name
        sheet.range(f"C{last_row}").value = work_order_id
        sheet.range(f"D{last_row}").value = description
        sheet.range(f"E{last_row}").value = status

        # Increment row
        last_row += 1 
    
    # Save workbook after written data
    wb.save()

# Define function to fetch Maximo status using Selenium
def extract_maximo_status(browser, excel_file_path, excel_file_sheet):
    print("Extracting Maximo status...\n")
    # Connect to the workbook
    wb = xw.Book(excel_file_path)
    sheet = wb.sheets[excel_file_sheet]

    # Find the last used row in the Excel sheet
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    
    actions = ActionChains(browser) 
    wait = WebDriverWait(browser, 20)
    browser.get('https://prod.manage.prod.iko.max-it-eam.com/maximo')   
    browser.maximize_window()

    # Extract login information
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
        
    # Ensure History? section is NULL
    history = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tfrow_[C:20]_txt-tb")))
    #history.click()
    history.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    history.send_keys('Y', Keys.ENTER)
    time.sleep(5)
    print('History updated...')
    
    # Row 1 contains headers
    for row_num in range(2, last_row + 1):  
        if sheet.range(f'F{row_num}').value != 'CLOSE':
            # Work order is in column C
            work_order_id = sheet.range(f'C{row_num}').value

            if work_order_id:
                # Search for work order using Selenium
                searchWO_number = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tfrow_[C:1]_txt-tb")))
                searchWO_number.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                searchWO_number.send_keys(work_order_id)
                searchWO_number.send_keys(Keys.ENTER)
                time.sleep(5)
                
                try:
                    status = wait.until(EC.element_to_be_clickable((By.ID, "m6a7dfd2f_tdrow_[C:13]-c[R:0]")))
                    maximo_status = status.text
                    print(work_order_id, maximo_status)
                    
                    # Write Maximo status to Excel
                    sheet.range(f'F{row_num}').value = maximo_status
                
                except:
                    # If work order is not found in Maximo, status is "DNE"
                    sheet.range(f'F{row_num}').value = "DNE"
                    print(work_order_id, 'DNE')

            else:
                # If work order ID is missing, status is "Not Sure"
                sheet.range(f'F{row_num}').value = "NOT SURE"
                print(work_order_id, 'NOT SURE')
        
    # Save the Excel file
    wb.save()

main()