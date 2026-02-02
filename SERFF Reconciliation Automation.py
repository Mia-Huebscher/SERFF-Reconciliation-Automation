import requests
import time
import base64
import pyautogui
import numpy as np
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
from urllib.parse import urlencode
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def get_user_credentials(filename):
    with open(filename, 'r') as file:
        username = file.readline().strip()
        password = file.readline().strip()
    return username, password

def get_table_data(fee_type, table):
    '''
    Extracts data from an HTML table that sits to the right of a fee type category name
    '''
    xpath = f"//td[text()='{fee_type}']/following-sibling::td"
    fee= table.find_element(By.XPATH, xpath)
    return float(fee.text[1:])


if __name__ == '__main__':
    # Open a Chrome browser and navigate to SERFF
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    prefs = {
        "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
    } # Add these preferences to default the print option as Save as PDF
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(url=('https://login.serff.com/serff/signin.do'))

    # Login to SERFF with credentials
    username, password = get_user_credentials('credentials.txt')
    username_input = driver.find_element(By.ID, 'userName')
    username_input.send_keys(username)
    password_input = driver.find_element(By.ID, 'password')
    password_input.send_keys(password)
    driver.find_element(By.CLASS_NAME, 'inputSubmit').click()
    time.sleep(7)

    # Open the csv file downloaded from SERF containing the month's billing report
    billing_report = pd.read_csv('SERFF-Billing-Report.csv')
    # billing_report.drop('Unnamed: 15', axis=1, inplace=True)

    # Add 2 blank rows to the bottom of the dataframe for the "totals" data points
    for _ in range(2):
        billing_report.loc[len(billing_report)] = [""] * len(billing_report.columns)

    # Add state fees column
    billing_report['State Fees'] = np.zeros(len(billing_report))

    # Loop through the SERF tracking numbers
    for row_idx, row_data in billing_report.iterrows():
        serff_num = row_data['SERFF Tracking #']
        if serff_num != '':
            if serff_num[0] == 'P':

                # Ensure program is in the right browser window
                windows = driver.window_handles
                driver.switch_to.window(windows[0])
                
                # Input number in search box (upper, right corner)
                search_input = driver.find_element(By.ID, 'searchField')
                search_input.send_keys(serff_num)

                # Click Search button -> SERFF Tracking Number
                driver.find_element(By.ID, 'searchButtonBox').click()
                search_options = driver.find_element(By.ID, 'trackingSearchOptions')
                serff_tracking_num = search_options.find_element(By.TAG_NAME, 'li')
                serff_tracking_num.find_element(By.TAG_NAME, 'a').click()

                # Navigate to filing fees tab and extract state fee data
                driver.find_element(By.ID, 'tab-fees').click()
                active_tab = driver.find_element(By.ID, 'activeTab')

                # Navigate to the full fee table (second table in fees tab) and obtain its data
                tables = driver.find_elements(By.CLASS_NAME, 'dataTable')
                try:
                    try:
                        fee_table = tables[-1]
                    except:
                        fee_table = tables[0]
                    rows = fee_table.find_elements(By.TAG_NAME, 'tr')
                    
                    # Loop through the table to aggregate the necessary state fee data points
                    state_fees = 0
                    print(len(rows))
                    for row in rows[2:4]:
                        fee_info = row.find_elements(By.TAG_NAME, 'td')
                        state_fees += float(fee_info[1].text.strip()[1:])
                    print(f'{serff_num}: {state_fees}')

                    # Input total state fee amount into State Fees column
                    billing_report.loc[row_idx, 'State Fees'] = f'${state_fees}'

                    '''
                    Download third page of each PDF
                    '''
                    # Click the PDF Pipeline Button
                    driver.find_element(By.ID, 'generatePDF').click()

                    # Wait for the new window to open
                    time.sleep(3)

                    # Switch to the new pop-up window
                    windows = driver.window_handles
                    driver.switch_to.window(windows[-1])

                    # Click the filing information check box and generate pdf
                    driver.find_element(By.ID, 'include_filingInfo').click()
                    driver.find_element(By.ID, 'performPdfGeneration').click()

                    # Click the download button manually
                    try:
                        page_length_element = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, 'pagelength'))
                        )
                        page_length = page_length_element
                        print(f'Page length: {page_length}')
                    except:
                        print('Could not find page length element')
                        page_length = '3'
                    time.sleep(2)
                    pyautogui.hotkey('ctrl', 'p')
                    time.sleep(5)
                    for _ in range(6):
                        pyautogui.press('tab')    
                    time.sleep(3)
                    pyautogui.press('space') # Opens up the dropdown
                    pyautogui.press('down', presses=3) # Navigates to Custom
                    pyautogui.press('enter') # Chooses Custom
                    time.sleep(2)
                    pyautogui.press(page_length)
                    time.sleep(5)
                    pyautogui.press('tab', presses=2) # Navigates to the Save button
                    pyautogui.press('enter') # Presses Save
                    time.sleep(5)
                    pyautogui.press('enter') # Presses Save
                    time.sleep(3)
                except:
                    pass

            # Compute the totals for fees in the last 2 rows
            elif serff_num[0] == 'T':
                billing_report.loc[row_idx, 'Payment Method'] = 'Total'
                transaction_fees = [float(fee[1:]) for fee in billing_report['Amount'][:-2] if fee is not np.nan]
                billing_report.loc[row_idx, 'Amount'] = sum(transaction_fees)
                billing_report.loc[row_idx, 'Instance Name'] = 'Total'
                try:
                    state_fees = [float(fee[1:]) for fee in billing_report['State Fees'][:-3] if fee is not np.nan]
                except:
                    state_fees = [0]
                billing_report.loc[row_idx, 'State Fees'] = sum(state_fees)
    
    billing_report.loc[len(billing_report) - 2, 'State Fees'] = ''
    billing_report.loc[len(billing_report) - 1, 'Instance Name'] = 'Combined Total'
    billing_report.loc[len(billing_report) - 1, 'State Fees'] = sum(transaction_fees) + sum(state_fees)

    # Save updated dataframe into a new Excel file
    today = date.today()
    last_month = today - relativedelta(months=1)
    billing_report.to_excel(f'SERFF Fee Reconciliation {last_month.strftime('%m-%Y')}.xlsx', 
                            sheet_name=f'SERFF Fee Reconciliation {last_month.strftime('%m-%Y')}', index=False)
    