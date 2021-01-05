#TODO
#Clear the landing pad
#Dynamically pathed downloading
#variables for mouse positions
#rescale window for other window
    #code in to pull from latest download versus statically named file

from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui as pag
import pandas as pd
import time, os, shutil, openpyxl, datetime, re, os.path

sel_wait_time = 0.2
gui_wait_time = 0.5
switch_wait_time = 2
long_wait_time = 5
downloads = ""
accela_link = ""
outlook_link = ""
username = ""
password = ""
outlook_username = 
outlook_password = 
reviewer_report_xls = ''
all_plans_report_xls = ""
today = datetime.datetime.today()
column_widths_1 = {'A':20,'B':20,'C':50,'D':35,'E':25,'F':1,'G':30,'H':1,'I':20,'J':20,'K':20}
column_widths_2 = {'A':20,'B':20,'C':13,'D':24,'E':60,'F':1,'G':25,'H':30,'I':1,'J':15,'K':12, 'L':12}

def clear_downloads():
    for root, dirs, files in os.walk(downloads):
        for file in files:
            os.remove(os.path.join(root, file))

def login_and_download_reports():
    
    Planning = (47,195)
    review_PM = (120,558)
    exit_button = (888,88)
    review_reviewer = (175,573)
    
    
    driver = webdriver.Chrome()
    driver.get(accela_link)
    report_frame = driver.find_element_by_tag_name("iframe")
    driver.switch_to.frame(report_frame)
    agency_input = driver.find_element_by_id("servProvCode"); time.sleep(0.5)
    agency_input.send_keys("CVB"); time.sleep(sel_wait_time)
    actions_1 = ActionChains(driver) 
    actions_1.send_keys(Keys.TAB)
    actions_1.pause(sel_wait_time)
    actions_1.send_keys(username)
    actions_1.pause(sel_wait_time)
    actions_1.send_keys(Keys.TAB)
    actions_1.pause(sel_wait_time)
    actions_1.send_keys(password)
    actions_1.pause(sel_wait_time)
    actions_1.send_keys(Keys.TAB)
    actions_1.pause(sel_wait_time)
    actions_1.send_keys(Keys.ENTER)
    actions_1.perform(); time.sleep(gui_wait_time)
    driver.execute_script('''window.open("https://ceav9.vbgov.com/portlets/reports/dailyTree.do?mode=enter&module=Planning&reportPortletNo=3#");''')
    time.sleep(switch_wait_time)
    pag.moveTo(Planning[0], Planning[1], gui_wait_time)
    pag.click()
    pag.moveTo(review_PM[0], review_PM[1], gui_wait_time)
    pag.click()
    pag.moveTo(exit_button[0], exit_button[1], gui_wait_time)
    pag.click()
    pag.moveTo(review_reviewer[0], review_reviewer[1], gui_wait_time)
    pag.click()
    pag.moveTo(exit_button[0], exit_button[1], gui_wait_time)
    pag.click()
    time.sleep(gui_wait_time)
    
def prepare_spreadsheets():

    plans_under_review = 0
    late_reviews = 0
    late_letters = 0
    na_plans = ['VAR', 'PM'] 
    
    temp_report_1 = pd.read_excel(reviewer_report_xls)
    temp_report_1.to_excel("temp_report_1.xlsx")
    temp_report_2 = pd.read_excel(all_plans_report_xls)
    temp_report_2.to_excel("temp_report_2.xlsx")

    #Preparing the reviewer report
    wb = openpyxl.load_workbook('temp_report_1.xlsx')
    ws = wb['Sheet1']
    
    for row in ws.iter_rows():
        if row[3].value is None:
            continue
        else:
            if na_plans[0] in row[3].value or na_plans[1] in row[3].value:
                continue
            if isinstance(row[11].value, datetime.datetime):
                if today >= row[11].value:
                    x = row[10].value 
                    y = row[11].value 
                    late_reviews += 1
                    for i in range(2,12):
                        row[i].style = 'Accent1'
                    row[10].value = x
                    row[11].value  = y
    ws.delete_rows(1,4)
    ws.delete_cols(1,2)
    for col, width in column_widths_1.items():
        ws.column_dimensions[col].width = width
               
    ws['A1'].value = "Plan Revew Status Report"
    ws['A1'].style = 'Headline 1'
    wb.save('C:\\Users\\TCRhodes\\Desktop\\Launch&LandingPad\\Plan Review Status Report.xlsx')
    wb.close()
    
    #Preparing letter report
    wb = openpyxl.load_workbook('temp_report_2.xlsx')
    ws = wb['Sheet1']
    ws.delete_rows(1,4)
    for row in ws.iter_rows():
        if row[3].value is None:
            continue
        else:
            if na_plans[0] in row[3].value or na_plans[1] in row[3].value:
                continue
            if row[13].value != "Letter Due":
                if row[13].value is not None:
                    date = datetime.datetime.strptime(row[13].value, '%m/%d/%Y').date()
                    plans_under_review += 1
                    if today.date() > date:
                        x = row[12].value 
                        y = row[13].value 
                        late_letters += 1
                        for i in range(2,14):
                            row[i].style = 'Accent1'
                        row[12].value = x
                        row[13].value  = y

    for col, width in column_widths_2.items():
        ws.column_dimensions[col].width = width
    ws.delete_cols(1,2)
    
    ws['A1'].value = "Review Letter Status Report"
    ws['A1'].style = 'Headline 1'
    wb.save('C:\\Users\\TCRhodes\\Desktop\\Launch&LandingPad\\Review Letter Status Report.xlsx')
    wb.close()
    print(plans_under_review, " Plans Under Review")
    print(late_letters, " Past Target Due Date")
    return [plans_under_review, late_reviews, late_letters]
    
def write_email(report_data):
    email_entrybox = (390,432)
    next_button = (656,559)
    password_entrybox = (639,377)
    no_button = (546,557)
    new_button = (111,203)
    to_button = (694,265)
    cc_line = (697,351)
    subject_line = (624,376)
    body_line = (868,282)
    attach_button = (746,200)
    browse_button = (746,234)
    
    driver = webdriver.Chrome()
    driver.get(outlook_link); 
    pag.moveTo(email_entrybox[0], email_entrybox[1], switch_wait_time)
    pag.click()
    time.sleep(gui_wait_time)
    pag.typewrite(outlook_username, interval=0.03)
    pag.typewrite(['enter'])
    pag.moveTo(password_entrybox[0], password_entrybox[1], long_wait_time)
    pag.click()
    pag.typewrite(outlook_password, interval=0.03)
    pag.typewrite(['enter'])
    pag.moveTo(no_button[0], no_button[1], switch_wait_time)
    pag.click()
    pag.moveTo(new_button[0], new_button[1], switch_wait_time)
    pag.click()
    pag.moveTo(to_button[0], to_button[1], switch_wait_time)
    pag.click()
    pag.typewrite("", interval = 0.06)
    time.sleep(switch_wait_time)
    pag.typewrite(['enter'])
    pag.moveTo(cc_line[0], cc_line[1], gui_wait_time)
    pag.click(); time.sleep(gui_wait_time)
    pag.typewrite("", interval = 0.06); time.sleep(gui_wait_time)
    pag.typewrite(['enter']); time.sleep(gui_wait_time)
    pag.typewrite("", interval = 0.06); time.sleep(gui_wait_time)
    pag.typewrite(['enter']); time.sleep(gui_wait_time)
    pag.typewrite(['tab']); time.sleep(gui_wait_time)
    pag.typewrite('Report of Submittals under Review', interval = 0.03); time.sleep(switch_wait_time)
    pag.typewrite(['tab']); time.sleep(switch_wait_time)
    pag.typewrite('Carrie,\n\nThere are currently {} plans under active review. Of the {} submittals, {} are past their due date. Please let me know if you need anything else, thanks!\n\n This email was automatically generated from script on {}'.format(report_data[0], report_data[0], report_data[2], datetime.datetime.today().strftime("%A, %d. %B %Y %I:%M%p")), interval = 0.03); time.sleep(switch_wait_time)

def main():
        clear_downloads()
        login_and_download_reports()    
        report_data = prepare_spreadsheets()
        write_email(report_data)
        
if __name__ == "__main__":
    main()    
    

