#Selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys

from datetime import datetime, timedelta
import time

import os
import win32com.client as win32
import excel_macro as exm

import comvida_notification_functions as CV

wait = WebDriverWait(CV.driver,30)

#department Alog lists (order matters)
TH_dept = ["ACT", "CA", "FS", "HK", "LA", "LPN", "RN", "SCREEN"] 
VH_dept = TH_dept
TC_dept = ["ACT ST", "COOK", "FS", "HK", "LA", "LPN", "PWA", "REC", "SCREEN"]

#department email lists (order matters)
TH_email_dept = ["RN"]
VH_email_dept = ["LPN", "RN"]
TC_email_dept = ["LPN"]

#email recipient lists (order matters)
TH_email_shifts = ["RCCDAY", "RCCDWE", "RCCEVE", "RCCEwk", "RNN"]
VH_email_shifts = ["LPND1", "LPND2", "LPNE1", "LPNN1", "RN12D"]
TC_email_shifts = ["LPND", "LPNE", "LPNN"]

#TH night shifts (order matters)
TH_night_shifts = ["1LPNN", "1-N1", "2-N2", "2-NF2", "3-N3", "4-N4", "NR", "OrienN", "ORLPNN", "ORRNN", "RNN"]

assignType=[".Sched", "LATE", "Exchanged", "OT"] 

Alog_locations = [["Home", TH_dept], ["Valhaven", VH_dept], ["Court", TC_dept]]



def comvida_attendance_log_daily(location, dept, offset=1):
    url = "https://adaptiveems.com/" + location + "/SS51/Staff%20Scheduling/SSReports/AttendanceLog" 
    CV.comvida_login(url=url) #login to location specific comvida
        
    #input todays date and tomorrows date as start/end
    al_start_date_element = CV.driver.find_element(By.ID,"rptSSDateRange_DateFrom_I")
    al_start_date_element.send_keys(11*Keys.BACKSPACE)
    al_start_date_element.send_keys(CV.relative_date_today(days=offset))

    al_end_date_element = CV.driver.find_element(By.ID, "rptSSDateRange_DateTo_I")
    al_end_date_element.send_keys(11*Keys.BACKSPACE)
    al_end_date_element.send_keys(CV.relative_date_today(days=offset))

    #input paid and counted 
    CV.hover_click_element(CV.driver.find_element(By.ID,"cbPaid_B-1Img")) #click paid dropdown
    CV.hover_click_element(CV.driver.find_element(By.ID,"cbPaid_DDD_L_LBI1T0")) #click paid list item

    CV.hover_click_element(CV.driver.find_element(By.ID, "cbCounted_B-1Img")) #click counted dropdown
    CV.hover_click_element(CV.driver.find_element(By.ID,"cbCounted_DDD_L_LBI1T0" )) #click counted item
    
    for x in dept:
        scroll_and_click_id("DeptSelected", x) #department click list
    
    for x in assignType:
        scroll_and_click_id("AssignType", x) #assignType click list
    
    CV.hover_click_element(CV.driver.find_element(By.ID, "btnView_CD")) #click view report
    CV.hover_click_element(wait.until(CV.EC.element_to_be_clickable((By.ID, "btnXLSX")))) # wait then click export to excel
    
    while not os.path.exists("./SSAttendanceLog.xlsx"): #wait until downloaded
        time.sleep(.5)
    
    for attempt in range(10):
        try:
            os.rename('SSAttendanceLog.xlsx', "Reports/Unprocessed/"+ report_filename(location, offset)) #rename and move to reports folder
        except FileExistsError:
            try: 
                os.remove("Reports/Unprocessed/"+ report_filename(location, offset))
            except FileNotFoundError:
                os.remove('SSAttendanceLog.xlsx')
        else:
            break

    exm.attendance_log(file_path ="Reports/Unprocessed/"+ report_filename(location, offset), 
                       save_file_path="Reports/Processed/"+report_filename(location, offset),offset=offset)
    
    # if location =="TH":
    #     CV.hover_click_element(CV.driver.find_element(By.ID, "btnBack_CD")) #click back (returns to report checklist page)

    CV.driver.delete_all_cookies()

def comvida_night_attendance_log():

    save_pdf_button = "DocumentViewer_Splitter_Toolbar_Menu_DXI9_T"
    
    CV.comvida_login(url="https://adaptiveems.com/home/SS51/Staff%20Scheduling/SSReports/AttendanceLog")

    al_start_date_element = CV.driver.find_element(By.ID,"rptSSDateRange_DateFrom_I")
    al_start_date_element.send_keys(11*Keys.BACKSPACE)
    al_start_date_element.send_keys(CV.relative_date_today())

    al_end_date_element = CV.driver.find_element(By.ID, "rptSSDateRange_DateTo_I")
    al_end_date_element.send_keys(11*Keys.BACKSPACE)
    al_end_date_element.send_keys(CV.relative_date_today())

    #input paid and counted 
    CV.hover_click_element(CV.driver.find_element(By.ID,"cbPaid_B-1Img")) #click paid dropdown
    CV.hover_click_element(CV.driver.find_element(By.ID,"cbPaid_DDD_L_LBI1T0")) #click paid list item

    CV.hover_click_element(CV.driver.find_element(By.ID, "cbCounted_B-1Img")) #click counted dropdown
    CV.hover_click_element(CV.driver.find_element(By.ID,"cbCounted_DDD_L_LBI1T0" )) #click counted item

    for shift in TH_night_shifts:
        scroll_and_click_id("Shift", shift, .02) #shift click list
    
    for x in assignType:
        scroll_and_click_id("AssignType", x) #assignType click list
    
    CV.hover_click_element(CV.driver.find_element(By.ID, "btnView_CD")) #click view report
    CV.hover_click_element(wait.until(CV.EC.element_to_be_clickable((By.ID, save_pdf_button)))) # wait then click save pdf

    while not os.path.exists("./SSAttendanceLog.pdf"): #wait until downloaded
        time.sleep(.5)
    
    for attempt in range(10):
        try:
            os.rename('SSAttendanceLog.pdf', "Reports/TH_Nights/"+ report_filename("Home", 0, ".pdf")) #rename and move to reports folder
        except FileExistsError:
            try: 
                os.remove("Reports/TH_Nights/"+ report_filename("Home", 0, ".pdf"))
            except FileNotFoundError:
                os.remove('SSAttendanceLog.pdf')
        else:
            break

def scroll_and_click_id(cat_prefix, checkbox_dept, scroll_delay = 0.1):
    
    full_checkbox_id = "lst" + cat_prefix + "_" + checkbox_dept + "_D"
    full_scrollbar_id = "lst" + cat_prefix + "_D"

    try:
            # Locate the scrollable element
            scrollable_element = CV.driver.find_element(By.ID, full_scrollbar_id)

            while True:
                try:
                    # Check if the checkbox is visible
                    target_element = CV.driver.find_element(by=By.ID, value=full_checkbox_id)

                    # Scroll into view if needed
                    time.sleep(scroll_delay)
                    CV.driver.execute_script("arguments[0].scrollIntoView(true);", target_element)

                    target_element.click() # Perform the click action
                
                    break
                except Exception as scroll_error: # Scroll down inside the scrollable element
                    CV.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + arguments[0].offsetHeight;", scrollable_element)


    except Exception as e:
        print(f"An error occurred: {e}")

def report_filename(location, offset = 1, filetype = ".xlsx"):
    offset_date = datetime.now() + timedelta(days=offset)
    file_date = offset_date.strftime( "%B %d, %Y")

    match location:
        case "Home":
            return "TH Attendance Log - " + file_date + filetype
        case "Valhaven":
            return "VH Attendance Log - " + file_date + filetype
        case "Court":
            return "TC Attendance Log - " + file_date + filetype

def main():
    for x in Alog_locations:
        comvida_attendance_log_daily(*x, offset=1)

    comvida_night_attendance_log()
    # comvida_attendance_log_daily("Home", TH_dept, offset=2)
    CV.driver.quit()

if __name__ == "__main__":
    main()





