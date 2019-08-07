from selenium import webdriver
from selenium.webdriver import Firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC,\
    expected_conditions
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from os import devnull


import re
from _datetime import datetime, timedelta

from tkinter import Tk, Button, Label, constants, Checkbutton, BooleanVar, Text,\
    Scrollbar, Frame, StringVar, Radiobutton
from pyautogui import moveTo
from time import sleep
from selenium.webdriver.support.expected_conditions import alert_is_present
from lib2to3.pgen2.tokenize import Ignore
# import HelperFunctions

containersMoved = []
def popUpOKLeft(text1, text2="", textSize = 16):
    bgC = "lavender"
    top = Tk()
    top.config(bg = bgC)
    L1 = Label(top, text=text1, bg = bgC, padx = 20)
    L1.config(font=("serif", textSize))
    L1.grid(row=0, column=0, sticky=constants.W+constants.E)
    L1 = Label(top, text=text2, bg = bgC, padx = 20, justify=constants.LEFT)
    L1.config(font=("serif", textSize))
    L1.grid(row=1, column=0, sticky=constants.W + constants.E)
    def callbackOK():
#         sys.exit()
        top.destroy()
        
    MyButton = Button(top, text="OK", command=callbackOK)
    MyButton.grid(row=2, column=0, sticky=constants.W+constants.E, padx = 20, pady = (0,20))
    MyButton.config(font=("serif", 30), bg="green")
      
    top.update()
    
    w = top.winfo_width() # width for the Tk root
    h = top.winfo_height() # height for the Tk root
       
    ws = top.winfo_screenwidth() # width of the screen
    hs = top.winfo_screenheight() # height of the screen
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    
    top.geometry('%dx%d+%d+%d' % (w, h, x, y))
    top.update()
    moveTo(MyButton.winfo_width()/2 + MyButton.winfo_rootx(), MyButton.winfo_height()/2 + MyButton.winfo_rooty())
    top.lift()
    top.attributes('-topmost',True)
    top.after_idle(top.attributes,'-topmost',False)
    top.mainloop()


def getDaysForward(daysForward, weekend=False):
    target_Date_dateTime =datetime.now() + timedelta(days = daysForward)
    
    if target_Date_dateTime.weekday()>4 and not weekend:
        return False
    
    month = str(target_Date_dateTime.month)
    thisDay = str(target_Date_dateTime.day)
    
    if len(month)==1:
        month = "0" + month
    if len(thisDay)==1:
        thisDay = "0" + thisDay
    
    return str(target_Date_dateTime.year) + "-" + month + "-" + thisDay

def setupCn():
    fp = FirefoxProfile();
    fp.set_preference("webdriver.load.strategy", "unstable");
#     fp.set_preference("XRE_NO_WINDOWS_CRASH_DIALOG=1")
     
    driver = Firefox(firefox_profile=fp, log_path=devnull)
    driver.get("http://cn.ca/")
#     driver.set_window_position(1920, 0)
#     sleep(30)
    driver.maximize_window()
    
    driver.implicitly_wait(40)
#     f=open(r"J:\LOCAL DEPARTMENT\Automation - DO NOT MOVE\CN Login.txt", 'r')
    f=open(r"C:\Automation\CN Login.txt", 'r')
    read = f.readline()
    m = re.search("username: *", read)
    username = read[m.end():].rstrip()
    read = f.readline()
    m = re.search("password: *", read)
    password = read[m.end():].rstrip()
    f.close()
    
    driver.find_element_by_class_name("lbl").click()
    
    driver.find_element_by_id("login_usernameNew").send_keys(username)
    driver.find_element_by_id("login_passwordNew").send_keys(password)
    driver.find_element_by_id("loginform_enterbutton").click()
    
    return driver

def queryDate():
    dates=[]
#     returnDates = []
#     dateButtons = []
    today = datetime.now()
    
    
    
    daysForward = ""
    if today.weekday()<2:
        daysForward = 3
    else:
        daysForward = 5
    #remove the +1 for actual use
    for i in range(daysForward):
        thisDay = getDaysForward(i, True)
        if thisDay:
            dates.append(thisDay)
    
    bgC = "lavender"
    top = Tk()
    top.config(bg = bgC)
    L1 = Label(top, text="Please select the date to run on\nPlease only select one date.", bg = bgC, padx = 20)
    L1.config(font=("serif", 16))
    L1.grid(row=0, column=0, sticky=constants.W+constants.E)
    
    dateVar=StringVar()
    dateVar.set(dates[-1])
    f = Frame(top)
    f.grid(row=1, column=0)
    print("here")
    for date in dates:
        Radiobutton(f, text=date, variable=dateVar, value=date, bg=bgC, font = ("serif", 16)).pack()
    
#         check = BooleanVar()
#         checkButton = Checkbutton(top, text=date, variable=check, font=("serif", 16), bg=bgC)
#         checkButton.grid(row=i, column=0, sticky=constants.W+constants.E)
#         if i==daysForward:
#             check.set(True)
#         i+=1
#         dateButtons.append((check, date))
        
        
    def callbackDates():
#         for button, date in dateButtons:
#             if button.get():
#                 returnDates.append(date) 
        top.destroy()
    
    MyButton = Button(top, text="OK", command=callbackDates)
    MyButton.grid(row=2, column=0, sticky=constants.W+constants.E, padx = 20, pady = (0,20))
    MyButton.config(font=("serif", 30), bg="green")
      
    top.update()
    
    w = top.winfo_width() # width for the Tk root
    h = top.winfo_height() # height for the Tk root
       
    ws = top.winfo_screenwidth() # width of the screen
    hs = top.winfo_screenheight() # height of the screen
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    
    top.geometry('%dx%d+%d+%d' % (w, h, x, y))
    top.update()
    
    top.lift()
    top.attributes('-topmost',True)
    top.after_idle(top.attributes,'-topmost',False)
    moveTo(MyButton.winfo_width()/2 + MyButton.winfo_rootx(), MyButton.winfo_height()/2 + MyButton.winfo_rooty())
    
    top.mainloop()
    
    return dateVar.get()

def getBetterRVs(driver, date):
    driver.implicitly_wait(500)
#     sleep(30)
    driver.implicitly_wait(5)
    try:
        driver.find_element_by_css_selector("fa-icon[class='fa-fw ng-fa-icon'").click()
    except:
        print("no")
        pass

#     driver.switch_to.frame("menuHeader")
#     driver.find_element_by_id("id41").click()
    driver.implicitly_wait(1)
#     try:
#         driver.find_element_by_id("id41").click()
#     except:
#         driver.implicitly_wait(600)
# #         driver.find_element_by_id("tools").click()
#         driver.find_element_by_id("tools").click()
#         driver.find_element_by_id("tools").click()
#         driver.find_element_by_id("tools").click()
# #         sleep(5)
# #         driver.find_element_by_class_name("tools selected").click()
#         driver.switch_to_default_content()
#         driver.switch_to_frame("content1")
# #         sleep(500)
# #         driver.find_element_by_css_selector(r'a[href^="top.frames[0].openTab(\'id41\');"]').click()
# #         print(driver.find_element_by_css_selector('a[onclick*="top.frames[0].openTab(\'id41\');"]').text)
#         driver.find_element_by_css_selector('a[onclick*="top.frames[0].openTab(\'id41\');"]').click()
#         driver.switch_to_default_content()
#         driver.switch_to.frame("menuHeader")
#         driver.find_element_by_id("id41").click()
    try:
        driver.find_element_by_css_selector("ci-tools-standalone-menu[class='ci-recent-tools-container ng-tns-c5-3 ng-star-inserted']").click()
        driver.find_element_by_css_selector("a[href='#/tools/gate-appointment-inquiry']").click()
    except:
        pass

    driver.switch_to_frame(driver.find_element_by_css_selector("iframe[name='ci-tools-frame']"))
#     i = 2
#     found = False
#     driver.implicitly_wait(0)
#     while not found:
#         try:
#             driver.switch_to_default_content()
#             currentFrame = driver.find_element_by_css_selector("frame[name='content" + str(i) + "']")
#             driver.switch_to_frame(currentFrame)
#             driver.find_element_by_css_selector("form[action='AppointmentQuery']")
#             found = True
#         except:
#             if i<30:
#                 i+=1
#             else:
#                 i=2
    driver.implicitly_wait(600)
#     print(sel.first_selected_option())
#     print(sel.first_selected_option().name())
#     while sel.first_selected_option() != 
    
    
    timeOrders=[]
#     for date in dates:
    sel =Select(driver.find_element_by_name("terminalId")) 
    sel.select_by_visible_text("BRAMPTON INTERMODAL")
    elem= driver.find_element_by_name("appointmentDate")
    elem.clear()
    elem.send_keys(date)
#     for day in dates:
#         print(day)
    
    driver.find_element_by_id("btn_23").click()
    
    try:
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
    except:
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
        
    badTimes = [
        "00",
        "20",
        "19",
        "18"
        ]
    
    ignore_list = []
    
    done = False
    
    def check_RVs():
        nonlocal rows
        i=0
        while i < len(rows):
            cells = rows[i].find_elements_by_css_selector("td")
            time = cells[3].text[:2]
            if time != "" and int(time)>16 or time=="00":
                contNum =cells[8].text
                if cells[5].text=="Pickup" and not contNum in ignore_list:
                    cells[6].click()
                    driver.find_element_by_name("Modify").click()
                    if "No record of equipment" in driver.page_source:
    #                     print("no record")
                        elem1 = driver.find_element_by_name("Cancel")
                        elem1.click()
                        if time in badTimes:
                            elem2 = driver.find_element_by_name("Cancel")
                            while elem1.id == elem2.id:
                                elem2= driver.find_element_by_name("Cancel")
                            elem2.click()
                            while not "Gate Appointment Inquiry" in driver.page_source:
                                driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
                            wait = WebDriverWait(driver, 10000)
                            wait.until(lambda driver: "Gate Appointment Inquiry" in driver.page_source)
                            rows = driver.find_element_by_css_selector("table[class='TableStandardBG']").find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
                            break
    #                     else:
    #                         ignore_list.append(contNum)
                    elif("Equipment entered is not destined" in driver.page_source):
                        ignore_list.append(contNum)
                    driver.find_element_by_name("Cancel").click()
                    while not "Gate Appointment Inquiry" in driver.page_source:
                        driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
                    wait = WebDriverWait(driver, 100)
                    wait.until(lambda driver: "Gate Appointment Inquiry" in driver.page_source)
                    rows = driver.find_element_by_css_selector("table[class='TableStandardBG']").find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
            i+=1
        return True
    
    
    while not done:
        done = check_RVs()

    time_Order = ["05","06","07","08","04","09","11","12","13","14","10","16","18","19","20","00"] 
    preference_Order = list(time_Order)
    try:
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
    except:
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
        
    last_time = ""
    
    for row in reversed(rows):
#         time1 = datetime.now()
        cells = row.find_elements_by_css_selector("td")
#         time2 = datetime.now()
        
        if cells[5].text=="Pickup":
            last_time = cells[3].text[:2]
            break
    
    for time in list(time_Order):
        if int(last_time)<int(time):
            time_Order.remove(time)
    
    timeOrders.append(time_Order)
    
    driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()

    first = True
    sel =Select(driver.find_element_by_name("terminalId")) 
    sel.select_by_visible_text("BRAMPTON INTERMODAL")
    elem= driver.find_element_by_name("appointmentDate")
    elem.clear()
    elem.send_keys(date)
    
    driver.find_element_by_id("btn_23").click()
    
    time_Order = timeOrders.pop(0)
    cur_time = time_Order.pop()
    
    def get_worst_RV(driver):
        nonlocal cur_time
        
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
        while True:
            if (int(cur_time)<18 and cur_time!="00"):
                popUpOKLeft("Done")
                driver.quit()
                exit()
            for row in rows:
                cells = row.find_elements_by_css_selector("td")
                if cells[5].text=="Pickup":
                    time = cells[3].text[:2]
                    contNum =cells[8].text
                    if time=="00" and not contNum in ignore_list:
                        cells[6].click()
                        driver.find_element_by_name("Modify").click()
                        return contNum
                    elif not contNum in ignore_list:
                        break
            for row in reversed(rows):
                cells = row.find_elements_by_css_selector("td")
                if cells[5].text=="Pickup":
                    time = cells[3].text[:2]
                    contNum =cells[8].text
                    if time==cur_time and not contNum in ignore_list:
                        cells[6].click()
                        driver.find_element_by_name("Modify").click()
                        return contNum
            cur_time = time_Order.pop()
            
    driver.implicitly_wait(600)
    rv = get_worst_RV(driver)
    daysForward = ""
    day = datetime.now()
    if day.weekday()<2:
        daysForward = 3
    else:
        daysForward = 5
    target_Date = getDaysForward(daysForward)
    
#     target_Date = "2018-06-20"
    
    
    def take_appointment_time(target_Date, contNum):
#         print(contNum)
#         print("Old time: " + cur_time)
#         print("New time: " + buttonTime+ "  and date: " + target_Date)
        foundAGoodOne = False
        switch = False
        Select(driver.find_element_by_id("alternateDate")).select_by_visible_text(target_Date)
        
        driver.find_element_by_css_selector('input[onclick="beforeSubmit();saveActionAndSubmit(\'CheckAvailability\', \'GateAppointmentForm\', \'actionId\')"]').click()
        driver.implicitly_wait(0)
        try:
            times = driver.find_elements_by_name("alternateTimeChecked")
            driver.implicitly_wait(600)
            buttonTime = "00"
            button4=""
            button10=""
            for button in times:
                buttonTime =button.get_attribute('value')
                if buttonTime=="00:00":
                    continue
                if buttonTime=="04:00":
                    button4=button
                    continue
                if int(buttonTime[:2])<9:
                    button4=""
                    button.click()
                    break
                if button4:
                    button4.click()
                    break
                if buttonTime=="09:00":
                    button.click()
                    break
                if buttonTime=="10:00":
                    button10=button
                if int(buttonTime[:2])<16:
                    button10=""
                    button.click()
                    break
                if button10:
                    button10.click()
                    break
                if buttonTime=="16:00":
                    button.click()
                    break
                else:
                    button.click()
                    break
            if button4:
                button4.click()
            if button10 and not button4:
                button10.click()
            if preference_Order.index(buttonTime[:2])<preference_Order.index(cur_time):
                driver.find_element_by_name("Save").click()
                wait = WebDriverWait(driver, 100)
                wait.until(alert_is_present())
                driver.switch_to_alert().accept()
                foundAGoodOne = True
                switch = True
            elif("The equipment's ETA of" in driver.page_source):
                ignore_list.append(contNum)
                foundAGoodOne = True
                driver.find_element_by_name("Cancel").click()
                driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
            else:
                driver.find_element_by_name("Cancel").click()
                driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
            if switch:
                driver.switch_to_default_content()
                driver.switch_to_frame(driver.find_element_by_css_selector("iframe[name='ci-tools-frame']"))
#                 i = 2
#                 found = False
#                 driver.implicitly_wait(0)
#                 while not found:
#                     try:
#                         driver.switch_to_default_content()
#                         currentFrame = driver.find_element_by_css_selector("frame[name='content" + str(i) + "']")
#                         driver.switch_to_frame(currentFrame)
#         #                 driver.find_element_by_css_selector("form[action='AppointmentQuery']")
#                         driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']")
#                         found = True
#                     except:
#                         if i<30:
#                             i+=1
#                         else:
#                             i=2
                            
                
                driver.implicitly_wait(600)
                if "No appointments were available for the window" in driver.page_source:
                    driver.find_element_by_name("Modify").click()
                    return take_appointment_time(target_Date, contNum)
                else:
                    print(contNum)
                    print("Old time: " + cur_time)
                    print("New time: " + buttonTime+ "  and date: " + target_Date)
                    driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
        except:
            if not "Appointment Detail" in driver.page_source:
                driver.find_element_by_name("Cancel").click()
            driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
        return foundAGoodOne
    
    
    if first:
        go = False
        while datetime.now().minute <45:
    #         print(datetime.now().minute)
            True
    
        while not go:
            wait = WebDriverWait(driver, 100)
            elem = wait.until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, 'input[onclick="beforeSubmit();saveActionAndSubmit(\'CheckAvailability\', \'GateAppointmentForm\', \'actionId\')"]')))
            elem.click()
            sel = Select(driver.find_element_by_id("alternateDate"))
            option = sel.options[-1]
    #         driver.find_element_by_css_selector('input[onclick="beforeSubmit();saveActionAndSubmit(\'CheckAvailability\', \'GateAppointmentForm\', \'actionId\')"]').click()
    #         option = sel.options[-1]
    #         print(target_Date)
            if option.text==target_Date:
                go = True
            
                take_appointment_time(target_Date, rv)
        
    while True:
        day = datetime.now()
        if day.weekday()<2:
            daysForward = 3
        else:
            daysForward = 5
        
        for i in reversed(range(daysForward)):
            target_Date = getDaysForward(i+1)
#             print(target_Date)
            if target_Date:
                appointmentsLeft = True
                while appointmentsLeft:
                    contNum = get_worst_RV(driver)
                    appointmentsLeft = take_appointment_time(target_Date, contNum)
#                 if appointmentsLeft:
#                     print(contNum)
        
    
        
def time_test():
    
#     a = range(10000)
    time1 = datetime.now()
    
    time2 = datetime.now()
    print(time1)
    print(time2)
    print(time2.microsecond-time1.microsecond)
    
    exit()
    
if __name__ == '__main__':
#     time_test()
    driver = setupCn()
#     driver2 = setupCn()
    date = queryDate()
# #     dates = "2018-06-25"
# #     containers = queryContainers()
    getBetterRVs(driver, date)
    
# pyinstaller "C:\Users\ssleep\workspace\RV Changer\Automator\__init__.py" --distpath "J:\Spencer\RV Changer" --noconsole -y
# pyinstaller "C:\Users\spencer\workspaceseaport\programs\RV Changer\Automator\__init__.py" --distpath "c:\users\Spencer\RV Changer" -y  