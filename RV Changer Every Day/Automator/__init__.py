from selenium import webdriver
from selenium.webdriver import Firefox, chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC,\
    expected_conditions
from selenium.common.exceptions import TimeoutException,\
    SessionNotCreatedException, InvalidSessionIdException
from selenium.webdriver.common.by import By
from os import devnull

import re
from _datetime import datetime, timedelta

from tkinter import Tk, Button, Label, constants, Checkbutton, BooleanVar, Text,\
    Scrollbar, StringVar, Frame, Radiobutton
from pyautogui import moveTo
from time import sleep
from selenium.webdriver.support.expected_conditions import alert_is_present
from lib2to3.pgen2.tokenize import Ignore
# import HelperFunctions
from sys import exc_info
from selenium.webdriver.firefox.options import Options
from sys import exit
import atexit
# import sys
import os
from getpass import fallback_getpass
# from HelperFunctions import popUpOK

# import monkey
import signal
import subprocess
import sys
from _signal import SIGTERM, SIGINT
import psutil
import time
import win32gui
import win32process
from win32con import SW_MINIMIZE
# from signal import pthread_sigmask
# process="No Proc"

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


def getDaysForward(daysForward, weekdaysOK=False):
    target_Date_dateTime =datetime.now() + timedelta(days = daysForward)
    
    if target_Date_dateTime.weekday()>4 and not weekdaysOK:
        return False
    
    month = str(target_Date_dateTime.month)
    thisDay = str(target_Date_dateTime.day)
    
    if len(month)==1:
        month = "0" + month
    if len(thisDay)==1:
        thisDay = "0" + thisDay
    
    return str(target_Date_dateTime.year) + "-" + month + "-" + thisDay

def queryDatesAndTimes(allConts, malport):
    dates=[]
    returnDates = []
    dateButtons = []
    dateValues=[]
    returnTimes = []
    timeButtons = []
    timeValues=[]
    today = datetime.now()
    
    
    daysForward = ""
    
    if today.weekday()<2:
        daysForward = 3
    else:
        daysForward = 5

    for i in range(daysForward+1):
        thisDay = getDaysForward(i, True)
        if thisDay:
            dates.append(thisDay)
    
    bgC = "lavender"
    top = Tk()
    top.config(bg = bgC)
    L1 = Label(top, text="Please select the acceptable dates/times\nfor the new RVs\n(Hold shift to select all since last selection)", bg = bgC, padx = 20)
    L1.config(font=("serif", 16))
    L1.grid(row=0, column=0, sticky=constants.W+constants.E, columnspan = 2)
    
    start = [0]
    
    i=1

    def selectstart(start, chkbuttons, event):
        start[0]= chkbuttons.index(event.widget)
        
    def selectrange(start, chkbuttons, event):
            startLocal = start[0]
            end = chkbuttons.index(event.widget)
            sl = slice(min(startLocal, end)+1, max(startLocal, end))
            for cb in chkbuttons[sl]:
                cb.toggle()
            start[0] = startLocal
    
    for date in dates:
        check = BooleanVar()
        checkButton = Checkbutton(top, text=date, variable=check, bg="royal blue", font=("serif", 12))
        checkButton.grid(row=i, column=0, sticky=constants.W+constants.E, pady=0)
#         if malport:
#             checkButton.grid(columnspan=2)
        if i>1 and allConts:
            check.set(True)
        i+=1
        checkButton.bind("<Button-1>", lambda event: selectstart(start, dateButtons, event))
        checkButton.bind("<Shift-Button-1>", lambda event: selectrange(start, dateButtons, event))
        dateButtons.append(checkButton)
        dateValues.append((check, date))
    
    if not malport:
        time_Order = ["04", "05","06","07","08","09","10","11","12","13","14","16","18","19","20", "00"]
    else:
#         time_Order = ["06","11","14","16","18", "00"]
#         f=open(r"C:\Automation\CNPort.txt", 'r')
        time_Order=[]
        f=open(r"J:\LOCAL DEPARTMENT\Automation - DO NOT MOVE\Malport Hours.txt", 'r')
#         f=open(r"C:\Automation\Malport Hours.txt", 'r')
        #     read = f.readline()
        #     m = re.search("username: *", read)
        #     username = read[m.end():].rstrip()
        
        read = f.readline()
        while read != "":
            if read == "\n":
                continue
            else:
                read = read.strip()
                if len(read)==1:
                    read = "0"+read
                if len(read)==2:
                    read+=":00"
                if len(read)==4 and not ":" in read:
                    read=read[:2]+":"+read[2:]
                elif len(read)==4:
                    read = "0"+read
            time_Order.append(read)
            read = f.readline()
        f.close()
#         time_Order = ["06","08","10","12","14","16","18","20", "00"]

        
    for i in range(len(time_Order)):
        check=BooleanVar()
        if len(time_Order[i])==2:
            cb = Checkbutton(text=time_Order[i]+":00",padx=0,pady=0,bd=0, variable=check, bg="dark violet", font=("serif", 12))
        else:
            cb = Checkbutton(text=time_Order[i],padx=0,pady=0,bd=0, variable=check, bg="dark violet", font=("serif", 12))
        cb.grid(row = i+1, column = 1, sticky=constants.W+constants.E+constants.N+constants.S)
        cb.bind("<Button-1>", lambda event: selectstart(start, timeButtons, event))
        cb.bind("<Shift-Button-1>", lambda event: selectrange(start, timeButtons, event))
        timeValues.append((check, time_Order[i]))
        timeButtons.append(cb)
        
        
        
    def callbackDates():
        for button, date in dateValues:
            if button.get():
                returnDates.append(date)
        if len(returnDates)<1:
            popUpOKLeft("Please select target dates for the new RV(s)")
        for button, time in timeValues:
            if button.get():
                returnTimes.append(time)  
        
        top.destroy()
        
    
        
    MyButton = Button(top, text="OK", command=callbackDates)
    MyButton.grid(row=i+2, column=0, sticky=constants.W+constants.E, padx = 20, pady = (0,20), columnspan = 2)
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
    if not malport:
        moveTo(MyButton.winfo_width()/2 + MyButton.winfo_rootx(), MyButton.winfo_height()/2 + MyButton.winfo_rooty())
    
    top.mainloop()
#     if not malport:
    if len(returnTimes)<1:
        returnTimes = ["04", "05","06","07","08","09","10","11","12","13","14","16"]
#     else:
#         if len(returnTimes)<1:
#             returnTimes = ["06","08","10","12","14","16"]
#     else:
#         returnTimes = ["00"]
    
    return returnDates, returnTimes

def queryContainers():
    containers = []
    
    bgC = "lavender"
    top = Tk()
    top.config(bg = bgC)
    L1 = Label(top, text="Please enter the containers to run on, and\nwhat day their current RVs are", bg = bgC, padx = 20)
    L1.config(font=("serif", 16))
    L1.grid(row=0, column=0, sticky=constants.W+constants.E)
    
    S1=Scrollbar(top, orient='vertical')
    S1.grid(row=1, column=1, sticky=constants.N + constants.S)
    S2=Scrollbar(top, orient='horizontal')
    S2.grid(row=2, column=0, sticky=constants.E + constants.W)
    
    T1 = Text(top, height = 20, width = 97, xscrollcommand = S2.set, yscrollcommand=S1.set, wrap = constants.NONE)
    T1.grid(row=1, column=0)
    T1.insert("end", "ALL")
    
    f1 = Frame(top)
    f1.grid(row=4, column=0) 
    
    checkHeadless = BooleanVar()
    checkHeadless.set(True)
    cb = Checkbutton(f1, text="Run in background?", variable=checkHeadless, bg="brown1", font=("serif", 12))
    cb.grid(row=0, column=0, pady=(10,0), padx = 20)
    
    checkDelivery = BooleanVar()
    cb = Checkbutton(f1, text="Delivery RVs?", variable=checkDelivery, bg="brown1", font=("serif", 12))
    cb.grid(row=0, column=1, pady=(10,0), padx = 20)
    
    checkallowSameDayAsETA = BooleanVar()
    cb = Checkbutton(f1, text="Try for RVs for the same\nday as ETA?", variable=checkallowSameDayAsETA, bg="brown1", font=("serif", 12))
    cb.grid(row=0, column=2, pady=(10,0), padx = 20)
    
    checkMalport = BooleanVar()
    cb = Checkbutton(f1, text="Malport?", variable=checkMalport, bg="brown1", font=("serif", 12))
    cb.grid(row=0, column=3, pady=(10,0), padx = 20)
    
    def callbackCont():
        if date.get()=="":
            popUpOKLeft("Please select the day the current RV is on")
        else:
            if T1.get("1.0", constants.END).strip()=="":
                popUpOKLeft("Please list the target containers or RV #s (for deliveries)")
            else:
                containers.append(T1.get("1.0", constants.END).splitlines())
                top.destroy()
        
    dates=[]
            
    daysForward = ""
    
    if datetime.now().weekday()<2:
        daysForward = 3
    else:
        daysForward = 5

    if (datetime.now().hour>7 and datetime.now().minute>46) or datetime.now().hour>8:
        daysForward+=1
    for i in range(daysForward):
        dates.append(getDaysForward(i, True))    
    
    date=StringVar()
    f = Frame(top)
    f.grid(row=3, column=0)
    for dateStr in dates:
        Radiobutton(f, text=dateStr, variable=date, value=dateStr, indicatoron=0, bg="royal blue", font = ("serif", 16)).pack(side="left")
    
    MyButton = Button(top, text="OK", command=callbackCont)
    MyButton.grid(row=5, column=0, sticky=constants.W+constants.E, padx = 20, pady = 10)
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
    
    return containers[0], date.get(), checkHeadless.get(), checkDelivery.get(), checkallowSameDayAsETA.get(), checkMalport.get()


def setupCn(headlessBrowser, process="", globalPort=""):
    
#     webdriver.common.service.Service.start = monkey.start
    
    options = Options()
    options.set_headless(headless=headlessBrowser)
    fp = FirefoxProfile();
    fp.set_preference("webdriver.load.strategy", "unstable");
    if globalPort=="":
        try:
            f=open(r"C:\Automation\CNPort.txt", 'r')
        #     read = f.readline()
        #     m = re.search("username: *", read)
        #     username = read[m.end():].rstrip()
            read = f.readline()
            port=int(read)+1
            if port == 9999:
                port=1000
    #     m = re.search("password: *", read)
    #     password = read[m.end():].rstrip()
            f.close()
            os.remove(r"C:\Automation\CNPort.txt")
        except Exception:
            port=4444    
        f=open(r"C:\Automation\CNPort.txt", 'w+')
        f.write(str(port))
        f.close()
#     port = 4444
    failed=True
    while failed:
        try:
            if process=="":
                subprocess.Popen("start geckodriver -p "+str(port), shell=True)
                newestGecko=""
                found = False
                while not found:
                    for proc in psutil.process_iter():
                        if "geckodriver" in proc.name().lower():
    #                         p = psutil.Process(proc.pid)
                            if newestGecko=="":
                                newestGecko=proc
                            elif proc.create_time()>newestGecko.create_time():
                                newestGecko=proc
                    if newestGecko!="" and (time.time()-newestGecko.create_time()<5):
                        found=True
                process=newestGecko
#                 print(process)
                def get_hwnds_for_pid (pid):
                    def callback (hwnd, hwnds):
                        if win32gui.IsWindowVisible (hwnd) and win32gui.IsWindowEnabled (hwnd):
                            _, found_pid = win32process.GetWindowThreadProcessId (hwnd)
                            if found_pid == pid:
                                hwnds.append (hwnd)
                            return True
                    
                    hwnds = []
                    win32gui.EnumWindows (callback, hwnds)
                    return hwnds
            
                for hwnd in get_hwnds_for_pid(process.pid):
                    win32gui.ShowWindow(hwnd, SW_MINIMIZE)
            if globalPort!="":
                port=globalPort
            driver=webdriver.Remote("http://127.0.0.1:"+str(port),desired_capabilities=webdriver.DesiredCapabilities.FIREFOX, options=options)
            failed=False
            globalPort=port
        except Exception:
#             p=psutil.Process(process.pid)
            process.kill()
            port+=1
#     print(process)
#     process.kill()
#     driver = Firefox(firefox_profile=fp, log_path=devnull, firefox_options=options)
#     driver = webdriver.Chrome()
#     driver = webdriver.Ie()
    driver.get("http://cn.ca/")
    driver.maximize_window()
    
    driver.implicitly_wait(100)
    
    f=open(r"J:\LOCAL DEPARTMENT\Automation - DO NOT MOVE\CN Login.txt", 'r')
#     f=open(r"C:\Automation\CN Login.txt", 'r')
    read = f.readline()
    m = re.search("username: *", read)
    username = read[m.end():].rstrip()
    read = f.readline()
    m = re.search("password: *", read)
    password = read[m.end():].rstrip()
    f.close()    
    
    driver.find_element_by_class_name("lbl").click()
    driver.find_element_by_id("login_usernameNew").clear()
    driver.find_element_by_id("login_usernameNew").send_keys(username)
    driver.find_element_by_id("login_passwordNew").send_keys(password)
    driver.find_element_by_id("loginform_enterbutton").click()
    
    return driver, process, globalPort


def getBetterRVs(driver, containers, date, acceptableTimes, targetDates, headlessBrowser, delivery, allowSameDayAsETA, malport, process, globalPort):
    def exit_hander():
        print("Quitting")
#         sleep(20)
        try:
            driver.quit()
        except Exception:
            pass
        try:process.kill()
        except Exception:
            pass
#         sys.exit()
    atexit.register(exit_hander)
    messages=False
    allConts = containers[0]=="ALL"
    weekend = not getDaysForward(int(date[-2:]) - datetime.now().day)
    
    driver.switch_to_default_content()
    driver.switch_to.frame("menuHeader")
    driver.implicitly_wait(1)
    try:
        driver.find_element_by_id("id41").click()
    except Exception:
        driver.implicitly_wait(600)
#         driver.find_element_by_id("tools").click()
        driver.find_element_by_id("tools").click()
        driver.find_element_by_id("tools").click()
        driver.find_element_by_id("tools").click()
#         sleep(5)
#         driver.find_element_by_class_name("tools selected").click()
        driver.switch_to_default_content()
        driver.switch_to_frame("content1")
#         sleep(500)
#         driver.find_element_by_css_selector(r'a[href^="top.frames[0].openTab(\'id41\');"]').click()
#         print(driver.find_element_by_css_selector('a[onclick*="top.frames[0].openTab(\'id41\');"]').text)
        driver.find_element_by_css_selector('a[onclick*="top.frames[0].openTab(\'id41\');"]').click()
        driver.switch_to_default_content()
        driver.switch_to.frame("menuHeader")
        driver.find_element_by_id("id41").click()

    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[name='content1']"))
    i = 2
    found = False
    driver.implicitly_wait(0)
    while not found:
        try:
            driver.switch_to_default_content()
            currentFrame = driver.find_element_by_css_selector("frame[name='content" + str(i) + "']")
            driver.switch_to_frame(currentFrame)
            driver.find_element_by_css_selector("form[action='AppointmentQuery']")
            found = True
        except Exception:
            if i<30:
                i+=1
            else:
                i=2
    driver.implicitly_wait(600)
    sel =Select(driver.find_element_by_name("terminalId")) 
    if not malport:
        sel.select_by_visible_text("BRAMPTON INTERMODAL")
    else:
        sel.select_by_visible_text("MALPORT YARD")
    
    elem= driver.find_element_by_name("appointmentDate")
    elem.clear()
    elem.send_keys(date)
    
    ignore_list=[]
    
    driver.find_element_by_id("btn_23").click()
    failed = True
    while failed:
        try:
            table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
            rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
            failed = False
        except Exception:
            pass
#         except Exception:
#             table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
#             rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
        
    badTimes = [
        "00",
        "20",
        "19",
        "18"
        ]
    
    if not delivery:
    
        def check_RVs():
            nonlocal rows
            i=0
            while i < len(rows):
                cells = rows[i].find_elements_by_css_selector("td")
                time = cells[3].text[:2]
                if time != "" and int(time)>16 or time=="00" or weekend:
                    contNum =cells[8].text
                    if cells[5].text=="Pickup" and not contNum in ignore_list:
                        cells[6].click()
                        driver.find_element_by_name("Modify").click()
                        try:
                            driver.switch_to_alert().accept()
                            alert=True
                        except Exception:
                            alert=False
                        if alert:
                            failed = True
                            while failed:
                                try:
                #                 driver.switch_to_default_content()
                #                 driver.switch_to_frame(driver.find_element_by_css_selector("frame[name='content1']"))
                                    j = 2
                                    found = False
                                    driver.implicitly_wait(0)
                                    while not found:
                                        try:
                                            driver.switch_to_default_content()
                                            currentFrame = driver.find_element_by_css_selector("frame[name='content" + str(j) + "']")
                                            driver.switch_to_frame(currentFrame)
                                            driver.find_element_by_id("alternateDate")
                                            found = True
                                        except Exception:
                                            if j<30:
                                                j+=1
                                            else:
                                                j=2
                                                     
                                         
                                    driver.implicitly_wait(60)
#                                     Select(driver.find_element_by_id("alternateDate")).select_by_visible_text(target_Date)
                                    failed = False
                #                 Select(driver.find_element_by_id("alternateDate")).select_by_visible_text(target_Date)
                #                 failed = True
                                except Exception:
                                    print(sys.exc_info())
                                    driver.refresh()
                        if "No record of equipment" in driver.page_source:
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
    
        done = False
        
        
        while not done:
            done = check_RVs()
#         
        
        
        
    time_Order = ["05","06","07","08","04","09","11","12","13","14","10","16","18","19","20","00"] 
    preference_Order = list(time_Order)
    try:
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
    except Exception:
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
            
    if not delivery:
        last_time = ""
        
        for row in reversed(rows):
            cells = row.find_elements_by_css_selector("td")
            
            if (cells[5].text=="Pickup" and not delivery) or (cells[5].text=="Delivery" and delivery):
                last_time = cells[3].text[:2]
                break
        
        for time in list(time_Order):
            if int(last_time)<int(time):
                time_Order.remove(time)
                
        cur_time = time_Order.pop()
    
    
        def get_worst_RV(driver):
            nonlocal cur_time
            nonlocal allConts
            found = False
            while not found:
                try:
                    table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
                    rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
                    found=True
                except Exception:
                    pass
            while True:
                if (int(cur_time)<18 and cur_time!="00") and allConts and not weekend:
                    if (messages):
                        popUpOKLeft("Done, but check the console\nwindow for messages")
                    else:
                        popUpOKLeft("Done")
#                     driver.quit()
                    exit()
                for row in rows:
                    cells = row.find_elements_by_css_selector("td")
                    if cells[5].text=="Pickup":
                        time = cells[3].text[:2]
                        contNum =cells[8].text
                        if time=="00":
                            if not contNum in ignore_list and (allConts or contNum in containers or contNum[:4]+contNum[5:11] in containers):
                                cells[6].click()
                                driver.find_element_by_name("Modify").click()
                                return contNum
                        else:
                            break
                for row in reversed(rows):
                    cells = row.find_elements_by_css_selector("td")
                    if (cells[5].text=="Pickup" and not delivery) or (cells[5].text=="Delivery" and delivery):
                        time = cells[3].text[:2]
                        contNum =cells[8].text
                        if time==cur_time and not contNum in ignore_list and (allConts or contNum in containers or contNum[:4]+contNum[5:11] in containers):
                            cells[6].click()
                            driver.find_element_by_name("Modify").click()
                            return contNum
                if len(time_Order)==0:
#                     targetDateString = ""
#                     for targetDateX in date:
#                         targetDateString=targetDateString+targetDateX+", "
#                     targetDateString=targetDateString[:-2]
                    for cont in containers:
                        if (not cont in ignore_list):
                            print("Could not find "+cont+" on "+date)
                            messages=True
#                     [i for i in ignore_list if not i in containers or containers.remove(i)]
                    
                    if (messages):
                        popUpOKLeft("Done, but check the console\nwindow for messages")
                    else:
                        popUpOKLeft("Done")
#                     driver.quit()
                    exit()
                else:
                    cur_time = time_Order.pop()
                
                
        driver.implicitly_wait(600)
    else:
        cur_time = "00"
    
    
    def get_RV(rv):
        found=False
        table = driver.find_element_by_css_selector("table[class='TableStandardBG']")
        rows = table.find_element_by_css_selector("table[id='listingTable']>tbody").find_elements_by_css_selector("tr")
        for row in rows:
            cells = row.find_elements_by_css_selector("td")
            if rv in cells[6].text:
                cells[6].click()
                driver.find_element_by_name("Modify").click()
                found=True
                break
        if not found:
            print("Could not find " + rv + " on date: " +str(date))
            messages=True
        return found
    
    def take_appointment_time(target_Date, contNum, rv=False):
        nonlocal allConts
        foundAGoodOne = False
        switch = False
        failed = True
        try:
            driver.switch_to_alert().accept()
            alert=True
        except Exception:
            alert=False
        if alert:
            while failed:
                try:
#                 driver.switch_to_default_content()
#                 driver.switch_to_frame(driver.find_element_by_css_selector("frame[name='content1']"))
                    i = 2
                    found = False
                    driver.implicitly_wait(0)
                    while not found:
                        try:
                            driver.switch_to_default_content()
                            currentFrame = driver.find_element_by_css_selector("frame[name='content" + str(i) + "']")
                            driver.switch_to_frame(currentFrame)
                            driver.find_element_by_id("alternateDate")
                            found = True
                        except Exception:
                            if i<30:
                                i+=1
                            else:
                                i=2
                                     
                         
                    driver.implicitly_wait(60)
                    Select(driver.find_element_by_id("alternateDate")).select_by_visible_text(target_Date)
                    failed = False
#                 Select(driver.find_element_by_id("alternateDate")).select_by_visible_text(target_Date)
#                 failed = True
                except Exception:
#                     driver.refresh()
                    pass
        try:
            Select(driver.find_element_by_id("alternateDate")).select_by_visible_text(target_Date)
        except Exception:
            print("Failed @alternateDate, restarting")
            return -1
        failed = True
        while failed:
            try:
                wait = WebDriverWait(driver, 100)
                elem = wait.until(expected_conditions.element_to_be_clickable((By.CSS_SELECTOR, 'input[onclick="beforeSubmit();saveActionAndSubmit(\'CheckAvailability\', \'GateAppointmentForm\', \'actionId\')"]')))
                elem.click()
                failed=False
            except Exception:
                pass
        
        driver.implicitly_wait(0)
        try:
            times = driver.find_elements_by_name("alternateTimeChecked")
        except Exception:
            times = []
        if not times==[]:
            try:
                driver.implicitly_wait(600)
                buttonTime = "00"
                button4=""
                button10=""
                for button in times:
                    buttonTime =button.get_attribute('value')
                    if acceptableTimes and not buttonTime[:2] in acceptableTimes:
                        continue
                    if buttonTime=="00:00":
                        if buttonTime[:2] in acceptableTimes:
                            button.click()
                            break
                        else:
                            continue
                    if buttonTime=="04:00":
                        button4=button
                        continue
                    if int(buttonTime[:2])<9:
                        button.click()
                        button4=""
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
                    buttonTime="04:00"
                if button10 and not button4:
                    button10.click()
                    buttonTime="10:00"
                if (preference_Order.index(buttonTime[:2])<preference_Order.index(cur_time) 
                        and (not acceptableTimes or buttonTime[:2] in acceptableTimes)
                        and not (buttonTime=="20:00")
                        and not ((buttonTime=="19:00" or buttonTime=="18:00") and not target_Date==date)) \
                    or \
                    (not allConts
                        and (not acceptableTimes or buttonTime[:2] in acceptableTimes)):
    #                 print("clicking")medu9350677
                    driver.find_element_by_name("Save").click()
                    wait = WebDriverWait(driver, 100)
                    wait.until(alert_is_present())
                    driver.switch_to_alert().accept()
                    foundAGoodOne = True
                    switch = True
                    try:
                        wait = WebDriverWait(driver, 1)
                        wait.until(alert_is_present())
                        driver.switch_to_alert().accept()
#                         sleep(10)
                    except Exception:
                        pass
                elif((not allowSameDayAsETA and "The equipment's ETA of" in driver.page_source) or "time of " + target_Date + " " + acceptableTimes[-1] in driver.page_source):
                    if len(targetDatesTemp)>1:
                        targetDatesTemp.remove(target_Date)
                    else:
                        ignore_list.append(contNum)
                        print("ETA for "+contNum+" is on or after the last acceptable day. \nIf you want the program to try to get an RV for whatever times are allowed anyway,\n run again but check the \"same day as ETA\" option")
                        messages=True
                        driver.find_element_by_name("Cancel").click()
                        driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
                        return True
                if switch:
                    driver.switch_to_default_content()
                    driver.switch_to_frame(driver.find_element_by_css_selector("frame[name='content1']"))
                    i = 2
                    found = False
                    driver.implicitly_wait(0)
                    while not found:
                        try:
                            driver.switch_to_default_content()
                            currentFrame = driver.find_element_by_css_selector("frame[name='content" + str(i) + "']")
                            driver.switch_to_frame(currentFrame)
                            driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']")
                            found = True
                        except Exception:
                            if i<30:
                                i+=1
                            else:
                                i=2
                                 
                     
                    driver.implicitly_wait(600)
                    if "No appointments were available for the window" in driver.page_source:
                        driver.find_element_by_name("Modify").click()
                        foundAGoodOne = False
                        print(datetime.now())
                        print("Missed: " + buttonTime)
                    else:
                        if not allConts:
                            ignore_list.append(contNum)
                        print(datetime.now())
                        print(contNum)
                        print("Old time: " + cur_time + "  and date: " + date)
                        print("New time: " + buttonTime+ "  and date: " + target_Date + "\n")
                        driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
            except Exception:
                driver.find_element_by_name("Cancel").click()
                driver.find_element_by_css_selector("img[src='/ImxEbusWeb/images/english/Back.gif']").click()
                print(exc_info())
                messages=True
                if (messages):
                    popUpOKLeft("Done, but check the console\nwindow for messages")
                else:
                    popUpOKLeft("Done")
                exit()
        return foundAGoodOne
    
    startTime = datetime.now()
    
    rvs = list(containers)
    rv = rvs[0]
    while len(rvs)>0:
        if datetime.now()-startTime>timedelta(seconds=60):
            driver.close()
            print("Restarting...   " + str(datetime.now()))
            driver, process, globalPort = setupCn(headlessBrowser, process, globalPort)
            driver.switch_to.frame("menuHeader")
            driver.implicitly_wait(1)
            try:
                driver.find_element_by_id("id41").click()
            except Exception:
                driver.implicitly_wait(600)
        #         driver.find_element_by_id("tools").click()
                driver.find_element_by_id("tools").click()
                driver.find_element_by_id("tools").click()
                driver.find_element_by_id("tools").click()
                driver.find_element_by_id("tools").click()
                driver.find_element_by_id("tools").click()
        #         sleep(5)
        #         driver.find_element_by_class_name("tools selected").click()
                driver.switch_to_default_content()
                driver.switch_to_frame("content1")
        #         sleep(500)
        #         driver.find_element_by_css_selector(r'a[href^="top.frames[0].openTab(\'id41\');"]').click()
        #         print(driver.find_element_by_css_selector('a[onclick*="top.frames[0].openTab(\'id41\');"]').text)
                try:
                    driver.find_element_by_css_selector('a[onclick*="top.frames[0].openTab(\'id41\');"]').click()
                except Exception:
                    driver.execute_script("arguments[0].click();", driver.find_element_by_css_selector('a[onclick*="top.frames[0].openTab(\'id41\');"]'))
                driver.switch_to_default_content()
                driver.switch_to.frame("menuHeader")
                driver.find_element_by_id("id41").click()
            driver.switch_to_default_content()
            driver.switch_to_frame(driver.find_element_by_css_selector("frame[name='content1']"))
            i = 2
            found = False
            driver.implicitly_wait(0)
            while not found:
                try:
                    driver.switch_to_default_content()
                    currentFrame = driver.find_element_by_css_selector("frame[name='content" + str(i) + "']")
                    driver.switch_to_frame(currentFrame)
                    driver.find_element_by_css_selector("form[action='AppointmentQuery']")
                    found = True
                except Exception:
                    if i<30:
                        i+=1
                    else:
                        i=2
            driver.implicitly_wait(600)
            sel =Select(driver.find_element_by_name("terminalId")) 
            if not malport:
                sel.select_by_visible_text("BRAMPTON INTERMODAL")
            else:
                sel.select_by_visible_text("MALPORT YARD")
            elem= driver.find_element_by_name("appointmentDate")
            elem.clear()
            elem.send_keys(date)
            
            driver.find_element_by_id("btn_23").click()
            
            startTime = datetime.now()
        if not delivery:
            contNum = get_worst_RV(driver)
            targetDatesTemp = targetDates
            gotRV = False
            while not gotRV:
                if datetime.now()-startTime>timedelta(seconds=60):
                    break
                for day in targetDatesTemp:
                    gotRV = take_appointment_time(day, contNum)
                    if gotRV:
                        break
            
            if not allConts:
                allCont = True
                for container in containers:
                    allCont = allCont and container in ignore_list
                if allCont:
                    if (messages):
                        popUpOKLeft("Done, but check the console\nwindow for messages")
                    else:
                        popUpOKLeft("Done")
#                     driver.quit()
                    exit()
        else:
            targetDatesTemp = targetDates
            gotRV = False
            if not get_RV(rv):
                messages=True
                rvs.remove(rv)
                if len(rvs)>0:
                    rv = rvs[0]
                else:
                    if (messages):
                        popUpOKLeft("Done, but check the console\nwindow for messages")
                    else:
                        popUpOKLeft("Done")
#                             driver.quit()
                    exit()
                continue
            while not gotRV:
                if datetime.now()-startTime>timedelta(seconds=60):
                    break
                for day in targetDatesTemp:
                    gotRV = take_appointment_time(day, "None", rv)
                    if gotRV:
                        rvs.remove(rv)
                        if len(rvs)>0:
                            rv = rvs[0]
                        else:
                            if (messages):
                                popUpOKLeft("Done, but check the console\nwindow for messages")
                            else:
                                popUpOKLeft("Done")
#                             driver.quit()
                            exit()
                        break
    
    
if __name__ == '__main__':
    print("IF RUNNING IN BACKGROUND DO NOT EXIT THIS WINDOW")
    print("HIT \"CONTROL-C\" TO END THE PROGRAM, AND THEN EXIT THE WINDOW\n") 
    
    containers, date, headless, delivery, allowSameDayAsETA, malport = queryContainers()
#     if containers =="" or date=="":
#         containers, date, headless, delivery, allowSameDayAsETA, malport = queryContainers()
    targetDates, times = queryDatesAndTimes(containers[0]== "ALL", malport)

    if not containers[0]=="ALL" and not delivery:
        for i in range(len(containers)):
            if containers[i]!="":
                containers[i]=containers[i].strip().upper()
                if containers[i][4]!=" ":
                    containers[i] = containers[i][:4]+" " + containers[i][4:]
                containers[i] = containers[i][:11]
                while containers[i][5]=="0":
                    containers[i] = containers[i][:5]+containers[i][6:]
            else:
                del containers[i]
    elif delivery:
        for i in range(len(containers)):
            if containers[i]!="":
                containers[i]=containers[i].strip().upper()
            else:
                del containers[i]
                
                
    conts = ""
    for container in containers:
        conts=conts+container+", "
    conts=conts[:-2]
    
    targetDateString = ""
    for targetDateX in targetDates:
        targetDateString=targetDateString+targetDateX+", "
    targetDateString=targetDateString[:-2]
    
    timeString = ""
    for timeX in times:
        timeString=timeString+timeX+", "
    timeString=timeString[:-2]
    
    
    print("The program is running on the following containers:")
    print(conts)
    print("Which have RVs on "+ date+"\n")
    print("Looking for new RVs on " + targetDateString)
    print("At " +timeString+"\n")
    
    if headless:
        print("Setting up background process, do not exit")
    
    driver, process, globalPort = setupCn(headless)
    if headless:
        print("Setup complete, feel free to exit with \"CTRL-C\"")
    
    def interrupt_handler(a,b):
        print("Quitting...")
        try:
            driver.quit()
        except Exception:
            pass
        try:process.kill()
        except Exception:
            pass
        exit()
        
    signal.signal(signal.SIGINT, interrupt_handler)
#     pthread_sigmask(signal.SIG_IGN, signal.SIGINT)
    repeat = True
    count=0
    while repeat:
        try:
            if getBetterRVs(driver, containers, date, times, targetDates, headless, delivery, allowSameDayAsETA, malport, process, globalPort)!=-1:
#                 getBetterRVs(driver, containers, date, times, targetDates, headless, delivery, allowSameDayAsETA, malport, process, globalPort)
                repeat=False
#             else:
        except InvalidSessionIdException:
            exit()
        except SystemExit:
            exit()
        except SessionNotCreatedException:
            exit()
        except KeyboardInterrupt:
            exit()
        except Exception:
#             popUpOK("FAILED")
            print("Encountered an error:")
            print(exc_info())
            if count>60:
                print("Repeated error, quitting")
                try:
                    driver.quit()
                except Exception:pass
                exit()
            count+=1
            print("Restarting")
            driver.quit()
            driver, process, globalPort = setupCn(headless)
            
#     except Exception:
#         try:
#             driver.quit()
#         except Exception:
#             True
#         print(exc_info())
#         sleep(1000)
    
#     pyinstaller "C:\Users\ssleep\workspace\RV Changer Every Day\Automator\__init__.py" --distpath "J:\Spencer\RV Changer Constant" -y