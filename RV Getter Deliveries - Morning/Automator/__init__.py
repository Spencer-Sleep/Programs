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
import sys
import re
from _datetime import datetime, timedelta

from tkinter import Tk, Button, Label, constants, Checkbutton, BooleanVar, Text,\
    Scrollbar, StringVar, Frame, Radiobutton
from pyautogui import moveTo
from time import sleep
from selenium.webdriver.support.expected_conditions import alert_is_present
from lib2to3.pgen2.tokenize import Ignore
import HelperFunctions
from sys import exc_info
from selenium.webdriver.firefox.options import Options
from openpyxl.reader.excel import load_workbook

def setupCn(headlessBrowser):
    options = Options()
    options.set_headless(headless=headlessBrowser)
    fp = FirefoxProfile();
    fp.set_preference("webdriver.load.strategy", "unstable");
#     fp.set_preference("XRE_NO_WINDOWS_CRASH_DIALOG=1")
     
    driver = Firefox(firefox_profile=fp, log_path=devnull, firefox_options=options)
    driver.get("http://cn.ca/")
#     driver.set_window_position(1920, 0)
#     sleep(30)
    driver.maximize_window()
    
    driver.implicitly_wait(100)
    
    f=open(r"J:\LOCAL DEPARTMENT\Automation - DO NOT MOVE\CN Login.txt", 'r')
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

class Booking:
    def __init__(self):
        self.bookingNumber = ""
        self.timeRange = ""
        self.num = ""

def loadInfo(specificPath):
    routingBook = load_workbook(specificPath)
    infoSheet = routingBook.get_active_sheet()
    
    bookings = []
    
    for row in infoSheet.rows:
        if row[0].value != "None" and not "Booking" in str(row[0].value):
            booking = Booking()
            booking.bookingNumber=row[0].value
            
            timesAvailable = [0, 4, 5,6,7,8,9,10,11,12,13,14,16,18,19,20]
            time1 = timesAvailable.index(int(row[1].value))
            time2 = timesAvailable.index(int(row[2].value))
            booking.timeRange=timesAvailable[time1:time2]
            if len(row)<4 or not row[3] or row[3].value=="None": 
                booking.num=1
            else:
                booking.num=row[3].value
            
            bookings.append(booking)
    
    return bookings

def getRVs(bookings):
    daysForward = ""
    
    if datetime.now().weekday()<2:
        daysForward = 3
    else:
        daysForward = 5
    
    target_Date_dateTime =datetime.now() + timedelta(days = daysForward)
    
    month = str(target_Date_dateTime.month)
    thisDay = str(target_Date_dateTime.day)
    
    if len(month)==1:
        month = "0" + month
    if len(thisDay)==1:
        thisDay = "0" + thisDay
    
if __name__ == '__main__':
    sys.argv = r"a C:\Users\ssleep\Documents\TEST.xlsx".split()
    
    
    specificPath = ''
    for i in range(len(sys.argv)):
        if i!=0:
            specificPath+=sys.argv[i]
            if i != len(sys.argv) - 1:
                specificPath+=" "

    driver = setupCn(False)
    bookings = loadInfo(specificPath)
#     for booking in bookings:
#         print(booking.bookingNumber)
#         print(booking.timeRange)
#         print(booking.num)
#     getRVs(driver, bookings)
    
# pyinstaller "C:\Users\ssleep\workspace\RV Changer\Automator\__init__.py" --distpath "J:\Spencer\RV Changer" --noconsole -y

