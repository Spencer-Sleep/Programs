from win32api import GetKeyState
from tkinter import Button, Tk, Label, Entry
from sys import argv

from os import listdir
from os import path

from openpyxl import load_workbook

from selenium import webdriver
from selenium.webdriver import firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
# import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import re
from os import devnull
from tkinter.constants import CURRENT

class Container(object):
    number=""
    size=""

def setupEterm():
    fp = FirefoxProfile();
    fp.set_preference("webdriver.load.strategy", "unstable");
     
    driver = webdriver.Firefox(firefox_profile=fp, log_path=devnull)
    driver.get("http:/etermsys.com/")
    
    # time.sleep(10)
    driver.switch_to.frame("sideNavBar")
    elem = driver.find_element_by_name("UserName")
    elem.clear()
    f=open(r"C:\Automation\Eterm Login.txt", 'r')
    read = f.readline()
    m = re.search("username: *", read)
    read = read[m.end():]
    elem.send_keys(read)
    elem = driver.find_element_by_name("Password")
    read = f.readline()
    m = re.search("password: *", read)
    read = read[m.end():]
    elem.send_keys(read)
    f.close()
    elem = driver.find_element_by_name("button1")
    elem.click()
    
    
    driver.implicitly_wait(10)
    
    driver.switch_to_default_content()
    driver.switch_to_frame("main")
    select = Select(driver.find_element_by_name("ddl_terminal_select"))
    select.select_by_visible_text("Seaport Intermodal")
    elem = driver.find_element_by_name("submit_eTERM")
    elem.click()
    
    
    
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
    elem = driver.find_element_by_css_selector('a[href*="Gate/VirtualArrive/VirtualArrive.asp"')
    elem.click()
    return driver

def arrive(container, driver):
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
    elem = driver.find_element_by_name("container_prefix_dof")
    elem.send_keys(container.number[:4])
    elem = driver.find_element_by_name("container_number_dof")
    elem.send_keys(container.number[4:11])
    select = Select(driver.find_element_by_name("ddlLoadStatus_dof"))
    select.select_by_visible_text("Load")
    select = Select(driver.find_element_by_name("lineid"))
    select.select_by_visible_text("Seaport Intermodal")
    elem = driver.find_element_by_name("ddlSzTyCnt")
    elem.send_keys(container.size)
    
    elem = driver.find_element_by_name("cargo_weight")
    elem.send_keys(1)
    
    elem = driver.find_element_by_name("eqcomments")
    elem.send_keys("DELL MIDDLETOWN")
    
    elem = driver.find_element_by_name("bkg_nbr_dof")
    elem.send_keys("middletown")
    select = Select(driver.find_element_by_name("Line"))
    select.select_by_visible_text("Seaport Intermodal")
    
    elem = driver.find_element_by_name("Submit")
    elem.click()
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: "Equipment is already on Terminal" in driver.page_source or "Equipment is already on facility" in driver.page_source or EC.element_to_be_clickable(driver.find_element_by_name("Close")))
    if "Equipment is already on Terminal" in driver.page_source or "Equipment is already on facility" in driver.page_source:
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
        elem = driver.find_element_by_css_selector('a[href*="Gate/VirtualArrive/VirtualArrive.asp"')
        elem.click()
    else:
        elem = driver.find_element_by_name("Close")
        elem.click()
        
def loadInfo(specificPath):
    dellLog = load_workbook(specificPath)
    
    currentSheet =dellLog.worksheets[0]
    numCol=""
    sizeCol=""
    
    containers = []
    
    for cell in next(currentSheet.rows):
        if cell.value == "CONTAINER#":
            numCol = cell.col_idx - 1
        elif cell.value == "SIZE":
            sizeCol = cell.col_idx - 1

    
    hiddenRows = []
    for rowNum, rowDimension in currentSheet.row_dimensions.items():
        if rowDimension.hidden == True:
            hiddenRows.append(rowNum)
    i=1
    for row in currentSheet.rows: 
        if row[numCol].value == "CONTAINER#":
            continue
        i=i+1
        if not i in hiddenRows:
            if row[numCol].value:
                container = Container()
                container.number = str(row[numCol].value)
                container.size = str(row[sizeCol].value)
                containers.append(container)
    

    return containers

if __name__ == '__main__':
    
#     argv = r"a J:\All motor routings\2017\Week 29\Hapag-Lloyd".split()
#     argv = r"a C:\Users\ssleep\Documents\Programming\Hapag Dispatchmate\Thursday\LCBO\601331975 PARS MANIFESTS.pdf".split()
    
#     argv=r"a C:\Users\ssleep\Downloads\DELL LOG.xlsx".split()
    
    driver = setupEterm()
    
    specificPath = ''
    for i in range(len(argv)):
        if i!=0:
            specificPath+=argv[i]
            if i != len(argv) - 1:
                specificPath+=" "

    containers = loadInfo(specificPath)
    
    for container in containers:
        arrive(container, driver)