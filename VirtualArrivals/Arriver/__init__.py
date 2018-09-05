from selenium import webdriver
from selenium.webdriver import firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
# import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


from os import listdir
from os import path

from sys import argv

from PyPDF2 import PdfFileReader

import re
from os import devnull

driver = ""
testfile = ""


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
    
def recursiveArrive(specificPath):
    for filename in listdir(specificPath):
        if "PARS MANIFESTS" in filename and filename[-4:] == ".pdf" or filename[-4:] == ".PDF":
            arrive(specificPath+'\\'+filename)
        elif(path.isdir(specificPath+"\\"+filename) and not filename=="Flattened"):
            recursiveArrive(specificPath+"\\"+filename)

def arrive(specificPath):
    pdfFileObj = open(specificPath, 'rb')
    pdfReader = PdfFileReader(pdfFileObj)
    
    fields = pdfReader.getFields()
#     print(len(fields)-15)


    for i in range(len(fields)-15):
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))

        containerNumber = ""
        size = ""
        weight = ""
        otherInfo = ""
        consignee = ""
        if i == 0:
#             prefix = str(i) + "."
            wo = fields["WO"].value
            containerNumber = fields["Container Row1"].value
            size = fields["SizeRow1"].value
            weight = float(fields["Weight KGRow1"].value)
            otherInfo = fields["Other info"].value
            consignee = fields["Consignee"].value
        else:
            for j in list(fields.keys()):
                if j==str(i):
                    for k in list(fields[j]["/Kids"]):
                        try:
                            if(k.getObject()['/T']=="WO"):
                                wo=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Container Row1"):
                                containerNumber=k.getObject()['/V']
                            elif(k.getObject()['/T']=="SizeRow1"):
                                size=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Weight KGRow1"):
                                weight=float(k.getObject()['/V'])
                            elif(k.getObject()['/T']=="Other info"):
                                otherInfo=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Consignee"):
                                consignee=k.getObject()['/V']
                        except KeyError:
                            True
        elem = driver.find_element_by_name("container_prefix_dof")
        elem.send_keys(containerNumber[:4])
        elem = driver.find_element_by_name("container_number_dof")
        elem.send_keys(containerNumber[4:11])
        select = Select(driver.find_element_by_name("ddlLoadStatus_dof"))
        select.select_by_visible_text("Load")
        select = Select(driver.find_element_by_name("lineid"))
        select.select_by_visible_text("Hapag-Lloyd Container Line")
        elem = driver.find_element_by_name("ddlSzTyCnt")
        elem.send_keys(size)
        elem = driver.find_element_by_name("cargo_weight")
        elem.send_keys(str(weight))
        select = Select(driver.find_element_by_name("ddWeightUnits"))
        select.select_by_visible_text("Kgs")
        elem = driver.find_element_by_id("CkbCR")
        elem.click()
        try:
            if not ("LCBO" in consignee or "LIQUOR CONTROL" in consignee):
                elem = driver.find_element_by_id("CkbFR")
                elem.click()
        except:
            elem = driver.find_element_by_id("CkbFR")
            elem.click()
        reservation = "import"
        
        if size== "20R86" or size == "40R96":
            m = re.search("Temperature: ", otherInfo)
            n = re.search(r"\.\d+ C", otherInfo[m.end():])
#             print(str(m.end()) + "  " + str(n.start()))
            reservation += otherInfo[m.end():n.start() + m.end()]+"c"
        
        elem = driver.find_element_by_name("bkg_nbr_dof")
        elem.send_keys(reservation)
        select = Select(driver.find_element_by_name("Line"))
        select.select_by_visible_text("Hapag-Lloyd Container Line")
        
        elem = driver.find_element_by_name("Submit")
        elem.click()
        wait = WebDriverWait(driver, 10)
        wait.until(lambda driver: "Equipment is already on Terminal" in driver.page_source or EC.element_to_be_clickable(driver.find_element_by_name("Close")))
        if "Equipment is already on Terminal" in driver.page_source:
            f=open(testfile, "a+")
            f.write("WO: " + wo + "          " + "Container: " + containerNumber + "\n")
            f.close()
            driver.switch_to_default_content()
            driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
            elem = driver.find_element_by_css_selector('a[href*="Gate/VirtualArrive/VirtualArrive.asp"')
            elem.click()
        else:
            elem = driver.find_element_by_name("Close")
            elem.click()
        
        
if __name__ == '__main__':
    
#     argv = r"a J:\All motor routings\2017\Week 29\Hapag-Lloyd".split()
#     argv = r"a C:\Users\ssleep\Documents\Programming\Hapag Dispatchmate\Thursday\LCBO\601331975 PARS MANIFESTS.pdf".split()
    
    driver = setupEterm()
    
    specificPath = ''
    for i in range(len(argv)):
        if i!=0:
            specificPath+=argv[i]
            if i != len(argv) - 1:
                specificPath+=" "
 
                     
    testfile = specificPath[:specificPath.rfind('\\')] + "\\" + "Containers already in Eterm.txt"
    if path.isdir(specificPath):
        recursiveArrive(specificPath)
    elif "PARS MANIFESTS" in specificPath and (specificPath[-4:] == ".pdf" or specificPath[-4:] == ".PDF"):
        arrive(specificPath)