
from pyotrs.lib import Client
from warnings import filterwarnings
import sys
from HelperFunctions import popUpOK
from os.path import isdir

from exchangelib import DELEGATE, Account, Credentials, ewsdatetime
from exchangelib.attachments import FileAttachment, ItemAttachment
from exchangelib.items import Message
from exchangelib import Mailbox

from tkinter import Tk, Button, Label, constants, Checkbutton, BooleanVar, Text,\
    Scrollbar, StringVar, Frame, Radiobutton
from PyPDF2.pdf import PdfFileWriter, PdfFileReader
from time import sleep
import os

from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver import Firefox
from os.path import devnull
from re import compile, findall, search
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from exchangelib.extended_properties import ExtendedProperty
from HelperFunctions import popUpOK
from exchangelib import Q
from datetime import timedelta
from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone
from ContainerSizeInfo import standardSize

from selenium.webdriver.firefox.options import Options

import warnings

def sendRelease(container, driver, account):
    success=True
    print(container)
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='topnavframe.asp']"))
    
    driver.find_element_by_css_selector('a[href*="MenuNavFrame.asp?MenuID=5"').click()
    
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
    
    driver.find_element_by_css_selector('a[href*="inventory/udsearch.asp"').click()
    
    
    failed = True
    while failed:
        try:
            driver.switch_to_default_content()
            driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
            driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='udSearchForm.asp?func=']"))
            failed = False
        except:
            pass
    
    elem = driver.find_element_by_name("unitprefix")
    elem.clear()
    elem.send_keys(container[:4])
    elem = driver.find_element_by_name("unitnumber")
    elem.clear()
    elem.send_keys(container[4:11])
    
    driver.find_element_by_css_selector("Input[class='Button'][name='Submit']").click()
    
#     driver.switch_to_default_content()
#     driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
# #     frame = driver.find_element_by_css_selector("frame[src='udSearchResult.asp']")
#     driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='udSearchResult.asp']"))
#     wait = WebDriverWait(driver, 10)
#     wait.until(lambda driver: "Your query did not return any results" in driver.page_source or
#                 EC.element_to_be_clickable(driver.find_element_by_css_selector("a[href*='unitDisposition.asp?eqid']")))
    found=False
    while not found:
        try:
#             print(driver.page_source)
            driver.switch_to_default_content()
            driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
            driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='udSearchResult.asp']"))
            if ("Your query did not return any results" in driver.page_source):
                found=True
            driver.implicitly_wait(0)
            driver.find_element_by_css_selector("a[href*='unitDisposition.asp?eqid']")
            found=True
            driver.implicitly_wait(60)
        except:
            driver.implicitly_wait(60)
    if("Your query did not return any results" in driver.page_source):
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='topnavframe.asp']"))
        driver.find_element_by_css_selector('a[href*="MenuNavFrame.asp?MenuID=1"').click()
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
        driver.find_element_by_css_selector("a[href='Gate/VirtualArrive/VirtualArriveSearch.asp']").click()
        
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='VirtualArriveForm.asp']"))
        driver.find_element_by_id("1").send_keys(container[:4])
        driver.find_element_by_id("2").send_keys(container[4:11])
        driver.find_element_by_css_selector("Input[class='Button'][name='Submit']").click()
        
        found = False
        while not found:
            try:
                driver.switch_to_default_content()
                driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
                driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='VirtualArriveResult.asp']"))
                if ("Your query did not return any results" in driver.page_source):
                    found=True
                driver.implicitly_wait(0)
                driver.find_element_by_css_selector("a[href*='javascript:EditVirtualArrive(']")
                found=True
                driver.implicitly_wait(60)
            except:
                driver.implicitly_wait(60)
        
        if (not "Your query did not return any results" in driver.page_source):
            driver.find_element_by_css_selector("a[href*='javascript:EditVirtualArrive(']").click()
            
            parentWindow = driver.current_window_handle;
            handles =  driver.window_handles;
            for windowHandle in handles:
                if(not windowHandle==parentWindow):
                    driver.switch_to.window(windowHandle);
                    if(driver.find_element_by_id("CkbFR").get_attribute("checked")=="true"):
                        driver.find_element_by_id("CkbFR").click()
                        driver.find_element_by_css_selector("Input[class='Button'][name='Submit']").click()
                        wait = WebDriverWait(driver, 10)
                        wait.until(lambda driver: "Your information has been saved." in driver.page_source)
                        
                    driver.close(); 
                    driver.switch_to.window(parentWindow);
        else:
            print(container + " NOT IN SYSTEM")
            success=False
            
    else:
        driver.find_element_by_css_selector("a[href*='unitDisposition.asp?eqid']").click()
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
        elem =driver.find_element_by_css_selector("input[name='frelease'][value='1']")
        if elem.get_attribute("checked")!="true":
            elem.click()
#             driver.find_element_by_name("comments").send_keys("\n"+transaction)
            driver.find_element_by_css_selector("Input[class='Button'][name='Submit']").click()
            
        if driver.find_element_by_css_selector("input[name='lfd']").get_attribute("value")=="":
            m = Message(
                account=account, 
                subject='Missing LFD for container: ' + container, 
                body='Missing LFD for container: ' + container, 
                to_recipients=[Mailbox(email_address='yard@seaportint.com')]
            )
            m.send()
            
    return success

def setupEterm():
#     fp = FirefoxProfile();
#     fp.set_preference("webdriver.load.strategy", "unstable");
#     driver = Firefox(log_path=devnull)
#     driver = Firefox(firefox_profile=fp, log_path=devnull)
    print("Starting Firefox in background...")
#     print("Starting Firefox")
    
    options = Options()
    options.set_headless(headless=True)
#     options.set_headless()
#     options.set_headless(headless=False)
    fp = FirefoxProfile();
    fp.set_preference("webdriver.load.strategy", "unstable");
       
    driver = Firefox(firefox_profile=fp, log_path=devnull, firefox_options=options)
    driver.get("http:/etermsys.com/")
    
    driver.maximize_window()
    
    # time.sleep(10)
    driver.implicitly_wait(20)
    driver.switch_to.frame("sideNavBar")
    try:
        elem = driver.find_element_by_name("UserName")
        elem.clear()
        f=open(r"C:\Automation\Eterm Login.txt", 'r')
        read = f.readline()
        m = search("username: *", read)
        read = read[m.end():]
        elem.send_keys(read)
        elem = driver.find_element_by_name("Password")
        read = f.readline()
        m = search("password: *", read)
        read = read[m.end():]
        elem.send_keys(read)
        f.close()
        elem = driver.find_element_by_name("button1")
        elem.click()
    except:
        driver.switch_to_default_content()
        driver.switch_to_frame("main")
        wait = WebDriverWait(driver, 1000)
        wait.until((EC.element_to_be_clickable((By.NAME, "ddl_terminal_select"))))
        
    
    
    driver.switch_to_default_content()
    driver.switch_to_frame("main")
    select = Select(driver.find_element_by_name("ddl_terminal_select"))
    select.select_by_visible_text("Seaport Intermodal")
    elem = driver.find_element_by_name("submit_eTERM")
    elem.click()
    
    return driver    

def getCNs(driver, account):
    bgC = "lavender"
    top = Tk()
    top.config(bg = bgC)
    L1 = Label(top, text="Please enter the Container Numbers to release", bg = bgC, padx = 20)
    L1.config(font=("serif", 16))
    L1.grid(row=0, column=0, sticky=constants.W+constants.E, columnspan=3)
    
    S1=Scrollbar(top, orient='vertical')
    S1.grid(row=1, column=2, sticky=constants.N + constants.S)
    S2=Scrollbar(top, orient='horizontal')
    S2.grid(row=2, column=0, sticky=constants.E + constants.W, columnspan=2)
    
    T1 = Text(top, height = 20, width = 97, xscrollcommand = S2.set, yscrollcommand=S1.set, wrap = constants.NONE)
    T1.grid(row=1, column=0, columnspan=2)
    
    def callbackCont():
        if T1.get("1.0", constants.END).strip()=="":
            popUpOK("Please list the container numbers to release")
        else:
            pattern = compile(r"[a-zA-Z]{4}[0-9]{6,7}")
            deleteNumber=1
            cNumbers=T1.get("1.0", constants.END).splitlines()
            for container in cNumbers:
                container=container.replace(" ", "")
#                 if len(container)>9:
                if not pattern.match(container):
                    print("Invalid container number: \n" + container)
                    deleteNumber+=1
                else:                
                    if sendRelease(container, driver, account):
    #                     T1.delete("1.0", "2.0")
                        T1.delete(str(deleteNumber)+".0", str(deleteNumber+1)+".0")
                    else:
                        deleteNumber+=1
#             top.destroy()

    def quitThis():
        driver.quit()
        sys.exit()

    MyButton = Button(top, text="OK", command=callbackCont)
    MyButton.grid(row=5, column=0, sticky=constants.W+constants.E, padx = 20, pady = 10)
    MyButton.config(font=("serif", 30), bg="green")
    
    MyButton = Button(top, text="QUIT", command=quitThis)
    MyButton.grid(row=5, column=1, sticky=constants.W+constants.E, padx = 20, pady = 10)
    MyButton.config(font=("serif", 30), bg="red")
      
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
    
    top.mainloop()
    
#     return cNumbers[0]

if __name__ == '__main__':
    
    warnings.filterwarnings("ignore")
    
    driver = setupEterm()
    
    credentials = Credentials(
    username='ssleep@seaportint.com',  # Or myusername@example.com for O365
    password='ss#99PASS'
    )
    account = Account(
        primary_smtp_address='torrns@seaportint.com', 
        credentials=credentials, 
        autodiscover=True, 
        access_type=DELEGATE
    )
    cNumbers = getCNs(driver, account)
    
    driver.quit()
    
#     pyinstaller "C:\Users\ssleep\workspace\Freight Releaser\Automator\__init__.py" --distpath "J:\Spencer\Freight Releaser" -y