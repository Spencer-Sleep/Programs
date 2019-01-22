# import email
# import imaplib
# import os
from exchangelib import DELEGATE, Account, Credentials, ewsdatetime
from exchangelib.attachments import FileAttachment, ItemAttachment
from exchangelib.items import Message
from time import sleep
# from datetime import 
# from img2pdf import 
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
import atexit
from sys import exc_info
from tkinter import Tk

import warnings

from tkinter import Tk, Button, Label, constants
from exchangelib import Mailbox
from urllib3.exceptions import InsecureRequestWarning


class Container(object):
    number=""
    size=""
    weight = ""
    CCN=""
        

def updateInfo(content, driver, done):
    print("\nUpdating Virtual Arrival Info for: ")
    pattern = compile(r"[a-zA-Z]{4}[0-9]{7}")
    matches = findall(pattern, content)
    for match in matches:
#         if match in done:
#             continue
#         else:
#             done.append(match)
#         if not popped:
#             popUpOK("HERE WE GO")
#             popped=True
        container = Container()
        container.number=match
        startIndex = content.find(match)
        startIndex = content[startIndex:].find("<span")+startIndex
        startIndex = content[startIndex:].find(">")+startIndex+1
        endIndex = content[startIndex:].find("<")+startIndex
        container.size = standardSize(content[startIndex:endIndex])
        startIndex = content[startIndex:].find("<span")+startIndex
        startIndex = content[startIndex:].find(">")+startIndex+1
        endIndex = content[startIndex:].find("<")+startIndex
        container.weight = content[startIndex:endIndex]
        for _ in range(4):
            startIndex = content[startIndex:].find("<span")+startIndex+1
        startIndex = content[startIndex:].find(">")+startIndex+1
        endIndex = content[startIndex:].find("<")+startIndex
        container.CCN = "\nCargo Control Number: 9082"+content[startIndex:endIndex]
        
        updateEterm(container, driver)

def updateEterm(container, driver):
    print(container.number)
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='topnavframe.asp']"))
    driver.find_element_by_css_selector('a[href*="MenuNavFrame.asp?MenuID=1"').click()
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
    driver.find_element_by_css_selector("a[href='Gate/VirtualArrive/VirtualArriveSearch.asp']").click()
    
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='VirtualArriveForm.asp']"))
    driver.find_element_by_id("1").send_keys(container.number[:4])
    driver.find_element_by_id("2").send_keys(container.number[4:11])
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
    #     wait = WebDriverWait(driver, 10)
#             wait.until(lambda driver: "Your query did not return any results" in driver.page_source or
#                     EC.element_to_be_clickable(driver.find_element_by_css_selector("a[href*='javascript:EditVirtualArrive(']")))
    if (not "Your query did not return any results" in driver.page_source):
        driver.find_element_by_css_selector("a[href*='javascript:EditVirtualArrive(']").click()
        parentWindow = driver.current_window_handle;
        handles =  driver.window_handles;
        for windowHandle in handles:
            if(not windowHandle==parentWindow):
                driver.switch_to.window(windowHandle);
    #             elem = driver.find_element_by_id("CkbCR")
                elem = driver.find_element_by_name("eqcomments")
                if not "Cargo Control Number" in elem.get_attribute('value'):
                    elem.send_keys(container.CCN)
            
                elem = driver.find_element_by_name("ddlSzTyCnt")
                if(elem.get_attribute("value")==""):
                    elem.send_keys(container.size)
                
                elem = driver.find_element_by_name("bkg_nbr_dof")
                if(elem.get_attribute("value")==""):
                    elem.send_keys("import")
                
                elem = driver.find_element_by_name("cargo_weight")
                if(elem.get_attribute("value")=="" or elem.get_attribute("value")=="0" or elem.get_attribute("value")=="1"):
                    elem.clear()
                    elem.send_keys(container.weight)
                    Select(driver.find_element_by_name("ddWeightUnits")).select_by_visible_text("Lbs")
                
                elem1 = Select(driver.find_element_by_name("Line"))
                elem2 = Select(driver.find_element_by_name("lineid"))
                try:
                    if(elem1.first_selected_option.text==""):
                        elem1.select_by_visible_text(elem2.first_selected_option.text)
                except:
                    elem1.select_by_visible_text(elem2.first_selected_option.text)
                
                failed = True
                while not failed:
                    try:
                        driver.find_element_by_css_selector("Input[class='Button'][name='Submit']").click()
                        wait = WebDriverWait(driver, 10)
                        wait.until(lambda driver: "Your information has been saved." in driver.page_source)
                        failed = False
                    except:
                        pass
                driver.close()
#                 print(parentWindow)
                driver.switch_to.window(parentWindow)
    
def sendRelease(container, transaction, driver, account):
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
                    elem = driver.find_element_by_id("CkbCR")
    #                 try:
    # eqcomments
                    if(elem.get_attribute("checked")=="true"):
                        driver.find_element_by_id("CkbCR").click()
                        driver.find_element_by_name("eqcomments").send_keys("\n"+transaction)
                    
                        elem = driver.find_element_by_name("bkg_nbr_dof")
                        if(elem.get_attribute("value")==""):
                            elem.send_keys("import")
                        elem1 = Select(driver.find_element_by_name("Line"))
                        elem2 = Select(driver.find_element_by_name("lineid"))
                        try:
                            if(elem1.first_selected_option.text==""):
                                elem1.select_by_visible_text(elem2.first_selected_option.text)
                        except:
                            elem1.select_by_visible_text(elem2.first_selected_option.text)
                        if driver.find_element_by_name("cargo_weight").get_attribute("value")=="":    
                            driver.find_element_by_name("cargo_weight").send_keys("0")
                        Select(driver.find_element_by_name("ddWeightUnits")).select_by_visible_text("Kgs")
                        driver.find_element_by_css_selector("Input[class='Button'][name='Submit']").click()
                        wait = WebDriverWait(driver, 10)
                        wait.until(lambda driver: "Your information has been saved." in driver.page_source)
                        
                        
                    driver.close(); 
                    driver.switch_to.window(parentWindow);
    else:
        driver.find_element_by_css_selector("a[href*='unitDisposition.asp?eqid']").click()
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
        elem =driver.find_element_by_css_selector("input[name='crelease'][value='1']")
        if elem.get_attribute("checked")!="true":
            elem.click()
            driver.find_element_by_name("comments").send_keys("\n"+transaction)
            driver.find_element_by_css_selector("Input[class='Button'][name='Submit']").click()
            
        if driver.find_element_by_css_selector("input[name='lfd']").get_attribute("value")=="":
            m = Message(
                account=account, 
                subject='Missing LFD for container: ' + container, 
                body='Missing LFD for container: ' + container + "\n\n\nTHIS IS AN AUTOMATED MESSAGE, DO NOT REPLY", 
                to_recipients=[Mailbox(email_address='nicholas.cardin@seaportint.com'), Mailbox(email_address='terminal@seaportint.com'),]
            )
            m.send()
            
    return True
    
def release(content, driver, account):
    transactionIndex = content.find("Transaction:")
    transaction = content[transactionIndex: content[transactionIndex:].find("\n")+transactionIndex]
    ccnIndex =  content.find("Cargo Control Number:")
    releaseIndex = content.find("Release Office: ")
    ccn = content[ccnIndex+22:releaseIndex].strip()
    transaction = transaction+"\n"+content[ccnIndex: content[ccnIndex:].find("\n")+ccnIndex]
    deliveryIndex = content.find("Delivery Instructions")-2
    
    containerIndex = content.find("Container ID(s):")+17
    containers = []
    
    
    
    if containerIndex<deliveryIndex:
        while containerIndex<deliveryIndex:
            containers.append(content[containerIndex:content[containerIndex:].find(",")+containerIndex].strip())
            containerIndex = content[containerIndex:].find(",")+1+containerIndex
        print("\nReleasing the following containers: ")
        for cont in containers:
            sendRelease(cont, transaction, driver, account)
    else:
        m = Message(
            account=account, 
            subject='CCN is Released: ' + ccn, 
            body='PB #(if Seaport CCN): ' + ccn[-6:] + "\n\nThe RNS release message did not have a container number in it\n\n\nTHIS IS AN AUTOMATED MESSAGE, DO NOT REPLY", 
            to_recipients=[Mailbox(email_address='terminal@seaportint.com'),]
        )
        m.send()
        print("\nEmail with no container number")
        
    
def setupEterm():
#     fp = FirefoxProfile();
#     fp.set_preference("webdriver.load.strategy", "unstable");
#     driver = Firefox(log_path=devnull)
#     driver = Firefox(firefox_profile=fp, log_path=devnull)
    print("Starting Firefox in background...")
    
    options = Options()
    options.set_headless(headless=True)
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

class Flag(ExtendedProperty):
    property_tag = 0x1090
    property_type = 'Integer'

if __name__ == '__main__':
    driver = setupEterm()
    
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)
    
    def doStuff(credentials, yardAccount, account, done):
#         while(True):
        account.inbox.refresh()
        yardAccount.inbox.refresh()
        print("Unread Emails in RNS Queue: " + str(account.inbox.unread_count))
        for item in account.inbox.filter(is_read=False, sender="cadex@custombroker.com"):
            if("Goods Released" in item.body):
                release(item.body, driver, account)
            item.flag=1
            item.is_read=True
#             else:
#                 print("Unrecognized RNS Message: \n" + item.body)
#                 popUpOK("Unrecognized RNS Message, see inbox or\nconsole window for details",16,top)
#                 item.is_read=True
#                 item.flag=2
#             item.flag=1
#             item.is_read=True
            item.save()
        q = (Q(subject__contains="Back Haul") | Q(subject__icontains="BH ")) & Q(sender="irfan.ghansar@mdstrucking.net") & ~Q(flag=1)
        for item in yardAccount.inbox.filter(q):
            updateInfo(item.body, driver, done)
#             print(item.flag)
            item.flag=1
            item.save()
#             sleep(30)
#             top.after(3000, doStuff())
        print("Done\n")
        top.after(60000, doStuff, credentials, yardAccount, account, done)
            
    bgC = "lavender"
    top = Tk()
    top.config(bg = bgC)
    
    def callbackOK():
#         sys.exit()
        driver.quit()
        top.destroy()
        exit()
        
    MyButton = Button(top, text="QUIT", command=callbackOK)
    MyButton.grid(row=0, column=0, sticky=constants.W+constants.E, padx = 20, pady = (0,20))
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
    
    yardAccount = Account(
    primary_smtp_address='toryard@seaportint.com', 
    credentials=credentials, 
    autodiscover=True, 
    access_type=DELEGATE
    )
    done = []
    tz = EWSTimeZone.localzone()
    Message.register('flag', Flag)
    
    
    top.after(0, doStuff, credentials, yardAccount, account, done)
    top.mainloop() 
#         print("quitting")
#         sleep(2)
#         print("beofre")
#         try:
# #             driver.close()
#             driver.quit()
#         except:
#             print(exc_info())
#         print("aftr")
#         sleep(10)
# pyinstaller "C:\Users\ssleep\workspace\RNS Releaser\Automator\__init__.py" --distpath "J:\Spencer\RNS Releaser" -y