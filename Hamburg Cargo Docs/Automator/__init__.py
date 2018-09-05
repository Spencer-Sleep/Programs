from selenium import webdriver
from selenium.webdriver import firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC,\
    expected_conditions, wait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from os import devnull, path

from concurrent.futures import ThreadPoolExecutor 

from threading import Lock

import re

from pyautogui import click, moveTo
from pyautogui import press

from os import listdir
from openpyxl import load_workbook

from itertools import islice

from tkinter import Button, Tk, Label, Entry, constants
# from pyautogui import click
from time import sleep

from datetime import datetime

from win32api import GetKeyState
from future.backports.http.client import GONE

from sys import argv
from sys import exit
from tkinter.constants import CURRENT
from _datetime import timedelta
from HelperFunctions import done
from selenium.webdriver.common import by
import PyPDF2
import HelperFunctions



class AnyEc:
    """ Use with WebDriverWait to combine expected_conditions
        in an OR.
    """
    def __init__(self, *args):
        self.ecs = args
    def __call__(self, driver):
        for fn in self.ecs:
            try:
                if fn(driver): return True
            except:
                pass

CONTAINERNUMBER= "Container"
CCN= "CCN"
WEIGHT= "Weight"
PIECECOUNT="Piece"
TERMINAL = "Terminal"
POL = "Port of Loading"
DESCRIPTION = "Description"            
            
class Container(object):
    def __init__(self):
        self.properties = {CONTAINERNUMBER: "",
               CCN: "",
#                WEIGHLBS: "",    
               WEIGHT: "",
               PIECECOUNT: "",
               TERMINAL: "",
               POL: "",
               DESCRIPTION:""}

def popUpOKLeft(text1, text2, textSize = 16):
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

def changeContants(container, ccn, weight, piececount, terminal):
    CONTAINERNUMBER= container
    CCN= ccn
    WEIGHT= weight
    PIECECOUNT=piececount
    TERMINAL = terminal

def loadContainerInfoHapag(specificPath):
    containers = []
    for filename in listdir(specificPath):
        if "PARS MANIFESTS" in filename and filename[-4:] == ".pdf" or filename[-4:] == ".PDF":
            containers = containers + loadContainerInfoHapagRecurse(specificPath+'\\'+filename)
        elif(path.isdir(specificPath+"\\"+filename) and not filename=="Flattened"):
            containers = containers + loadContainerInfoHapag(specificPath+"\\"+filename)
    
    return containers

def loadContainerInfoHapagRecurse(specificPath):
    pdfFileObj = open(specificPath, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    
    fields = pdfReader.getFields()
#     print(len(fields)-15)

    containers = []
    for i in range(len(fields)-15):
        container = Container()
#         print(i)
#         wo = ""
#         terminal = ""
#         containerNumber = ""
#         size = ""
#         weight = ""
        
        if i == 0:
#             prefix = str(i) + "."
            container.properties[PIECECOUNT] = fields["QtyRow1"].value    
            container.properties[POL] = fields["undefined"].value
            container.properties[TERMINAL] = fields["Port of discharge"].value
            container.properties[CONTAINERNUMBER] = fields["Container Row1"].value
            container.properties[WEIGHT] = fields["Weight KGRow1"].value
            container.properties[DESCRIPTION] = fields["Description of goods"].value
#             container.properties[PIECECOUNT] = float(fields["Weight KGRow1"].value)
        else:
            for j in list(fields.keys()):
                if j==str(i):
                    for k in list(fields[j]["/Kids"]):
                        try:
                            if(k.getObject()['/T']=="undefined"):
                                container.properties[POL]=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Port of discharge"):
                                container.properties[TERMINAL]=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Container Row1"):
                                container.properties[CONTAINERNUMBER]=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Weight KGRow1"):
                                container.properties[WEIGHT]=float(k.getObject()['/V'])
                            elif(k.getObject()['/T']=="Description of goods"):
                                container.properties[DESCRIPTION]=k.getObject()['/V']
                            elif(k.getObject()['/T']=="QtyRow1"):
                                container.properties[PIECECOUNT]=k.getObject()['/V']
                        except KeyError:
                            True
        containers.append(container)
    return containers


def loadContainerInfo(specificPath):
    routingBook = load_workbook(specificPath)
    routing = routingBook.active
    
    containers = []
    
    colDict = {CONTAINERNUMBER: "",
               CCN: "",
#                WEIGHLBS: "",    
               WEIGHT: "",
#                PIECES: "",
               PIECECOUNT: "",
               TERMINAL: ""} 
    
    
    
    CSX = False
    
    for cell in next(routing.rows):
        for contProperty in colDict:
            if cell.value and contProperty in cell.value:
                colDict[contProperty] = cell.col_idx - 1
        if cell.value and "Transport_Mode" in cell.value:
            if not (cell.offset(1,0).value and "SPT_TOR" in cell.offset(1,0).value):
                CSX = True
            else:
                colDict[TERMINAL] = cell.col_idx+2
                
    popUpMessage = ""
    for column, content in [(CONTAINERNUMBER, "container number"), (WEIGHT, "weight"), (CCN, "CCN")]:
        if colDict[column]=="":
            popUpMessage = popUpMessage + "Could not find a column named: \"" + column + "\",\n which should contain the " + content + ".\n\n"
    if colDict[TERMINAL]=="" and not CSX:
        popUpMessage = popUpMessage + "Could not find a column named: \"Terminal\", which should contain the terminal.\nIf this is supposed to be a CSX vessel, there should instead be a column named \"Transport_Mode\"\n"
    if popUpMessage != "":
        HelperFunctions.popUpOK(popUpMessage)
        driver.quit()
        exit()
                
#     changeContants("Container #", "CCN", "Weight", "Piece", "")
    
#     changeContants("Container", "CCN", "Weight", "Piece", "")
    
    i=1
    hiddenRows = []
    pattern = re.compile(r"[A-Z]{4}[0-9]{7}")
    ccnPattern = re.compile(r"20C0(PARS)?[0-9]{6}")
    
    invalids = "The following values are invalid:"
    invalidConts = ""
    invalidWeight=""
    invalidPieces=""
    invalidCCN=""
    for rowNum, rowDimension in routing.row_dimensions.items():
        if rowDimension.hidden == True:
            hiddenRows.append(rowNum)
    for row in routing.rows:
        if not i in hiddenRows and i>1:
            container = Container()
            for contProperty in container.properties:
                if contProperty in colDict and colDict[contProperty]!="":
                    container.properties[contProperty] = str(row[colDict[contProperty]].value).upper()
            
            notBlank = False
            for prop in container.properties.values():
                notBlank = (prop != "NONE" and prop != "") or notBlank
            if notBlank:
                if CSX:
                    container.properties[TERMINAL] = "CSX"
                
                if container.properties[PIECECOUNT]:
                    container.properties[PIECECOUNT] = container.properties[PIECECOUNT].split(" ")[0]    
                
                if not pattern.fullmatch(container.properties[CONTAINERNUMBER]):
                    invalidConts = invalidConts + ("\nContainer number: " + container.properties[CONTAINERNUMBER])
                try:
                    int(float(container.properties[WEIGHT]))
                except:
                    invalidWeight = invalidWeight + "\nWeight for container " + container.properties[CONTAINERNUMBER]+ ": " +container.properties[WEIGHT]
                if not container.properties[PIECECOUNT]=="NONE":
                    try:
                        int(float(container.properties[PIECECOUNT]))
                    except:
                        invalidPieces = invalidPieces + "\nPiece count for container " + container.properties[CONTAINERNUMBER]+ ": " + container.properties[PIECECOUNT]
    #             print(container.properties[CCN])
    #             print(ccnPattern.fullmatch(container.properties[CCN]))
                
                if not ccnPattern.fullmatch(container.properties[CCN]):
                    invalidCCN = invalidCCN + "\nCCN for container " + container.properties[CONTAINERNUMBER]+ ": " + container.properties[CCN]
                
                containers.append(container)
        i+=1
    invalidExtras = invalidConts+invalidCCN + invalidWeight+invalidPieces
    if invalidExtras != "":
        popUpOKLeft(invalids, invalidExtras)
        exit()
        
    return containers
    
def setupPortal():
    fp = FirefoxProfile();
    fp.set_preference("webdriver.load.strategy", "unstable");
    
    driver = webdriver.Firefox(firefox_profile=fp, log_path=devnull)
    driver.get("http://www.cbsa-asfc.gc.ca/prog/manif/portal-portail/menu-eng.html")
    driver.maximize_window()
    driver.implicitly_wait(6000)
    
    elem = driver.find_element_by_css_selector('a[href*="https://apps-cbsa-asfc.fjgc-gccf.gc.ca/LCS/?l=eng&t=https://apps.cbsa-asfc.gc.ca/GCKey"]')
    elem.click()
    
    elem = driver.find_element_by_id("token1")
    f=open(r"C:\Automation\CBSA Login.txt", 'r')
    read = f.readline()
    m = re.search("username: *", read)
    read = read[m.end():]
    elem.send_keys(read)
    
    elem = driver.find_element_by_id("token2")
    read = f.readline()
    m = re.search("password: *", read)
    read = read[m.end():]
    elem.send_keys(read)
    f.close()
    
    elem = driver.find_element_by_css_selector('[title="Connect to the GCKey service"]')
    elem.click()
    
    
    elem = driver.find_element_by_id("continue")
    elem.click()
    
    
    elem = driver.find_element_by_name("_acceptEvent")
    elem.click()
    
    return driver


def makeCargoDocs(containers, driver):
    for container in containers:
        while not "Filter Submitted Documents list to view the following:" in driver.page_source:
            driver.find_element_by_id("tradeDocumentsTab").click()
            
            
        driver.find_element_by_name("_create").click()
        
        Select(driver.find_element_by_id("docTypeSelected")).select_by_visible_text("Highway Cargo Document")
        driver.find_element_by_id("submitButton").click()
        
        if container.properties[CCN] != "":
            driver.find_element_by_id("ccnDocumentNumberForm.documentNumberWithoutClientCode").send_keys(container.properties[CCN][4:])
        else:
            driver.find_element_by_id("ccnDocumentNumberForm.documentNumberWithoutClientCode").send_keys("20C0PARS")
            while not GetKeyState(13)<0:
                True
        
        fiveDays = datetime.now()+timedelta(days=5)
        
        currentDate = str(fiveDays.year)
        
        if len(str(fiveDays.month))==1:
            currentDate+="0"
        currentDate+=str(fiveDays.month) 
        
        if len(str(fiveDays.day))==1:
            currentDate+="0" 
        currentDate+=str(fiveDays.day)
        
        driver.find_element_by_id("datePicker").send_keys(currentDate)
        
        Select(driver.find_element_by_id("csaShipmentOptionId")).select_by_visible_text("No")
        
        Select(driver.find_element_by_id("cargoGeneralForm.consolidationIndicator")).select_by_visible_text("No")
        
#         driver.find_element_by_css_selector('a[class*="toggle-link-collapse"]').click()
        driver.find_element_by_css_selector('a[aria-controls="toggle-content-id2"]').click()
#         sleep(5)
        driver.find_element_by_id("containerIdentifier").send_keys(container.properties[CONTAINERNUMBER])
#         driver.fin
        driver.find_element_by_id("cargoPortsTabBottom").click()
        
        city = "New Jersey"
        state= "New Jersey"
        crossing = "427"
        
        if container.properties[TERMINAL] == "PACKER":
            city = "Philadelphia"
            state = "Pennsylvania"
        elif container.properties[TERMINAL] == "NYCT":
            city = "New York"
            state = "New York"
        elif container.properties[TERMINAL] == "CSX":
            city = "Buffalo"
            state = "New York"
            crossing = "410"
        
        driver.find_element_by_id("cargoPortsForm.placeOfReceiptByCarrier.city").send_keys(city)
        
        
        clicked = False
        while not clicked:
            try:
                Select(driver.find_element_by_id("countrySelected2")).select_by_visible_text("United States")
                clicked=True
            except:
                pass
        
        Select(driver.find_element_by_id("stateSelected2")).select_by_visible_text(state)
        
        driver.find_element_by_id("firstPortOfArrival").send_keys(crossing)
        
        crossing2 = crossing
        
        if (not "PARS" in container.properties[CCN]) and container.properties[DESCRIPTION]=="":
            crossing2 = "495"
            
        driver.find_element_by_id("portOfDestinationExit").send_keys(crossing2)
        
        if crossing2=="495":
            driver.find_element_by_id("cargoPortsForm.portOfDestinationExitSublocation").send_keys("5279")
        
        elem = driver.find_element_by_id("cargoPortsForm.foreignPortPlaceOfLoading.city")
        while elem.id != driver.switch_to.active_element.id:
            elem.click()
            
        if container.properties[POL]!="":
            elem.send_keys(container.properties[POL])
        else:
            while not GetKeyState(13)<0:
                True
        
        press("tab")
        driver.find_element_by_id("countrySelected").click()
        
        while not GetKeyState(13)<0:
            True
        
        elem = driver.find_element_by_id("stateSelected")
        
        if elem.is_enabled():
            elem.click()
            while not GetKeyState(13)<0:
                True
        
        driver.find_element_by_id("cargoAddressesTabTop").click()
        
        elem = driver.find_element_by_id("shipperName")
#         elem.equal(driver.switchTo().activeElement())
        while elem.id != driver.switch_to.active_element.id:
            try:
                elem.click()
            except:
                pass
            
        moveTo(900, 500)
            
#         quitThis = False    
#         
#         lock = Lock()
        
#         def wait_and_click():    
#             while not GetKeyState(13)<0:
#                 if quitThis:
#                     return
#                 True
#             
#             lock.acquire()
#             elem = driver.find_element_by_id("consigneeName")
# #         elem.equal(driver.switchTo().activeElement())
#             while elem.id != driver.switch_to.active_element['value'].id:
#                 elem.click()
#             driver.execute_script("window.scrollTo(0, 1080)") 
#             lock.release()
#             
#             while not GetKeyState(13)<0:
#                 if quitThis:
#                     return
#                 True
#             
# #            
#             page2 = False
#             while not page2:
#                 if quitThis:
#                     return
#                 lock.acquire()
#                 driver.find_element_by_id("cargoCargoDetailsTabBottom").click()
#                 page2 = "Total Cargo Weight:" in driver.page_source
#                 lock.release()
#             
#         with ThreadPoolExecutor() as executor:
#             executor.submit(wait_and_click)
#             
# #             wait = WebDriverWait(driver, 100000000)
# #             wait.until(lambda driver: lock.acquire() and "Total Cargo Weight:" in driver.page_source and lock.release())
#             page = False
#             while not page:
#                 lock.acquire()
#                 page = "Total Cargo Weight:" in driver.page_source
#                 lock.release()
#             quitThis = True
        def consignee_and_shipper(driver):
            while not GetKeyState(13)<0:
                if "Total Cargo Weight:" in driver.page_source:
                    return
                
#             if not GetKeyState(17)<0:
#                 sleep(3)
#                 if GetKeyState(17)<0:
#                     while not GetKeyState(13)<0:
#                         if "Total Cargo Weight:" in driver.page_source:
#                             return
            
            press("tab")
            elem = driver.find_element_by_id("consigneeName")
            while elem.id != driver.switch_to.active_element.id:
                try:
                    elem.click()
                except:
                    pass
                            
            driver.execute_script("window.scrollTo(0, 1080)")
        
            while not GetKeyState(13)<0:
                if "Total Cargo Weight:" in driver.page_source:
                    return
            
#             if not GetKeyState(17)<0:
#                 sleep(3)
#                 if GetKeyState(17)<0:
#                     while not GetKeyState(13)<0:
#                         if "Total Cargo Weight:" in driver.page_source:
#                             return
            press("tab")
            while not "Total Cargo Weight:" in driver.page_source:
                clicked = False
                while not clicked:
                    try:
                        driver.find_element_by_id("cargoCargoDetailsTabBottom").click()
                        clicked=True
                    except:
                        pass
            
        try:
            consignee_and_shipper(driver)
        except:
            raise
            wait = WebDriverWait(driver, 100000000)
            wait.until(lambda driver: "Total Cargo Weight:" in driver.page_source)
        
        weight = container.properties[WEIGHT]
        dot = weight.find(".")
        if dot>0:
            weight = weight[:dot]
        
        
        driver.find_element_by_id("cargoCargoDetailsForm.totalCargoWeight").send_keys(weight)
        
        Select(driver.find_element_by_id("cargoCargoDetailsForm.totalCargoWeightUnitOfMeasure")).select_by_visible_text("KILOGRAM")
        
        driver.find_element_by_name("_addCargoDetailsInformation").click()
        
        elem = driver.find_element_by_id("cargoQuantity")
        while elem.id != driver.switch_to.active_element.id:
            try:
                elem.click()
            except:
                pass
            
        if container.properties[PIECECOUNT] != "NONE":
            elem.send_keys(container.properties[PIECECOUNT])
        else:
            while not GetKeyState(13)<0:
                True
        
        driver.find_element_by_id("cargoQuantityUnitOfMeasure").click()
        
        while not GetKeyState(13)<0:
            True
        
        if container.properties[DESCRIPTION]!="":
            driver.find_element_by_id("cargoDescription").send_keys(container.properties[DESCRIPTION])
        driver.find_element_by_id("cargoDescription").click()
        
        while not GetKeyState(13)<0:
            True
        
        driver.find_element_by_name("_save").click()
#         while()
        driver.implicitly_wait(0)
        proceed = False
        while not proceed:
            try:
                if not "All errors must be corrected" in driver.page_source:
                    driver.find_element_by_id("buttonPortalOk")
                proceed = True
            except:
                try:
                    driver.find_element_by_name("_checkForErrorPostButton").click()
                except:
                    pass
                
        driver.implicitly_wait(6000000)
        
        clicked = False
        while not clicked:
            try:
                driver.find_element_by_id("buttonPortalOk").click()
                clicked=True
            except:
                pass
        
#         exit()
        
#         driver.find_element_by_id("_submitToCBSA").click()
        clicked = False
        while not clicked:
            try:
                driver.find_element_by_id("buttonPortalYes").click()
                clicked=True
            except:
                pass
        
        
        
    
if __name__ == '__main__':
    
#     argv = r"a J:\All motor routings\2018\Week 30\MAERSK\PARS\NORTHERN MONUMENT 1806\NORTHERN MONUMENT 1806.xlsx".split()
#     argv = r"a C:\Users\ssleep\Documents\MSC ANNICK N010.xlsx".split()
#     argv = r"a J:\Running Routing by Vessel\MSC FIAMMETTA-MSC-IN PROGRESS.xlsx".split()
    
    specificPath = ''
    for i in range(len(argv)):
        if i!=0:
            specificPath+=argv[i]
            if i != len(argv) - 1:
                specificPath+=" "
    
    
    driver = setupPortal()
    
    if "hapag" in specificPath.lower():
        containers = loadContainerInfoHapag(specificPath)
    else:
        containers = loadContainerInfo(specificPath)
    
    makeCargoDocs(containers, driver)
    
    done()
    
#     pyinstaller "C:\Users\ssleep\workspace\Hamburg Cargo Docs\Automator\__init__.py" --distpath "J:\Spencer\Cargo Doc Helper" --noconsole -y