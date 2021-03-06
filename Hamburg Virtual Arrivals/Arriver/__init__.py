# from selenium import webdriver
from selenium.webdriver import Firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
# import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By 
# 
from os import listdir
from os import path

from sys import argv, exc_info

from PyPDF2 import PdfFileReader

import re
from os import devnull

from tkinter import Button, Tk, Label, Entry, Text
# from pyautogui import click
from time import sleep
from tkinter import constants
from tkinter import Scrollbar
from tkinter import Toplevel
from tkinter import Radiobutton
from tkinter import Frame 
from tkinter import StringVar

from pyautogui import click
from pyautogui import moveTo
import string
from tkinter.test.test_tkinter.test_widgets import LabelTest

from win32api import GetAsyncKeyState
from ContainerSizeInfo import standardSize

driver = ""
testfile = ""

class Container(object):
    
    eta = ''
    vessel=''
    voyage=''
    workOrder=''
    portOfLoading=""
    portOfDischarge=""
    description=""
    containerNumber=""
    quantity=""
    packageType=""
    size=""
    unknownSize=False
    weight=""
    otherInfo=""
    shipper=""
    consignee=""
    reefer=""
    overweight=""
    overweightTier=0
    BOL=""
    temp=""
    alreadyReservation = False
    
    def __init__(self, containerNumber, size, BOL, weight=-1, temp=-9999):
        self.containerNumber = containerNumber
        self.size = size
        self.BOL = BOL
        self.weight = weight
        self.temp = temp

def popUp(top2, top, w=300, h=90, widget=""):
    top2.lift()
    top2.attributes('-topmost',True)
    top2.after_idle(top2.attributes,'-topmost',False)
    
    # get screen width and height
    ws = top2.winfo_screenwidth() # width of the screen
    hs = top2.winfo_screenheight() # height of the screen
    
    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    
    # set the dimensions of the screen 
    # and where it is placed
    top2.geometry('%dx%d+%d+%d' % (w, h, x, y))
    
    if widget !="":
        top2.wait_visibility(widget)
        click(widget.winfo_rootx()+widget.winfo_width()/2, widget.winfo_rooty()+5+widget.winfo_height()/2)
    
    top.wait_window(top2)

def setupEterm():
    fp = FirefoxProfile();
#     fp.set_preference("webdriver.load.strategy", "unstable");
    
    driver = Firefox(firefox_profile=fp, log_path=devnull)
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
    
    
    
#     driver.switch_to_default_content()
#     driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
#     elem = driver.find_element_by_css_selector('a[href*="Gate/VirtualArrive/VirtualArrive.asp"')
#     elem.click()
    return driver
    
# def recursiveArrive(specificPath):
#     for filename in listdir(specificPath):
#         if "PARS MANIFESTS" in filename and filename[-4:] == ".pdf" or filename[-4:] == ".PDF":
#             arrive(specificPath+'\\'+filename)
#         elif(path.isdir(specificPath+"\\"+filename) and not filename=="Flattened"):
#             recursiveArrive(specificPath+"\\"+filename)
def setupLine(line):
    if T1.get(str(line)+".0", str(line)+".end").lstrip()=="":
        return
    def callbackGetCorrect(top2):
        T1.tag_remove("bad", str(line)+'.0', str(line)+".end")
        top2.destroy()
        
    def callbackGetCorrectEntry(top2, value):
        value[0] = E1.get()
        T1.tag_remove("bad", str(line)+'.0', str(line)+".end")
        top2.destroy()
        
    containerIndex = T1.search('\t', str(line) + ".0")
    
    container = T1.get(str(line)+".0", containerIndex).upper()
    container = [container]

    
    m = re.search(r'[A-Z]{4}\d{7}', container[0])
    if not m:
        T1.tag_add("bad", str(line)+'.0', str(line)+".end")
        
        top2 = Toplevel(top)
        L1 = Label(top2, text='Invalid Container Number: ' + container[0] + "\nPlease correct Container Number for the underlined line then hit \"OK\"")
        L1.grid(row=0, column=0)
        E1 = Entry(top2, bd = 5)
        E1.grid(row=1, column=0)
        E1.insert(0, container[0])
        MyButton4 = Button(top2, text="OK", width=10, command=lambda: callbackGetCorrectEntry(top2, container))
        MyButton4.grid(row=2, column=0)
        popUp(top2, top, 400, widget=E1)
        T1.delete(str(line)+".0", containerIndex)
        T1.insert(str(line) + ".0", container[0])
        setupLine(line)
        return
    
    
    periodIndex = containerIndex.find(".") + 1
    sizeIndex = T1.search('\t', str(line) + "." + str(int(containerIndex[periodIndex:])+1), str(line) + ".end")
    size = T1.get(str(line) + "." + str(int(containerIndex[periodIndex:])+1), sizeIndex)
    
    size = [size]
    sizeStandard = standardSize(size[0])
#     if (size[0] != "20DC" and size[0] != "20RF" and size[0] != "40DC" and size[0] != "40HC" and size[0] != "40RH" and size[0] != "40OT" and size[0] != "20OT" and size[0] != "40OH"):
#         and size[0] != "20dc" and size[0] != "20rf" and size[0] != "40dc" and size[0] != "40hc" and size[0] != "40rh" and size[0] != "40ot" and size[0] != "20ot" and size[0] != "40oh"
#         and size[0] != "20D86" and size[0] != "20R86" and size[0] != "40D86" and size[0] != "40D96" and size[0] != "40R96" and size[0] != "40O86" and size[0] != "20O86" and size[0] != "40O96"
#         and size[0] != "20st" and size[0] != "20dv" and size[0] != "20R86" and size[0] != "40st" and size[0] != "40D96" and size[0] != "40R96" and size[0] != "40O86" and size[0] != "20O86" and size[0] != "40O96":
    if not sizeStandard:            
        T1.tag_add("bad", str(line)+'.0', str(line)+".end")
        top2 = Toplevel(top)
        L1 = Label(top2, text='Invalid Size: ' + size[0] + " for Container: " + container[0] + "\nPlease enter correct size: ")
        L1.grid(row=0, column=0)
        E1 = Entry(top2, bd = 5)
        E1.grid(row=1, column=0)
        MyButton4 = Button(top2, text="OK", width=10, command=lambda: callbackGetCorrectEntry(top2, size))
        MyButton4.grid(row=2, column=0)
        popUp(top2, top, widget= E1)
        T1.delete(str(line) + "." + str(int(containerIndex[periodIndex:])+1), sizeIndex)
        T1.insert(str(line) + "." + str(int(containerIndex[periodIndex:])+1), size[0])
        setupLine(line)
        return
    BOLIndex = T1.search('\t', str(line) + "." + str(int(sizeIndex[periodIndex:])+1))
    consigneeIndex = T1.search('\t', str(line) + "." + str(int(BOLIndex[periodIndex:])+1), str(line) + ".end")
    weight = ""
    weightIndex = ""
    if consigneeIndex !="":
        weightIndex = T1.search('\t', str(line) + "." + str(int(consigneeIndex[periodIndex:])+1), str(line) + ".end")
        if weightIndex != "":
            weight = T1.get(str(line) + "." + str(int(consigneeIndex[periodIndex:])+1), weightIndex)
        else:
            weight = T1.get(str(line) + "." + str(int(consigneeIndex[periodIndex:])+1), str(line) + ".end")
    weight = [weight]
    if not re.fullmatch(r'\d+\.*\d*', weight[0]):    
        T1.tag_add("bad", str(line)+'.0', str(line)+".end")
        top2 = Toplevel(top)
        labelText=""
        if  consigneeIndex == "":
            labelText = "Please enter a weight for container: "  + container[0]
        else:
            labelText = 'Invalid Weight: ' + weight[0] + " for Container: " + container[0] + "\nPlease enter correct weight: "
        L1 = Label(top2, text=labelText)
        L1.grid(row=0, column=0)
        E1 = Entry(top2, bd = 5)
        E1.grid(row=1, column=0)
        MyButton4 = Button(top2, text="OK", width=10, command=lambda: callbackGetCorrectEntry(top2, weight))
        MyButton4.grid(row=2, column=0)
        popUp(top2, top, h=120, widget = E1)
        if  consigneeIndex == "":
            T1.insert(str(line) + ".end" , '\t' + weight[0])
        elif weightIndex=="":
            T1.delete(str(line) + "." + str(int(consigneeIndex[periodIndex:])+1), str(line) + ".end")
            T1.insert(str(line) + "." + str(int(consigneeIndex[periodIndex:])+1), weight[0])
        else:
            T1.delete(str(line) + "." + str(int(consigneeIndex[periodIndex:])+1), weightIndex)
            T1.insert(str(line) + "." + str(int(consigneeIndex[periodIndex:])+1), weight[0])
        setupLine(line)
        return
    
    if sizeStandard=="40R96" or sizeStandard=="20R86":
        temp = ""
        tempIndex = ""
        if weightIndex !="":
            tempIndex = T1.search('\t', str(line) + "." + str(int(weightIndex[periodIndex:])+1), str(line) + ".end")
            if tempIndex != "":
                temp = T1.get(str(line) + "." + str(int(weightIndex[periodIndex:])+1), tempIndex)
            else:
                temp = T1.get(str(line) + "." + str(int(weightIndex[periodIndex:])+1), str(line) + ".end")
        temp = [temp]
        if (not re.fullmatch(r'\-?\d+\.*\d*', temp[0])) or float(temp[0]) <-30 or float(temp[0]) >30:   
            T1.tag_add("bad", str(line)+'.0', str(line)+".end")
            top2 = Toplevel(top)
            labelText=""
            if  weightIndex == "":
                labelText = "Please enter a temperature for container: "  + container[0]
            else:
                labelText = 'Invalid temperature: ' + temp[0] + " for Container: " + container[0] + "\nPlease enter correct temperature: "
            L1 = Label(top2, text=labelText)
            L1.grid(row=0, column=0)
            E1 = Entry(top2, bd = 5)
            E1.grid(row=1, column=0)
            MyButton4 = Button(top2, text="OK", width=10, command=lambda: callbackGetCorrectEntry(top2, temp))
            MyButton4.grid(row=2, column=0)
            popUp(top2, top, h=120, widget = E1)
            if  weightIndex == "":
                T1.insert(str(line) + ".end" , '\t' + temp[0])
            elif tempIndex=="":
                T1.delete(str(line) + "." + str(int(weightIndex[periodIndex:])+1), str(line) + ".end")
                T1.insert(str(line) + "." + str(int(weightIndex[periodIndex:])+1), temp[0])
            else:
                T1.delete(str(line) + "." + str(int(weightIndex[periodIndex:])+1), tempIndex)
                T1.insert(str(line) + "." + str(int(weightIndex[periodIndex:])+1), temp[0])
            setupLine(line)
            return
#        
#     if [0].temp < -500:
#         top2 = Toplevel(top)
#         L1 = Label(top2, text="Please enter Temperature for " + reservation[0].containerNumber + ":")
#         L1.grid(row=0, column=0)
#         E1 = Entry(top2, bd = 5)
#         E1.grid(row=1, column=0)
#         
#         def callbackTemp(reservation):
#             for container in reservation:
#                 container.temp = E1.get()
#             top2.destroy() 
#         
#         MyButton4 = Button(top2, text="OK", width=10, command=lambda: callbackTemp(reservation))
#         MyButton4.grid(row=2, column=0)
        
        
    
    
    
#     if [0].temp < -500:
#         top2 = Toplevel(top)
#         L1 = Label(top2, text="Please enter Temperature for " + reservation[0].containerNumber + ":")
#         L1.grid(row=0, column=0)
#         E1 = Entry(top2, bd = 5)
#         E1.grid(row=1, column=0)
#         
#         def callbackTemp(reservation):
#             for container in reservation:
#                 container.temp = E1.get()
#             top2.destroy() 
#         
#         MyButton4 = Button(top2, text="OK", width=10, command=lambda: callbackTemp(reservation))
#         MyButton4.grid(row=2, column=0)
#         
#         top2.lift()
#         top2.attributes('-topmost',True)
#         top2.after_idle(top2.attributes,'-topmost',False)
#         
#         w = 240 # width for the Tk root
#         h = 90 # height for the Tk root
#         
#         # get screen width and height
#         ws = top2.winfo_screenwidth() # width of the screen
#         hs = top2.winfo_screenheight() # height of the screen
#         
#         # calculate x and y coordinates for the Tk root window
#         x = (ws/2) - (w/2)
#         y = (hs/2) - (h/2)
#         
#         # set the dimensions of the screen 
#         # and where it is placed
#         top2.geometry('%dx%d+%d+%d' % (w, h, x, y))
#         
#         
# #                     def click2():
# #                         click((ws/2) - (w/2) + E1.winfo_rootx()+1, (hs/2) - (h/2)+E1.winfo_rooty()+1)
#         top2.wait_visibility(E1)
#         click(E1.winfo_rootx()+E1.winfo_width()/2, E1.winfo_rooty()+5+E1.winfo_height()/2)
#         
# #                     top2.mainloop()
#         
# #                     print(container.temp)
#         top.wait_window(top2)

def setupText():
    tempIndex = T1.search('\tEquipment already on terminal', "1.0")
    periodIndex = tempIndex.find(".") + 1
    while tempIndex != '':
        T1.delete(tempIndex, tempIndex[0:periodIndex-1]+"."+str(int(tempIndex[periodIndex:])+30))
        tempIndex = T1.search('\tEquipment already on terminal', "1.0")
        periodIndex = tempIndex.find(".") + 1
    tempIndex = T1.search('\tEquipment already on facility', "1.0")
    periodIndex = tempIndex.find(".") + 1
    while tempIndex != '':
        T1.delete(tempIndex, tempIndex[0:periodIndex-1]+"."+str(int(tempIndex[periodIndex:])+30))
        tempIndex = T1.search('\tEquipment already on facility', "1.0")
        periodIndex = tempIndex.find(".") + 1
    numRows = len(T1.get("1.0", constants.END).split('\n'))
    for tag in T1.tag_names():
        T1.tag_delete(tag)
    T1.tag_configure("bad", foreground = 'red', underline=1)
    T1.tag_configure("done", foreground = 'green', underline=1)
    for i in range(numRows):
        if i != 0:
            setupLine(i)
    T1.update_idletasks()
    
    

def setupContainers(containerStrings, containers, reservations):
    for containerString in containerStrings:
#             container = container.split(' ')
#             for x in container:
#                 print(x)
        weight = -1
        temp = -9999
        container = ""
        if containerString != "":
            containerString = containerString.split('\t')
            if len(containerString) >3:
                containerNumber = containerString[0]
                size = containerString[1]
#                 if size=='20DC':
#                     size="20D86"
#                 elif size=='20RF':
#                     size="20R86"
#                 elif size=='20OT':
#                     size="20O86"
#                 elif size=='40DC':
#                     size="40D86"
#                 elif size=='40HC':
#                     size="40D96"
#                 elif size=='40RH':
#                     size="40R96"
#                 elif size=='40OT':
#                     size="40O86"
#                 elif size=='40OH':
#                     size="40O96"
                size = standardSize(size)
                BOL = containerString[2]
                
                weight = containerString[4]
                if len(containerString)>5:
                    temp = containerString[5]
                
                container = Container(containerNumber, size, BOL, weight, temp)
                
                containers.append(container)
    
    
        
        if container != "":
            found =False
            for reservation in reservations:
                if container.BOL==reservation[0].BOL:
                    found = True
                    reservation.append(container)
            if found==False:
                reservations.append([container])
                
                
def reserve(reservation):
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='topnavframe.asp']"))
    driver.find_element_by_css_selector('a[href*="MenuNavFrame.asp?MenuID=4"').click()
    
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
    driver.find_element_by_css_selector('a[href*="reservations/addupdatebooking.asp"').click()

    driver.switch_to_default_content()            
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
    driver.find_element_by_name("ShipRef").send_keys(reservation[0].BOL)
    
    Select(driver.find_element_by_name("Direction")).select_by_visible_text("Inbound")
    
    Select(driver.find_element_by_name("ConvType")).select_by_visible_text("Truck")
    
    Select(driver.find_element_by_id("RT")).select_by_visible_text("Import")
    
    if steamShipLine.get()=="HAM":
        Select(driver.find_element_by_name("LineID")).select_by_visible_text("Hamburg Sud")
    elif steamShipLine.get()=="MSC":
        Select(driver.find_element_by_name("LineID")).select_by_visible_text("Mediterranean Shipping Co.")
    elif steamShipLine.get()=="CMA":
        Select(driver.find_element_by_name("LineID")).select_by_visible_text("CMA/CGM Can")
    
    Select(driver.find_element_by_name("LEStatus")).select_by_visible_text("Load")
    
    driver.find_element_by_name("Submit").click()
    
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: "Reservation Number already exists." in driver.page_source or EC.element_to_be_clickable(driver.find_element_by_name("polcd")))
    
    stopThis = [False]
    
    if "Reservation Number already exists." in driver.page_source:
        stopThis = [True]
        top2 = Toplevel(top)
        L1 = Label(top2, text="Reservation " + reservation[0].BOL + " already exists.")
        L1.grid(row=0, column=0, columnspan=2)
        
        def callbackContinue(stopThis):
            stopThis[0] = False
            top2.destroy() 
        
        def callbackStop():
            top2.destroy()
             
        
        MyButton4 = Button(top2, text="Proceed anyway", width=14, command=lambda: callbackContinue(stopThis))
        MyButton4.grid(row=2, column=0)
        
        MyButton5 = Button(top2, text="Stop", width=10, command=lambda: callbackStop())
        MyButton5.grid(row=2, column=1)
        
        top2.lift()
        top2.attributes('-topmost',True)
        top2.after_idle(top2.attributes,'-topmost',False)
        
        w = 240 # width for the Tk root
        h = 60 # height for the Tk root
        
        # get screen width and height
        ws = top2.winfo_screenwidth() # width of the screen
        hs = top2.winfo_screenheight() # height of the screen
        
        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        
        # set the dimensions of the screen 
        # and where it is placed
        top2.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        
#                     def click2():
#                         click((ws/2) - (w/2) + E1.winfo_rootx()+1, (hs/2) - (h/2)+E1.winfo_rooty()+1)
        top2.wait_visibility(MyButton5)
        moveTo(MyButton5.winfo_rootx()+MyButton5.winfo_width()/2, MyButton5.winfo_rooty()+MyButton5.winfo_height()/2)
        
#                     top2.mainloop()
        
#                     print(container.temp)
        top.wait_window(top2)
        if not stopThis[0]:
            return True
        
        
    if stopThis[0]:
        return False
    
    driver.find_element_by_name("polcd").send_keys("CATOR")
    
    if reservation[0].size=="20R86" or reservation[0].size=="40R96":
        driver.find_element_by_name("ReeferFlg").click()
        
        driver.find_element_by_name("SetTempNbr").send_keys(str(reservation[0].temp))
        
        Select(driver.find_element_by_name("SetTempUnits")).select_by_value("C")
    
    twentyStan = 0
    twentyReef = 0
    twentyOpen = 0
    twentyTank = 0
    fourtyStan = 0
    fourtyOpen = 0
    fourtyOpenHigh = 0
    fourtyReef = 0
    fourtyHigh = 0
    
    for container in reservation:
        if container.size=='20D86':
            twentyStan+=1
        elif container.size=='20O86':
            twentyOpen+=1
        elif container.size=='20R86':
            twentyReef+=1
        elif container.size=='20T86':
            twentyTank+=1
        elif container.size=='40D86':
            fourtyStan+=1
        elif container.size=='40O86':
            fourtyOpen+=1
        elif container.size=='40O96':
            fourtyOpenHigh+=1
        elif container.size=='40D96':
            fourtyHigh+=1
        elif container.size=='40R96':
            fourtyReef+=1
    
    diffTypes = 1
    
    if twentyStan>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(twentyStan)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("20D86")    
        diffTypes+=1
    if twentyOpen>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(twentyOpen)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("20O86")    
        diffTypes+=1
    if twentyReef>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(twentyReef)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("20R86")
        diffTypes+=1
    if twentyTank>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(twentyTank)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("20T86")
        diffTypes+=1
    if fourtyStan>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(fourtyStan)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("40D86")
        diffTypes+=1
    if fourtyOpen>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(fourtyOpen)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("40O86")
        diffTypes+=1
    if fourtyOpenHigh>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(fourtyOpenHigh)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("40O96")
        diffTypes+=1
    if fourtyHigh>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(fourtyHigh)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("40D96")
        diffTypes+=1
    if fourtyReef>0:
        driver.find_element_by_name("EqQty" + str(diffTypes)).send_keys(fourtyReef)
        driver.find_element_by_name("EqType"+ str(diffTypes)).send_keys("40R96")
        diffTypes+=1
    if GetAsyncKeyState(27) < 0:
        return False
    driver.find_element_by_name("Submit").click()
    return True
#         lastIndex = "a"
#         for container in reservation:
#             T1.insert(lastIndex[0]+".end", "\t Reservation already existed")
#             container.alreadyReservation = True
#             lastIndex = T1.search(reservation[0], constants.END)

def arrive(container):
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='topnavframe.asp']"))
    
    driver.find_element_by_css_selector('a[href*="MenuNavFrame.asp?MenuID=1"').click()
    
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
    
    driver.find_element_by_css_selector('a[href*="Gate/VirtualArrive/VirtualArrive.asp"').click()
    
    driver.switch_to_default_content()
    driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='portals/portal.asp']"))
    
    driver.find_element_by_name("container_prefix_dof").send_keys(container.containerNumber[:4])
    driver.find_element_by_name("container_number_dof").send_keys(container.containerNumber[4:11])
    
    select = Select(driver.find_element_by_name("ddlLoadStatus_dof"))
    select.select_by_visible_text("Load")
    
    select = Select(driver.find_element_by_name("lineid"))
    
    if steamShipLine.get()=="HAM":
        select.select_by_visible_text("Hamburg Sud")
    elif steamShipLine.get()=="MSC":
        select.select_by_visible_text("Mediterranean Shipping Co.")
    elif steamShipLine.get()=="CMA":
        select.select_by_visible_text("CMA/CGM Can")
    
    elem = driver.find_element_by_name("ddlSzTyCnt")
    elem.send_keys(container.size)
    
    elem = driver.find_element_by_name("cargo_weight")
    elem.send_keys(str(container.weight))
    
    select = Select(driver.find_element_by_name("ddWeightUnits"))
    select.select_by_visible_text("Kgs")
    
    driver.find_element_by_id("CkbCR").click()
    driver.find_element_by_id("CkbFR").click()
    
    
    #         if not ("LCBO" in consignee or "Liquor Control" in consignee):
    #             elem = driver.find_element_by_id("CkbFR")
    #             elem.click()
        
    elem = driver.find_element_by_name("bkg_nbr_dof")
    elem.send_keys(container.BOL)
    
    
    select = Select(driver.find_element_by_name("Line"))
    if steamShipLine.get()=="HAM":
        select.select_by_visible_text("Hamburg Sud")
    elif steamShipLine.get()=="MSC":
        select.select_by_visible_text("Mediterranean Shipping Co.")
    elif steamShipLine.get()=="CMA":
        select.select_by_visible_text("CMA/CGM Can")
      
    if GetAsyncKeyState(27) < 0:
        return False
    
    elem = driver.find_element_by_name("Submit")
    elem.click()
    wait = WebDriverWait(driver, 10)
    wait.until(lambda driver: "Equipment is already on Terminal" in driver.page_source or
                "Equipment is already on facility." in driver.page_source or
                EC.element_to_be_clickable(driver.find_element_by_name("Close")))
    errorMess = ""
    if "Equipment is already on Terminal" in driver.page_source :
        errorMess = "Equipment already on terminal"
    if "Equipment is already on facility." in driver.page_source:
        errorMess = "Equipment already on facility"
    if errorMess != "":
        searchIndex = T1.search(container.containerNumber, "1.0")
        periodIndex = searchIndex.find(".")+1
        T1.insert(searchIndex[0:periodIndex-1] + ".end", "\t" + errorMess)
        T1.tag_add("bad", searchIndex[0:periodIndex-1] + ".0", searchIndex[0:periodIndex-1] + ".end")
        T1.update_idletasks()
        driver.switch_to_default_content()
        driver.switch_to_frame(driver.find_element_by_css_selector("frame[src='MenuNavFrame.asp?MenuID=10']"))
        elem = driver.find_element_by_css_selector('a[href*="Gate/VirtualArrive/VirtualArrive.asp"')
        elem.click()
    else:
        elem = driver.find_element_by_name("Close")
        elem.click()
        searchIndex = T1.search(container.containerNumber, "1.0")
        periodIndex = searchIndex.find(".")+1
        T1.tag_add("done", searchIndex[0:periodIndex-1] + ".0", searchIndex[0:periodIndex-1] + ".end")
        T1.update_idletasks()
    
    return True
        
if __name__ == '__main__':
    
#     argv = r"a C:\Users\ssleep\Documents\Programming\Hapag Dispatchmate\601346204 PARS MANIFESTS.pdf".split()
    argv = r"a J:\All motor routings\2018\Week 52\HAPAG".split()
    try:
        driver = setupEterm()
    except:
        print(exc_info())
        sleep(50)

    
#     specificPath = ''
#     for i in range(len(argv)):
#         if i!=0:
#             specificPath+=argv[i]
#             if i != len(argv) - 1:
#                 specificPath+=" "
 
                     
#     testfile = specificPath[:specificPath.rfind('\\')] + "\\" + "Containers already in Eterm.txt"
#     if path.isdir(specificPath):
#         recursiveArrive(specificPath)
#     elif "PARS MANIFESTS" in specificPath and (specificPath[-4:] == ".pdf" or specificPath[-4:] == ".PDF"):
#         arrive(specificPath)
    
    top = Tk()
    L1 = Label(top, text="Please enter Info, with this format:")
    L1.grid(row=0, column=0)
    L2 = Label(top, text="Seperate data fields with 'tabs' NOT SPACES")
    L2.grid(row=1, column=0)
    L3 = Label(top, text="Container#        Size        BOL#            Consignee        Weight        Temp (if reefer):")    
    L3.grid(row=2, column=0)
    
    S1=Scrollbar(top, orient='vertical')
    S1.grid(row=3, column=1, sticky=constants.N + constants.S)
    S2=Scrollbar(top, orient='horizontal')
    S2.grid(row=4, column=0, sticky=constants.E + constants.W)
#     S2.pack(side=constants.BOTTOM, fill=constants.X)
    
    T1 = Text(top, height = 33, width = 97, xscrollcommand = S2.set, yscrollcommand=S1.set, wrap = constants.NONE)
    T1.grid(row=3, column=0)
    
    steamShipLine=StringVar()
    
    f = Frame(top)
    f.grid(row=5, column=0)
    R1 = Radiobutton(f, text="Hamburg Sud", variable=steamShipLine, value="HAM").pack(side="left")
    R2 = Radiobutton(f, text="MSC", variable=steamShipLine, value="MSC").pack(side="left")
    R3 = Radiobutton(f, text="CMA", variable=steamShipLine, value="CMA").pack(side="left")
#     R1.grid(row=5, column=0)
#     R2.grid(row=5, column=1)
#     R3.grid(row=5, column=2)
    
    S1.config(command=T1.yview)
    S2.config(command=T1.xview)
    
    
    
    def callbackArrive():
#         print(T1.get("0.0", constants.END))
#         setupText()
        if steamShipLine.get()=="":
            top2 = Tk()
            L2 = Label(top2, text="Please select steamship line")
            L2.config(font=("Courier", 16))
            L2.grid(row=0, column=0)
             
            def callbackStop():
                top2.destroy()
                
            MyButton4 = Button(top2, text="OK", width=10, command=callbackStop)
            MyButton4.grid(row=1, column=0)
            
            # get screen width and height
            ws = top2.winfo_screenwidth() # width of the screen
            hs = top2.winfo_screenheight() # height of the screen
            
            w = 400
            h = 100
            
            # calculate x and y coordinates for the Tk root window
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            
            # set the dimensions of the screen 
            # and where it is placed
            top2.geometry('%dx%d+%d+%d' % (w, h, x, y))
            
            popUp(top2, top, w, h)

        setupText()
         
        containerStrings = T1.get("0.0", constants.END).split("\n")
        containers = []
        reservations = []
         
        setupContainers(containerStrings, containers, reservations)
        
        for reservation in reservations:
            if reserve(reservation):
                for container in reservation:
                    if not arrive(container):
                        return
            else:
                return
            
    
#     def callbackTest():
#         lastIndex = T1.search('a67SSZIA5998X', constants.END)
# #         T1.config(fg="red")
# #         T1.insert(lastIndex[0]+".end", "\t Reservation already exists")
# #         T1.config(fg="black")
#         print(lastIndex)
#         lastIndex = T1.search('a67SSZIA5998X', str(lastIndex[0]) + "." + str(int(lastIndex[periodIndex:])+1))
# #         T1.insert(lastIndex[0]+".end", "\t Reservation already exists")
#         print(lastIndex)
#         lastIndex = T1.search('a67SSZIA5998X', str(lastIndex[0]) + "." + str(int(lastIndex[periodIndex:])+1))
#         print(lastIndex)
#         T1.insert(lastIndex[0]+".end", "\t Reservation already exists")
#         lastIndex = T1.search('a67SSZIA5998X', lastIndex[0]+".end")
#         T1.insert(lastIndex[0]+".end", "\t Reservation already exists")
#         lastIndex = T1.search('a67SSZIA5998X', lastIndex[0]+".end")
#         T1.insert(lastIndex[0]+".end", "\t Reservation already exists")
    
    
#     MyButton5 = Button(top, text="Arrive", width=30, command=lambda:setupLine(1))
    MyButton5 = Button(top, text="Arrive", width=30, command=callbackArrive)
    MyButton5.grid(row=6, column=0)
    
    top.lift()
    top.attributes('-topmost',True)
    top.after_idle(top.attributes,'-topmost',False)
    
    w = 800 # width for the Tk root
    h = 800 # height for the Tk root
    
    # get screen width and height
    ws = top.winfo_screenwidth() # width of the screen
    hs = top.winfo_screenheight() # height of the screen
    
    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    
    # set the dimensions of the screen 
    # and where it is placed
    top.geometry('%dx%d+%d+%d' % (w, h, x, y))

    top.mainloop()
    
#     pyinstaller "C:\Users\ssleep\workspace\Hamburg Virtual Arrivals\Arriver\__init__.py" --distpath "J:\Spencer\Hamburg Sud Virtual Arriver" -y --noconsole