from selenium.webdriver import Ie, ie
from selenium.webdriver import Firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver import FirefoxProfile
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pywinauto.win32structures import RECT
from pyautogui import click
from pyautogui import moveTo

from re import search
from threading import Thread

from pywinauto import Application
from pywinauto import handleprops

from tkinter import Button, Tk, Label, Entry, Text
from time import sleep
from tkinter import constants
from tkinter import Scrollbar
from tkinter import Toplevel
from selenium import webdriver

from win32api import GetKeyState
from distutils.command.clean import clean
from _datetime import datetime, timedelta

from pynput.mouse import Controller

import win32gui
from win32con import SWP_SHOWWINDOW

from sys import exc_info
import HelperFunctions
from os.path import devnull
import atexit

boo = [False]

class Container(object):
    
    description = ""
    weight = ""
    pieces = ""
    terminal = ""
    containerNumber = ""
    bond = ""
    PB = ""
    size = ""
    customer = ""
    extraText = ""
    
def popUp(top, top2 = "", w=300, h=90, widget=""):
    if not top2=="":
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
    else:
        top.lift()
        top.attributes('-topmost',True)
        top.after_idle(top.attributes,'-topmost',False)
        
        # get screen width and height
        ws = top.winfo_screenwidth() # width of the screen
        hs = top.winfo_screenheight() # height of the screen
        
        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        
        # set the dimensions of the screen 
        # and where it is placed
        top.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        if widget !="":
            top.wait_visibility(widget)
            moveTo(widget.winfo_rootx()+widget.winfo_width()/2, widget.winfo_rooty()+5+widget.winfo_height()/2)
        top.mainloop()
    
def setupPortal():

#     driver = Firefox(log_path=devnull)

    
# ...
# capabilities.setCapability("requireWindowFocus", true);
# WebDriver driver = new InternetExplorerDriver(capabilities);
# r"C:\Automation\IEDriverServer.exe"

#     options = ()
#     options.IntroduceInstabilityByIgnoringProtectedModeSettings = True
#     options.IgnoreZoomLevel = True
# #     options.UnhandledPromptBehavior = un.Accept
#     options.EnablePersistentHover = True
#     options.EnableNativeEvents = False
#     options.EnsureCleanSession = True    // this cleansession did the trick

    def exit_hander():
#         sleep(20)
#         driver.quit()
        print("Error:")     
    atexit.register(exit_hander)
    capabilities = DesiredCapabilities.INTERNETEXPLORER
#     capabilities['INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS']=True
    capabilities['requireWindowFocus']=True
    driver = Ie(capabilities=capabilities)
#     driver = Ie()
    driver.get("https://ace.cbp.dhs.gov/")
    driver.set_window_position(1920, 0)
    driver.maximize_window()
    driver.implicitly_wait(100)
    f=open(r"C:\Automation\ACE Login.txt", 'r')
    read = f.readline()
    m = search("username: *", read)
    read = read[m.end():].strip()

    elem = driver.find_element_by_name('username')
    elem.clear()
    elem.send_keys(read)
    read = f.readline()
    m = search("password: *", read)
    read = read[m.end():].strip()
    elem =driver.find_element_by_name('password')
    elem.clear()
    elem.send_keys(read)
    f.close()
#     sleep(100)
    elem = driver.find_element_by_name('Login')
#     wait = WebDriverWait(driver, 100)
#     wait.until(EC.element_to_be_clickable(elem)) 
#     elem.send_keys(Keys.ENTER)
    elem.send_keys(Keys.ENTER)
    
#     driver.find_element_by_class_name("wpsToolBar")
    driver.find_element_by_css_selector("button[accesskey='T']").send_keys(Keys.ENTER)
    driver.find_element_by_id("clayView:ns_7_M3ULU7BUVD0M2HFG8D10000000_00000:_idsc00001:_idsc00003_2:_idsc00007").send_keys(Keys.ENTER)
#     sleep(5)
    driver.implicitly_wait(5)
#     try:
#         driver.find_element_by_css_selector('img[href="/ace1/wps/myportal/!ut/p/c5/hZHBcoIwFEW_xS94D2iELhHEgMQqCAIbhhF04oi6yFjh6ytTbZkRMFmenNx3E0jgvk_Zle8zwc-n7AgxJGpqUGZQRwmZo7gEZWv-ufIcVcIZQlhmN17yusghgmTcexTxzuNxij1Lx8ZmSuAG6iQITWQytWY2kR781x7gHbbhP2-HDUT4kfoH7cIqEbm1cV0fal8Wpi4Jc4WLOq5Evl7moRdMdNPaFt-3v0RpimgTDW3Z0diXha3Ed20HJhq2N81HvPq01bfhQ6_Z5ZNW_oP3tuvy1X9_Qc9lAZdSq_iReDvf5sv9aPQD-GwhZg!!/dl3/d3/L0lKQSEvd0pDQUFBISEvWVB3IS9ub3JtYWw!/"]').send_keys(Keys.ENTER)
#     except:
#         True
    minimized = False
    try:
        driver.execute_script("arguments[0].click();", driver.find_element_by_css_selector('img[alt="Minimize Manifest"]'))
#         driver.find_element_by_css_selector('a[href="/ace1/wps/myportal/!ut/p/c5/hZHRboIwFIafxSc4h7IKu8QyLIw6hIHADSGRmRpRLxo3-_STzG0kA9Zefv36n7-FEm77WF_krlbydKwPUEBpVYwLxgMzE4EZUiTe8-M6DiwDlwhZW3_IVupmCzmU89GjiDdezCscWQ52tjDTMLUWaeaiINxb-tS48y97gg_YLPm-HTaQ40OV7O2zuKo81OzyutcJUa5jqHD9LnShxSqKtlmcLhyXKGzYT6LxhOhTG30S2OLFw17if20nJpq2N91H_PV5r2_Hp15zyKe9_DsfbTfkW7_-ip_aBs6tfZUHGr8lvox2s9knEEREDw!!/dl3/d3/L0lKQSEvd0pDQUFBISEvWVB3IS9taW5pbWl6ZWQ!/"]').send_keys(Keys.ENTER)
        minimized = True
    except:
        True
    driver.implicitly_wait(100)
    driver.find_element_by_id("clayView:ns_7_CHMCHJ3VMJ3L502FK9QRJ71003_00000:accountListForm:_idsc00058_0:_idsc00065").send_keys(Keys.ENTER)
    driver.find_element_by_id("clayView:ns_7_CHMCHJ3VMJ3L502FK9QRJ71003_00000:accountListForm:_idsc00058_1:_idsc00068").send_keys(Keys.ENTER)
    driver.implicitly_wait(2)
#     driver.implicitly_wait(1)
    if not minimized:
        try:
#         driver.execute_script("arguments[0].click();", driver.find_element_by_css_selector('img[alt="Minimize Manifest"]'))
            driver.execute_script("arguments[0].click();", driver.find_element_by_css_selector('img[alt="Minimize Manifest"]'))
        except:
            True
    
        
    driver.implicitly_wait(100)
    return driver
    
    
def callbackBookTE(driver):
#     if "Create Standard Shipment for another Carrier" in driver.page_source:
#         driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_Create_Standard_Shipment").send_keys(Keys.ENTER)
#     elif not "Either an ID or Full Shipper information is required for a Shipper." in driver.page_source:
#         while not "Either an ID or Full Shipper information is required for a Shipper." in driver.page_source:
#             top = Tk()
#             L2 = Label(top, text="Please navigate to shipment page the hit \"OK\"")
#             L2.grid(row=0, column=0)
#             
#             def callbackStop():
#                 top.destroy()
#                
#             MyButton4 = Button(top, text="OK", width=10, command=callbackStop)
#             MyButton4.grid(row=1, column=0)
#         
#             popUp(top)
    driver.implicitly_wait(100)
    container = Container()
    setupDM(container)
    try:
        portCode = [""]
        if container.terminal=="311":
            if "PACKER" in container.extraText:
                portCode[0] = "1101"
            elif "PNCT" in container.extraText or "APM" in container.extraText or "MAHER" in container.extraText:
                portCode[0] = "4601"
            else:
                top = Tk()
                
                L1 = Label(top, text="Is the terminal Packer?")
                L1.grid(row=0, column=0, columnspan = 2)
                L1.config(font=("Courier", 24))
                
                def callbackPackerYes(portCode):
                    portCode[0] = "1101"
                    top.destroy()
                def callbackPackerNo(portCode):
                    portCode[0] = "4601"
                    top.destroy()
                
                MyButton4 = Button(top, text="Yes", width=20, command=lambda: callbackPackerYes(portCode))
                MyButton4.grid(row=2, column=0)
                MyButton4.config(font=("Courier", 24))
                
                MyButton5 = Button(top, text="No", width=20, command=lambda: callbackPackerNo(portCode))
                MyButton5.grid(row=2, column=1)
                MyButton5.config(font=("Courier", 24))
                top.lift()
                top.attributes('-topmost',True)
                top.after_idle(top.attributes,'-topmost',False)
                
                # get screen width and height
                ws = top.winfo_screenwidth() # width of the screen
                hs = top.winfo_screenheight() # height of the screen
                
                w = 800
                h = 150
                
                # calculate x and y coordinates for the Tk root window
                x = (ws/2) - (w/2)
                y = (hs/2) - (h/2)
                
                # set the dimensions of the screen 
                # and where it is placed
                top.geometry('%dx%d+%d+%d' % (w, h, x, y))
                moveTo(946, 614)
                top.mainloop()
        elif container.terminal=="309":
            portCode[0] = "1101"
        else:
            portCode[0] = "4601"
        
        elem = driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_SHIPMENTTYPE")
        while "Either an ID or Full Shipper information is required for a Shipper." in driver.page_source:
            try:
                elem.send_keys(Keys.ENTER)
            except:
                pass
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_SHIPMENTTYPE")).select_by_visible_text("Prefiled Inbond")
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_null").send_keys(Keys.ENTER)
        elem = driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_userEnteredSCN")
        elem.clear()
        elem.send_keys("801" + str(container.PB))
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_POINTOFLOADINGQLFR")).select_by_visible_text("Schedule K")
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_POINTOFLOADING").send_keys("80107")
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_FDACONFIRMATIONIND")).select_by_visible_text("No")
         
        try:
            driver.find_element_by_css_selector("input[value='Find Shipper']").send_keys(Keys.ENTER)
            sleep(3)
            wait = WebDriverWait(driver, 100000000)
            wait.until(lambda driver: "Either an ID or Full Shipper information is required for a Shipper." in driver.page_source)
        except:
            driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_null").send_keys(Keys.ENTER)
            wait = WebDriverWait(driver,10)
            wait.until(EC.alert_is_present());
            driver.switch_to_alert().accept()
        address1 = ""
        address2 = ""
        city = ""
        country = ""
        stateProv = ""
        zipPost = ""
    #     print(container.terminal) 
        if container.terminal=="311":
            address1="C/O CSX BUFFALO"
            address2 = "257 LAKE AVE"
            city = "BUFFALO"
            country = "USA"
            stateProv = "New York"
            zipPost = "14206"
        elif container.terminal=="305":
            address1 = "C/O ASI TERMINAL"
            address2 = "MARSH ST"
            city = "NEWARK"
            country = "USA"
            stateProv = "New Jersey"
            zipPost = "07100"
        elif container.terminal=="306":
            address1 = "C/O APM TERMINALS"
            address2 = "5080 MCLESTER STEET"
            city = "NEWARK"
            country = "USA"
            stateProv = "New Jersey"
            zipPost = "07100"
        elif container.terminal=="664":
            address1 = "C/O NEW YORK CONTAINER TERMINAL"
            address2 = "WESTERN AVE"
            city = "STATEN ISLAND"
            country = "USA"
            stateProv = "New York"
            zipPost = "10303"
        elif container.terminal=="309":
            address1 = "C/O PACKER TERMINAL"
            address2 = "3301 S COLUMBUS BLVD"
            city = "PHILADELPHIA"
            country = "USA"
            stateProv = "Pennsylvania"
            zipPost = "19148"
        elif container.terminal=="330":
            address1 = "C/O MAHER TERMINAL"
            address2 = "1260 CORBIN STREET"
            city = "NEWARK"
            country = "USA"
            stateProv = "New Jersey"
            zipPost = "07201"
          
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_CONSIGNEE_ADDRESS_STREET").send_keys(address1)
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_CONSIGNEE_ADDRESS_STREET2").send_keys(address2)
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_CONSIGNEE_ADDRESS_CITY").send_keys(city)
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_CONSIGNEE_ADDRESS_COUNTRY")).select_by_visible_text(country)
    #         try:
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_CONSIGNEE_ADDRESS_REGION")).select_by_visible_text(stateProv)
    #         except:
    #             wait = WebDriverWait(driver, 100000000)
    #             wait.until(lambda driver: "Create Standard Shipment for another Carrier" in driver.page_source)
    #             return
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_CONSIGNEE_ADDRESS_ZIP").send_keys(zipPost)
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_STANDARDSHIPMENT_CONSIGNEE_NAME").send_keys(Keys.ENTER)
         
        while not GetKeyState(13)<0:
            if GetKeyState(27)<0:
                raise AssertionError
                 
    #             True
              
         
         
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_standardShipmentEquipmentType")).select_by_visible_text("Create One Time")
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_standardShipmentEquipmentType").send_keys(Keys.ENTER)
          
          
        sizeSelect = ""
        if container.size=="40":
            sizeSelect = "40ft ClosedTopSeaCnt"
        elif container.size=="20":
            sizeSelect = "20ft ClosedTopSeaCnt"
          
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_SHIPMENTEQUIPMENT_TYPE")).select_by_visible_text(sizeSelect)
    #     print(container.containerNumber[:4] + container.containerNumber[5:11] + container.containerNumber[12:13])
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_SHIPMENTEQUIPMENT_TRANSPORTID").send_keys(container.containerNumber)
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_null").send_keys(Keys.ENTER)
         
         
    #     sleep(10)
        driver.find_element_by_xpath("//form[@name='PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_createStandardShipment']/table/tbody/tr[2]/td/fieldset[6]/table/tbody/tr/td/a").send_keys(Keys.ENTER)
    # #     
        container.pieces = container.pieces.replace(',', "")
        elem = driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDCOMMODITY_SHIPMENTQUANTITY")
        elem.clear()
        elem.send_keys(container.pieces)
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDCOMMODITY_QUANTITYUOM").send_keys(Keys.ENTER)
    #       
        while not GetKeyState(13)<0:
            if GetKeyState(27)<0:
                raise AssertionError
    #       
     
        elem = driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDCOMMODITY_WEIGHT")
        elem.clear()
        container.weight = container.weight.replace(',', "")
        index = container.weight.rfind(".")
        if index>0:
            container.weight=container.weight[:index]
        elem.send_keys(container.weight)
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDCOMMODITY_WEIGHTUOM")).select_by_visible_text('Kilograms')
           
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDCOMMODITY_DESCRIPTION").send_keys(container.description)
        elem = driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDCOMMODITY_VALUE")
        elem.clear()
        elem.send_keys(Keys.ENTER)
            
        while not GetKeyState(13)<0:
            if GetKeyState(27)<0:
                raise AssertionError
            
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_AvailableINBONDCOMMODITY_HTSNUMS").send_keys(Keys.ENTER)
        sleep(1)
            
        while not GetKeyState(13)<0:
            if GetKeyState(27)<0:
                raise AssertionError
                
                
        HS = driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_AvailableINBONDCOMMODITY_HTSNUMS").get_attribute("value")
        zeroes=""
        if len(HS)<10:
            for _ in range(10 - len(HS)):
                zeroes+="0"
         
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_AvailableINBONDCOMMODITY_HTSNUMS").send_keys(zeroes)
                 
         
        driver.execute_script("arguments[0].click();", driver.find_element_by_css_selector('img[src*="/ace1/wps/PA_Shipment/images/right_single.gif"]'))
    #         driver.find_element_by_css_selector('img[src*="/ace1/wps/PA_Shipment/images/right_single.gif"]').send_keys(Keys.ENTER)
         
    #         driver.find_element_by_css_selector('img[src*="/ace1/wps/PA_Shipment/images/right_single.gif"]').click()
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_null").send_keys(Keys.ENTER)
    #     
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_standardShipmentEquipmentType")).select_by_visible_text("Conveyance")
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_standardShipmentEquipmentType").send_keys(Keys.ENTER)
    #         didntwork = True
    #         while didntwork:
    #             try:
        wait = WebDriverWait(driver, 100000000)
        wait.until(lambda driver: "Conveyance</t" in driver.page_source)
        
        Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDSHIPMENT_ENTRYTYPE")).select_by_visible_text("Transportation and Exportation")
    #                 while not "Transportation and Exportation" in elem.first_selected_option.text:
    #                     failed = True
    #                     while failed:
    #                         try:
    #     Select(elem).select_by_visible_text("Transportation and Exportation")
    #                             failed = False
    #                         except:
    #                             pass
    #                     elem = Select(driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDSHIPMENT_ENTRYTYPE"))
    #                 
    #                 didntwork=False
    #             except:
    #                 pass
        
        
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDSHIPMENT_INBONDDESTINATION").send_keys(portCode[0])
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDSHIPMENT_BONDEDCARRIER").send_keys("98-066177700")
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDSHIPMENT_INBONDNUMBER").send_keys(container.bond)
        
        date = datetime.now()
        date = (date + timedelta(days=14)).strftime('%m/%d/%Y')
        
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDSHIPMENT_ESTDATEOFUSDEPARTURE").send_keys(date)
        driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_INBONDSHIPMENT_FOREIGNPORTOFDESTINATION").send_keys(Keys.ENTER)
        wait.until(lambda driver: "Create Standard Shipment for another Carrier" in driver.page_source)
    except AssertionError:
        pass
    except:
        top = Tk()
        L1 = Label(top, text="Something went wrong. Either complete the rest of this T&E manually,\n or cancel and restart.")
        L1.config(font=("Courier", 30))
        L1.grid(row=0, column=0)
        L2 = Label(top, text=exc_info())
#         L2.config(font=("Courier", 30))
        L2.grid(row=1, column=0)
         
        def callbackDM():
            top.destroy()
         
        MyButton4 = Button(top, text="OK", width=20, command=callbackDM)
        MyButton4.grid(row=2, column=0)
        MyButton4.config(font=("Courier", 30))
        popUp(top, w=1700, h=200, widget = MyButton4)
        
#     while not GetKeyState(13)<0 and not "Create Standard Shipment for another Carrier" in driver.page_source:
#         if GetKeyState(13)<0:
#             driver.find_element_by_id("PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_PC_7_CHMCHJ3VMJ3L502FK9QRJ710G2000000_null").send_keys(Keys.ENTER)
#     while not "Create Standard Shipment for another Carrier" in driver.page_source:
#         try:
    wait = WebDriverWait(driver, 100000000)
    wait.until(lambda driver: "Create Standard Shipment for another Carrier" in driver.page_source)
#         except:
#             alertObj = driver.switch_to.alert
#             alertObj.accept()
#             True
#         try:
#             alertObj = driver.switch_to.alert
#             alertObj.accept()
#         except: 
#             True

        
#     FOCUs doES NOT WORK. LOOK FOR KEYLISTENER
    
#     print(elem.get_focus())
#     wait = WebDriverWait(driver, 100000000)
#     wait.until(lambda driver: not elem.get_focus())
    
    
#     while True:
#         sleep(60)
    
# def windowEnumerationHandler(hwnd, top_windows):
#     top_windows.append((hwnd, GetWindowText(hwnd)))    

def setupDM(container):
    rects = readRects()
    app = Application(backend="win32").connect(path = r"C:\DM54_W16\DM54_W16.exe")
    
#     top_windows = []
#     EnumWindows(windowEnumerationHandler, top_windows)
#     for i in top_windows:
#         if 'Dispatch-Mate' in i[1]:
#             SetWindowPos(i[0], None, 0, 0, 1920, 1080, SWP_SHOWWINDOW)
#             SetForegroundWindow(i[0])
            
    
    winChildren = ""
    
    dialogs = app.windows()
    
    
    
    click(50, 350)
    fore = win32gui.GetForegroundWindow()
    DMFore = "Dispatch-Mate" in win32gui.GetWindowText(fore)
    while not DMFore:
        top = Tk()
        L1 = Label(top, text="Please maximize DispatchMate and the PB in the left monitor")
        L1.grid(row=0, column=0)
         
        def callbackDM():
            top.destroy()
         
        MyButton4 = Button(top, text="OK", width=10, command=callbackDM)
        MyButton4.grid(row=1, column=0)
     
        popUp(top, w=350, h=50, widget = MyButton4)
         
        click(50, 350)
        fore = win32gui.GetForegroundWindow()
        DMFore = "Dispatch-Mate" in win32gui.GetWindowText(fore)
#         
#     moveTo(2000, 550)
#     def windowEnumerationHandler(hwnd, top_windows):
#         top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))
#  
# #     results = []
#     top_windows = []
#     win32gui.EnumWindows(windowEnumerationHandler, top_windows)
#     for i in top_windows:
#         if "Dispatch-Mate" in i[1].lower():
# #             win32gui.ShowWindow(i[0],5)
#             win32gui.SetForegroundWindow(i[0])
#             break


    for x in dialogs:
        if handleprops.classname(x) == "WinDevObject":
            winChildren = handleprops.children(x)
            topWindow = x
            break
    
#     description = ""
#     weight = ""
#     pieces = ""
#     terminal = ""
#     containerNumber = ""
#     bond = ""
#     PB = ""
#     sleep(5)
#     topWindowWrap = app.window(handle=topWindow)
#     topWindowWrap.MoveWindow(0,0,1920,1080)
#     topWindowWrap.Maximize()
    for x in winChildren:
#         if handleprops.classname(x)=="ListBox":
# #             for y in handleprops.children(x):
# #                 print(handleprops.text(x))
#             topWindowWrap = app.window(handle=topWindow)
#             boxWrap = topWindowWrap.child_window(handle=x).wrapper_object()
#         if handleprops.text(x)=="Description":
#             print("desc   " + str(handleprops.rectangle(x)))
#         if "CUT - 7/13" in handleprops.text(x):
#             print(handleprops.rectangle(x))
#         if handleprops.classname(x)=="ComboBox":
        if handleprops.rectangle(x)==RECT(232, 917, 443, 939):
            topWindowWrap = app.window(handle=topWindow)
            boxWrap = topWindowWrap.child_window(handle=x).wrapper_object()
#             print(boxWrap.texts())
#             print(boxWrap.selected_index())
#             print(len(boxWrap.children()))
#             print(boxWrap.get_properties()['texts'])
#             for y in boxWrap.texts():
#                 if y !="":
#                     print(boxWrap.ItemData(y))
            num = boxWrap.ItemData(boxWrap.texts()[1])
            if num <6:
                container.size = "20"
            else:
                container.size = "40"
#             else:
#                 container.size = "getInput"
#             if boxWrap.texts()[1]==r"\x0c":
#                 container.size = '40'
#             elif boxWrap.texts()[1]==r"\n":
#                 container.size = '40'
#             elif boxWrap.texts()[1]==r"\x04":
#                 container.size = '20'
#             elif boxWrap.texts()[1]==r"\x03":
#                 container.size = '20'
#             elif boxWrap.texts()[1]==r"\x0b":
#                 container.size = '40'
# #             elif boxWrap.texts()[1]==r"\r":
# #                 container.size = '45'
#             elif boxWrap.texts()[1]==r"\x07":
#                 container.size = '40'
# #         elif handleprops.rectangle(x)==RECT(197, 342, 256, 360):
#         elif handleprops.rectangle(x)==RECT(2090, 146, 2219, 168):
#             weight = x
# #         elif handleprops.rectangle(x)==RECT(512, 342, 569, 360):
#         elif handleprops.rectangle(x)==RECT(2089, 233, 2191, 255):
#             pieces = x
#         elif handleprops.rectangle(x)==RECT(625, 342, 722, 360):
#             containerNumber = x
#         elif handleprops.rectangle(x)==RECT(14, 295, 73, 313):
#             terminal = x
#         elif handleprops.rectangle(x)==RECT(1645, 136, 1818, 171):
#             bond = x
#         elif handleprops.rectangle(x)==RECT(182, 991, 250, 1015):
#             PB = x
        elif handleprops.rectangle(x)==rects[0]:
            container.description = handleprops.text(x)
#         elif handleprops.rectangle(x)==RECT(197, 342, 256, 360):
        elif handleprops.rectangle(x)==rects[1]:
            container.weight = handleprops.text(x)
#         elif handleprops.rectangle(x)==RECT(512, 342, 569, 360):
        elif handleprops.rectangle(x)==rects[2]:
            container.pieces = handleprops.text(x)
        elif handleprops.rectangle(x)==rects[3]:
            container.containerNumber = handleprops.text(x)
        elif handleprops.rectangle(x)==RECT(282, 295, 341, 313):
            container.terminal = handleprops.text(x)
        elif handleprops.rectangle(x)==RECT(1645, 136, 1818, 171):
            container.bond = handleprops.text(x)
        elif handleprops.rectangle(x)==RECT(182, 991, 250, 1015):
            container.PB = handleprops.text(x)
        elif handleprops.rectangle(x)==RECT(554, 295, 613, 313):
            container.customer = handleprops.text(x)
        elif handleprops.rectangle(x)==RECT(14, 891, 231, 985):
            container.extraText = handleprops.text(x)
        
#         Pb 362000 is used for widths
#         if handleprops.text(x)=="MAEU 463991 2":
#             print("desc")
#             print(handleprops.rectangle(x))
#         elif handleprops.text(x)=="7,660.00":
#             print("we")
#             print(handleprops.rectangle(x))
#         elif handleprops.text(x)=="6":
#             print("pi")
#             print(handleprops.rectangle(x))
#         elif handleprops.text(x)=="MAEU 411606 4":
#             print(handleprops.rectangle(x))
#     print(container.containerNumber)
    container.containerNumber = container.containerNumber.replace(' ', '')
    if (container.description == "" or
        container.weight == "" or
        container.pieces == "" or
        container.containerNumber == "" or
        container.terminal == "" or
        container.bond == "" or
        container.PB == ""):
        top = Tk()
        L0 = Label(top, text="Some of the information is missing. Either fill it in below, or ensure \n that your Dispatch-mate is properly formatted and hit \"Try again\". \n To do so, go to PB362000 and double click on the line between \n the boxes that read \"Description\" and \"Weight\"")
        L0.grid(row=0, column=0, columnspan=2)
        L1 = Label(top, text="Description:")
        L1.grid(row=1, column=0, sticky=constants.E)
        E1 = Entry(top, bd = 5)
        E1.grid(row=1, column=1)
        E1.insert(0, container.description)
        L2 = Label(top, text="Weight:")
        L2.grid(row=2, column=0, sticky=constants.E)
        E2 = Entry(top, bd = 5)
        E2.grid(row=2, column=1)
        E2.insert(0, container.weight)
        L3 = Label(top, text="Piece Count:")
        L3.grid(row=3, column=0, sticky=constants.E)
        E3 = Entry(top, bd = 5)
        E3.grid(row=3, column=1)
        E3.insert(0, container.pieces)
        L4 = Label(top, text="Container Number:")
        L4.grid(row=4, column=0, sticky=constants.E)
        E4 = Entry(top, bd = 5)
        E4.grid(row=4, column=1)
        E4.insert(0, container.containerNumber)
        L5 = Label(top, text="Consignee:")
        L5.grid(row=5, column=0, sticky=constants.E)
        E5 = Entry(top, bd = 5)
        E5.grid(row=5, column=1)
        E5.insert(0, container.terminal)
        L6 = Label(top, text="Bond #:")
        L6.grid(row=6, column=0, sticky=constants.E)
        E6 = Entry(top, bd = 5)
        E6.grid(row=6, column=1)
        E6.insert(0, container.bond)
        L7 = Label(top, text="PB #:")
        L7.grid(row=7, column=0, sticky=constants.E)
        E7 = Entry(top, bd = 5)
        E7.grid(row=7, column=1)
        E7.insert(0, container.PB)
        
        def callbackGoAhead(container):
            container.description = E1.get()
            container.weight = E2.get()
            container.pieces = E3.get()
            container.containerNumber = E4.get()
            container.terminal = E5.get()
            container.bond = E6.get()
            container.PB = E7.get()
            top.destroy()
        
        
        
        
        def checkStuff():
            app = Application(backend="win32").connect(path = r"C:\DM54_W16\DM54_W16.exe")
            winChildren = ""
            
            dialogs = app.windows()
            for x in dialogs:
                if handleprops.classname(x) == "WinDevObject":
                    winChildren = handleprops.children(x)
                    topWindow = x
                    break
            for x in winChildren:
                if handleprops.text(x)==E1.get():
                    rects[0]=handleprops.rectangle(x)
                if handleprops.text(x).replace(",", "")==E2.get():
                    rects[1]=handleprops.rectangle(x)
                if handleprops.text(x)==E3.get():
                    rects[2]=handleprops.rectangle(x)
                if handleprops.text(x)==E4.get():
                    rects[3]=handleprops.rectangle(x)
            writeRects(rects)
            
            
#             top.lift()
#             top.attributes('-topmost',True)
#             top.after_idle(top.attributes,'-topmost',False)
#             widget = E1
#             moveTo(widget.winfo_rootx()+widget.winfo_width()/2, widget.winfo_rooty()+5+widget.winfo_height()/2)

        def callbackDetect():
            top.after_idle(checkStuff)
            
        def callbackTryAgain(container):
            top.destroy()
            setupDM(container)
            
        MyButton4 = Button(top, text="Use these values", width=17, command=lambda: callbackGoAhead(container))
        MyButton4.grid(row=8, column=1)
        
        MyButton5 = Button(top, text="Try again", width=10, command=lambda: callbackTryAgain(container))
        MyButton5.grid(row=8, column=0)
        
        MyButton5 = Button(top, text="Set box locations: \n (copy description, weight, piece and cont# \n into the boxes above then hit this button)", width=10, command=callbackDetect)
        MyButton5.grid(row=9, column=0, columnspan=2, sticky=constants.W+constants.E)
        
        popUp(top, w=380, h=340, widget = E1)

def readRects():
    rects = []
    try:
        f=open("C:\Automation\TandE Rect Locations.txt", "r")
        for _ in range(4):
            rect = f.readline().split()
            rects.append(RECT(int(rect[0]), int(rect[1]), int(rect[2]), int(rect[3])))
        f.close()
    except:
#         rects = ([RECT(15, 344, 128, 357), RECT(2090, 146, 2219, 168), RECT(2089, 233, 2191, 255), RECT(755, 342, 842, 360)])
        rects = ([RECT(15, 344, 178, 357), RECT(2090, 146, 2219, 168), RECT(2089, 233, 2191, 255), RECT(755, 342, 842, 360)])
    return rects
        
def writeRects(rects):
    f=open("C:\Automation\TandE Rect Locations.txt", "w+")
    for rect in rects:
        f.write(str(rect.left)+" "+ str(rect.top)+" "+ str(rect.right)+" "+ str(rect.bottom)+"\n")
#         print(str(rect[0])+" "+ str(rect[1])+" "+ str(rect[2])+" "+ str(rect[3]))
    f.close()
         
if __name__ == '__main__':
    
    
    
#     rects=[]
#     
#     container = Container()
#     setupDM(container)
    driver = setupPortal()
         
    while True:
#         sleep(1)
        wait = WebDriverWait(driver, 100000000)
        wait.until(lambda driver: "Either an ID or Full Shipper information is required for a Shipper." in driver.page_source)
        callbackBookTE(driver)
#     time.sleep(5)
#     top = Tk()
#     
#     MyButton4 = Button(top, text="Do T&E", width=10, command=callbackBookTE(driver, top))
#     MyButton4.grid(row=0, column=0)
#     popUp(top, w=50, h=20, widget=MyButton4)
#     bookTE(driver)

# pyinstaller "C:\Users\ssleep\workspace\TandE\TandE\__init__.py" --distpath "J:\Spencer\TandE" --noconsole -y
