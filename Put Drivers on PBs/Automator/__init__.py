from pywinauto.win32structures import RECT
from pyautogui import click
from pyautogui import moveTo, press, typewrite, hotkey, position

from pywinauto import Application
from pywinauto import handleprops

from win32api import GetKeyState

import win32gui
from win32con import SWP_SHOWWINDOW

from tkinter import Button, Tk, Label, Entry, Text

from sys import argv
import os

from os import listdir
from openpyxl import load_workbook
from openpyxl.styles.fills import FILL_NONE
from _datetime import date, timedelta

from time import sleep
import HelperFunctions
import sys
from HelperFunctions import popUpOK

class Driver(object):
    PARS = ""
    driver = ""
    name = ""
    city = ""
    
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

def setupDM(folderPath, drivers):
#     app = Application(backend="win32").connect(path = r"C:\DM54_W16\DM54_W16.exe")
#     
# #     top_windows = []
# #     EnumWindows(windowEnumerationHandler, top_windows)
# #     for i in top_windows:
# #         if 'Dispatch-Mate' in i[1]:
# #             SetWindowPos(i[0], None, 0, 0, 1920, 1080, SWP_SHOWWINDOW)
# #             SetForegroundWindow(i[0])
#             
#     
#     winChildren = ""
#      
#     dialogs = app.windows()
#     
#     for x in dialogs:
#         if handleprops.classname(x) == "WinDevObject":
#             winChildren = handleprops.children(x)
#             topWindow = x
#             break
    
    clickTuple = False
    
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
        
#     i=0
    for driver in drivers:
#         if i>0: 
    
        click(327, 33)
         
        sleep(3)
         
        click(399, 281)
        
        if driver.PARS[-1]=="A" or driver.PARS[-1]=="B" or driver.PARS[-1]=="C":
            typewrite(driver.PARS[-7:-1])
        else:
            typewrite(driver.PARS[-6:])
         
        press("enter")
         
        sleep(3)
         
        click(1857, 100)
        
        sleep(0.5)
         
        click(569, 900)
         
        month = str(pickupDate.month)
        if len(month)<2:
            month = "0" + month
        typewrite(month)
        day = str(pickupDate.day)
        if len(day)<2:
            day = "0" + day
        typewrite(day)
        typewrite(str(pickupDate.year))
         
        click(569, 926)
         
        month = str(pickupDate.month)
        if len(month)<2:
            month = "0" + month
        typewrite(month)
        day = str(pickupDate.day)
        if len(day)<2:
            day = "0" + day
        typewrite(day)
        typewrite(str(pickupDate.year))
         
        click(695, 100)
         
        click(1542, 730)
         
        typewrite(str(driver.driver))
         
        press('enter')
        
        if str(driver.driver)[:3]!="801":
            click(1558,679)
            
            typewrite(driver.name)
         
            press('enter')
                
            sleep(0.5)    
            
            thru = False
            
            click(189, 854)
            hotkey('ctrl', 'a')
            hotkey('ctrl', 'c')
            clipTk=Tk()
            if "THRUWAY" in clipTk.clipboard_get():
                thru = True
                
            typewrite("TRUCK")
            press('tab')
            typewrite("0.55")
            press('tab')
            press('delete')
            press('tab')
            miles = "490"
            if str(driver.city)=="PACKER":
                miles = "507"
            if str(driver.city)=="NYCT":
                miles = "500"
            typewrite(miles)
            
            click(189, 890)
            hotkey('ctrl', 'a')
            hotkey('ctrl', 'c')
            if "THRUWAY" in clipTk.clipboard_get():
                thru = True
                
            click(189, 872)
            hotkey('ctrl', 'a')
            hotkey('ctrl', 'c')
            if "THRUWAY" in clipTk.clipboard_get():
                thru = True
            
            if thru:
                typewrite("COMPANY DR THRUWAY")
            
                press('tab')
            
                typewrite("1.00")
                press('tab')    
            
            
                press('delete')
                press('tab')
            
                if str(driver.city)=="PACKER":
                    typewrite("81.4")
                else:
                    typewrite("31.9")
                    
            else:
                press('delete')
                press('tab')
                press('delete')
                press('tab')
                press('delete')
                press('tab')
                press('delete')
            
            click(189, 890)
            hotkey('ctrl', 'a')
            hotkey('ctrl', 'c')
            if "THRUWAY" in clipTk.clipboard_get():
                thru = True
            clipTk.destroy()
            if thru:
                press('delete')
                press('tab')
                press('delete')
                press('tab')
                press('delete')
                press('tab')
                press('delete')
            
        if GetKeyState(145) < 0:
            exit()    
        
        sleep(2)
        
        click(1857, 120)
        click(1857, 120)
        
        sleep(3)
        
        click(1845, 200, button="right")
        
        sleep(1)
        
        click(1826, 312)
        
        sleep(0.3)
        
#         click(1600, 316)
#         
#         sleep(7)
        click(1600, 361)
        sleep(0.3)
        click(1772, 160)
             
#         sleep(5)
#         sleep(5)
        done=False
        while not done:
            try:
                if os.path.isfile(r"C:\Program Files\Microsoft Office 15\root\office15\outlook.exe"):
                    app = Application(backend="win32").connect(path = r"C:\Program Files\Microsoft Office 15\root\office15\outlook.exe")
                elif os.path.isfile(r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"):
                    app = Application(backend="win32").connect(path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE")
                elif os.path.isfile(r"C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"):
                    app = Application(backend="win32").connect(path = r"C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE")
    
#                 elif os.path.isfile(r"C:\Program Files\WindowsApps\Microsoft.Office.Desktop.Outlook_16040.10827.20138.0_x86__8wekyb3d8bbwe\Office16\outlook.exe"):
#                     app = Application(backend="win32").connect(path = r"C:\Program Files\WindowsApps\Microsoft.Office.Desktop.Outlook_16040.10827.20138.0_x86__8wekyb3d8bbwe\Office16\outlook.exe")    
                else:
                    directoryPath = "C:\Program Files\WindowsApps\\"
                    if os.path.isdir("C:\Program Files\WindowsApps\\"):
                        contents = os.listdir("C:\Program Files\WindowsApps\\")
                        outLookFolders=[]
                        for folder in contents:
                            if os.path.isdir("C:\Program Files\WindowsApps\\"+folder):
                                if "outlook" in folder.lower():
                                    outLookFolders.append(folder)
                    outlookPath = ""
                    for folder in outLookFolders:
                        contents = os.listdir("C:\Program Files\WindowsApps\\"+folder)
                        for officeFolder in contents:
                            if os.path.isdir(directoryPath+folder+"\\"+officeFolder):
                                if "office" in officeFolder.lower():
                                    contentsInner = os.listdir(directoryPath+folder+"\\"+officeFolder)
                                    for outlookProgram in contentsInner:
                                        if outlookProgram.lower()=="outlook.exe":
                                            outlookPath=directoryPath+folder+"\\"+officeFolder+"\\"+outlookProgram
                    if os.path.isfile(outlookPath):
                        app = Application(backend="win32").connect(path = outlookPath)
                    else:    
                        popUpOK("Could not find Outlook in \"C:\Program Files\Microsoft Office 15\root\office15\outlook.exe\" \n or \"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE\" \n or \"C:\Program Files (x86)\Microsoft Office\Office14\" \n or the \"C:\Program Files\WindowsApps\\\" folder")
                        exit()
                done=True
            except:
                pass
#             top_windows = []
#             EnumWindows(windowEnumerationHandler, top_windows)
#             for i in top_windows:
#                 if 'Dispatch-Mate' in i[1]:
#                     SetWindowPos(i[0], None, 0, 0, 1920, 1080, SWP_SHOWWINDOW)
#                     SetForegroundWindow(i[0])
                  
        winChildren = ""
            
        done = False
        
        while not done:
            dialogs = app.windows()
            topWindow = None
            for x in dialogs:
                if isinstance(handleprops.text(x), str) and not handleprops.text(x)==None:
                    try:
                        if "Carrier Confirmation" in handleprops.text(x):
                            winChildren = handleprops.children(x)
                            topWindow = x
                            break
                    except:
                        pass
                
            send = ""
            if topWindow==None:
                continue
                
            topWindowWrap = app.window(handle=topWindow)
            
            for x in winChildren:
#                 print(handleprops.text(x) + "   " + handleprops.classname(x))
                if handleprops.text(x)=="&Send":
                    send = x
                    done = True
#                 if handleprops.text(x)=="Fro&m":
#                     buttonWrap = topWindowWrap.child_window(handle=x).wrapper_object()
#                     buttonWrap.click()
#                     
                    
#                     
#                     if not clickTuple:
#                         moveTo(114, 221)
#                          
#                         while not GetKeyState(145)<0:
#                             True
#                          
#                         clickTuple = position()
#                      
# #                         else:                        
#                     click(clickTuple)
            
            if done==True:            
                buttonWrap = topWindowWrap.child_window(handle=send).wrapper_object()
                buttonWrap.click()
                
        
            
            
          
#             winChildren = handleprops.children(topWindow)
#               
#             for x in winChildren:
# #                 print(handleprops.text(x) + "   " + handleprops.classname(x))
#                 if handleprops.classname(x)=="NetUIHWND":
#                     topWindowWrap = app.window(handle=topWindow)
#                     netWrap = topWindowWrap.child_window(handle=x).wrapper_object()
#                     print(netWrap.Texts())
#               
#                     for y in handleprops.children(x):
#                         print(handleprops.text(y)) 
#       
#             app.top_window().window(title="From", control_type="Button").print_control_identifiers()
        
#             if not clickTuple:
#                 
#                 moveTo(113, 167)
#                 
#                 while not GetKeyState(13)<0:
#                     True
#                 
#                 clickTuple = position()
#                 
#             else:
#                 sleep(5)
#                 click((113, 167))
#                 
#             click(clickTuple)
        sleep(4)
#     exit()
            
#         i = i+1
    
                

def loadinfo(folderPath):
#     for x in listdir(folderPath):
#         if not x[0]=="~":
#             if "Imports" in x:
#                 imports = x
            
    imports = load_workbook(folderPath, data_only=True)
    colorSheet = imports['COLOUR KEY']
    driverSheet = imports['Drivers']
    imports = imports['Imports Sheet']
    
    driverCol = ""
    nameCol = ""
    parsCol = ""
    termCol = ""
    
    for cell in next(imports.rows):
        if cell.value == "DRIVER":
            driverCol = cell.col_idx - 1
        elif cell.value == "NAME":
            nameCol = cell.col_idx - 1
        elif cell.value == "PARS NUMBER":
            parsCol = cell.col_idx - 1
        elif cell.value == "Terminal":
            termCol = cell.col_idx - 1
    
    colors = ["","","","",""]
    for xRow in colorSheet.rows:
        for xCell in xRow:
            if xCell.value=="MONDAY":
                colors[0]=xCell.offset(0,2).fill
                break
            elif xCell.value=="TUESDAY":
                colors[1]=xCell.offset(0,2).fill
                break
            elif xCell.value=="WEDNESDAY":
                colors[2]=xCell.offset(0,2).fill
                break
            elif xCell.value=="THURSDAY":
                colors[3]=xCell.offset(0,2).fill
                break
            elif xCell.value=="FRIDAY":
                colors[4]=xCell.offset(0,2).fill
                break
    

    dayOfTheWeek = [-1]
    startAt = [""]
    top = Tk()
    L1 = Label(top, text="Please select which driver to start at \n as well as which day of the week to run on \n ")
    L1.config(font=("Courier", 16))
    L1.grid(row=0, column=0, columnspan=5)
    L2 = Label(top, text="Start at driver #:")
    L2.config(font=("Courier", 10))
    L2.grid(row=1, column=0, columnspan = 2)
    E1 = Entry(top, bd = 5, width = 39)
    E1.grid(row=1, column=2, columnspan = 2)
    L3 = Label(top, text="OR")
    L3.grid(row=2, column=1, columnspan=2)
    L3.config(font=("Courier", 20))
    top.lift()
    top.attributes('-topmost',True)
    top.after_idle(top.attributes,'-topmost',False)
      
    def callbackDay(day):
        startAt[0]=E1.get().strip()
        dayOfTheWeek[0] = day
        top.destroy()
      
      
    
    MyButton5 = Button(top, text="MONDAY", command=lambda: callbackDay(0))
    MyButton5.grid(row=4, column=0)
    MyButton5.config(font=("Courier", 16))
    MyButton6 = Button(top, text="TUESDAY", command=lambda: callbackDay(1))
    MyButton6.grid(row=4, column=1)
    MyButton6.config(font=("Courier", 16))
    MyButton7 = Button(top, text="WEDNESDAY",  command=lambda: callbackDay(2))
    MyButton7.grid(row=4, column=2)
    MyButton7.config(font=("Courier", 16))
    MyButton8 = Button(top, text="THURSDAY", command=lambda: callbackDay(3))
    MyButton8.grid(row=4, column=3)
    MyButton8.config(font=("Courier", 16))
    MyButton9 = Button(top, text="FRIDAY", command=lambda: callbackDay(4))
    MyButton9.grid(row=4, column=4)
    MyButton9.config(font=("Courier", 16)) 
    
    w = 620 # width for the Tk root
    h = 260 # height for the Tk root
       
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
    
    year = folderPath.split("\\")[-3]
    week = folderPath.split("\\")[-2].split(" ")[-1]
    pickupDate = date(int(year), 1, 1)
    
    oneDay = timedelta(1)
    oneWeek = timedelta(7)
    
    while pickupDate.weekday() != 0:
        pickupDate = pickupDate+oneDay
    
    pickupDate = pickupDate + (int(week)-1)*oneWeek
    
    pickupDate = pickupDate + (dayOfTheWeek[0])*oneDay
    
    todaysFill = colors[dayOfTheWeek[0]]
    
    drivers = []
    
    startFound = False
    if startAt[0] != "":
        for row in imports.rows:
            if row[0].fill.fgColor == todaysFill.fgColor:
                if (not startFound) and str(row[driverCol].value)==startAt[0]:
                    startFound = True
                if startFound:
                    driver = Driver()
                    driver.driver = row[driverCol].value
                    driver.name = row[nameCol].value
                    driver.PARS = row[parsCol].value.strip()
                    driver.city = row[termCol].value
                    drivers.append(driver)
                    
    else:
        for row in imports.rows:
            if row[0].fill.fgColor == todaysFill.fgColor:
                driver = Driver()
#                 print("name: " +str(nameCol))
                driver.name = row[nameCol].value
                driver.driver = row[driverCol].value
                driver.PARS = row[parsCol].value.strip()
                driver.city = row[termCol].value
                drivers.append(driver)
    
    nameDict = {}
    
    for row in driverSheet:
#         print(row[0].fill)
#         if row[0].fill.patternType==None:
        nameDict[row[1].value] = row[0].value
        nameDict[row[2].value] = row[0].value
#     print(list(nameDict.values()))
    driverNameSorted = list(nameDict.values())
    driverNameSorted.sort()
    alternate = False
    for driverName in driverNameSorted.copy():
        if alternate:
            driverNameSorted.remove(driverName)
        alternate = not alternate

    for driver in drivers:
        if not str(driver.driver)[:3]=="801":
            if driver.name in nameDict:
                driver.name = nameDict[driver.name]
#                 print(driver.name)
            elif not driver.name in driverNameSorted:
                if not driver.name:
                    driver.name = "NONE"
                top = Tk()
                L1 = Label()
                L1 = Label(top, text="Company Driver \"" + driver.name + "\" not recognized. \n Please select corresponding name:")
                L1.config(font=("Courier", 16))
                L1.grid(row=0, column=0)
#                 top.lift()
#                 top.attributes('-topmost',True)
#                 top.after_idle(top.attributes,'-topmost',False)
                  
#                 def callbackDriver(driverName, driver):
#                     print(driver.name)
#                     driver.name = driverName
#                     print(driver.name)
#                     top.destroy()
                
                
                
                def callbackDriver(driverNameSorted, q, driver):
                    driver.name = driverNameSorted[q]
                    top.destroy()
                
#                 listI = range(len(driverNameSorted))
                
                
                for i in range(len(driverNameSorted)):
                    MyButton = Button(top, text=driverNameSorted[i], command=lambda i=i: callbackDriver(driverNameSorted, i, driver), width=40)
                    MyButton.grid(row=i+1, column=0)
                  
                
                w = 500 # width for the Tk root
                h = i*20 + 200 # height for the Tk root
                   
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
                
#                 print(driver.name)
                
                
                
    return drivers, pickupDate


if __name__ == '__main__':
    
#     argv = r"a J:\Linehaul\Linehaul Drivers Weekly Reports\2019\2019\Week 32\Imports week 32.xlsx".split()

#     done=False
#     while not done:
#         try:
#             if os.path.isfile(r"C:\Program Files\Microsoft Office 15\root\office15\outlook.exe"):
#                 app = Application(backend="win32").connect(path = r"C:\Program Files\Microsoft Office 15\root\office15\outlook.exe")
#             elif os.path.isfile(r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"):
#                 app = Application(backend="win32").connect(path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE")
#             elif os.path.isfile(r"C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"):
#                 app = Application(backend="win32").connect(path = r"C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE")
# 
# #                 elif os.path.isfile(r"C:\Program Files\WindowsApps\Microsoft.Office.Desktop.Outlook_16040.10827.20138.0_x86__8wekyb3d8bbwe\Office16\outlook.exe"):
# #                     app = Application(backend="win32").connect(path = r"C:\Program Files\WindowsApps\Microsoft.Office.Desktop.Outlook_16040.10827.20138.0_x86__8wekyb3d8bbwe\Office16\outlook.exe")    
#             else:
#                 directoryPath = "C:\Program Files\WindowsApps\\"
#                 if os.path.isdir("C:\Program Files\WindowsApps\\"):
#                     contents = os.listdir("C:\Program Files\WindowsApps\\")
#                     outLookFolders=[]
#                     for folder in contents:
#                         if os.path.isdir("C:\Program Files\WindowsApps\\"+folder):
#                             if "outlook" in folder.lower():
#                                 outLookFolders.append(folder)
#                 outlookPath = ""
#                 for folder in outLookFolders:
#                     contents = os.listdir("C:\Program Files\WindowsApps\\"+folder)
#                     for officeFolder in contents:
#                         if os.path.isdir(directoryPath+folder+"\\"+officeFolder):
#                             if "office" in officeFolder.lower():
#                                 contentsInner = os.listdir(directoryPath+folder+"\\"+officeFolder)
#                                 for outlookProgram in contentsInner:
#                                     if outlookProgram.lower()=="outlook.exe":
#                                         outlookPath=directoryPath+folder+"\\"+officeFolder+"\\"+outlookProgram
#                 if os.path.isfile(outlookPath):
#                     app = Application(backend="win32").connect(path = outlookPath)
#                 else:    
#                     popUpOK("Could not find Outlook in \"C:\Program Files\Microsoft Office 15\root\office15\outlook.exe\" \n or \"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE\" \n or \"C:\Program Files (x86)\Microsoft Office\Office14\" \n or the \"C:\Program Files\WindowsApps\\\" folder")
#                     exit()
#             done=True
#         except:
#             pass
#     
#     winChildren = ""
#             
#     done = False
#     
#     while not done:
#         dialogs = app.windows()
#         topWindow = None
#         for x in dialogs:
#             if isinstance(handleprops.text(x), str) and not handleprops.text(x)==None:
#                 try:
#                     if "Carrier Confirmation" in handleprops.text(x):
#                         winChildren = handleprops.children(x)
#                         topWindow = x
#                         break
#                 except:
#                     pass
#             
#         send = ""
#         if topWindow==None:
#             continue
#             
#         topWindowWrap = app.window(handle=topWindow)
#         
#         for x in winChildren:
# #                 print(handleprops.text(x) + "   " + handleprops.classname(x))
#             if handleprops.text(x)=="&Send":
#                 send = x
#                 print("found Send")
#             if handleprops.text(x)=="Fro&m":
#                 print("found From")
#                 buttonWrap = topWindowWrap.child_window(handle=x).wrapper_object()
#                 buttonWrap.click()
#                 
# #                 done = True
#                 
# #                 if not clickTuple:
# #                     moveTo(114, 221)
# #                      
# #                     while not GetKeyState(145)<0:
# #                         True
# #                      
# #                     clickTuple = position()
# #                  
# # #                         else:                        
# #                 click(clickTuple)
#         
#         sleep(100)
#         if done==True:            
#             buttonWrap = topWindowWrap.child_window(handle=send).wrapper_object()
#             buttonWrap.click()
#     
#     
#     
#     exit()
    
    folderPath = ''
    for i in range(len(argv)):
        if i!=0:
            folderPath+=argv[i]
            if i != len(argv) - 1:
                folderPath+=" "
#     try:
    drivers, pickupDate = loadinfo(folderPath)
#     
#     
    setupDM(pickupDate, drivers)
    
    HelperFunctions.done()
#     except:
#         print(sys.exc_info())
#         sleep(100)
#     pyinstaller "C:\Users\ssleep\workspace\Put Drivers on PBs\Automator\__init__.py" --distpath "J:\Spencer\Linehaul Drivers on PBs" -y
#     pyinstaller "C:\Users\ssleep\workspace\Put Drivers on PBs\Automator\__init__.py" --distpath "J:\Spencer\Linehaul Drivers on PBs" --noconsole -y