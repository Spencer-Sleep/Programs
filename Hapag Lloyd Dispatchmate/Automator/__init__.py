from sys import argv, exc_info
from time import sleep
from PyPDF2 import PdfFileReader
from os import path
from os import listdir
# from pdfminer.pdfparser import PDFParser, PDFDocument #@UnresolvedImport
# from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter #@UnresolvedImport
# from pdfminer.converter import PDFPageAggregator #@UnresolvedImport
# from pdfminer.layout import LAParams, LTTextBox, LTTextLine #@UnresolvedImport

from pywinauto.keyboard import SendKeys
from tkinter import Button, Tk, Label, Entry
from datetime import datetime
from datetime import timedelta

from pyautogui import moveTo, click, press, typewrite
from pymsgbox import confirm
from win32gui import GetWindowText, GetForegroundWindow

import PyPDF2
from PIL.EpsImagePlugin import field
 
from threading import Thread
import msvcrt

import sys
from win32api import GetKeyState

from openpyxl import load_workbook

import ContainerSizeInfo

import DispatchmateLocations as loc
from HelperFunctions import done

# DISPTACHMATE LOCATIONS

# SHIPPERCODELOC = (48,301)
# CONSIGNEECODELOC = (301,301)
# CUSTOMERCODELOC = (564,301)
# 
# DESCRIPTIONLOC = (20,354)
# 
# EQUIPMENTLOC = (327, 900)
# DIRECTIONLOC = (400,922)
# DIVISIONLOC = (550,944)
# LANECODELOC = (568, 967)
# 
# HOUSELOC = (1807, 1000)
# 
# OKLOC = (1857, 127)
# DUPLICATELOC = (1845, 273)
# 
# ROUTINGTABLOC = (688, 96)
# DRIVERPAYOUTLOC = (132, 858)
#  
# RATINGTABLOC = (1153, 98)
# CUSTOMERCHARGELOC = (49,144)
# 
# NEWLOC = (1850,100)

contRates = {}
laneCodes = {}

def recursiveBookPB(specificPath):
#     print(specificPath)
#     sleep(5)
    for filename in listdir(specificPath):
        if "PARS MANIFESTS" in filename and filename[-4:] == ".pdf" or filename[-4:] == ".PDF":
            bookPB(specificPath+'\\'+filename)
        elif(path.isdir(specificPath+"\\"+filename) and not filename=="Flattened"):
            recursiveBookPB(specificPath+"\\"+filename)

    
    
def bookPB(specificPath): 
#     text = ""
    pdfFileObj = open(specificPath, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    
    fields = pdfReader.getFields()
#     print(len(fields)-15)


    for i in range(len(fields)-15):
#         print(i)
        wo = ""
        terminal = ""
        containerNumber = ""
        size = ""
        weight = ""
        
        description = ""
        if i == 0:
#             prefix = str(i) + "."
            wo = fields["WO"].value
            terminal = fields["Port of discharge"].value
            containerNumber = fields["Container Row1"].value
            size = fields["SizeRow1"].value
            weight = float(fields["Weight KGRow1"].value)
            description = fields["Description of goods"].value
        else:
            for j in list(fields.keys()):
                if j==str(i):
                    for k in list(fields[j]["/Kids"]):
                        try:
                            if(k.getObject()['/T']=="WO"):
                                wo=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Port of discharge"):
                                terminal=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Container Row1"):
                                containerNumber=k.getObject()['/V']
                            elif(k.getObject()['/T']=="SizeRow1"):
                                size=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Weight KGRow1"):
                                weight=float(k.getObject()['/V'])
                            elif(k.getObject()['/T']=="Description of goods"):
                                description=k.getObject()['/V']
                        except KeyError:
                            True

#         for j in list(fields.keys()):
#             print(j)
#         print(wo)
#         print(terminal)
#         print(containerNumber)
#         print(size)
#         print(weight)
        terminalNum = [""]
        if terminal == 'NYCT':
            terminalNum[0] = '664'
        elif terminal == 'GLOBAL':
            terminalNum[0] = '304'
        elif terminal == 'PACKER':
            terminalNum[0] = '309'
        elif terminal == 'APM':
            terminalNum[0] = '306'
        elif terminal == 'ASI':
            terminalNum[0] = '305'
        elif terminal == 'PNCT':
            terminalNum[0] = '310'
        elif terminal == 'MAHER':
            terminalNum[0] = '330'
        else:
            top = Tk()
            L1 = Label(top, text="Please enter Shipper (Terminal) # for container\n" + containerNumber + " at terminal \n\"" + terminal + "\":")
            L1.grid(row=0, column=0)
            E1 = Entry(top, bd = 5)
            E1.grid(row=1, column=0)
             
            def callbackTerminal(terminalNum):
                terminalNum[0] = E1.get()
                top.destroy()
             
            MyButton4 = Button(top, text="OK", width=10, command=lambda: callbackTerminal(terminalNum))
            MyButton4.grid(row=2, column=0)
             
            top.lift()
            top.attributes('-topmost',True)
            top.after_idle(top.attributes,'-topmost',False)
             
            w = 300 # width for the Tk root
            h = 150 # height for the Tk root
             
            # get screen width and height
            ws = top.winfo_screenwidth() # width of the screen
            hs = top.winfo_screenheight() # height of the screen
             
            # calculate x and y coordinates for the Tk root window
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
             
            # set the dimensions of the screen 
            # and where it is placed
            top.geometry('%dx%d+%d+%d' % (w, h, x, y))
             
             
            def click2():
                click(907, 561)
             
            top.after(10, click2)
            top.mainloop()
        
        if terminal == "PACKER" and weight>contRates[size].T4weight:
            top = Tk()
            L1 = Label(top, text="Container " + containerNumber + " is too heavy for PA \n Skipping container.")
            L1.grid(row=0, column=0)
             
            def callbackTerminal(terminalNum):
                top.destroy()
             
            MyButton4 = Button(top, text="OK", width=10, command=lambda: callbackTerminal(terminalNum))
            MyButton4.grid(row=1, column=0)
             
            top.lift()
            top.attributes('-topmost',True)
            top.after_idle(top.attributes,'-topmost',False)
             
            w = 300 # width for the Tk root
            h = 150 # height for the Tk root
             
            # get screen width and height
            ws = top.winfo_screenwidth() # width of the screen
            hs = top.winfo_screenheight() # height of the screen
             
            # calculate x and y coordinates for the Tk root window
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
             
            # set the dimensions of the screen 
            # and where it is placed
            top.geometry('%dx%d+%d+%d' % (w, h, x, y))
             
            moveTo(1001, 508)
#             def click2():
#                 click(1001, 508)
#              
#             top.after(10, click2)
            top.mainloop()
        
        click(loc.OKLOC)
        click(loc.NEWLOC)
        
        click(loc.SHIPPERCODELOC)
        typewrite(terminalNum[0])
        press('tab')
         
        click(loc.CONSIGNEECODELOC)
        typewrite("303")
        press('tab')
         
        click(loc.CUSTOMERCODELOC)
        typewrite("522")
        press('tab')
        typewrite(wo)
        
        click(loc.DESCRIPTIONLOC)
        press("tab", 7)
        typewrite("1")
         
        press("tab")
        typewrite("r")
         
        press("tab")
        typewrite(containerNumber)
     
        press("tab", 3)
        typewrite("HAPAG")
         
        click(loc.EQUIPMENTLOC)
        downAmount = 0
        if size == "20D86":
            downAmount = 5
        elif size == "20R86":
            downAmount = 4
        elif size == "40D86":
            downAmount = 8
        elif size == "40D96":
            downAmount = 11
        elif size == "40R96":
            downAmount = 13
        press('down', downAmount)
        press('enter')
         
        click(loc.DIRECTIONLOC)
        typewrite("imp")
        press("enter")
         
        click(loc.DIVISIONLOC)
        typewrite("usa")
        press("enter")
        
        try:
            click(loc.LANECODELOC)
            if description=="Coffee" and laneCodesCoffee[terminal]==556:
                typewrite("56")
                press('up')
            elif description=="Coffee":
                typewrite(str(laneCodesCoffee[terminal]))
            elif weight>contRates[size].T4weight and laneCodesThru[terminal]==555:
                typewrite("56")
                press("up", 2)
            elif weight>contRates[size].T4weight:
                typewrite(laneCodesThru[terminal])
            elif laneCodes[terminal]==554:
                typewrite("56")
                press('up', 3)
            else:
                typewrite(str(laneCodes[terminal]))
            press("enter")
        except:
            laneCode = [""]
            top = Tk()
            L1 = Label(top, text="Please enter Lane Code \n for terminal " + terminal + ":")
            L1.grid(row=0, column=0)
            E1 = Entry(top, bd = 5)
            E1.grid(row=1, column=0)
             
            def callbackTerminal(terminalNum):
                laneCode[0] = E1.get()
                top.destroy()
             
            MyButton4 = Button(top, text="OK", width=10, command=lambda: callbackTerminal(terminalNum))
            MyButton4.grid(row=2, column=0)
             
            top.lift()
            top.attributes('-topmost',True)
            top.after_idle(top.attributes,'-topmost',False)
             
            w = 300 # width for the Tk root
            h = 150 # height for the Tk root
             
            # get screen width and height
            ws = top.winfo_screenwidth() # width of the screen
            hs = top.winfo_screenheight() # height of the screen
             
            # calculate x and y coordinates for the Tk root window
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
             
            # set the dimensions of the screen 
            # and where it is placed
            top.geometry('%dx%d+%d+%d' % (w, h, x, y))
             
             
            def click2():
                click(935, 544)
             
            top.after(10, click2)
            top.mainloop()
            
            click(loc.LANECODELOC)
            if description == "Coffee":
                laneCode[0] = laneCode[0][:2] + "6"
            typewrite(str(laneCode[0]))
            press("enter")
            
            
        click(loc.HOUSELOC)
        typewrite("LONG")
         
        
#   *****************************
#        Routing Tab
#   *****************************
        overweight = False
        reefer = False
        click(loc.ROUTINGTABLOC)
         
         
#         if description == "COFFEE":
#             overweightTier = 1
#         else:
#             if size == "20D86":
#                 if weight>19050:
#                     overweight=True
#                 if weight<19640:
#                     overweightTier=0
#                 elif weight<20774:
#                     overweightTier=1
#                 elif weight<21908:
#                     overweightTier=2
#                 elif weight<24176:
#                     overweightTier=3
#                 else:
#                     overweightTier=4
#             elif size == "20R86":
#                 reefer=True
#                 if weight>18143:
#                     overweight=True
#                 if weight<18778:
#                     overweightTier=0
#                 elif weight<19912:
#                     overweightTier=1
#                 elif weight<21046:
#                     overweightTier=2
#                 elif weight<23314:
#                     overweightTier=3
#                 else:
#                     overweightTier=4
#             elif size == "40D86":
#                 if weight>19958:
#                     overweight=True
#                 if weight<18189:
#                     overweightTier=0
#                 elif weight<19323:
#                     overweightTier=1
#                 elif weight<20457:
#                     overweightTier=2
#                 elif weight<22725:
#                     overweightTier=3
#                 else:
#                     overweightTier=4
#             elif  size == "40D96":
#                 if weight>19958:
#                     overweight=True
#                 if weight<17962:
#                     overweightTier=0
#                 elif weight<19096:
#                     overweightTier=1
#                 elif weight<20230:
#                     overweightTier=2
#                 elif weight<22498:
#                     overweightTier=3
#                 else:
#                     overweightTier=4
#             elif  size == "40R96":
#                 reefer=True
#                 if weight>18143:
#                     overweight=True
#                 if weight<17418:
#                     overweightTier=0
#                 elif weight<18552:
#                     overweightTier=1
#                 elif weight<19686:
#                     overweightTier=2
#                 elif weight<21954:
#                     overweightTier=3
#                 else:
#                     overweightTier=4
         
         
        payout = 0
        OW = False
        
        
        if weight>contRates[size].P1weight:
            overweight = True
        if "R" in size:
            reefer = True
        
        
        if description == "COFFEE": 
            payout = contRates[size].T1payout
        else:
            if weight>contRates[size].T4weight:
                payout = contRates[size].T4payout
                OW = True
            elif weight>contRates[size].T3weight:
                payout = contRates[size].T3payout
            elif weight>contRates[size].T2weight:
                payout = contRates[size].T2payout
            elif weight>contRates[size].T1weight:
                payout = contRates[size].T1payout
        
        if payout != 0:
            click(loc.DRIVERPAYOUTLOC)
            typewrite("over")
            
            press("tab", 3)
              
#             f=open("J:\Spencer\Hapag-Lloyd Dispatchmate\Overweights.txt")
#             overweights = []
#             for i in range(6):
#                 overweights.append(f.readline())
            
            typewrite(str(payout))
            if OW:
                click(loc.DRIVERPAYOUTLOC[0], loc.DRIVERPAYOUTLOC[1]+19)
                typewrite("thru")
                press("tab", 3)
                typewrite(str(contRates[size].OWpayout))
            
         
         
#   *****************************
#        Rating Tab
#   *****************************
 
 
        click(loc.RATINGTABLOC)
         
        clickHeight = 0
         
#         click(89, clickHeight)
#         clickHeight += 20
#         typewrite("ont")
#         press("tab")
#         if terminalNum[0]=="309":
#             typewrite("946")
#         else:
#             typewrite("858")
         
        if(reefer):
            click(loc.CUSTOMERCHARGELOC[0], loc.CUSTOMERCHARGELOC[1]+ clickHeight)
            clickHeight += 19
            typewrite('re')
            press("down")
         
        if(overweight):
            click(loc.CUSTOMERCHARGELOC[0], loc.CUSTOMERCHARGELOC[1]+ clickHeight)
            clickHeight += 19
            typewrite("ov")
            press("tab")
            
        if terminalNum[0]=="664":
            click(loc.CUSTOMERCHARGELOC[0], loc.CUSTOMERCHARGELOC[1]+ clickHeight)
            clickHeight += 19
            typewrite("n")
            press("tab")
        
#         click(1871, 339)
         
#         click(1866, 100)
         
#         pbUp = False
#         pbDone = False
#         print(str(GetKeyState(27)))

        click(loc.OKLOC)

        if GetKeyState(27) < 0:
            sys.exit()
 
#         click(OKLOC)
#         click(NEWLOC)
         
#         while not pbDone:
#             window = GetForegroundWindow()
#             if "Dispatch-Mate" in GetWindowText(window) and pbUp:
#                 pbDone = True
#             elif "Available Load" in GetWindowText(window) and not "Dispatch-Mate" in GetWindowText(window):
#                 pbUp = True
         
#         click(1876, 101)
        
        
# def listenForExit():
#     i=0
#     while True:
#         print("AAAAAAAAAA")
#         key = getch()
# #         f=open(("testfile.txt"), "w")
# #         f.write(str(key) + "\n")
# #         f.close()
#         print(ord(key))
# #         if(ord(getch()) == 27):
# #             sys.exit()
#         sleep(1)
#    
if __name__ == '__main__':
#     argv = r"a J:\All motor routings\2018\Week 22\Hapag-Lloyd".split()
    
#     t = Thread(target=listenForExit)
#     t.setDaemon(True)
#     t.start()
    
#     key = msvcrt.getche()
#     print(ord(key))
    
    specificPath = ''
    for i in range(len(argv)):
        if i!=0:
            specificPath+=argv[i]
            if i != len(argv) - 1:
                specificPath+=" "
#     destinationOfFiles = workOrderLocation[:workOrderLocation.rfind('\\')] + "\\"
#     print(specificPath)
    
    values = ContainerSizeInfo.loadValues(True, True, "NJ HAPAG-LLOYD: $250", "", "", True)
    contRates = values[0]
    laneCodes = values[1]
    laneCodesCoffee = values[2]
    laneCodesThru = values[3]

    if path.isdir(specificPath):
        recursiveBookPB(specificPath)
    elif "PARS MANIFESTS" in specificPath and (specificPath[-4:] == ".pdf" or specificPath[-4:] == ".PDF"):
        bookPB(specificPath)
        
    done()