   
# import re
# import subprocess
# import sys
from sys import argv, exc_info
from time import sleep
from PyPDF2 import PdfFileReader
# import os
from os import path
from os import listdir
from pdfminer.pdfparser import PDFParser, PDFDocument #@UnresolvedImport
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter #@UnresolvedImport
from pdfminer.converter import PDFPageAggregator #@UnresolvedImport
from pdfminer.layout import LAParams, LTTextBox, LTTextLine #@UnresolvedImport
# from pywinauto import Application, findwindows, handleprops
# import pywinauto
# import win32 #@UnresolvedImport
# from win32 import win32process #@UnresolvedImport
# from _overlapped import NULL
# from pywinauto.win32structures import RECT

from pywinauto.keyboard import SendKeys
from tkinter import Button, Tk, Label, Entry
from datetime import datetime
from datetime import timedelta
# import win32api

# import pyautogui
from pyautogui import moveTo, click, press, typewrite
# import pymsgbox
from pymsgbox import confirm
from win32gui import GetWindowText, GetForegroundWindow
# from pywin import dialogs

import sys
from win32api import GetKeyState

#sys.argv = "a C:\\Users\\ssleep\\Documents\\Programming\\Dispatchmate automation\\TRHU3267880_B13.pdf".split()
# sys.argv = "a C:\\Users\\ssleep\\Documents\\Programming\\Dispatchmate automation\\APZU3703747_B13.pdf".split()



ERDList = ['','',False]


# print(workOrderLocation)



def recursiveBookPB(specificPath):
    for filename in listdir(specificPath):
        if filename[-4:] == ".pdf" or filename[-4:] == ".PDF":
            bookPB(specificPath+'\\'+filename)
        elif(path.isdir(specificPath+"\\"+filename)):
            recursiveBookPB(specificPath+"\\"+filename)
    
def bookPB(specificPath):    
#     pdfFileObj = open(specificPath, 'rb')
#     pdfReader = PdfFileReader(pdfFileObj)
#      
#     text=''
#     if pdfReader.isEncrypted:
#         pdfReader.decrypt('')
#     for i in range(pdfReader.getNumPages()):
# #         print(i)
#         pageObj = pdfReader.getPage(i)
#         text += pageObj.extractText()
    
    

    text = ""
 
    fp = open(specificPath, 'rb')
    parser = PDFParser(fp)
    doc = PDFDocument()
    parser.set_document(doc)
    doc.set_parser(parser)
    doc.initialize('')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
#     Process each page contained in the document.
    for page in doc.get_pages():
        interpreter.process_page(page)
        layout = device.get_result()
        for lt_obj in layout:
            if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                text+=lt_obj.get_text()
         
     
    text = text.split('\n')

#     shipper=''
# 
#     terminal = "APM"
#     terminalNum = '306'
#     
#     if text[77].find("SUNRISE METALS") != -1:
#         shipper='SUNRISE METALS'
     
    description = text[114]
     
    traceNumber = text[92][4:]
     
    weight = text[129]
     
#     value = text[126]
     
#     quantity=text[85]
     
    containerNumber=text[199]
    
    exporter = text[83]
    
    vessel = text[84]
    
    pieces = "1" 
    
    containerSize = '20D86'
    
       
    marking = "SUNRISE"
    

    
#   *****************************
#        Overview Tab
#   *****************************
    
    click(45, 305)
    typewrite("249")
    press('tab')
    
    click(299, 305)
    typewrite("306")
    press('tab')
    typewrite(traceNumber)
    
    click(565, 305)
    typewrite("249")
    press('tab')
    
    click(158, 353)
    typewrite(description)
    
    press("tab")
    typewrite(weight)
    
    press("tab", 6)
    typewrite(pieces)
    
    press("tab")
    typewrite("al")
    press('enter')
    
#     press("tab")
    typewrite(containerNumber)

    press("tab", 3)
    typewrite(marking)
    
   
    click(400, 911)
    downAmount = 0
    if containerSize == "20D86":
        downAmount = 5
    elif containerSize == "20R86":
        downAmount = 4
    elif containerSize == "40D86":
        downAmount = 8
    elif containerSize == "40D96":
        downAmount = 11
    elif containerSize == "40R96":
        downAmount = 13
    press('down', downAmount)
    press('enter')
    
    click(413, 930)
    typewrite("exp")
    press("enter")
    
    
    click(652, 957)
    typewrite("usa")
    press("enter")
    
    click(1876, 999)
    typewrite("LONG")
    
    if not ERDList[2]:
        top = Tk()
        L1 = Label(top, text="Please enter ERD:")
        L1.grid(row=0, column=0)
        E1 = Entry(top, bd = 5)
        E1.grid(row=0, column=1)
                
        def callbackERD(ERDList):
            ERDList[0] = E1.get()
            ERDtm = datetime.strptime(ERDList[0], "%m/%d")
            ERDList[1] = (ERDtm + timedelta(days=4)).strftime('%m/%d')    
            top.destroy()  
            
        def callbackERD2(ERDList):
            ERDList[0] = E1.get()
            ERDtm = datetime.strptime(ERDList[0], "%m/%d")
            ERDList[1] = (ERDtm + timedelta(days=5)).strftime('%m/%d')    
            top.destroy()       
                    
        def callbackERDForAll(ERDList):
            ERDList[0] = E1.get()
            ERDtm = datetime.strptime(ERDList[0], "%m/%d")
            ERDList[1] = (ERDtm + timedelta(days=4)).strftime('%m/%d')               
            ERDList[2] = True
            top.destroy()
        
        def callbackCut(ERDList, E2, top2):
            ERDList[1] = E2.get()
            top2.destroy()
            
        def callbackCutForAll(ERDList, E2, top2):
            ERDList[1] = E2.get()
            ERDList[2] = True
            top2.destroy()
        
        def callbackCutWindow(ERDList):
            ERDList[0] = E1.get()
            top.destroy()
            top2 = Tk()
            L2 = Label(top2, text="Please enter CutDate:")
            L2.grid(row=0, column=0)
            E2 = Entry(top2, bd = 5)
            E2.grid(row=0, column=1)
            
            MyButton4 = Button(top2, text="OK", width=10, command=lambda: callbackCut(ERDList, E2, top2))
            MyButton4.grid(row=1, column=1)
            
            MyButton5 = Button(top2, text="Use for all", width=10, command=lambda: callbackCutForAll(ERDList, E2, top2))
            MyButton5.grid(row=2, column=1)
            
            top2.lift()
            top2.attributes('-topmost',True)
            top2.after_idle(top2.attributes,'-topmost',False)
            
            w = 300 # width for the Tk root
            h = 150 # height for the Tk root
            
            # get screen width and height
            ws = top2.winfo_screenwidth() # width of the screen
            hs = top2.winfo_screenheight() # height of the screen
            
            # calculate x and y coordinates for the Tk root window
            x = (ws/2) - (w/2)
            y = (hs/2) - (h/2)
            
            # set the dimensions of the screen 
            # and where it is placed
            top2.geometry('%dx%d+%d+%d' % (w, h, x, y))
            
            
            def click2():
                click(1001, 508)
            
            top2.after(10, click2)
            top2.mainloop()
            
            
            
            
        
        MyButton1 = Button(top, text="Enter cut date", width=10, command=lambda: callbackCutWindow(ERDList))
        MyButton1.grid(row=3, column=1)
        
        MyButton2 = Button(top, text="Cut 4 days later", width=12, command=lambda: callbackERD(ERDList))
        MyButton2.grid(row=1, column=1)
        
        MyButton2 = Button(top, text="Cut 5 days later", width=12, command=lambda: callbackERD2(ERDList))
        MyButton2.grid(row=2, column=1)
        
        MyButton3 = Button(top, text="Use 4 later for all", width=15, command=lambda: callbackERDForAll(ERDList))
        MyButton3.grid(row=4, column=1)
        
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
        
        def click1():
            click(966, 510)
        
        top.after(10, click1)
        top.mainloop()
    
    click(113, 901)
    typewrite(vessel + " ")
    typewrite(exporter)
    press({"enter"}, 2)
    typewrite('ERD: ' + ERDList[0])
    press({"enter"}, 2)
    typewrite('CUT: ' + ERDList[1])
    
    
    
    
    
#   *****************************
#        Routing Tab
#   *****************************
    
    click(688, 96)
     
    click(132, 858)
    typewrite("Company dr th")
    click(354, 858)
    typewrite('29')    

    if GetKeyState(27) < 0:
        sys.exit()
    
    #click book
#     moveTo(1871, 339)
#     click(1871, 339)
     
        
    pbUp = False
    pbDone = False
    
    moveTo(964, 576)
    
    while not pbDone:
        window = GetForegroundWindow()
        if "Dispatch-Mate" in GetWindowText(window) and pbUp:
            pbDone = True
        elif "Available Load" in GetWindowText(window) and not "Dispatch-Mate" in GetWindowText(window):
            pbUp = True
    
    click(1876, 101)
    
    
   
if __name__ == '__main__':
#     argv = r"a C:\Users\ssleep\Documents\Programming\Dispatchmate automation\CMAU1422767_B13.pdf".split()
#     f=open(("testfile21.txt"), "w")
#     f.write("RAN\n\n")
#     f.close()
#     try:
    specificPath = ''
    for i in range(len(argv)):
        if i!=0:
            specificPath+=argv[i]
            if i != len(argv) - 1:
                specificPath+=" "
#     destinationOfFiles = workOrderLocation[:workOrderLocation.rfind('\\')] + "\\"
    if path.isdir(specificPath):
        recursiveBookPB(specificPath)
    elif specificPath[-4:] == ".pdf" or specificPath[-4:] == ".PDF":
        bookPB(specificPath)
#     except:
#         print(exc_info())
#         sleep(50)
            