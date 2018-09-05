from openpyxl import load_workbook

from sys import argv, exit

from pyautogui import moveTo, click, press, typewrite, hotkey

from win32api import GetKeyState

from tkinter import Button, Tk, Label, Entry

from time import sleep

from sys import exc_info

import ContainerSizeInfo

import DispatchmateLocations as loc
from DispatchmateLocations import CONSIGNEECODELOC, SHIPPERLOC, SHIPPERCODELOC,\
    RATINGTABLOC, CUSTOMERCHARGELOC, OKLOC, OVERVIEWTABLOC, ROUTINGTABLOC,\
    HAZMATLOC, HANDLINGLOC, DESCRIPTIONLOC, CUSTOMERCODELOC, DRIVERPAYOUTLOC
from HelperFunctions import done
import HelperFunctions

CONTAINERNUMBER = "Container"
BOOKING = "TPDoc"
VESSEL = "Vessel"
# WEIGHLBS = "Weight (lbs)"
WEIGHT = "Weight"
PIECES = "Piece Count"
SIZE = "Size-Type"
PARS = "Border"
HAZ = "Haz"


class Container(object):
    def __init__(self):
        self.properties = {CONTAINERNUMBER: "",
               BOOKING: "",
               VESSEL: "",
#                WEIGHLBS: "",    
               WEIGHT: "",
               PIECES: "",
               SIZE: "",
               PARS: "",
               HAZ:""}    
    


def loadInfo(workbookPath):
    
    typeOfManifest = [-1]
    skip = [0]
    startAt = [""]
    top = Tk()
    L1 = Label(top, text="Please select which container to start at or \n how many containers to skip as well as\n whether to do PARS, A8As, or both \n ")
    L1.config(font=("Courier", 16))
    L1.grid(row=0, column=0, columnspan=5)
    L2 = Label(top, text="Start at container #:")
    L2.config(font=("Courier", 10))
    L2.grid(row=1, column=0, columnspan = 2)
    E1 = Entry(top, bd = 5, width = 39)
    E1.grid(row=1, column=2, columnspan = 2)
    L3 = Label(top, text="OR")
    L3.grid(row=2, column=1, columnspan=2)
    L3.config(font=("Courier", 20))
    L4 = Label(top, text="# of containers to skip:")
    L4.grid(row = 3, column = 0, columnspan = 2)
    L4.config(font=("Courier", 10))
    E2 = Entry(top, bd = 5, width = 39)
    E2.grid(row=3, column=2, columnspan=2)
    E2.insert(0, "0")
    top.lift()
    top.attributes('-topmost',True)
    top.after_idle(top.attributes,'-topmost',False)
      
    def callbackType(typeOf):
        startAt[0]=E1.get().strip()
        skip[0]=E2.get().strip()
        typeOfManifest[0] = typeOf
        top.destroy()
      
      
    
    MyButton5 = Button(top, text="PARS", command=lambda: callbackType(0), width=10)
    MyButton5.grid(row=4, column=0)
    MyButton5.config(font=("Courier", 16))
    MyButton6 = Button(top, text="A8A", command=lambda: callbackType(1), width=10)
    MyButton6.grid(row=4, column=1)
    MyButton6.config(font=("Courier", 16))
    MyButton7 = Button(top, text="BOTH", command=lambda: callbackType(2), width=10)
    MyButton7.grid(row=4, column=2)
    MyButton7.config(font=("Courier", 16))
      
      
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
    
    routing = load_workbook(workbookPath)
    routing = routing.get_active_sheet()
    
    pars = []
    a8a = []
    
    colDict = {CONTAINERNUMBER: "",
               BOOKING: "",
               VESSEL: "",
#                WEIGHLBS: "",
               WEIGHT: "",
               PIECES: "",
               SIZE: "",
               PARS: "",
               HAZ: ""}              
    
    for cell in next(routing.rows):
        for contProperty in colDict:
            if contProperty in cell.value:
                colDict[contProperty] = cell.col_idx - 1
    
    i=0
    found = startAt[0]==""
    containerList = [""]
    for row in routing.rows:
        if i>int(skip[0]):
            container = Container()
            for contProperty in container.properties:
#                 if contProperty != WEIGHLBS:    
                container.properties[contProperty] = str(row[colDict[contProperty]].value)
            if found or container.properties[CONTAINERNUMBER]==startAt[0]:
                if not container.properties[CONTAINERNUMBER] in containerList: 
                    found = True
                    if container.properties[PARS] == 'PARS':
                        pars.append(container)
                    elif container.properties[PARS] == 'A8A':
                        a8a.append(container)
                    containerList.append(container.properties[CONTAINERNUMBER])
                
#             print(container.properties[CONTAINERNUMBER])
        i+=1
      
    return (pars, a8a, typeOfManifest)      

def bookPars(containers):
    click(loc.CONSIGNEECODELOC)
    typewrite("303")
    press('tab')
     
    click(loc.SHIPPERCODELOC)
    typewrite("311")
    press('tab')
     
    click(loc.CUSTOMERCODELOC)
    typewrite("1779")
    press('tab')

    click(loc.DESCRIPTIONLOC)
    press("tab", 8)
    typewrite("r")
    
    press("tab", 4)
    typewrite("MAERSK")
    
    click(loc.DIRECTIONLOC)
    typewrite("imp")
    press("enter")
    
    click(loc.DIVISIONLOC)
    typewrite("toronto li")
    press("enter")
    
    click(loc.LANECODELOC)
    typewrite("30")
    press("enter")
    
    click(loc.HOUSELOC)
    hotkey('ctrl', 'a')
    typewrite("REGIONAL")
    
    click(loc.CONSIGNEETRACELOC)
    typewrite(containers[0].properties[VESSEL])
    
#     click(1153, 98)
#     click(89, 144)
#     typewrite("on")
#     press("tab")
#     
#     click(248, 101)
    lastHaz = False
    for container in containers:
        click(loc.OKLOC)
        click(loc.DUPLICATELOC)
        
        if GetKeyState(27) < 0:
            exit()
        
        if lastHaz:
            click(DESCRIPTIONLOC)
            hotkey('ctrl', 'a')
            press("del")
            press("tab", 25)
            press("del", 5)
            click(HAZMATLOC)
            click(HANDLINGLOC)
            hotkey('ctrl', 'a')
            press("del")
            click(RATINGTABLOC)
            click(CUSTOMERCHARGELOC)
            hotkey('ctrl', 'a')
            press("del")
            press("tab")
            press("del")
            click(OVERVIEWTABLOC)
            lastHaz = False
        
        click(loc.BOLLOC)
        typewrite(container.properties[BOOKING])
        
        
        click(loc.DESCRIPTIONLOC)
        press("tab", 7)
        typewrite("1")
        
        press("tab", 2)
        typewrite(container.properties[CONTAINERNUMBER])
        
        
        click(loc.EQUIPMENTLOC)
        press('home')
        size = container.properties[SIZE]
        if size == "20 DRY 8 6":
            downAmount = 4
        elif size == "20 REEF 8 6":
            downAmount = 3
        elif size == "40 DRY 8 6":
            downAmount = 7
        elif size == "40 DRY 9 6":
            downAmount = 10
        elif size == "40 REEF 9 6":
            downAmount = 12
        elif size == "45 DRY 9 6":
            downAmount = 13
        

        press('down', downAmount)
        press('enter')
        
        if container.properties[HAZ]=="Y":
            lastHaz=True
            click(RATINGTABLOC)
            click(CUSTOMERCHARGELOC)
            typewrite("HAZ")
            press("tab")
            click(ROUTINGTABLOC)
            click(DRIVERPAYOUTLOC)
            typewrite("HAZ")
            press("tab")
            click(OVERVIEWTABLOC)
            click(HAZMATLOC)
            click(HANDLINGLOC)
            typewrite("HAZ")
            click(DESCRIPTIONLOC)
            press("tab", 25)
            HelperFunctions.popUpOK("THIS IS A HAZMAT LOAD\n PLEASE ENTER UN CODE\n THEN HIT \"OK\"", 16)
        
    click(loc.OKLOC)
    
    return lastHaz
        
def bookA8A(containers, lastHaz):
#     click(45, 305)
#     typewrite("311")
#     press('tab')
#      
#     click(299, 305)
#     typewrite("303")
#     press('tab')
    if typeOfManifest[0]==1:
        click(loc.CUSTOMERCODELOC)
        typewrite("1779")
        press('tab')

#     click(18, 351)
#     press("tab", 8)
#     typewrite("r")
    
#     press("tab", 4)
#     typewrite("MAERSK")
    
        click(loc.DIRECTIONLOC)
        typewrite("imp")
        press("enter")
        
        click(loc.DIVISIONLOC)
        typewrite("toronto li")
        press("enter")
        
        click(loc.LANECODELOC)
        typewrite("30")
        press("enter")
        
        click(loc.HOUSELOC)
        hotkey('ctrl', 'a')
        typewrite("REGIONAL")
        
        click(loc.CONSIGNEETRACELOC)
        typewrite(containers[0].properties[VESSEL])
        
        
    
#     click(1153, 98)
#     click(89, 144)
#     typewrite("on")
#     press("tab")
#     
#     click(248, 101)

    click(loc.DOCUMENTATIONANDCUSTOMSTABLOC)
    
    click(loc.LOCATIONOFGOODSLOC)
    typewrite("4")
    press("tab")
    typewrite("0")
    press("tab")
    typewrite("0")
    press("enter")
    
    click(loc.CROSSINGLOC)
    hotkey('ctrl', 'a')
    typewrite("0901")
    press("enter")
    
    click(loc.CONTROLLOC)
    typewrite("20C0")
    
    click(loc.OVERVIEWTABLOC)
    
    for container in containers:
        click(loc.OKLOC)
        click(loc.DUPLICATELOC)
        
        if lastHaz:
            click(DESCRIPTIONLOC)
            press("tab", 25)
            press("del", 5)
            click(HAZMATLOC)
            click(HANDLINGLOC)
            hotkey('ctrl', 'a')
            press("del")
            click(RATINGTABLOC)
            click(CUSTOMERCHARGELOC)
            hotkey('ctrl', 'a')
            press("del")
            press("tab")
            press("del")
            click(OVERVIEWTABLOC)
            lastHaz = False
        
        while not GetKeyState(45)<0:
#             if not GetKeyState(71)<0:
            True
        
        if GetKeyState(27) < 0:
            exit()
#         click(60, 134)
#         press("tab")
        
        click(loc.DESCRIPTIONLOC)
        press("tab", 1)
        typewrite(container.properties[WEIGHT])
        
        press("tab", 6)
        
#         click(550, 350)
        typewrite(container.properties[PIECES])
        
#         click(18, 351)
        press("tab", 1)
        typewrite("r")
        
        
        press("tab", 1)
        typewrite(container.properties[CONTAINERNUMBER])
        
        sizeCode = ''
        
        size = container.properties[SIZE]
        if size == "20 DRY 8 6":
            sizeCode = "20DC"
        elif size == "20 REEF 8 6":
            sizeCode = "20RF"
        elif size == "40 DRY 8 6":
            sizeCode = "40DC"
        elif size == "40 DRY 9 6":
            sizeCode = "40HC"
        elif size == "40 REEF 9 6":
            sizeCode = "40RH"
        elif size == "45 DRY 9 6":
            sizeCode = "45DC"
            
            
        press("tab", 3)
        typewrite("MAERSK" + sizeCode)
        
        click(loc.BOLLOC)
        typewrite(container.properties[BOOKING])
        
#         click(472, 304)
#         typewrite(container.properties[VESSEL])
        
#         click(18, 351)
#         press("tab", 7)
#         typewrite("1")
        
        
        
        
        click(loc.EQUIPMENTLOC)
        press('home')
        
        if size == "20 DRY 8 6":
            downAmount = 4
        elif size == "20 REEF 8 6":
            downAmount = 3
        elif size == "40 DRY 8 6":
            downAmount = 7
        elif size == "40 DRY 9 6":
            downAmount = 10
        elif size == "40 REEF 9 6":
            downAmount = 12
        elif size == "45 DRY 9 6":
            downAmount = 13
        

        press('down', downAmount)
        press('enter')
        
        if container.properties[HAZ]=="Y":
            lastHaz=True
            click(RATINGTABLOC)
            click(CUSTOMERCHARGELOC)
            typewrite("HAZ")
            press("tab")
            click(OVERVIEWTABLOC)
            click(HAZMATLOC)
            click(HANDLINGLOC)
            typewrite("HAZ")
            click(DESCRIPTIONLOC)
            press("tab", 25)
            HelperFunctions.popUpOK("THIS IS A HAZMAT LOAD\n PLEASE ENTER UN CODE\n THEN HIT \"OK\"", 16)
        
    click(loc.OKLOC)
        
        
if __name__ == '__main__':
#     workbookPath = r"C:\Users\ssleep\Documents\Programming\Old stuff\MSC GINA N003-MAERSK-IN PROGRESS.xlsx"
#     argv = r"a J:\Running Routing by Vessel\Transferred to RR\PRIMAVERA N009.xlsx".split()
    
#     try:


    workbookPath = ""
    for i in range(len(argv)):
        if i!=0:
            workbookPath+=argv[i]
            if i != len(argv) - 1:
                workbookPath+=" "
#         
    pars, A8A, typeOfManifest = loadInfo(workbookPath)
    lastHaz = False
    
    if typeOfManifest[0] == 0 or typeOfManifest[0] == 2:
        lastHaz = bookPars(pars)
    if typeOfManifest[0] > 0:
        bookA8A(A8A, lastHaz)
         
    done()


#     except:
#         print(exc_info())
#         sleep(50)
#         while True:
#             while not GetKeyState(45)<0:
# #             if not GetKeyState(71)<0:
#                 True

# pyinstaller "C:\Users\ssleep\workspace\Maersk DM\DMAutomater\__init__.py" --distpath "J:\Spencer\Maersk Dispatchmate" --noconsole -y