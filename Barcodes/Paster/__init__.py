from openpyxl import load_workbook
import pythoncom
import pyHook
from win32api import GetKeyState
from win32con import VK_CAPITAL
from pyautogui import click, press, hotkey, typewrite, keyDown, keyUp
from time import sleep
from pywinauto.win32functions import SendInput
import pyperclip
import ctypes
import time
from past.builtins.misc import unichr
import unicodedata
from HelperFunctions import done
SendInput = ctypes.windll.user32.SendInput

PUL = ctypes.POINTER(ctypes.c_ulong)
class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]

class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time",ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                 ("mi", MouseInput),
                 ("hi", HardwareInput)]

class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]
    
KEYEVENTF_UNICODE = 0x0004 
KEYEVENTF_KEYUP = 0x0002

def PressKey(KeyUnicode):

    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, KeyUnicode, KEYEVENTF_UNICODE, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def ReleaseKey(KeyUnicode):

    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, KeyUnicode, KEYEVENTF_UNICODE|KEYEVENTF_KEYUP, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

if __name__ == '__main__':
    print("CREATE THE FIRST BARCODE AS NORMAL AND COPY IT.\n")
    print("FOR SUBSEQUENT BARCODES, PASTE THE FIRST BARCODE")
    print("THEN HOVER THE CURSOR OVER IT AND PRESS \"CAPS LOCK\"")
#     print(chr(0xC3))
#     PressKey(0xAE)
#     ReleaseKey(0xAE)
#     exit()
    
    codeBook = load_workbook(r"C:\Automation\barcode-generator-excel.xlsm", data_only=True)
    
    codeSheet = codeBook.active
    
    for cell in next(codeSheet.rows):
        if cell.value and "Item Number" in cell.value:
            itemCol = cell.col_idx - 1
        if cell.value and "Barcode" in cell.value:
            barcodeCol = cell.col_idx - 1

    pars = []
    
    first = False
    for row in codeSheet.rows:
#         print(str(row[itemCol].value)[0:4])
        if str(row[itemCol].value)[0:4]=="20C0":
            pars.append((row[itemCol].value, row[barcodeCol].value))
    try:
        pars.pop(0)
    except:
        print("NO VALID PARS FOUND")
        sleep(100)
        exit()
    print("\n")
    print("PARS PASTED:")
#     print("\n")
    while len(pars)>0:
        if GetKeyState(27) < 0:
                exit()
                
        while not GetKeyState(20)<0:
            True
        
        if GetKeyState(27) < 0:
            exit()
        
            
        click(clicks=2)
        
        press("up")
        press("end")
        press("backspace", 20)
        
        if GetKeyState(VK_CAPITAL):
            press("capslock")
        
        parsX = pars.pop(0)
        typewrite(str(parsX[0]))
        
        press("right", 2)
        press("delete", 20)
        press("backspace")
        
        
        for i in range(len(str(parsX[1]))):
            s = str(parsX[1])[i]
            unicodedata.name(s)
            PressKey(ord(s))
            ReleaseKey(ord(s))

        print(parsX[0])            
        sleep(1)
    
    done()
    
#     pyinstaller "C:\Users\ssleep\workspace\Barcodes\Paster\__init__.py" --distpath "J:\Spencer\Barcode Paster" -y