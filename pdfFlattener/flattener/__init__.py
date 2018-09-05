
import os
import sys
import time
import win32api
import win32print
import PyPDF2
import re
import shutil
import ctypes

#     If Document did not print to PDF and it printed to another printer, open Adobe and print something with "Adobe PDF" as the printer. 
#     There is a bug that sometimes happens where Adobe ignores the default printer. If that doesn't work, contact Spencer for help.
#     If you are Spencer, good luck.

def recursivePrint(path):
    try:
        for filename in os.listdir(path):
            if filename[-4:] == ".pdf" or filename[-4:] == ".PDF":
                if (not PyPDF2.PdfFileReader(open(path+"\\"+filename, 'rb')).isEncrypted and not re.compile("(A|a)(R|r)(R|r)(I|i)(V|v)(A|a)(L|l) (N|n)(O|o)(T|t)(I|i)(C|c)(E|e)").search(filename)):
                    if not os.path.exists(path + "\\Flattened"):
                        os.makedirs(path + "\\Flattened")
                    if (not os.path.isfile(path+"\\Flattened\\"+filename)):
                        win32api.ShellExecute(0, "print", path+"\\"+filename, ".", ".", 0)
                        toPrint = True
                        while(not os.path.exists("C:\\Printed PDFs\\"+filename)):
                            time.sleep(0.1)
                            if toPrint:
                                print("Printing; " + filename)
                                toPrint=False
                        shutil.move("C:\\Printed PDFs\\"+filename, path+"\\Flattened\\"+filename)
            elif(os.path.isdir(path+"\\"+filename)):
                if (not filename == "Flattened"):
                    recursivePrint(path+"\\"+filename)
    except:
#     for i in range(1000):
#         print("Unexpected error:" + sys.exc_info()[0])
        print(sys.exc_info())
        time.sleep(5)



if __name__ == '__main__':
    path = ""

    sendPath = ""
        
    for i in range(len(sys.argv)):
        if i!=0:
            sendPath+=sys.argv[i]
            if i != len(sys.argv) - 1:
                sendPath+=" "
     
    if sendPath != "":
        path=sendPath
    else:
        path = os.getcwd()
     
     
    currentprinter = win32print.GetDefaultPrinter()
    win32print.SetDefaultPrinter("Adobe PDF")
    print ("Running")
    print ("If the PDF printed to another printer, close this window, open Adobe and print something with \"Adobe PDF\" as the printer, then try again.")
    print ("There is a bug that sometimes happens where Adobe ignores the default printer.\n")
    recursivePrint(path)
    win32print.SetDefaultPrinter(currentprinter)
    
