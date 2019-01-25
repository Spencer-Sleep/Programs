from pyotrs.lib import Client
from warnings import filterwarnings
import sys
from HelperFunctions import popUpOK
from os.path import isdir

from tkinter import Tk, Button, Label, constants, Checkbutton, BooleanVar, Text,\
    Scrollbar, StringVar, Frame, Radiobutton
from PyPDF2.pdf import PdfFileWriter, PdfFileReader
from time import sleep
import os

def getBOLs():
    bgC = "lavender"
    top = Tk()
    top.config(bg = bgC)
    L1 = Label(top, text="Please enter the WOs to fetch", bg = bgC, padx = 20)
    L1.config(font=("serif", 16))
    L1.grid(row=0, column=0, sticky=constants.W+constants.E)
    
    S1=Scrollbar(top, orient='vertical')
    S1.grid(row=1, column=1, sticky=constants.N + constants.S)
    S2=Scrollbar(top, orient='horizontal')
    S2.grid(row=2, column=0, sticky=constants.E + constants.W)
    
    T1 = Text(top, height = 20, width = 97, xscrollcommand = S2.set, yscrollcommand=S1.set, wrap = constants.NONE)
    T1.grid(row=1, column=0)
    bols=[]
    
    def callbackCont():
        if T1.get("1.0", constants.END).strip()=="":
            popUpOK("Please list the target WOs")
        else:
            bols.append(T1.get("1.0", constants.END).splitlines())
            top.destroy()

    MyButton = Button(top, text="OK", command=callbackCont)
    MyButton.grid(row=5, column=0, sticky=constants.W+constants.E, padx = 20, pady = 10)
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
    top.lift()
    top.attributes('-topmost',True)
    top.after_idle(top.attributes,'-topmost',False)
    
    top.mainloop()
    
    return bols[0]

def getPdf(tickets, WO):
    for ticket in tickets:
        ticket = client.ticket_get_by_id(ticket, True, True)
        for article in ticket.articles:
            for attachment in article.attachments:
                pdf = attachment
                if ".pdf" in attachment.Filename:
                    pdf.Filename = WO+".pdf"
                    pdf.save_to_dir(workOrderLocation)
                    return workOrderLocation+"\\"+pdf.Filename
                
    return False

def skipPopup(bol):
    skipThisContainer = [False]
    top = Tk()
    top.config(bg = "lavender")
    L1 = Label(top, text="WO " + bol + " has been fetched previously.", bg="lavender", font=("serif", 16))
    L1.grid(row=0, column=0, columnspan=2)
    
    def callbackSkip(skipThisContainer):
        skipThisContainer[0] = True
        top.destroy() 
    
    def callbackContinue():
        top.destroy()
         
    
    MyButton4 = Button(top, text="Proceed anyway", width=14, command=lambda: callbackContinue(), bg="green", font=("serif", 16))
    MyButton4.grid(row=2, column=0, padx=10, pady=10)
    
    MyButton5 = Button(top, text="Skip this WO", width=25, command=lambda: callbackSkip(skipThisContainer), bg="red", font=("serif", 16))
    MyButton5.grid(row=2, column=1, padx=10, pady=10)
    
    top.update()
    
    w = top.winfo_width() # width for the Tk root
    h = top.winfo_height() # height for the Tk root
       
    ws = top.winfo_screenwidth() # width of the screen
    hs = top.winfo_screenheight() # height of the screen
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    
    top.geometry('%dx%d+%d+%d' % (w, h, x, y))
    
    top.update()
    
    top.lift()
    top.attributes('-topmost',True)
    top.after_idle(top.attributes,'-topmost',False)
    
    top.mainloop()
    return skipThisContainer[0]



def fetchANs(bols, client,workOrderLocation):
#     pdfs=[]
    with open('J:\Spencer\Fetched Hapag.txt', 'r') as myfile:
        data=myfile.read().replace('\n', '')
    
    doneBols=[]
    failedPDFs=""
    for bol in bols:
        if bol !="" and not bol in doneBols:
            if bol in data and skipPopup(bol):
                continue
            doneBols.append(bol)
            tickets=client.ticket_search(Title="WOSD0001 - "+bol, From="@[Hh][Ll][Aa][Gg]")
            getPdf(tickets, bol)
            with open('J:\Spencer\Fetched Hapag.txt', 'a+') as myfile:
                myfile.write(bol+"\n")
#             print(pdf)
#             if pdf:
#                 pdfs.append(pdf)
#                 with open(workOrderLocation+"\\"+"DOs.pdf", 'wb') as f:
#                     input_streams = []
#                     try:
#                         # First open all the files, then produce the output file, and
#                         # finally close the input files. This is necessary because
#                         # the data isn't read from the input files until the write
#                         # operation. Thanks to
#                         # https://stackoverflow.com/questions/6773631/problem-with-closing-python-pypdf-writing-getting-a-valueerror-i-o-operation/6773733#6773733
#                         i=0
#                         for input_file in pdfs:
#                             f1=open(input_file, 'r+b')
#                             input_streams.append(f1)
#                             i+=1
#                         writer = PdfFileWriter()
#                         for reader in map(PdfFileReader, input_streams):
#                             for n in range(reader.getNumPages()):
#                                 writer.addPage(reader.getPage(n))
#                         writer.write(f)
#                     finally:
#                         for f in input_streams:
#                             f.close()
#             else:
#                 failedPDFs = "\n"+bol+failedPDFs
        
#     for pdf in pdfs:
#         os.remove(pdf)
        
    if failedPDFs!="":
        popUpOK("Could not find the following WOs: " + failedPDFs)
    
if __name__ == '__main__':
#     sys.argv=r"a C:\Users\ssleep\Documents\Maersk fetcher".split()
#     sys.argv=r"a J:\All motor routings\2019\Week 3\HAPAG".split()
    workOrderLocation = ''
    for i in range(len(sys.argv)):
        if i!=0:
            workOrderLocation+=sys.argv[i]
            if i != len(sys.argv) - 1:
                workOrderLocation+=" "
    
    filterwarnings("ignore")
    
    client = Client("https://core.seaportint.com/", "testadmin", "testpass")
    a = client.session_create()
    if(a):
        print("Connected to OTRS as Testadmin")
    
    if isdir(workOrderLocation):
        bols = getBOLs()
        for i in range(len(bols)):
            bols[i]=bols[i].replace("WOSD0001 - ","")
        fetchANs(bols, client, workOrderLocation)
    else:
        popUpOK("Please send a folder to the program")
    
    #pyinstaller "C:\Users\ssleep\workspace\Pull Hapag WO\Automator\__init__.py" --distpath "J:\Spencer\Pull Hapag WO" --noconsole -y