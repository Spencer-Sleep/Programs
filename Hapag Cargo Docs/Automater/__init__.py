from selenium import webdriver
from selenium.webdriver import firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from os import devnull

from pdfminer.pdfparser import PDFParser, PDFDocument #@UnresolvedImport
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter #@UnresolvedImport
from pdfminer.converter import PDFPageAggregator #@UnresolvedImport
from pdfminer.layout import LAParams, LTTextBox, LTTextLine #@UnresolvedImport

from PyPDF2 import PdfFileReader

from os import listdir
from sys import argv
from sys import path


pages = []

def setupPortal():
    fp = FirefoxProfile();
    fp.set_preference("webdriver.load.strategy", "unstable");
    
    driver = webdriver.Firefox(firefox_profile=fp, log_path=devnull)
    driver.get("http://www.cbsa-asfc.gc.ca/prog/manif/portal-portail/menu-eng.html")
    
    

def extract(objid, obj):
#     global pages
    pars = []
    if isinstance(obj, dict):
        # 'Type' is PDFObjRef type
#         if 'Type' in obj and obj['Type'].name == 'Page':
#             pages.append(objid)
        if not 'Type' in obj and obj['Type'].name == 'Page' and 'C' in obj:
            if 'Contents' in obj:
                pars.append(obj['Contents'])
#             for x in list(obj.keys()):
#                 print(x)
    return pars

def recursiveCargoDoc():
    for filename in listdir(specificPath):
        if "PARS MANIFESTS" in filename and filename[-4:] == ".pdf" or filename[-4:] == ".PDF":
            cargoDoc(specificPath+'\\'+filename)
        elif(path.isdir(specificPath+"\\"+filename) and not filename=="Flattened"):
            recursiveCargoDoc(specificPath+"\\"+filename)

def cargoDoc():
    fp = open(r"C:\Users\ssleep\Documents\Programming\Cargo Docker\Thursday\LCBO\601331975 PARS MANIFESTS.pdf", 'rb')
    parser = PDFParser(fp)
    doc = PDFDocument()
    parser.set_document(doc)
    doc.set_parser(parser)
    doc.initialize('')
    visited = set()
    
    pars = []
    
    for xref in doc.xrefs:
        for objid in xref.get_objids():
            if objid in visited: continue
            visited.add(objid)
            obj = doc.getobj(objid)
            if obj is None: continue
            pars = extract(objid,obj)

    pdfFileObj = open(specificPath, 'rb')
    pdfReader = PdfFileReader(pdfFileObj)
    
    fields = pdfReader.getFields()
#     print(len(fields)-15)


    for i in range(len(fields)-15):
        containerNumber = ""
        weight = ""
        consignee = ""
        shipper = ""
        eta = ""
        portOfLoading = ""
        portOfDischarge = ""
        description = ""
        if i == 0:
#             prefix = str(i) + "."
            containerNumber = fields["Container Row1"].value
            weight = float(fields["Weight KGRow1"].value)
            consignee = fields["Consignee"].value
            shipper = fields["Shipper"].value
            eta = fields["ETA DATE"].value
            portOfLoading = fields["undefined"].value
            portOfDischarge = fields["Port of Discharge"].value
            description = fields["Description of goods"].value
        else:
            for j in list(fields.keys()):
                if j==str(i):
                    for k in list(fields[j]["/Kids"]):
                        try:
                            if(k.getObject()['/T']=="WO"):
                                wo=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Container Row1"):
                                containerNumber=k.getObject()['/V']
                            elif(k.getObject()['/T']=="SizeRow1"):
                                size=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Weight KGRow1"):
                                weight=float(k.getObject()['/V'])
                            elif(k.getObject()['/T']=="Consignee"):
                                consignee=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Shipper"):
                                shipper=k.getObject()['/V']
                            elif(k.getObject()['/T']=="ETA DATE"):
                                eta=k.getObject()['/V']
                            elif(k.getObject()['/T']=="undefined"):
                                portOfLoading=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Port of Discharge"):
                                portOfDischarge=k.getObject()['/V']
                            elif(k.getObject()['/T']=="Description of goods"):
                                description=k.getObject()['/V']    
                        except KeyError:
                            True
        

if __name__ == '__main__':
    
    specificPath = ''
    for i in range(len(argv)):
        if i!=0:
            specificPath+=argv[i]
            if i != len(argv) - 1:
                specificPath+=" "
#     destinationOfFiles = workOrderLocation[:workOrderLocation.rfind('\\')] + "\\"
#     print(specificPath)
    if path.isdir(specificPath):
        recursiveCargoDoc(specificPath)
    elif "PARS MANIFESTS" in specificPath and (specificPath[-4:] == ".pdf" or specificPath[-4:] == ".PDF"):
        cargoDoc(specificPath)
