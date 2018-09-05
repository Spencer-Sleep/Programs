import PyPDF2
from fdfgen import forge_fdf
import re
import subprocess
import sys

workOrderLocation = sys.argv[1]
templateLocation = " C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\templates\\HL_TEMPLATETIER"
f=open(("testfile21.txt"), "w")
f.write("RAN\\n\\n")
f.write(workOrderLocation)
f.close()

class Container(object):
    
    eta = ''
    vessel=''
    voyage=''
    workOrder=''
    portOfLoading=""
    portOfDischarge=""
    description=""
    containerCode=""
    quantity=""
    packageType=""
    size=""
    weight=""
    otherInfo=""
    shipper=""
    consignee=""
    reefer=""
    overweight=""
    overweightTier=0
    
    
#def __init__()
    
    


    
if __name__ == '__main__':
    pdfFileObj = open(workOrderLocation, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    
    text = ''
    discharge = ''
    if pdfReader.isEncrypted:
        pdfReader.decrypt('')
    for i in range(pdfReader.getNumPages()):
        pageObj = pdfReader.getPage(i)
        text += pageObj.extractText()
    routingInfo = text.find('R o u t i n g - I n f o r m a t i o n 1a2b3n4m5k Pick up 1a2b3n4m5k ')
    #TODO IF NONE OF THESE FLAG IT AND DISPLAY AT END
    if 'GCT NEW YORK' in text[routingInfo+44:routingInfo+84]:
        discharge = 'NYCT'
    if 'GCT BAYONNE' in text[routingInfo+44:routingInfo+84]:
        discharge = 'GLOBAL'
    if 'GREENWICH TERMINALS LLC' in text[routingInfo+44:routingInfo+84]:
        discharge = 'PACKER'
    
    workOrder = text.find("Work Order:")
    m=re.compile('\D').search(text[workOrder+23:])
    y=m.start(0) + workOrder + 23
    workOrder=text[workOrder+23:y]
    
    #f=open(("testfile" + str(i) + ".txt"), "w")
    #f.write(containerText[i])
    #f.close()
    
    containerText = re.compile("CONTAINER 1a2b3n4m5k \\d+ 1a2b3n4m5k ").split(text)
    
    #for i in range(len(containerText)):
#        f=open(("testfile" + str(i) + ".txt"), "w")
        #f.write(containerText[i])
        #f.close()
    
    
    
    containerList = []
    i=1
    for containerTextX in containerText:
        
        if containerTextX[:10] != "Container:":
            continue
        
        containerX = Container()
        
        pageBreakLocation = containerTextX.find("HAPAG-LLOYD (CANADA) INC 1a2b3n4m5k 3400 BLVD. DE MAISONNEUVE WEST")
        
        if pageBreakLocation != -1:
            containerTextX = containerTextX[:pageBreakLocation] + containerTextX[pageBreakLocation+597:]
        
        x = containerTextX.find("Cargo Gross 1a2b3n4m5k DG 1a2b3n4m5k ")+37
        #m=re.compile("[a-zA-Z]").search(containerTextX[x:])
        #y=m.start(0) + x
        y=containerTextX[x:].find(" ") + x
        containerX.quantity=containerTextX[x:y]
        
        x=y+12
        y=containerTextX[x:].find(" ")+x
        containerX.packageType=containerTextX[x:y]
        
        descX = y
        
        m=re.compile('\\d+.\\d* 1a2b3n4m5k KGM 1a2b3n4m5k N').search(containerTextX)
        x=m.start(0)
        y=containerTextX.find(" 1a2b3n4m5k KGM 1a2b3n4m5k N")
        containerX.weight=containerTextX[x:y]
        
        descY = x
        
        descriptionMessy = containerTextX[descX:descY]
        containerX.description=descriptionMessy.replace(" 1a2b3n4m5k ", "")
        
        containerTextX=containerTextX.replace(" 1a2b3n4m5k ", "")
        
        #pageBreakLocation = containerTextX.find("HAPAG-LLOYD (CANADA) INC3400 BLVD. DE MAISONNEUVE WEST")
        
        #if pageBreakLocation != -1:
        #    containerTextX = containerTextX[:pageBreakLocation] + containerTextX[pageBreakLocation+429:]
        
        f=open(("testfile" + str(i) + ".txt"), "w+")
        f.write(containerTextX)
        f.close()
        
        
        containerX.containerCode=containerTextX[10:14]+containerTextX[16:23]
        w = float(containerX.weight)
        if containerTextX[27:36] == "20'X8'6\"G":
            containerX.size = "20D86"
            if w>19050:
                containerX.overweight='OW'
            if w<19640:
                containerX.overweightTier=0
            elif w<20774:
                containerX.overweightTier=1
            elif w<21908:
                containerX.overweightTier=2
            elif w<24176:
                containerX.overweightTier=3
            else:
                containerX.overweightTier=4
        elif containerTextX[27:36] == "20'X8'6\"R":    
            containerX.size = "20R86"
            containerX.reefer = "R"
            if w>18143:
                containerX.overweight='OW'
            if w<18778:
                containerX.overweightTier=0
            elif w<19912:
                containerX.overweightTier=1
            elif w<21046:
                containerX.overweightTier=2
            elif w<23314:
                containerX.overweightTier=3
            else:
                containerX.overweightTier=4
        elif containerTextX[27:36] == "40'X8'6\"G":    
            containerX.size = "40D86"
            if w>19958:
                containerX.overweight='OW'
            if w<18189:
                containerX.overweightTier=0
            elif w<19323:
                containerX.overweightTier=1
            elif w<20457:
                containerX.overweightTier=2
            elif w<22725:
                containerX.overweightTier=3
            else:
                containerX.overweightTier=4
        elif containerTextX[27:36] == "40'X9'6\"H":    
            containerX.size = "40D96"
            if w>19958:
                containerX.overweight='OW'
            if w<17962:
                containerX.overweightTier=0
            elif w<19096:
                containerX.overweightTier=1
            elif w<20230:
                containerX.overweightTier=2
            elif w<22498:
                containerX.overweightTier=3
            else:
                containerX.overweightTier=4
        elif containerTextX[27:36] == "40'X9'6\"R":    
            containerX.size = "40R96"
            containerX.reefer="R"
            if w>19958:
                containerX.overweight='OW'
            if w<16510:
                containerX.overweightTier=0
            elif w<17644:
                containerX.overweightTier=1
            elif w<18778:
                containerX.overweightTier=2
            elif w<21046:
                containerX.overweightTier=3
            else:
                containerX.overweightTier=4
        
        x=containerTextX.find("Voyage:") + 6
        m=re.compile("[a-zA-Z]").search(containerTextX[x:])
        y=m.start(0) + x
        z=containerTextX[y:].find('.') + y
        
        containerX.vessel = containerTextX[y:z]
        
        x=containerTextX.find("From:") + 5
        y=containerTextX.find("Arrival Date:")

        containerX.portOfLoading = containerTextX[x:y]
        
        x=y + 13
        y = containerTextX.find("Sched. Voy:")
        
        containerX.eta = containerTextX[x:y]
        
        x=y +11
        y = containerTextX.find("FOB Forw.:")
        
        containerX.voyage = containerTextX[x:y]
        
        x=containerTextX.find("Shipper:")+8
        y=containerTextX.find("Consignee:")
        
        containerX.shipper=containerTextX[x:y]
        
        x=y+10
        y=containerTextX.find("Packages")
        
        containerX.consignee=containerTextX[x:y]
        
        
        #TODO: MAYBE ADD IN DETECTION OF COMMON PACKAGE TYPES AND CARGOS
        
        
        
        otherInfoMessy=containerTextX[containerTextX.find("Seal - No:"):containerTextX.find("Shipper:")]
#         print(containerTextX + "\n\n\n\n\n\n")
#         print(len(containerTextX))
#         print(str(containerTextX.find("Seal - No:")) + "     " + str(containerTextX.find("Shipper:")))
        #TODO: ADD SPACES IN THE OTHER INFO TO IMPROVE READABILITY
        containerX.otherInfo=otherInfoMessy
        
        #datas = {
    #        'ETA DATE': str(containerX.eta),
#            'Vessel': str(containerX.vessel)
        #}
        
        
        #generated_pdf = pypdftk.fill_form("C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\Gen\\HL TEMPLATE.pdf", datas, str('C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\Gen\\HL TEMPLATE' + str(i) + '.pdf'))
        
        
        fields =     [('ETA DATE', containerX.eta),
                    ('Vessel', containerX.vessel),
                    ('Voy', containerX.voyage),
                    ('WO', workOrder),
                    ('undefined', containerX.portOfLoading),
                    ('Port of discharge', discharge),
                    ('Description of goods', containerX.description),
                    ('Container Row1', containerX.containerCode),
                    ('QtyRow1', containerX.quantity),
                    ('Pkg typeRow1', containerX.packageType),
                    ('SizeRow1', containerX.size),
                    ('Weight KGRow1', containerX.weight),
                    ('Other info', containerX.otherInfo),
                    ('Shipper', containerX.shipper),
                    ('Consignee', containerX.consignee)]
        fdf = forge_fdf("",fields,[],[],[])
        fdf_file = open("C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\Temp\\" + str(containerX.containerCode) ,"wb")
        fdf_file.write(fdf)
        fdf_file.close()
        subprocess.Popen("pdftk" + templateLocation + str(containerX.overweightTier) + containerX.overweight + containerX.reefer + ".pdf fill_form C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\Temp\\" + str(containerX.containerCode) + " output C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\Gen\\" + str(containerX.containerCode) + ".pdf").wait()
        #check_output("pdftk C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\HL TEMPLATE.pdf fill_form C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\Temp" + str(containerX.containerCode) + " C:\\Users\\ssleep\\Documents\\Programming\\Hapag-Lloyd_Automation\\Gen\\" + str(containerX.containerCode), True)
        i+=1
        
        
        #f=open(("testfileX" + str(i) + ".txt"), "w")
        #f.write(containerX.vessel)
        #f.write(str(x) + "    " + str(y) + "    " + str(z))
        #f.close()
        #i+=1
        
    #f=open("testfile2.txt", "w+")
    #f.write(discharge)
    #f.close()