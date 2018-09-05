from openpyxl import load_workbook

class ContainerSizeInfo:
    T1payout = ""
    T2payout = ""
    T3payout = ""
    OWpayout = ""
    THRUpayout = ""
    CompTHRUpayout = ""
    
    
    def __init__(self, code = "", T1weight = "", T2weight = "", T3weight = "", T4weight = "", P1weight = "", P2weight = ""):
        self.code = ""
        self.T1weight = ""
        self.T2weight = ""
        self.T3weight = ""
        self.T4weight = ""
        self.P1weight = ""
        self.P2weight = ""
        

def loadValues(AUTOMATIONVALUESPATH):
    workbook = load_workbook(AUTOMATIONVALUESPATH)
    payoutSheet = workbook.get_active_sheet("Driver Overweight Payouts")
    rateSheet = workbook.get_sheet_by_name("Customer Overweight Rates")
    
    TWENTYSTANDARD = "20D86"
    TWENTYREEFER = "20R86"
    FOURTYSTANDARD = "40D86"
    FOURTYHIGHCUBE = "40D96"
    FOURTYREEFER = "40R96"
    FOURTYFIVE = "45D96"
    
    TierType = "NJ Overweights"
    TierAmount = "Amount"
    
    colDict = {TWENTYSTANDARD: "",
               TWENTYREEFER: "",
               FOURTYSTANDARD: "",
               FOURTYHIGHCUBE: "",
               FOURTYREEFER: "",
               FOURTYFIVE: "",
               TierType: "",
               TierAmount: ""
               }    
    
    contRates = {TWENTYSTANDARD: "",
               TWENTYREEFER: "",
               FOURTYSTANDARD: "",
               FOURTYHIGHCUBE: "",
               FOURTYREEFER: "",
               FOURTYFIVE: ""} 
    
    
    for cell in next(payoutSheet.rows):
        for contSize in colDict:
            if contSize in cell.value:
                colDict[contSize] = cell.col_idx - 1
    
    for row in payoutSheet.rows:
        if "NJ HAPAG-LLOYD: $250" in row[0].value:
            for contSize in contRates:
                contRates[contSize] = ContainerSizeInfo(contSize, P1weight=row[colDict[contSize]].value)
                
    for cell in next(rateSheet.rows):
        for contSize in colDict:
            if contSize in cell.value:
                colDict[contSize] = cell.col_idx - 1            
    
    for row in payoutSheet.rows:
        if "CARGO TIER 1" in row[0].value:
            for contSize in contRates:
                contRates[contSize].T1weight=row[colDict[contSize]].value
                
        if "CARGO TIER 2" in row[0].value:
            for contSize in contRates:
                contRates[contSize].T2weight=row[colDict[contSize]].value
                
        if "CARGO TIER 3" in row[0].value:
            for contSize in contRates:
                contRates[contSize].T3weight=row[colDict[contSize]].value
                
        if "CARGO TIER 4" in row[0].value:
            for contSize in contRates:
                contRates[contSize].T4weight=row[colDict[contSize]].value
                
                
        if "Tier 1" in row[colDict[TierType]].value:
            ContainerSizeInfo.T1payout = row[colDict[TierAmount]].value
        if "Tier 2" in row[colDict[TierType]].value:
            ContainerSizeInfo.T2payout = row[colDict[TierAmount]].value
        if "Tier 3" in row[colDict[TierType]].value:
            ContainerSizeInfo.T3payout = row[colDict[TierAmount]].value
        if "Tier 4" in row[colDict[TierType]].value:
            ContainerSizeInfo.T4payout = row[colDict[TierAmount]].value
        if "Thruway surcharge" in row[colDict[TierType]].value:
            ContainerSizeInfo.OWpayout = row[colDict[TierAmount]].value 
        if "Company Driver Thruway" in row[colDict[TierType]].value:
            ContainerSizeInfo.CompTHRUpayout = row[colDict[TierAmount]].value 