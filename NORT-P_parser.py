#Import modules
import xml.etree.ElementTree as ET
import pandas as pd

#CONSTANTS
#This dictionary lists the ticker names, and the series ID numbers of investment instruments
fundHash = {'VEMIX': 'S000005786', 'VIIIX': 'S000002853', 'VTIVX': 'S000002574',
           'VMCPX': 'S000002844', 'VSCPX': 'S000002845', 'FSMDX': 'S000033637',
           'FSSNX': 'S000033638', 'VTSPX': 'S000038501', 'FXAIX': 'S000006027'}

#This dictionary lists the number of shares owned of each investment instrument
sharesHash = {'VEMIX': 62.01, 'VIIIX': 7.065, 'VTIVX': 0.045,
           'VMCPX': 4.66, 'VSCPX': 5.041, 'FSMDX': 91.872,
           'FSSNX': 112.97, 'VTSPX': 1197.552, 'FXAIX': 33.225}

#This is a prefix that seems to be built into all of the xml tag names
pT = "{http://www.sec.gov/edgar/nport}"

#This is an empty dictionary that describes what data elements should be extracted from the xml file
recordFeatures={'name': [], 'lei': [], 'title':[], 'cusip': [], 
                'balance':[], 'units':[], 'currencyConditional': ['curCd','exchangeRt'], 
                'valUSD': [], 'pctVal': [], 'payoffProfile': [], 'assetCat': [], 'issuerCat': [], 
                'invCountry': [], 'isRestrictedSec': [], 'fairValLevel': []}
				
#parseRecord(aNode, rF = recordFeatures)
#aNode: XML node that represents an individual investment instrument (XML tag invstOrSec)
#rF: Empty dictionary describing what data elements to extract from the XML records  
#Returns a dictionary of data values for the individual investment record
def parseRecord(aNode, rF = recordFeatures):
    #parseValue(k, v, rH, partStr = "")
    #k: Key value that designates either the tag name or the next-level node
    #v: Empty list (if it's the tag name) or list of 2nd-level tags to extract
    #rH: Dictionary to return, will populate with data values
    #partStr: partial string - not currently implemented, but would be needed for deeper nodes
    #No return value
    def parseValue(k, v, rH, partStr = ""):
        #Empty list means the key is the XML tag name
        if len(v) == 0:
            try:
                #Extract the node text
                rH[k] = aNode.find(partStr+pT+k).text
            except AttributeError: #This item is missing
                try:
                    if k == 'issuerCat': #The issuer category had a backup field
                        rH[k] = aNode.find(pT+'issuerConditional').get('issuerCat')
                except KeyError: #Otherwise it's not found
                    #print(f"Attribute not found {rH[k]}: {partStr+pT+k}")
                    rH[k] = ""
        #If the list is not empty, we need to go down a level and extract the items
        else:
            #Each item in the list is a sub-value
            for sV in v:
                try:
                    #Get the value from the sub-node
                    rH[sV] = aNode.find(partStr+pT+k).get(sV)
                except AttributeError: #Otherwise it's not found
                    #print(f"Attribute not found {rH['name']}: {partStr+pT+k}")
                    rH[sV] = ""
    #Initialize an empty dictionary            
    returnHash = {}

    #The ID record is unique in that it has several different potential tag types
    idRecord = aNode.find(pT+'identifiers')[0]
    returnHash['IDtype'] = idRecord.tag.split("}")[1]
    returnHash['ID'] = idRecord.attrib['value']

    #Call parseValue for each value in the record features dictionary
    for k, v in rF.items():
        parseValue(k, v, returnHash)
    
    return returnHash				
   
fundDFhash = {}
#For each fund in the list
for aFund, sID in fundHash.items():
    xmlFN = f"dataFiles/{sID}.xml" #Load XML file
    xmlTree = ET.parse(xmlFN)
    rootNode = xmlTree.getroot()
    #Get a list of all the investment instruments in the XML file
    allRecs = rootNode.findall("./"+pT+"formData/"+pT+"invstOrSecs/")
    #Call parseRecord function for each record
    parsedRecs = [parseRecord(aRec) for aRec in allRecs]
    #Transpose the dictionaries to call DataFrame constructor
    df = pd.DataFrame({k: [rec[k] for rec in parsedRecs] for k in parsedRecs[0].keys()})

    #Convert missing data
    df = df.replace("N/A", None)
    
    #Convert to numeric data types
    df['valUSD'] = df['valUSD'].astype(float)
    df['balance'] = df['balance'].astype(float)
    df['pctVal'] = df['pctVal'].astype(float)

    #Calculating average price per share f
    df['avgPricePerShare'] = df['valUSD']/df['balance']
    df['amtInvested'] = df['avgPricePerShare']*abs(df['pctVal'])*sharesHash[aFund]
    fundDFhash[aFund] = df
    print(f"{aFund}: contains {df.shape[0]} investment instruments")
    
#Try to join them up into a common sheet
#Starter set of columns
summarySheet = list(fundDFhash.values())[0][['ID', 'name']]
#I'm not totally sure what to join on here, because it doesn't seem like there is a reliable unique ID
for name, df in fundDFhash.items():
    df = df[['ID', 'name', 'balance', 'valUSD', 'pctVal', 'avgPricePerShare', 'amtInvested']]
    df.columns = ['ID', 'name'] + [c + "_" + name for c in df.columns if not c in ["ID", "name"]]
    summarySheet = pd.merge(summarySheet, df, how = 'outer', on = ['ID', 'name'])
    
#Dump output to Excel
with pd.ExcelWriter('output.xlsx') as writer:  
    summarySheet.to_excel(writer, sheet_name='Summary')
    for name, df in fundDFhash.items():
        df.to_excel(writer, sheet_name = name)    