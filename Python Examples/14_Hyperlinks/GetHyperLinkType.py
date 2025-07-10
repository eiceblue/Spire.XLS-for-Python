from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/HyperlinksSample2.xlsx"
outputFile = "GetHyperLinkType.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Iterate all hyperlinks
sb = []
for item in sheet.HyperLinks:
    #Get hyperlink address
    address = item.Address
    #Get hyperlink type
    type = item.Type
    sb.append("Link address: " + address)
    sb.append("Link type: " + str(type))
    sb.append("")
AppendAllText(outputFile, sb)
workbook.Dispose()

