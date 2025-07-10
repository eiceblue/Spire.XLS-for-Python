from spire.xls import *
from spire.xls.common import *

def AppendText(fname:str,text:str):
    fp = open(fname,"w")
    fp.write(text + "\n")
    fp.close()
inputFile = "./Demos/Data/ReadHyperlinks.xlsx"
outputFile = "ReadHyperlinks.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
address1 = sheet.HyperLinks[0].Address
address2 = sheet.HyperLinks[1].Address
AppendText(outputFile, address1 + "\r\n" + address2)
workbook.Dispose()

