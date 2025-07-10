from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "GetNamedRangeAddress.txt"

sb = []
#Create a workbook and load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get specific named range by index
NamedRange = workbook.NameRanges[0]
#Get the address of the named range
address = NamedRange.RefersToRange.RangeAddress
sb.append("The address of the named range " + NamedRange.Name + " is " + address)
#Save and launch result file
AppendAllText(outputFile, sb)
workbook.Dispose()

