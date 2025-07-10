from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	

inputFile = "./Demos/Data/AccessDocumentProperties.xlsx"
outputFile = "AccessDocumentProperties.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Create string builder
builder = []
#Get all document properties
properties = workbook.CustomDocumentProperties
#Access document property by property name
property1 = properties["Editor"]
obj = property1.Value
builder.append(property1.Name + " " + String(obj).Value)
#Access document property by property index
property2 = properties[0]
obj2 = property2.Value
builder.append(property2.Name + " " + String(obj2).Value)
#Save to txt file
AppendAllText(outputFile, builder)
workbook.Dispose()

