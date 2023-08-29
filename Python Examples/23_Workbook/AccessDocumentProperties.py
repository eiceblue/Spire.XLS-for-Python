import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


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
File.AppendAllText(outputFile, builder)
workbook.Dispose()

