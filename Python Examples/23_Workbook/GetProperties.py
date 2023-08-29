import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "GetProperties.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the general excel properties
properties1 = workbook.DocumentProperties
sb = []
sb.append("Excel Properties:")
for i, unusedItem in enumerate(properties1):
    name = properties1[i].Name
    obj = properties1[i].Value
    t = properties1[i].PropertyType
    value = None
    if t == PropertyType.Double:
        value = Double(obj).Value
    elif t == PropertyType.DateTime:
        value = DateTime(obj).ToLongDateString()
    elif t == PropertyType.Bool:
        value = Boolean(obj).Value
    elif t == PropertyType.Int:
        value = Int32(obj).Value
    elif t == PropertyType.Int32:
        value = Int32(obj).Value
    else:
        value = String(obj).Value
    sb.append(name + ": " + str(value))
sb.append("")
#Get the custom properties
properties2 = workbook.CustomDocumentProperties
sb.append("Custom Properties:")
for i, unusedItem in enumerate(properties2):
    name = properties2[i].Name
    t = properties2[i].PropertyType
    obj = properties2[i].Value
    value = None
    if t == PropertyType.Double:
        value = Double(obj).Value
    elif t == PropertyType.DateTime:
        value = DateTime(obj).ToLongDateString()
    elif t == PropertyType.Bool:
        value = Boolean(obj).Value
    elif t == PropertyType.Int:
        value = Int32(obj).Value
    elif t == PropertyType.Int32:
        value = Int32(obj).Value
    else:
        value = String(obj).Value
    sb.append(name + ": " + str(value))
#Save the document
File.AppendAllText(outputFile, sb)
workbook.Dispose()
