from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AccessDocumentProperties.xlsx"
outputFile = "LinkToContentProperty.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Add a custom document property
workbook.CustomDocumentProperties.Add("Test", "MyNamedRange")
#Get the added document property
properties = workbook.CustomDocumentProperties
property = properties["Test"]
#Link to content 
property.LinkToContent = True
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

