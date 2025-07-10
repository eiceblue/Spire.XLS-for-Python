from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/templateAz.xlsx"
outputFile = "RemoveCustomProperties.xlsx"

#Create a workbook
workbook = Workbook()
#Load a excel document
workbook.LoadFromFile(inputFile)
#Retrieve a list of all custom document properties of the Excel file
customDocumentProperties = workbook.CustomDocumentProperties
#Remove "Editor" custom document property
customDocumentProperties.Remove("Editor")
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

