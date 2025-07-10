from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ToPDF.xlsx"
outputFile = "SpecifyFontDirectory.pdf"

#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#Specify font directory
workbook.CustomFontFileDirectory= [("./Demos/Data/Fonts/")]
#convert to PDF file
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()