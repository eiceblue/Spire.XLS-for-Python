from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ModifyHyperlink.xlsx"
outputFile = "ModifyHyperlink.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the collection of all hyperlinks in the worksheet
sheet = workbook.Worksheets[0]
#Change the values of TextToDisplay and Address property 
links = sheet.HyperLinks
links[0].TextToDisplay = "Spire.XLS for .NET"
links[0].Address = "http://www.e-iceblue.com/Introduce/excel-for-net-introduce.html"
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

