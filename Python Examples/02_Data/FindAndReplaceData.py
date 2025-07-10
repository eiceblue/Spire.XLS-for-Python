from spire.xls.common import *
from spire.xls import *



inputFile = "./Demos/Data/CreateTable.xlsx"
outputFile = "FindAndReplaceData.xlsx"

#Create a workbook
workbook = Workbook()

#Load the Excel document from disk
workbook.LoadFromFile(inputFile)

#Get the first worksheet
worksheet = workbook.Worksheets[0]

#Find the "Brazil" string
ranges = worksheet.FindAllString("Area", False, False)
#Traverse the found ranges
for range in ranges:
    #Replace it with "China"
    range.Text = "Area Code"
    #Highlight the color
    range.Style.Color = Color.get_Yellow()
    
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

