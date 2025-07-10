from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SplitExcelDataIntoMultipleCols.xlsx"
outputFile = "SplitExcelDataIntoMultipleCols.xlsx"

#Create a workbook.
workbook = Workbook()

#Load the file from disk.
workbook.LoadFromFile(inputFile)

#Get the first worksheet.
sheet = workbook.Worksheets[0]

#Split data into separate columns by the delimited characters – space.
splitText = None
text = None
i = 1
while i < sheet.LastRow:
    text = sheet.Range[i + 1,1].Text
    splitText = text.split(' ')
    j = 0
    while j < len(splitText):
        sheet.Range[i + 1,1 + j + 1].Text = splitText[j]
        j += 1
        i += 1
        
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

