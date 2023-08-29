from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_3.xlsx"
outputFile = "RetrieveAndExtractData.xlsx"

# Create a new workbook instance and get the first worksheet.
newBook = Workbook()
newSheet = newBook.Worksheets[0]

#Create a new workbook instance and load the sample Excel file.
workbook = Workbook()
workbook.LoadFromFile(inputFile)

#Get the first worksheet.
sheet = workbook.Worksheets[0]

#Retrieve data and extract it to the first worksheet of the new excel workbook.
i = 1
columnCount = len(sheet.Columns)
cells = sheet.Columns[0].Cells
for range in cells:
            if range.Text == "teacher":
                sourceRange = sheet.Range[range.Row,1,range.Row,columnCount]
                destRange = newSheet.Range[i,1,i,columnCount]
                sheet.Copy(sourceRange, destRange, True)
                i += 1
#Save to file.
newBook.SaveToFile(outputFile, ExcelVersion.Version2013)
newBook.Dispose()

