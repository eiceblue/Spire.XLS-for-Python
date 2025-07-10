from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_3.xlsx"
outputFile = "FilterCellsByCellColor.xlsx"


#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Create an auto filter in the sheet and specify the range to be filterd
sheet.AutoFilters.Range = sheet.Range["G1:G19"]
#Get the coloumn to be filterd
filtercolumn = sheet.AutoFilters[0]
#Add a color filter to filter the column based on cell color
sheet.AutoFilters.AddFillColorFilter(filtercolumn, Color.get_Red())
#Filter the data.
sheet.AutoFilters.Filter()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

