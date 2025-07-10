from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/FilterCellsByString.xlsx"
outputFile = "FilterCellsByString_out.xlsx"

#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#get the first worksheet
sheet=workbook.Worksheets[0]
#filter cells data which strat with "South"
sheet.AutoFilters.Range = sheet.Range["D1:D24"]
filtercolumn = sheet.AutoFilters[0]
strCrt = String("South*")
sheet.AutoFilters.CustomFilter(filtercolumn, FilterOperatorType.Equal, strCrt)
sheet.AutoFilters.Filter()
#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()