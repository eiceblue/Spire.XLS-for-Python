from spire.common import *
from spire.xls import *


inputFile = "./Demos/Data/DataSorting.xls"
outputFile = "DataSorting.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
worksheet = workbook.Worksheets[0]
workbook.DataSorter.SortColumns.Add(2, OrderBy.Ascending)
workbook.DataSorter.SortColumns.Add(3, OrderBy.Ascending)
workbook.DataSorter.Sort(worksheet["A1:E19"])
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
