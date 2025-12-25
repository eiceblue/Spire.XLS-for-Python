from spire.xls import *


inputFile = "./Demos/Data/AddReportFilter.xlsx"
outputFile = "AddReportFilter_output.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]

# Access the first Pivot Table in the worksheet
pt = sheet.PivotTables[0]
# create a report filter for the field "Product"
reportFilter = PivotReportFilter("Product", True)
# Add the report filter to the pivot table
pt.ReportFilters.Add(reportFilter)
# Save the updated workbook to a new file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()