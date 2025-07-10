from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ParetoChart.xlsx"
outputFile = "CreateParetoChart.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet = workbook.Worksheets[0]

#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as Pareto
officeChart.ChartType = ExcelChartType.Pareto

#set data range in the worksheet   
officeChart.DataRange = sheet["A2:B8"]

officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 12
officeChart.PrimaryCategoryAxis.IsBinningByCategory = True

officeChart.PrimaryCategoryAxis.OverflowBinValue = 5
officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1

#set color of Pareto line      
officeChart.Series[0].ParetoLineFormat.LineProperties.Color = Color.get_Blue()

#gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6

#set the chart title
officeChart.ChartTitle = "Expenses"

#hide the legend
officeChart.HasLegend = False

#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()