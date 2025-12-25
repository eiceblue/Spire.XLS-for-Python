from spire.xls import *
from spire.xls.common import *
        
# Create a workbook
workbook = Workbook()

# Load the document from disk
workbook.LoadFromFile("Data/MoveChartsheet.xlsx")

# Get the count of the chartsheets 
print(len(workbook.Chartsheets))


