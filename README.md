# Spire.XLS-for-Python
# A professional Excel development component that can be used to create, read, write and convert Excel files on any Python platform.

[![Foo](https://i.imgur.com/pivJdXR.png)](https://www.e-iceblue.com/Introduce/xls-for-python.html)

[Product Page](https://www.e-iceblue.com/Introduce/xls-for-python.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-xls-f4.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

[Spire.XLS for Python](https://www.e-iceblue.com/Introduce/xls-for-python.html) is a professional Excel Python API that can be used to create, read, write, and convert Excel files in any type of Python applications without installing Microsoft Office.

The API supports both the old Excel 97-2003 format (.xls) and the new Excel 2007, Excel 2010, Excel 2013, Excel 2016 and Excel 2019 (.xlsx, .xlsb, .xlsm), along with Open Office(.ods) format. It features fast and reliably compared with developing your own spreadsheet manipulation solution or using Microsoft Automation.

### 100% Standalone Python API
Spire.XLS for Python is a 100% standalone Excel Python class library without requiring Microsoft Excel or Microsoft Office to be installed on the system.

### Freely Operate Excel Files
Create/Save/Merge/Split/Get Excel files.
Encrypt/Decrypt Excel files, add/delete digital signature, tracking changes, lock/unlock worksheets.
Create/Add/Rename/Edit/Delete/Move worksheets.
Insert/Modify/Remove hyperlinks.
Add/Remove/Change/Hide/Show comments in Excel.
Merge/Unmerge Excel cells, freeze/unfreeze Excel panes, insert/delete Excel rows and columns.
Add/Read/Calculate/Remove Excel formulas.
Create/Refresh pivot table.
Apply/Remove conditional format in Excel.
Add/Set/Change Excel header and footer.

### Easily Manipulate Cells & Excel Calculation Engine at Runtime
Developers can easily manipulate Excel cells and Evaluate formula value in Python at runtime. Super-fast, scalable excel calculation engine is compatible with the 97-2003/2007/2010/2013/2016/2019 Excel. Cell Styles are supported by this Excel Python API, such as cell merging/unmerging, text wrapping/unwrapping, text alignment, rotation, interior, borders, lock/unlock and etc. Font formats, like setting font type, size, color, bold, italic, strikeout and underlining etc. is also fully supported. Conditional formatting, text search and replace, filter and data validation can be applied to cells as easily as you expect.

### Powerful & High Quality Excel File Conversion
Convert Excel to PDF/Excel to HTML/Excel to XML/Excel to CSV/Excel to Image/Excel to XPS/Excel to SVG
Convert CSV to Excel/CSV toPDF/Datatable
Convert selected range of cells to PDF
Convert XLS to XLSM and maintain macro
Convert Excel to OpenDocument Spreadsheet(.ods) format
Save Excel chart sheet to SVG/Image
Convert HTML to Excel

### Chart, Data and other Elements
Spire.XLS for Python provides a wide range of Chart: Pie Chart, Bar Chart, Column Chart, Line Chart, Radar Chart and etc. This Excel Python API also supports data transportation between database and Excel in Python. Hyperlinks and templates are also supported by Spire.XLS for Python.

## Examples

### Create an Excel File in Python
```Python
from spire.common import *
from spire.xls import *


outputFile = "CreateAnExcelWithFiveSheet.xlsx"

workbook = Workbook()
workbook.CreateEmptySheets(5)
for i in range(0, 5):
    sheet = workbook.Worksheets[i]
    sheet.Name = "Sheet" + str(i)
    for row in range(1, 151):
        for col in range(1, 51):
            sheet.Range[row,col].Text = "row" + str(row) + " col" + str(col)

workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
```

### Convert Excel to PDF in Python
```Python
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ToPDF.xlsx"
outputFile = "ToPDF.pdf"

#create a workbook
workbook = Workbook()
#load a excel document
workbook.LoadFromFile(inputFile)
workbook.ConverterSetting.SheetFitToPage = True
#convert to PDF file
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()
```

### Convert Excel to Image in Python
```Python
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SheetToImage.xlsx"
outputFile = "SheetToImage.png"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn).Save(outputFile)
workbook.Dispose()
```

[Product Page](https://www.e-iceblue.com/Introduce/xls-for-python.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-xls-f4.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)
