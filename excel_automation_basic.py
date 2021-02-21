
import xlwings as xw

## Create blank workbook
#wb = xw.Book()
## Save excel file
#wb.save("test.xlsx")

## Second example
wbTest = xw.Book(r'E:\python_keyword_args\py_xlwings\test.2.xlsx')

## Reading sheets using index number or name
ws = wbTest.sheets['Tab1']
ws1 = wbTest.sheets['Tab2']
ws2 = wbTest.sheets['Tab3']

## Print name
print(ws.name)

## Manipulate cells 
ws.range('A1').value = 'Hello'
## Manipulate ranges of cells
ws.range('A1:E20').value = 100
## Clear all contents
ws.cells.clear_contents()

## Different methods for referening cells (x, y) rows and columns
ws.cells(1, 1).value = 100
ws.cells(1,'B').value = 200

## Create simple table
ws.range('A1').value = [['Col A', 'Col B'], [10, 20], [30, 40]]
## Create data from left to right
ws.range('A1').value = [100, 200, 300]

## Create data from top to bottom
ws.range('A1').options(transpose=True).value = [100, 200, 300, 400, 500]
## Read data from excel worksheet
print(ws.range('A1').value)
print(ws.range('A1').expand().value)
