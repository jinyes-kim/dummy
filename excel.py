import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

# Write
wb = excel.Workbooks.Add()
ws = wb.WorkSheets("Sheet1")
ws.Cells(1, 1).Value = "Hello World"
ws.Range("A1:A10").Interior.ColorIndex = 27
wb.SaveAs("C:\\Jinyes_Trading\\test.xlsx")
excel.Quit()


# Read
wb = excel.Workbooks.Open("C:\\Jinyes_Trading\\test.xlsx")
ws = wb.ActiveSheet
print(ws.Cells(1,1).Value)
excel.Quit()
