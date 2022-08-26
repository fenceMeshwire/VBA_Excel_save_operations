Option Explicit

Sub copy_multiple_sheets_to_wkb()

Worksheets(Array("Sheet1", "Sheet2", "Sheet3")).Copy
  
With ActiveWorkbook
  .SaveAs Filename:=ThisWorkbook.Path & "\result.xlsx", FileFormat:=xlOpenXMLWorkbook
  .Close SaveChanges:=False
End With
  
End Sub
