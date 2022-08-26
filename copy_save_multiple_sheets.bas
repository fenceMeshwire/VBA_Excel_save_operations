Option Explicit

Sub copy_save_multiple_sheets_to_new_workbook()

Dim dteDate As Date
Dim intDay, intMonth As Integer
Dim strDay, strMonth, strDate, strPath As String

dteDate = Date
intDay = CInt(Day(dteDate))
intMonth = CInt(Month(dteDate))
If intDay < 10 Then strDay = "0" & CStr(intDay) Else strDay = CStr(intDay)
If intMonth < 10 Then strMonth = "0" & CStr(intMonth) Else strMonth = CStr(intMonth) 

strDate = Year(dteDate) & strMonth & strDay
strPath = ThisWorkbook.Path & "\" & strDate & "_result.xlsx"

' This method works with the visible ThisWorkbook.WorkSheets("Sheet1").name
' Multiple WorkSheets can be copied as an array
Worksheets(Array("Sheet1", "Sheet2", "Sheet3")).Copy
  
With ActiveWorkbook
  .SaveAs Filename:=strPath, FileFormat:=xlOpenXMLWorkbook
  .Close SaveChanges:=False
End With
  
End Sub
