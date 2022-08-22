Option Explicit

Sub copy_save_worksheet()

Dim dteDate As Date
Dim intDay, intMonth As Integer
Dim intFeedback As Integer
Dim strDay, strMonth As String
Dim strPath As String
Dim strReportType, strWorkbookName As String

intFeedback = MsgBox("Select YES or NO for differentiation.", vbYesNo)

If intFeedback = 6 Then strReportType = "Option_1"
If intFeedback = 7 Then strReportType = "Option_2"

dteDate = Date
intDay = CInt(Day(dteDate))
intMonth = CInt(Month(dteDate))
If intDay < 10 Then strDay = "0" & CStr(intDay) Else strDay = CStr(intDay)
If intMonth < 10 Then strMonth = "0" & CStr(intMonth) Else strMonth = CStr(intMonth)

strWorkbookName = Year(dteDate) & strMonth & strDay & _
  "_" & strReportType

' Assemble the absolute path to the WorkBook:
strPath = ThisWorkbook.Path & "\" & strWorkbookName & ".xlsx"

' Copy the desired WorkSheet
Sheet1.Copy

' Silently save and close the new WorkBook thus created.
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs strPath
ActiveWorkbook.Close
Application.DisplayAlerts = True

End Sub
