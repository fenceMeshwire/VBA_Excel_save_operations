Option Explicit

Sub copy_save_worksheet()

Dim dteDate As Date
Dim intFeedback As Integer
Dim strPath As String
Dim strReportType, strWorkbookName As String

intFeedback = MsgBox("Select YES or NO for differentiation.", vbYesNo)

If intFeedback = 6 Then strReportType = "Option_1"
If intFeedback = 7 Then strReportType = "Option_2"

dteDate = Date
strWorkbookName = Year(dteDate) & Month(dteDate) & Day(dteDate) & _
  "_" & strReportType
  
strPath = ThisWorkbook.Path & "\" & strWorkbookName & ".xlsx"

' Copy the desired WorkSheet
Sheet1.Copy

' Silently save and close the new WorkBook thus created.
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs strPath
ActiveWorkbook.Close
Application.DisplayAlerts = True

End Sub
