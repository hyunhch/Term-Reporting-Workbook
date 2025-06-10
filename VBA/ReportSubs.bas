Attribute VB_Name = "ReportSubs"
Option Explicit

Sub ReportClearTotals()
'Only clears the totals. Called when clearing the roster and clearing the entire report

    Dim ReportSheet As Worksheet
    Dim DelRange As Range
    Dim HeaderRefRange As Range
    Dim LastString As String
    
    Set ReportSheet = Worksheets("Report Page")
    Set HeaderRefRange = Range("ReportColumnNamesList")
    
    'We go from "Total" to the end of the table, definted by the last cell in the list
    LastString = HeaderRefRange.Rows(HeaderRefRange.Rows.Count).Value
    
    Set DelRange = FindTableHeader(ReportSheet, "Total", LastString)
    
    'Delete the row beneath the header
    DelRange.Offset(1, 0).ClearContents

End Sub

Sub ReportClearAll()
'Removes everything, including the totals

    Dim ReportSheet As Worksheet
    Dim DelRange As Range
    Dim ReportTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    Call UnprotectSheet(ReportSheet)
    
    'Verify that the table is there
    If CheckTable(ReportSheet) > 2 Then
        Call CreateReportTable
        
        GoTo Footer
    End If
    
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Clear and remake the table
    Set DelRange = ReportTable.Range
    
    Call RemoveTable(ReportSheet)
    
    DelRange.ClearContents
    DelRange.ClearFormats
    
    Set ReportTable = CreateReportTable
    
Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
