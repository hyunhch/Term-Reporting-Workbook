Attribute VB_Name = "ReportSubs"
Option Explicit

Sub ReportClearTotals()
'Only clears the totals. Called when clearing the roster and clearing the entire report

    Dim ReportSheet As Worksheet
    Dim DelRange As Range
    Dim HeaderRefRange As Range
    Dim LastString As String
    Dim ReportTable As ListObject
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure there's a table
    If CheckReport(ReportSheet) > 2 Then
        Call MakeReportTable
        
        GoTo Footer
    End If

    'Clear the second row
    Set ReportTable = ReportSheet.ListObjects(1)
    
    ReportTable.DataBodyRange.ClearContents

Footer:

End Sub

Sub ReportCoverInfo(ReportSheet As Worksheet)
'Pulls in information from the CoverSheet

    Dim CoverSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim HeaderString As String
    Dim ValueString As String
    Dim ReportTable As ListObject
    Dim CoverArray As Variant
    
    If ReportSheet.ListObjects.Count <> 1 Then
        GoTo Footer
    End If
    
    Set ReportTable = ReportSheet.ListObjects(1)
    
    CoverArray = GetCoverInfo
        If IsEmpty(CoverArray) Or Not IsArray(CoverArray) Then
            GoTo Footer
        End If
    
    For i = 1 To UBound(CoverArray, 2)
        HeaderString = CoverArray(1, i)
        ValueString = CoverArray(2, i)
        
        Set c = ReportTable.HeaderRowRange.Find(HeaderString, , xlValues, xlWhole)
            If c Is Nothing Then
                GoTo NextHeader
            End If
        
        c.Offset(1, 0).Value = ValueString
        c.EntireColumn.AutoFit
NextHeader:
    Next i
        
Footer:

End Sub
