Attribute VB_Name = "RecordsSubs"
Option Explicit

Function RecordsClear(RecordsSheet As Worksheet) As Long
'Wipes all names off the RecordsSheet
'Returns the number of removed students

    Dim RecordsNameRange As Range
    Dim i As Long
    
    i = 0
    
    'Break if there are no students
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo NumberRemoved
    End If
    
    'Define the range of names and delete
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    
    i = RecordsNameRange.Cells.Count
    RecordsNameRange.EntireRow.Delete
    
    'Wipe the Report
    Call ReportClearTotals

NumberRemoved:
    RecordsClear = i

Footer:

End Function
