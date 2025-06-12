Attribute VB_Name = "CheckSubs"
Option Explicit

Function CheckAttendance(RecordsSheet As Worksheet, NameCell As Range, Optional CountAbsent As String) As Long
'Returns 1 if the passed student was present for anything, 0 otherwise
'Passing "Absent" will consider both present (1) and absent (0) as attending

    Dim AttendanceRange As Range
    Dim i As Long
    
    Set AttendanceRange = FindRecordsAttendance(RecordsSheet, NameCell)
    
    'Either sum or count cells with a number in them
    If Not AttendanceRange Is Nothing Then
        i = WorksheetFunction.Sum(AttendanceRange)
    ElseIf Not AttendanceRange Is Nothing And CountAbsent = "Absent" Then
        i = WorksheetFunction.CountA(AttendanceRange)
    Else
        GoTo Footer
    End If

    'Return a binary answer
    If i > 0 Then
        i = 1
    End If
    
    CheckAttendance = i

Footer:
    
End Function

Function CheckCover() As Long
'Returns 1 if all of the information is filled out on the CoverSheet

    Dim CoverSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim SearchString As String
    Dim SearchArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    ReDim SearchArray(1 To 3)
    SearchArray(1) = "Name"
    SearchArray(2) = "Date"
    SearchArray(3) = "Center"

    CheckCover = 0

    For i = 1 To UBound(SearchArray)
        SearchString = SearchArray(i)
        Set c = CoverSheet.Range("A:A").Find(SearchString, , xlValues, xlWhole).Offset(0, 1)
        
        If Len(c.Value) < 1 Then
            GoTo Footer
        End If
    Next i

    'If nothing failed
    CheckCover = 1

Footer:

End Function

Function CheckRA(DirectorySheet As Worksheet) As Long
'Add check for RA name and email


End Function

Function CheckRecords(RecordsSheet As Worksheet) As Long
'Checks if there are any students on the Records Page
'Three possibilities:
    '1 Students and recorded activities
    '2 Students but no activities
    '3 Neither activities nor students

    Dim AttendanceRange As Range
    
    Set AttendanceRange = FindRecordsAttendance(RecordsSheet)

    If AttendanceRange Is Nothing Then
        CheckRecords = 3
    ElseIf IsChecked(AttendanceRange, "All") = False Then 'Students but no saved activities
        CheckRecords = 2
    Else
        CheckRecords = 1 'Students and saved activities
    End If

Footer:

End Function

Function CheckReport(ReportSheet As Worksheet) As Long
'Ensures that at least the totals row has been filled out
'Three possibilities:
    '1 Totals and activities
    '2 Totals only
    '3 Empty
    '4 No table or only headers. This shouldn't happnen

    Dim TotalHeader As Range
    Dim c As Range
    Dim TotalsColumn As Range
    Dim ReportTable As ListObject
    
    'Check if there's a table with rows. There always should be
    If CheckTable(ReportSheet) > 3 Then
        CheckReport = 4
        
        GoTo Footer
    End If
    
    'See if there is a value in the row below the Totals column
    Set ReportTable = ReportSheet.ListObjects(1)
    Set TotalHeader = FindTableHeader(ReportSheet, "Total")
    
    If Not Len(TotalHeader.Offset(1, 0)) > 0 Then
        CheckReport = 3
        
        GoTo Footer
    Else
        CheckReport = 2
    End If
    
    'See of there are any activities
    Set c = ReportTable.ListColumns("Total").DataBodyRange.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If Not c.Address = TotalHeader.Offset(1, 0).Address Then
        CheckReport = 1
    End If

Footer:

End Function

Function CheckTable(TargetSheet As Worksheet) As Long
'Checks that there is a table, that there is at least one list row, and that there is at least one row checked
'Report sheet will need an additional check since there are two rows at the top
    '1 -> Table, rows, checks
    '2 -> Table, rows
    '3 -> Table
    '4 -> None
'Return a null value if there's an error

    Dim TargetCheckRange As Range
    Dim i As Long
    Dim j As Long
    Dim TargetTable As ListObject
    Dim TableHasCheck As Boolean
    
    'Is there a table
    If TargetSheet.ListObjects.Count < 1 Then
        i = 4
        GoTo Footer
    End If
    
    'Are there rows
    If TargetSheet.Name = "Report Page" Then 'Two rows at the top for the ReportSheet
        j = 2
    Else
        j = 1
    End If
    
    Set TargetTable = TargetSheet.ListObjects(1)
    If TargetTable.ListRows.Count < j Then
        i = 3
        GoTo Footer
    End If
    
    'Are there checks or is a table without that column
    If TargetTable.HeaderRowRange.Find("Select", , xlValues, xlWhole) Is Nothing Then
         i = 2
        GoTo Footer
    End If
    
    TableHasCheck = IsChecked(TargetTable.ListColumns("Select").DataBodyRange)
    If TableHasCheck = False Then
        i = 2
        GoTo Footer
    End If
    
    i = 1
    
Footer:
    CheckTable = i

End Function
