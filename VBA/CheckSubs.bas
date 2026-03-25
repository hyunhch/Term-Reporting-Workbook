Attribute VB_Name = "CheckSubs"
Option Explicit

Function CheckCover() As Long
'Returns 1 if all of the information is filled out on the CoverSheet
'Returns 0 if one or more missing

    Dim CoverSheet As Worksheet
    Dim RefRange As Range
    Dim SearchCell As Range
    Dim c As Range
    Dim i As Long
    Dim SearchString As String
    Dim SearchArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RefRange = Range("CoverTextList")
        If RefRange Is Nothing Then
            GoTo Footer
        End If
        
    CheckCover = 0
    
    'The reference table tells each string, one to the right is where it's placed on the CoverSheet
    'The cell one to the right on the CoverSheetshould have something in it
    For Each c In RefRange
        If c.Value = "Title" Or c.Value = "Version" Then
            GoTo NextSearch
        End If
        
        Set SearchCell = CoverSheet.Range(c.Offset(0, 1).Value)
        
        SearchString = SearchCell.Offset(0, 1).Value
        If Not Len(SearchString) > 0 Then
            GoTo Footer
        End If
        
NextSearch:
    Next c
    
    'If nothing failed
    CheckCover = 1

Footer:

End Function

Function CheckRA(DirectorySheet As Worksheet) As Long
'Add check for RA name and email


End Function

Function CheckRecords(RecordsSheet As Worksheet) As Long
'Two possibilities:
'1 - Students
'2 - No students

    Dim LRowCell As Range

    Set LRowCell = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If LRowCell Is Nothing Then
        GoTo Footer
    ElseIf LRowCell.Value = "H BREAK" Then
        CheckRecords = 2
    Else
        CheckRecords = 1
    End If

Footer:

End Function

Function CheckReport(ReportSheet As Worksheet) As Long
'Ensures that at least the totals row has been filled out
'Four possibilities:
    '1 Totals
    '2 Empty
    '3 No ListRow
    '4 No table

    Dim c As Range
    Dim ReportTable As ListObject
    
    'Check if there's a table. There always should be
    If Not ReportSheet.ListObjects.Count > 0 Then
        CheckReport = 4
        
        GoTo Footer
    End If
    
    'Check for a ListRow
    Set ReportTable = ReportSheet.ListObjects(1)
    
    If Not ReportTable.ListRows.Count > 0 Then 'Maybe have an option for if there are too many rows. There should only be one
        CheckReport = 3
        
        GoTo Footer
    End If
    
    'See if there is a value in the row below the Totals column
    Set c = ReportTable.ListColumns("Total").DataBodyRange
    
    If Not Len(c.Value) > 0 Then
        CheckReport = 2
    Else
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
'Return an null value if there's an error

    Dim TargetCheckRange As Range
    Dim i As Long
    Dim TargetTable As ListObject
    Dim TableHasCheck As Boolean
    
    'Is there a table
    If Not TargetSheet.ListObjects.Count > 0 Then
        i = 4
        
        GoTo Footer
    End If
    
    If TargetSheet.Name = "Report Page" Then 'Use separate function
        Err.Raise vbObjectError + 513, , "Wrong function"
        
        GoTo Footer
    End If
    
    Set TargetTable = TargetSheet.ListObjects(1)
    
    'Are there rows
    If Not TargetTable.ListRows.Count > 0 Then
        i = 3
    'Are there checks
    ElseIf IsChecked(TargetTable.ListColumns("Select").DataBodyRange) = False Then
        i = 2
    Else
        i = 1
    End If
    
Footer:
    CheckTable = i

End Function
