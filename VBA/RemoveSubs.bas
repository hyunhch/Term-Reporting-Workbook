Attribute VB_Name = "RemoveSubs"
Option Explicit

Function RemoveBadRows(TargetSheet As Worksheet, TargetRange As Range, SearchRange As Range, Optional SearchType As String) As Long
'Removes duplicate students and blank rows in a given range
'Passing "Duplicate" or "Blank" will restrict deletions to those instances
'Returns the number of duplicates removed
'Returns nothing on error

    Dim c As Range
    Dim d As Range
    Dim DelRange As Range
    Dim i As Long
    
    i = 0
    
    If Not SearchType = "Duplicate" Then
        Set c = FindBlanks(SearchRange)
    End If
    
    If Not SearchType = "Blank" Then
        Set d = FindDuplicate(SearchRange)
    End If
    
    If Not c Is Nothing Then
        Set DelRange = c
    End If
    
    If Not d Is Nothing Then
        Set DelRange = BuildRange(d, DelRange)
        
        i = i + d.Cells.Count
    End If
        
    'Nothing to remove
    If DelRange Is Nothing Then
        RemoveBadRows = 0
        
        GoTo Footer
    End If

    'Delete
    'i = DelRange.Cells.Count
    Call RemoveRows(TargetSheet, TargetRange, DelRange)
    
    RemoveBadRows = i

Footer:

End Function

Function RemoveFromRecords(RecordsSheet As Worksheet, DelRange As Range) As Long
'Returns the number of removed students
'Can take a list of students from the RosterSheet or directly on the RecordsSheet

    Dim RosterSheet As Worksheet
    Dim RosterDelRange As Range
    Dim RecordsNameRange As Range
    Dim RecordsDelRange As Range
    Dim i As Long
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Check if there are any students on the RecordsSheet. Break if not
    If Not CheckRecords(RecordsSheet) = 1 Then
        GoTo Footer
    End If
    
    'Check if DelRange is on the RecordsSheet
    If DelRange.Worksheet.Name = "Records Page" Then
        Set RecordsDelRange = DelRange
    
        GoTo RemoveStudents
    End If

    'Otherwise, match from the RosterSheet to the RecordsSheet
    Set RosterDelRange = NudgeToHeader(RosterSheet, DelRange, "First")
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsDelRange = FindName(RosterDelRange, RecordsNameRange)
        If RecordsDelRange Is Nothing Then
            GoTo Footer
        End If
        
    i = RecordsDelRange.Cells.Count

RemoveStudents:
    Call UnprotectSheet(RecordsSheet)
    Call RemoveRows(RecordsSheet, RecordsNameRange.Resize(RecordsNameRange.Rows.Count, 2), RecordsDelRange)

    'Remove and duplicates and empty rows
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
        If Not RecordsNameRange Is Nothing Then
            Call RemoveBadRows(RecordsSheet, RecordsNameRange.Resize(RecordsNameRange.Rows.Count, 2), RecordsNameRange)
        End If
        
    'Retabulate
    Call TabulateReportTotals

    'Return
    RemoveFromRecords = i

Footer:
    
End Function

Function RemoveFromRoster(RosterSheet As Worksheet, RosterDelRange As Range, RosterTable As ListObject) As Long
'Remove from the RecordsSheet
'Remove from the RosterSheet
'Retabulate everything
'Returns the number of removed students **from RecordsSheet**
'Returns 0 if there is nothing to remove, returns nothing on error

    Dim RecordsSheet As Worksheet
    Dim NudgeDelRange As Range
    Dim i As Long
   
    Set RecordsSheet = Worksheets("Records Page")
    
    'If the entire roster is being removed, we can skip several steps
    If RosterDelRange.Cells.Count = RosterTable.ListRows.Count Then
        i = RecordsClear(RecordsSheet) 'Also wipes the ReportSheet
        RosterTable.DataBodyRange.EntireRow.Delete
    
        GoTo NumberRemoved
    End If
    
    'Nudge to the first name column
    Set NudgeDelRange = NudgeToHeader(RosterSheet, RosterDelRange, "First")
        If NudgeDelRange Is Nothing Then
            GoTo Footer
        End If
    
    'Match names to the RecordsSheet, if applicable
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo RosterDelete
    End If
    
    i = RemoveFromRecords(RecordsSheet, RosterDelRange) 'This retabulates
        
RosterDelete:
    Call RemoveRows(RosterSheet, RosterTable.DataBodyRange, RosterDelRange)
    
    'Remake the table
    Set RosterTable = MakeRosterTable(RosterSheet)
    
    Call TableFormat(RosterSheet, RosterTable)

NumberRemoved:
    RemoveFromRoster = i

Footer:

End Function

Sub RemoveRows(TargetSheet As Worksheet, TargetRange As Range, DelRange As Range)
'TargetRange is the bounds of what to delete, done this was to avoid some errors with tables
'For tables, it should be the DataBodyRange
'DelRange will be in a column. Each passed cell in DelRange will be colored and sorted
'TargetRange and DelRange should always be in the same column, but will always be nudged into column A
'Sorted range has each row removed
'Does not remake a table

    Dim NudgeSortRange As Range
    Dim NudgeDelRange As Range
    Dim SortedDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim TargetTable As ListObject
    
    'Verify passed variables
    If TargetRange Is Nothing Or DelRange Is Nothing Then
        GoTo Footer
    'Is the DelRange entirely inside the Target Range?
    ElseIf Intersect(TargetRange, DelRange).Cells.Count <> DelRange.Cells.Count Then
        GoTo Footer
    'Does the TargetRange includ column A?
    ElseIf Not TargetRange.Columns(1).Column = 1 Then
        GoTo Footer
    End If
    
    'Nudge amd define sort range
    Set NudgeDelRange = NudgeToColumn(TargetSheet, DelRange, 1)
    Set NudgeSortRange = TargetRange.Columns(1)
    
    'Remove tables, if one exists
    Call UnprotectSheet(TargetSheet)
    Call RemoveTable(TargetSheet)
    
    'If everyone is checked, delete everything
    If NudgeDelRange.Address = NudgeSortRange.Address Then
        Set SortedDelRange = TargetRange
        
        GoTo DeleteRows
    End If
    
    'Flag each row to be deleted and sort
    NudgeDelRange.Interior.Color = vbRed
    
    With TargetSheet.Sort
        .SortFields.Clear
        .SortFields.Add2(NudgeSortRange, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0, 0)
        .SetRange TargetRange
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Find the bounds of colored cells
    Set c = TargetRange.Rows(1)
    
    For Each d In TargetRange.Columns(1).Rows
        If d.Interior.Color <> vbRed Then
            Set SortedDelRange = TargetSheet.Range(c, d.Offset(-1, 0)) 'The last row will never be colored because this step is skipped if all rows are being deleted
            
            GoTo DeleteRows
        End If
    Next d
    
DeleteRows:
    SortedDelRange.Delete Shift:=xlUp

Footer:

End Sub

Sub counttest()
    Dim ws As Worksheet
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    
    Set ws = Worksheets("Sheet1")
    Set c = ws.Range("A7:I17")
    'Set c = ws.Range("A1:E5")
    'Set d = ws.Range("A7:A17")
    Set d = ws.Range("A7,A9,A17")

    'i = d.Cells.Count
    'j = Intersect(c, d).Cells.Count
    
    'Debug.Print "DelRange: " & i
    'Debug.Print "Intersect: " & j

    'If Intersect(ws.Range("A:A"), c) Is Nothing Then
        'Debug.Print "Fail"
    'Else
        'Debug.Print "Pass"
    'End If
    
    'Debug.Print c.Columns(1).Address

   'For Each d In c.Columns(1).Rows
        'Debug.Print d.Address
    'Next d

    Call RemoveRows2(ws, c, d)


End Sub

Sub RemoveRowsOLD(TargetSheet As Worksheet, SearchRange As Range, SortRange As Range, DelRange As Range)
'SearchRange is the bound of what to delete, done to avoid some errors with tables
'SortRange is the column being sorted, usually "Select" or "First"
'Del range are the cells in SortRange to delete. The row is removed
'Needs to be passed the SearchRange to sort, i.e. a table DataBodyRange

    Dim SortDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim TargetTable As ListObject
    Dim HasTable As Boolean
    
    Call UnprotectSheet(TargetSheet)

    'I don't think removing that table is needed since I'm defining a number of cells to be deleted rather than the entire row. Need to test
    'Remove any table and formatting
    If TargetSheet.ListObjects.Count > 0 Then
        HasTable = True
        
        'Nudge to the select column
        Set SortRange = NudgeToHeader(TargetSheet, SortRange, "Select")
        Set DelRange = NudgeToHeader(TargetSheet, DelRange, "Select")
        
        Call RemoveTable(TargetSheet)
    End If
    
    SearchRange.FormatConditions.Delete
    
    'Flag each row to be deleted
    DelRange.Interior.Color = vbRed
    
    'Sort by color
    With TargetSheet.Sort
        .SortFields.Clear
        .SortFields.Add2(SortRange.Offset, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0, 0)
        .SetRange SearchRange
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Find the bounds of the red cells
    'Not looking at contents because the sub can be called to delete any row
    Set c = SearchRange.Rows(1)
    
    'For Each d In SortRange.Cells 'This is giving me the wrong row. I'm not sure why
        'If d.Interior.Color <> vbRed Then
            'Set d = SearchRange.Rows(d.Row - 1)
            'Exit For
        'End If
    'Next d
    
    For i = c.Row To SearchRange.Rows(SearchRange.Rows.Count + 1).Row 'In case every row is checked
        Set d = TargetSheet.Cells(i, SortRange.Column)
        If d.Interior.Color <> vbRed Then
            Set d = d.Offset(-1, 0)
            Exit For
        End If
    Next i
    
    'Make a range and delete
    Set SortDelRange = TargetSheet.Range(c, d)
    SortDelRange.Delete Shift:=xlUp
    
    'Put the table back in, if applicable
    If HasTable = False Then
        GoTo Footer
    End If
    
    If TargetSheet.Name = "Report Page" Then
        Set TargetTable = MakeReportTable
        Call TableFormatReport(TargetSheet, TargetTable)
    Else
        Set TargetTable = MakeTable(TargetSheet)
        Call TableFormat(TargetSheet, TargetTable)
    End If
    
Footer:

End Sub

Sub RemoveTable(TargetSheet As Worksheet)
'Unlists all table objects and removes formatting

    Dim OldTableRange As Range
    Dim OldTable As ListObject
    
    Call UnprotectSheet(TargetSheet)
    
    For Each OldTable In TargetSheet.ListObjects
        Set OldTableRange = OldTable.Range
        
        OldTable.Unlist
        OldTableRange.FormatConditions.Delete
        OldTableRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    Next OldTable

End Sub
