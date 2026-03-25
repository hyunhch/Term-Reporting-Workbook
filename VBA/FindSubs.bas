Attribute VB_Name = "FindSubs"
Option Explicit

Function FindAll(SearchRange As Range, SearchTerm As Variant) As Range
'Returns all cells containing the SearchTerm inside the SearchRange
'Returns nothing if Term not found or on error

    Dim c As Range
    Dim d As Range
    Dim ReturnRange As Range
    
    'Make sure there is a valid range passed
    If SearchRange Is Nothing Then
        GoTo Footer
    ElseIf Not SearchRange.Cells.Count > 0 Then
        GoTo Footer
    End If

    'Loop through the SearchRange
    For Each c In SearchRange
        If c.Value = SearchTerm Then
            Set ReturnRange = BuildRange(c, ReturnRange)
        End If
    Next c

    'Return
    If Not ReturnRange Is Nothing Then
        Set FindAll = ReturnRange
    End If

Footer:

End Function

Function FindBlanks(SearchRange As Range) As Range
'Finds empty cells in a column of a larger range, returns as a range
'Returns nothing if there are no blanks

    Dim DelRange As Range
    Dim c As Range
    
    'Build range of blanks
    For Each c In SearchRange.Cells
        If c.Value = "" Then
            Set DelRange = BuildRange(c, DelRange)
        End If
    Next c
    
    'If there are no blanks
    If DelRange Is Nothing Then
        GoTo Footer
    End If

    'REturn
    Set FindBlanks = DelRange
    
Footer:

End Function

Function FindChecks(SearchRange As Range, Optional SearchType As String) As Range
'Returns a range that contains an "a", not used on RecordsSheet
'Passing "Absent" returns the range of all blank boxes
'Passing "First" returns only the first found "a"
'Returns nothing on error

    Dim SearchSheet As Worksheet
    Dim NudgedRange As Range
    Dim CheckedRange As Range
    Dim c As Range
    
    If SearchRange Is Nothing Then
        GoTo Footer
    End If
    
    Set SearchSheet = Worksheets(SearchRange.Worksheet.Name)
    Set NudgedRange = NudgeToHeader(SearchSheet, SearchRange, "Select")
    
    For Each c In NudgedRange
        Select Case SearchType
        
            Case "First"
                If c.Value = "a" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                    
                    GoTo ReturnRange
                End If
                
            Case "Absent"
                If c.Value <> "a" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
                
            Case Else
                If c.Value = "a" Then
                    Set CheckedRange = BuildRange(c, CheckedRange)
                End If
        End Select
    Next c
    
    If CheckedRange Is Nothing Then
        GoTo Footer
    End If

ReturnRange:
    Set FindChecks = CheckedRange
    
Footer:
    
End Function

Function FindDuplicate(SourceRange As Range) As Range
'Intermediate function that determines OS because MacOS doesn't support dictionaries
'Returns the range of all duplicates in the range
'Returns nothing if no duplicates are found

    If Application.OperatingSystem Like "*Mac*" Then
        Set FindDuplicate = FindDuplicateMac(SourceRange)
    Else
        Set FindDuplicate = FindDuplicateWin(SourceRange)
    End If

Footer:

End Function

Function FindDuplicateMac(SourceRange As Range) As Range
'Returns the range of all duplicates in the range
'Returns nothing if no duplicates are found

    Dim DuplicateRange As Range
    Dim CompareRange As Range
    Dim LCell As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim SourceString As String
    Dim CompareString As String

    'Return nothing if it's a single cell
    If Not SourceRange.Cells.Count > 1 Then
        GoTo Footer
    End If

    Set LCell = SourceRange.Rows(SourceRange.Rows.Count)
    
    'Loop through and build range of duplicates
    i = 1
    For Each c In SourceRange.Resize(SourceRange.Cells.Count - 1, 1).Cells 'Stop at the 2nd to last cell
        Set CompareRange = Range(c.Offset(1, 0), LCell)
        SourceString = c.Value & " " & c.Offset(0, 1).Value
        
        For Each d In CompareRange
            CompareString = d.Value & " " & d.Offset(0, 1).Value
            
            If SourceString = CompareString Then
                Set DuplicateRange = BuildRange(d, DuplicateRange)
            End If
        Next d
    Next c

    'Return
    If Not DuplicateRange Is Nothing Then
        Set FindDuplicateMac = DuplicateRange
    End If

Footer:

End Function

Function FindDuplicateWin(SourceRange As Range) As Range
'Returns the range of all duplicates in the range
'Returns nothing if no duplicates are found

    Dim DuplicateRange As Range
    Dim c As Range
    Dim NameString As String
    Dim NameDict As Object
    
    Set NameDict = CreateObject("Scripting.Dictionary")
    NameDict.CompareMode = vbTextCompare

    'Loop through passed range, read into dictionary
    For Each c In SourceRange.Cells

        If Not Len(c.Value) > 0 Then
            GoTo NextName
        End If
    
        NameString = Trim(c.Value) & " " & Trim(c.Offset(0, 1).Value)
        If Not NameDict.Exists(NameString) Then
            NameDict.Add NameString, c
        Else
            Set DuplicateRange = BuildRange(c, DuplicateRange)
        End If
NextName:
    Next c

    'Return
    If Not DuplicateRange Is Nothing Then
        Set FindDuplicateWin = DuplicateRange
    End If

Footer:

End Function

Function FindLastRow(TargetSheet As Worksheet, Optional TargetHeader As String) As Range
'Returns the a cell in the last used row
'Returns the "Select" column by default, the specified column if a string is passed
'Returns nothing on error

    Dim c As Range
    Dim d As Range
    Dim HeaderString As String
    Dim TargetTable As ListObject

    Set TargetTable = TargetSheet.ListObjects(1)
    
    'If a header was passed
    If Len(TargetHeader) > 0 Then
        HeaderString = TargetHeader
    Else
        HeaderString = "Select"
    End If
    
    Set c = TargetTable.ListColumns(HeaderString).Range
    Set d = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If c Is Nothing Then
        GoTo Footer
    ElseIf d Is Nothing Then
        GoTo Footer
    End If
    
    Set FindLastRow = TargetSheet.Cells(d.Row, c.Column)

Footer:
    
End Function

Function FindName(SourceRange As Range, TargetRange As Range) As Range
'Intermediate function that determines OS because MacOS doesn't support dictionaries
'Returns a range of all matching names in the TargetRange
'Returns nothing if no matches found

    If Application.OperatingSystem Like "*Mac*" Then
        Set FindName = FindNameMac(SourceRange, TargetRange)
    Else
        Set FindName = FindNameWin(SourceRange, TargetRange)
    End If

Footer:

End Function

Function FindNameMac(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all matching names in the TargetRange
'Returns nothing if no matches found
'MacOS doesn't support dictionaries

    Dim MatchRange As Range
    Dim c As Range
    Dim d As Range
    Dim SourceName As String
    Dim TargetName As String
    
    'Loop through the SourceRange, only looking for a single match
    For Each c In SourceRange
        SourceName = c.Value & " " & c.Offset(0, 1).Value
        
        For Each d In TargetRange
            TargetName = d.Value & " " & d.Offset(0, 1).Value
        
            If SourceName = TargetName Then
                Set MatchRange = BuildRange(d, MatchRange)
                
                GoTo NextName
            End If
        Next d
NextName:
    Next c

    'Return
    Set FindNameMac = MatchRange

Footer:

End Function

Function FindNameWin(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all matching names in the TargetRange
'Returns nothing if no matches found

    Dim MatchRange As Range
    Dim MatchCell As Range
    Dim c As Range
    Dim d As Range
    Dim NameString As String
    Dim NameDict As Object
    
    Set NameDict = CreateObject("Scripting.Dictionary")
    NameDict.CompareMode = vbTextCompare 'Case insensitive

    'Loop through source range, read all unique names into dictionary
    For Each c In SourceRange
        If Len(c.Value) < 1 Then
            GoTo NextName1
        End If
        
        NameString = Trim(c.Value) & " " & Trim(c.Offset(0, 1).Value) 'Remove whitespace
        If Not NameDict.Exists(NameString) Then
            NameDict.Add NameString, c
        End If
NextName1:
    Next c
    
    'Loop through target range, find any matches
    For Each d In TargetRange
        If Len(d.Value) < 1 Then
            GoTo NextName2
        End If
    
        NameString = Trim(d.Value) & " " & Trim(d.Offset(0, 1).Value)
        If NameDict.Exists(NameString) Then
            Set MatchCell = d
            Set MatchRange = BuildRange(MatchCell, MatchRange)
            
            If SourceRange.Cells.Count = 1 Then 'So we don't loop the entire list of names if we are only looking for one
                GoTo ReturnRange
            End If
        End If
NextName2:
    Next d

ReturnRange:
    If Not MatchRange Is Nothing Then
        Set FindNameWin = MatchRange
    End If
Footer:

End Function

Function FindRecordsName(RecordsSheet As Worksheet, Optional NameCell As Range) As Range
'Returns the entire range of names if nothing passed
'Returns the "H BREAK" padding cell if there are no names
'Returns cell with the student's first name if NameCell is passed
'Returns nothing if a range is passed and a match not found

    Dim FCell As Range
    Dim LCell As Range
    Dim MatchCell As Range
    Dim NameRange As Range
    
    'If the headers are missing, put them back in
    Set FCell = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
        If FCell Is Nothing Then
            Call SetupRecordsText
            
            Set FCell = RecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
        End If

    'If there are no names
    Set LCell = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If LCell.Value = "H BREAK" Then
            Set FindRecordsName = LCell
            GoTo Footer
        End If
    
    Set NameRange = RecordsSheet.Range(FCell.Offset(1, 0), LCell)
    
    'If a name is passed
    If Not NameCell Is Nothing Then
        Set MatchCell = FindName(NameCell, NameRange)
        If Not MatchCell Is Nothing Then
            Set FindRecordsName = MatchCell
            GoTo Footer
        'If a match isn't found
        Else
            GoTo Footer
        End If
    End If
    
    'Entire range
    Set FindRecordsName = NameRange
        
Footer:

End Function

Function FindRemovedStudents(RecordsSheet As Worksheet, RosterSheet As Worksheet) As Range
'Checks if there are students on the RecordsSheet but not on the RosterSheet
'Returns a range of names if any are found
'Returns nothing if none are found or there are no students on the RosterSheet

    Dim RecordsNameRange As Range
    Dim RosterNameRange As Range
    Dim MissingRange As Range
    Dim RosterTable As ListObject
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    
    'If there are no students on the records sheet
    If RecordsNameRange.Cells.Count = 1 Then
        If RecordsNameRange.Value = "H BREAK" Then
            GoTo Footer
        End If
    End If

    'Find names, if any
    Set MissingRange = FindUnique(RecordsNameRange, RosterNameRange)
    
    If Not MissingRange Is Nothing Then
        Set FindRemovedStudents = MissingRange
    End If

Footer:

End Function

Function FindTableHeader(TargetSheet As Worksheet, StartString As String, Optional EndString As String) As Range
'Returns the cell containing the passed string in a table's header
'Returns all header cells between two strings if a second one is passed
'Returns nothing if the header isn't found

    Dim StartCell As Range
    Dim EndCell As Range
    Dim TargetTable As ListObject
    
    'Make sure there's a table
    If TargetSheet.ListObjects.Count < 1 Then
        GoTo Footer
    End If

    Set TargetTable = TargetSheet.ListObjects(1)
    Set StartCell = TargetTable.HeaderRowRange.Find(StartString, , xlValues, xlWhole)

    If StartCell Is Nothing Then
        GoTo Footer
    End If

    If Not Len(EndString) > 0 Then
        'Return one cell
        Set FindTableHeader = StartCell
    Else
        'Return a range
        Set EndCell = TargetTable.HeaderRowRange.Find(EndString, , xlValues, xlWhole)
        
        If EndCell Is Nothing Then
            GoTo Footer
        End If
        
        Set FindTableHeader = TargetSheet.Range(StartCell, EndCell)
    End If
    
Footer:

End Function

Function FindTableRange(TargetSheet As Worksheet) As Range
'Returns the range that will be used to create a table
'Returns empty if there's an error

    Dim FCell As Range
    Dim LCell As Range
    Dim LRow As Long
    Dim LCol As Long
    Dim SearchString As String
    
    On Error GoTo Footer
    
    '"Select" in the first column for the roster, "Center" for the report
    If TargetSheet.Name = "Report Page" Then
        SearchString = "Center"
    Else
        SearchString = "Select"
    End If
    
    Set FCell = TargetSheet.Range("A:A").Find(SearchString, , xlValues, xlWhole)
    LCol = FCell.EntireRow.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Which column to search can change so search all cells
    LRow = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Set LCell = TargetSheet.Cells(LRow, LCol)
    Set FindTableRange = TargetSheet.Range(FCell, LCell)

Footer:

End Function

Function FindTabulateRange(RosterSheet As Worksheet, RecordsSheet As Worksheet, LabelCell As Range) As Range
'Finds all students marked present on the Records Sheet
'Returns where the same students are on the Roster sheet for tabulation

    Dim RosterNameRange As Range
    Dim AttendanceRange As Range
    Dim PresentRange As Range
    Dim TabulateRange As Range
    Dim c As Range
     
    'Find the activity and identify students who were present
    Set PresentRange = FindPresent(RecordsSheet, LabelCell)
    
    'If nothing was found on the Records sheet
    If PresentRange Is Nothing Then
        GoTo Footer
    End If
    
    'Match the students to the Roster sheet
    Set RosterNameRange = RosterSheet.ListObjects(1).ListColumns("First").DataBodyRange
    Set TabulateRange = FindName(PresentRange, RosterNameRange)
    
    If TabulateRange Is Nothing Then 'This shouldn't happen
        GoTo Footer
    End If
    
    Set FindTabulateRange = TabulateRange
    
Footer:
    
End Function

Function FindUnique(SourceRange As Range, TargetRange As Range) As Range
'Intermediate function that determines OS because MacOS doesn't support dictionaries
'Returns a range of all non-matching names. Names in the source range but not the target range
'Returns nothing if no matches found

    If Application.OperatingSystem Like "*Mac*" Then
        Set FindUnique = FindUniqueMac(SourceRange, TargetRange)
    Else
        Set FindUnique = FindUniqueWin(SourceRange, TargetRange)
    End If

Footer:

End Function

Function FindUniqueMac(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all non-matching names. Names in the source range but not the target range
'Returns nothing if no matches found
'MacOS doesn't support dictionaries

    Dim UniqueRange As Range
    Dim c As Range
    Dim d As Range
    Dim SourceName As String
    Dim TargetName As String
    
    'Loop through the SourceRange, only looking for a single match
    For Each c In SourceRange
        SourceName = c.Value & " " & c.Offset(0, 1).Value
        
        For Each d In TargetRange
            TargetName = d.Value & " " & d.Offset(0, 1).Value
        
            If SourceName = TargetName Then
                GoTo NextName
            End If
        Next d
        
        'Unique name
        Set UniqueRange = BuildRange(c, UniqueRange)
NextName:
    Next c

    'Return
    Set FindUniqueMac = UniqueRange

Footer:

End Function

Function FindUniqueWin(SourceRange As Range, TargetRange As Range) As Range
'Returns a range of all non-matching names. Names in the source range but not the target range
'Returns nothing if no matches found

    Dim NoMatchRange As Range
    Dim c As Range
    Dim d As Range
    Dim NameString As String
    Dim NameDict As Object
    
    Set NameDict = CreateObject("Scripting.Dictionary") '
    NameDict.CompareMode = vbTextCompare

    'Loop through source range, read all unique names into dictionary
    For Each c In TargetRange
        If Len(c.Value) < 1 Then
            GoTo NextTargetName
        End If
        
        NameString = c.Value & " " & c.Offset(0, 1).Value
        If Not NameDict.Exists(NameString) Then
            NameDict.Add NameString, c
        End If
NextTargetName:
    Next c

    'Loop through target range, find those that don't match
    For Each c In SourceRange
        If Len(c.Value) < 1 Then
            GoTo NextSourceName
        End If
    
        NameString = c.Value & " " & c.Offset(0, 1).Value
        If Not NameDict.Exists(NameString) Then
            Set NoMatchRange = BuildRange(c, NoMatchRange)
        End If
NextSourceName:
    Next c
    
    'Return
    If Not NoMatchRange Is Nothing Then
        Set FindUniqueWin = NoMatchRange
    End If
Footer:

End Function

Function FindUsedRange(TargetSheet As Worksheet) As Range
'Returns the range between the first cell with text and last cell with text
'Does not consider buttons, empty table rows, formatting, etc.
'Returns nothing on error or if the sheet is blank

    Dim c As Range
    Dim d As Range
    Dim FRow As Long
    Dim FCol As Long
    Dim LRow As Long
    Dim LCol As Long
    
    'Make sure there's something on the sheet
    With TargetSheet
        If Not WorksheetFunction.CountA(.Cells) > 0 Then
            GoTo Footer
        End If
        
        'Find bounds
        LRow = .Cells.Find("*", .Cells(.Rows.Count, .Columns.Count), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        FRow = .Cells.Find("*", .Cells(.Rows.Count, .Columns.Count), SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
        LCol = .Cells.Find("*", .Cells(.Rows.Count, .Columns.Count), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        FCol = .Cells.Find("*", .Cells(.Rows.Count, .Columns.Count), SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
        
        'Return
        Set c = .Cells(FRow, FCol)
        Set d = .Cells(LRow, LCol)
        
        Set FindUsedRange = TargetSheet.Range(c, d)
    End With

Footer:

End Function
