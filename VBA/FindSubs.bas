Attribute VB_Name = "FindSubs"
Option Explicit

Function FindActivitySheet(SearchString As String, Optional SearchBook As Workbook) As Worksheet
'Returns the activity sheet with the passed label
'Will search in the passed workbook

    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
    
    If Not SearchBook Is Nothing Then
        Set TargetBook = SearchBook
    Else
        Set TargetBook = ThisWorkbook
    End If
    
ActivitySheet:
    For Each TargetSheet In TargetBook.Sheets
        If TargetSheet.Range("A1").Value = "Practice" And _
        Not TargetSheet.Range("1:1").Find(SearchString, , xlValues, xlWhole) Is Nothing Then
            Set FindActivitySheet = TargetSheet
            GoTo Footer
        End If
    Next TargetSheet

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

Function FindChecks(TargetRange As Range, Optional SearchType As String) As Range
'Returns a range that contains all cells that are not empty or 0
'Passing "Absent" returns absent students (cells with 0)

    Dim CheckedRange As Range
    Dim c As Range
    
    If SearchType <> "Absent" Then
        For Each c In TargetRange
            If c.Value <> "" And c.Value <> "0" Then 'Ignore empty spaces and absenses on the Records sheet
                Set CheckedRange = BuildRange(c, CheckedRange)
            End If
        Next c
    Else
        For Each c In TargetRange
            If c.Value = "0" Then  'Only get absenses
                Set CheckedRange = BuildRange(c, CheckedRange)
            End If
        Next c
    End If
    
    'Return range
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

Function FindPresent(RecordsSheet As Worksheet, LabelCell As Range, Optional OperationString As String) As Range
'Returns the range of all present students given the passed cell
'Returns nothing if there are no students recorded as present, or if the activity isn't found
'Returns absent students if "Absent" is passed
'Returns both absent and present if "All" is passed

    Dim RecordsNameRange As Range
    Dim RecordsAttendanceRange As Range
    Dim c As Range
    Dim d As Range
    Dim e As Range
    Dim IsPresent As Boolean
    Dim IsAbsent As Boolean
    
    'Make sure there are both students and activities
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If
    
    'Find the vertical range containing attendance information
    Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
    
    If RecordsAttendanceRange Is Nothing Then
        GoTo Footer
    End If

    'Check that there are students to return
    IsPresent = IsChecked(RecordsAttendanceRange)
    IsAbsent = IsChecked(RecordsAttendanceRange, "Absent")
    
    'No student attendance, absent or present
    If IsPresent = False And IsAbsent = False Then 'This checks the contents of the range, not if the range exists
        GoTo Footer
    'No absent students
    ElseIf OperationString = "Absent" And IsAbsent = False Then
        GoTo Footer
    'No present students
    ElseIf Len(OperationString) < 1 And IsPresent = False Then
        GoTo Footer
    End If
    
    'Define the range of names and grab all that were present/absent
    Set RecordsNameRange = FindRecordsName(RecordsSheet) 'Should always be in the A column, but making it programmatic
    Set c = FindChecks(RecordsAttendanceRange)
    Set d = FindChecks(RecordsAttendanceRange, "Absent")
    
    'Return
    If Len(OperationString) < 1 Then
        If Not c Is Nothing Then
            Set FindPresent = c.Offset(0, -RecordsAttendanceRange.Column + RecordsNameRange.Column)
        End If
          
    ElseIf OperationString = "Absent" Then
        If Not d Is Nothing Then
            Set FindPresent = d.Offset(0, -RecordsAttendanceRange.Column + RecordsNameRange.Column)
        End If
         
    ElseIf OperationString = "All" Then
        If Not c Is Nothing Then
            Set e = BuildRange(c, e)
        End If
        
        If Not d Is Nothing Then
            Set e = BuildRange(d, e)
        End If
        
        Set FindPresent = e.Offset(0, -RecordsAttendanceRange.Column + RecordsNameRange.Column)
    End If
    
Footer:

End Function

Function FindRecordsAttendance(RecordsSheet As Worksheet, Optional NameCell As Range, Optional LabelCell As Range) As Range
'Returns the intersection of all rows containing students and all columns containing activities
'Passing a cell with a name will return the attendance for just that student
'Passing a cell with a label will return the Attendance for that activity
'Passing both returns the specific cell of the passed student's attendance for the passed activity
'Returns nothing if there are either no students or no activities

    Dim RecordsNameRange As Range
    Dim RecordsLabelRange As Range
    Dim IntersectRange As Range
    Dim MatchCell As Range
    Dim c As Range
    Dim d As Range
    
    'Make sure there are students. This should have already been verified in a parent sub
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
        If RecordsNameRange(1, 1).Value = "H BREAK" Then
            GoTo Footer
        End If
    
    'Make sure there are activities
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet) 'This should always be present
        If RecordsLabelRange(1, 1).Value = "V BREAK" Then
            GoTo Footer
        End If
    
    'First, define c and d as the entire ranges
    Set c = RecordsNameRange
    Set d = RecordsLabelRange
    
    'If anything was passed, redefine the ranges
    If Not NameCell Is Nothing Then
        Set c = FindRecordsName(RecordsSheet, NameCell)
        
        If c Is Nothing Then
            GoTo Footer
        End If
    End If
    
    If Not LabelCell Is Nothing Then 'This can be done with a .Find, might change later
        Set d = FindRecordsLabel(RecordsSheet, LabelCell)
        
        If d Is Nothing Then
            GoTo Footer
        End If
    End If
    
    'Return the range
    Set IntersectRange = Intersect(c.EntireRow, d.EntireColumn)
        If IntersectRange Is Nothing Then
            GoTo Footer
        End If

    Set FindRecordsAttendance = IntersectRange
    
Footer:

End Function

Function FindRecordsLabel(RecordsSheet As Worksheet, Optional LabelCell As Range) As Range
'Returns the range of all activity labels
'If there are no activities, returns the "V BREAK" padding cell
'Returns the cell containing the label if LabelCell is passed
'Returns nothing if LabelCell is passed and a match not found

    Dim FCell As Range
    Dim LCell As Range
    Dim LabelRange As Range
    
    'Define the range of labels. Activity names are used as labels in term reports
    'If the headers are missing, put them back in
    Set FCell = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
        If FCell Is Nothing Then
            Call RecordsSheetText
        End If

    'If no activities. This should never happen for a term report
    Set LCell = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        If LCell.Value = "V BREAK" Then
            Set FindRecordsLabel = FCell
            GoTo Footer
        End If
    
    Set LabelRange = RecordsSheet.Range(FCell.Offset(0, 1), LCell) 'One past the padding cell
    
    'If a name is passed
    If Not LabelCell Is Nothing Then
        Set FCell = LabelRange.Find(LabelCell.Value, , xlValues, xlWhole)
        If Not FCell Is Nothing Then
            Set FindRecordsLabel = FCell
            GoTo Footer
        'If a match isn't found
        Else
            GoTo Footer
        End If
    End If
    
    'Entire range
    Set FindRecordsLabel = LabelRange

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
            Call RecordsSheetText
            
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

Function FindReportLabel(ReportSheet As Worksheet, Optional LabelCell As Range) As Range
'Returns the cell containing the passed label
'Returns the range of all labels if a string isn't passed
'Returns nothing if LabelCell is passed and a match not found
'On the term report, we look for the "Practice" rather than the "Label"

    Dim LabelColumn As Range
    Dim c As Range
    Dim ReportTable As ListObject

    'Make sure there is a table on the page
    If Not ReportSheet.ListObjects.Count > 0 Then
        Call CreateReportTable
    End If
    
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Unlike the weekly reports, all activities will be listed
    Set LabelColumn = ReportTable.ListColumns("Practice").DataBodyRange
   
    'If no string is passed
    If LabelCell Is Nothing Then
        Set c = LabelColumn.Offset(1, 0).Resize(c.Rows.Count - 1, 1)
    Else
        Set c = LabelColumn.Find(LabelCell.Value, , xlValues, xlWhole)
    End If

    'If it's not found, return nothing
    If c Is Nothing Then
        GoTo Footer
    End If
    
    'If found
    Set FindReportLabel = c

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
    
    On Error GoTo Footer
    
    'All tables used will have a cell with "Select" in the 1st column
    Set FCell = TargetSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
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
    If Not WorksheetFunction.CountA(TargetSheet.Cells) > 0 Then
        GoTo Footer
    End If
    
    'Find bounds
    LRow = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    FRow = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
    LCol = TargetSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    FCol = TargetSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
    
    'Return
    Set c = TargetSheet.Cells(FRow, FCol)
    Set d = TargetSheet.Cells(LRow, LCol)
    
    Set FindUsedRange = TargetSheet.Range(c, d)

Footer:

End Function
