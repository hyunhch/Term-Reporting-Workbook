Attribute VB_Name = "CopySubs"
Option Explicit

Function CopyMissing(ActivitySheet As Worksheet, LabelCell As Range, i As Long) As Range
'Copies over any students saved on the RecordsSheet that are not on the ActivitySheet
'If i = 3, then there are no list rows
'Returns a range of added students
'Returns nothing otherwise

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim ActivityNameRange As Range
    Dim RecordsAttendanceRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim RosterTable As ListObject
    Dim ActivityTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set ActivityTable = ActivitySheet.ListObjects(1) 'This should only be called when a table already exists
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    
    'Grab all students marked present OR absent for the activity
    Set RecordsAttendanceRange = FindPresent(RecordsSheet, LabelCell, "All")
        If RecordsAttendanceRange Is Nothing Then
            GoTo Footer
        End If
        
    'Compare with the names on the ActivitySheet IF there are rows
    Set c = RecordsAttendanceRange.Offset(0, -RecordsAttendanceRange.Column + 1) 'Names in the first column
    
    If i = 3 Then 'No ListRows
        Set d = c
    Else
        Set d = FindUnique(c, ActivityNameRange)
            If d Is Nothing Then
                GoTo Footer
            End If
    End If
        
    'Find on the Roster Sheet and copy over
    Set c = Nothing
    Set c = FindName(d, RosterNameRange) 'These should always be present
    Set CopyRange = Intersect(c.EntireRow, RosterTable.DataBodyRange)
    
    Set d = ActivityTable.ListColumns("First").Range.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set PasteRange = d.Offset(0, -1)
        If i = 3 Then
            Set PasteRange = PasteRange.Offset(1, 0) 'One below the headers
        End If
    
    Set CopyMissing = CopyRows(RosterSheet, CopyRange, ActivitySheet, PasteRange)
    Call CreateTable(ActivitySheet)

Footer:

End Function

Function CopyRows(SourceSheet As Worksheet, SourceRange As Range, TargetSheet As Worksheet, TargetRange As Range) As Range
'Copies over each row of the SourceRange, starting at the TargetRange
'Checking for duplicates, etc. should be done in parent function
'Returns the first column of the copied range

    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    
    i = 0
    For Each c In SourceRange.Rows '.SpecialCells(xlCellTypeVisible) This needs to be passed beforehand
        Set CopyRange = c
        Set d = TargetRange.Resize(1, c.Columns.Count)
        Set PasteRange = d.Offset(i, 0)
        Set ReturnRange = BuildRange(TargetRange.Offset(i, 0), ReturnRange) 'Maybe change to PasteRange so we get the entire row
        
        PasteRange.Value = CopyRange.Value
        i = i + 1
    Next c
    
    Set CopyRows = ReturnRange
    
End Function

Function CopyToRecords(RosterSheet As Worksheet, RecordsSheet As Worksheet) As Range
'Copies new students from the RosterSheet to the RecordsSheet
'Duplicates are ignored, blank rows and duplicates are deleted
'Returns the range of copied names

    Dim FullNameRange As Range
    Dim DelRange As Range
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim c As Range
    Dim d As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim RosterTable As ListObject
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    
    'Find the bottom of the name list
    Set c = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set PasteRange = c.Offset(1, 0)
    
    'If there are no names, simply copy everything
    If c.Value = "H BREAK" Then
        Set d = RosterNameRange
    Else
        'Find all non-duplicative students and copy over
        Set d = FindUnique(RosterNameRange, RecordsNameRange)
    End If

    If Not d Is Nothing Then
        Set CopyRange = Union(d, d.Offset(0, 1)) 'first and last names
        Set CopyToRecords = CopyRows(RosterSheet, CopyRange, RecordsSheet, PasteRange)
    End If
    
    'Delete blanks and duplicates
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set FullNameRange = RecordsNameRange.Resize(RecordsNameRange.Rows.Count, 2) 'Both columns
    
    Call RemoveDupeBlank(RecordsSheet, FullNameRange, RecordsNameRange)
       
Footer:

End Function

Function CopyToActivity(RosterSheet As Worksheet, ActivitySheet As Worksheet, Optional CopyNames As Range) As Range
'Copies non-duplicative selected students from the RosterSheet to an ActivitySheet
'Returns the first names of copied students
'Returns nothing if no students unique students are checked
'Passing a range of names skips finding checks and only copies those names

    Dim c As Range
    Dim d As Range
    Dim RosterNameRange As Range
    Dim ActivityNameRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim i As Long
    Dim RosterTable As ListObject
    Dim ActivityTable As ListObject
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    
    'Check if the ActivityTable already has students
    i = CheckTable(ActivitySheet)
    
    If i > 3 Then 'No table, something has gone wrong
        GoTo Footer
    End If
        
    'Define where to begin pasting, one row under the table
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set PasteRange = FindTableHeader(ActivitySheet, "First").Offset(ActivityTable.Range.Rows.Count, 0)

    'If there weren't any students, we can skip matching for duplicates
    If i > 2 Then
        Set c = FindChecks(RosterNameRange)
    Else
        Set d = FindChecks(RosterNameRange)
        Set c = FindUnique(c, RosterTable.ListColumns("First").DataBodyRange)
    End If
    
    If c Is Nothing Then
        GoTo Footer
    End If
    
    'Resize and copy over
    Set d = RosterTable.DataBodyRange.Resize(RosterTable.DataBodyRange.Rows.Count, RosterTable.ListColumns.Count - 1).Offset(0, 1) 'Everything but the first column
    Set CopyRange = Intersect(c.EntireRow, d)
    
    'Return
    Set CopyToActivity = CopyRows(RosterSheet, CopyRange, ActivitySheet, PasteRange)
    
Footer:

End Function

Function CopyToReport(ReportSheet As Worksheet, PasteCell As Range, PasteArray As Variant) As Range
'Copies values passed in the array to the row of the PasteCell
'PasteArray has the header names in (i, 1), the values in (i, 2)
'Can pass any column in the report header

    Dim ReportHeaderRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim OtherIndex As Long
    Dim OtherString As String
    Dim HeaderString As String
    Dim ReportTable As ListObject

    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportHeaderRange = ReportTable.HeaderRowRange
    
    Call UnprotectSheet(ReportSheet)
    
    'Loop through the paste array and find corresponding headers
    For i = LBound(PasteArray) To UBound(PasteArray)
        HeaderString = PasteArray(i, 1)
        
        'Grab the index of the "Other" category, if any
        If InStr(1, HeaderString, "Other") > 0 Then
            OtherIndex = i
            OtherString = PasteArray(OtherIndex, 1)
        End If
        
        'Find the header, insert value if found
        Set c = ReportHeaderRange.Find(HeaderString, , xlValues, xlWhole)
        
        If Not c Is Nothing Then
            'Paste under the matching header
            Set d = ReportSheet.Cells(PasteCell.Row, c.Column)
            d.Value = PasteArray(i, 2)
            Set ReturnRange = BuildRange(d, ReturnRange)
            
            'Get rid of zeroes
            If d.Value = 0 Then
                d.ClearContents
            End If
        Else
            'Sum up elements that aren't found in the header
            If IsNumeric(PasteArray(i, 2)) = True Then
                j = j + PasteArray(i, 2) 'Using this so we don't run into problems with strings
            End If
        End If
    Next i

    'Skip adding up "others" if no other category was found, such as with Low Income
    If OtherIndex < 1 Then
        GoTo ReturnRange
    End If

    'Push all leftover elements into the "Other" category
    'This will allow the list of categories to change in the future
    If j > 0 Then
        PasteArray(OtherIndex, 2) = PasteArray(OtherIndex, 2) + j
        Set c = ReportHeaderRange.Find(OtherString, , xlValues, xlWhole)
        Set d = ReportSheet.Cells(PasteCell.Row, c.Column)
        
        d.Value = PasteArray(OtherIndex, 2)
        Set ReturnRange = BuildRange(d, ReturnRange)
    End If

ReturnRange:
    If Not ReturnRange Is Nothing Then
        Set CopyToReport = ReturnRange
    End If

Footer:

End Function
