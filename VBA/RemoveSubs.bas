Attribute VB_Name = "RemoveSubs"
Option Explicit

Sub RemoveDupeBlank(TargetSheet As Worksheet, FullRange As Range, SearchRange As Range)

    Dim c As Range
    Dim d As Range
    Dim NameRange As Range
    Dim DelRange As Range
   
    Set NameRange = SearchRange.Columns(1)
    Set c = FindBlanks(NameRange)
    Set d = FindDuplicate(NameRange)
    
    If Not c Is Nothing Then
        Set DelRange = c
    End If
    
    If Not d Is Nothing Then
        Set DelRange = BuildRange(d, DelRange)
    End If
    
    If Not DelRange Is Nothing Then
        Call RemoveRows(TargetSheet, FullRange, SearchRange, DelRange)
    End If

End Sub

Function RemoveFromActivity(ActivitySheet As Worksheet, DelRange As Range) As Long
'Remove the selected students, retabulate, return the number of removed students
'Returns nothing on error
'Called whenever a student is removed from the RecordsSheet

    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim ActivityNameRange As Range
    Dim ActivityDelRange As Range
    Dim ActivityLabelCell As Range
    Dim RecordsNameRange As Range
    Dim RecordsDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim ActivityTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    
    'Make sure there's a table with students
    i = CheckTable(ActivitySheet)
        If i > 2 Then
            RemoveFromActivity = 0
            
            GoTo Footer
        End If

    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    Set ActivityLabelCell = ActivitySheet.Range("A:A").Find("Practice", , xlValues, xlWhole).Offset(0, 1) 'This should always be present
    
    'If the range is passed from the same page, skip matching
    If DelRange.Worksheet.Name = ActivitySheet.Name Then
        Set ActivityDelRange = DelRange
    
        GoTo UpdateAttendance
    End If
    
    'Otherwise match students, if any. Break if there are no matches
    Set ActivityDelRange = FindName(DelRange, ActivityNameRange)
        If ActivityDelRange Is Nothing Then
            RemoveFromActivity = 0
            
            GoTo Footer
        End If

UpdateAttendance:
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsDelRange = FindName(ActivityDelRange, RecordsNameRange)
    
    If Not RecordsDelRange Is Nothing Then
        Set c = FindRecordsLabel(RecordsSheet, ActivityLabelCell)  'This should always be present
        Set d = RecordsDelRange.Offset(0, c.Column - RecordsDelRange.Column)
        
        d.ClearContents
    End If

RemoveStudents:
    i = ActivityDelRange.Cells.Count
    
    Call UnprotectSheet(ActivitySheet)
    Call RemoveRows(ActivitySheet, ActivityTable.DataBodyRange, ActivityNameRange, ActivityDelRange)
    RemoveFromActivity = i 'Doing this afterward in case there's an error
    
    'Repull the attendance and save the activity
    Call ActivityPullAttendance(ActivitySheet, ActivityLabelCell)
    Call ActivitySave(ActivitySheet, ActivityLabelCell)
    
    'Retabulate activity
    Call UnprotectSheet(ReportSheet)
    Call TabulateActivity(ActivityLabelCell)

Footer:

End Function

Sub RemoveFromRecords(RecordsSheet As Worksheet, DelRange As Range, Optional ShowPrompt As String)
'Prompts for exporting a student, then deletes
'Passing "Yes" will prompt for exporting

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim RosterSheet As Worksheet
    Dim RecordsNameRange As Range
    Dim RecordsDelRange As Range
    Dim RecordsFullRange As Range
    Dim MissingStudentRange As Range
    Dim c As Range
    Dim ExportConfirm As Long
    Dim i As Long
    Dim j As Long
    Dim ExportSheetArray() As Variant
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Check if there are any students on the RecordsSheet. Break if not
    If CheckRecords(RecordsSheet) > 2 Then
        GoTo Footer
    End If

    'Check if there's a table on the RosterSheet
    i = CheckTable(RosterSheet)
        If Not i > 2 Then
            Set RosterTable = RosterSheet.ListObjects(1)
        End If

    'Define ranges
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
        If RecordsNameRange(1, 1).Value = "H BREAK" Then 'This shouldn't happen
            GoTo Footer
        End If

    Set c = RecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set RecordsFullRange = RecordsNameRange.Resize(RecordsNameRange.Rows.Count, c.Column)
    
    'If the DelRange is from the RecordsSheet, we don't need to do anything
    If DelRange.Worksheet.Name = "Records Page" Then
        Set RecordsDelRange = DelRange
    Else
        Set RecordsDelRange = FindName(DelRange, RecordsNameRange)
    End If
    
    If RecordsDelRange Is Nothing Then
        GoTo Footer
    End If
    
    'Also look for missing students, to avoid having two different exports
    If Not i > 2 Then
        Set MissingStudentRange = FindUnique(RecordsNameRange, RosterTable.ListColumns("First").DataBodyRange)
    
        If Not MissingStudentRange Is Nothing Then
            Set RecordsDelRange = BuildRange(MissingStudentRange, RecordsDelRange)
        End If
    End If
    
    'Prompt for export
    If ShowPrompt <> "Yes" Then
        GoTo SkipExport
    End If

    j = RecordsDelRange.Cells.Count
    ExportConfirm = MsgBox(j & " students are no longer on your roster. " & _
        "Do you want to save a copy of these students' attendance before removing them?", vbQuestion + vbYesNo + vbDefaultButton2)
        
    If ExportConfirm = vbYes Then 'Pass range of students for exporting
        ReDim ExportSheetArray(1 To 3)
            ExportSheetArray(1) = "Cover"
            ExportSheetArray(2) = "Roster"
            ExportSheetArray(3) = "Simple"
            
        If Not i > 2 Then
            ReDim Preserve ExportSheetArray(1 To 4)
              ExportSheetArray(4) = "Detailed"
        End If

        Set OldBook = ThisWorkbook
        Set NewBook = ExportMakeBook(RecordsDelRange, ExportSheetArray)
        
        Call ExportLocalSave(OldBook, NewBook)
        OldBook.Activate
    End If
    
SkipExport:
    'Define names + attendance, remove extra students
    Call UnprotectSheet(RecordsSheet)
    Call RemoveRows(RecordsSheet, RecordsFullRange, RecordsNameRange, RecordsDelRange)
    
    'Retabulate
    Call TabulateAll

Footer:
    
End Sub

Function RemoveFromReport(LabelCell As Range) As Long
'Clears all numbers from the row of the passed activity
'Returns 1 if successful, 0 of the activity isn't found (shouldn't happen), nothing on error

    Dim ReportSheet As Worksheet
    Dim ReportLabelCell As Range
    Dim ReportDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim ReportTable As ListObject

    Set ReportSheet = Worksheets("Report Page")
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Find the activity
    Set ReportLabelCell = FindReportLabel(ReportSheet, LabelCell)
        If ReportLabelCell Is Nothing Then
            RemoveFromReport = 0
            
            GoTo Footer
        End If

    'Define the rest of the row. Deleting everything from "Notes" column until the end
    Set c = ReportTable.HeaderRowRange.Find("Notes")
    Set d = ReportSheet.Cells(ReportLabelCell.Row, c.Column)
    Set ReportDelRange = d.Resize(1, ReportTable.ListColumns.Count - d.Column)
    
    'Remove
    Call UnprotectSheet(ReportSheet)
    ReportDelRange.ClearContents
    
    'Clear checks
    ReportTable.ListColumns("Select").DataBodyRange.ClearContents

    RemoveFromReport = 1

Footer:

End Function

Function RemoveFromRoster(RosterSheet As Worksheet, RosterDelRange As Range, RosterTable As ListObject) As Long
'Remove from the RecordsSheet, prompting for an export
'Remove from any activities
'Remove from the RosterSheet
'Retabulate everything
'Returns the number of removed students
'Returns 0 if there is nothing to remove, returns nothing on error

    Dim OldBook As Workbook
    'Dim NewBook As Workbook
    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim RecordsDelRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim DelConfirm As Long
    'Dim ExportConfirm As Long
    'Dim SheetArray() As Variant
    
    Set OldBook = ThisWorkbook
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    
    'Check the RecordsSheet. If there are no students, break
    i = CheckRecords(RecordsSheet) > 2
        If i Then
            RemoveFromRoster = 0
                
            GoTo SkipRecords
        End If

    'Match names on the RecordsSheet
    Set RecordsNameRange = FindRecordsName(RecordsSheet)

    'If everyone is being deleted, we can skip matching on the RecordsSheet
    If RosterDelRange.Address = RosterNameRange.Address Then
        Set RecordsDelRange = RecordsNameRange
    Else
        Set RecordsDelRange = FindName(RosterDelRange, RecordsNameRange)
        
        'If there are no matches, skip exporting and removing from records
        If RecordsDelRange Is Nothing Then
            RemoveFromRoster = 0
            
            GoTo SkipRecords
        End If
    End If
        
    'Confirm deletion
    j = RecordsDelRange.Cells.Count
    DelConfirm = MsgBox("Are you sure you want to remove these " & j & " students?" & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    
    If DelConfirm <> vbYes Then
        RemoveFromRoster = 0
    
        GoTo Footer
    End If
        
    'Promp for exporting
    'ExportConfirm = MsgBox("Do you want to save a copy of these students' attendance before removing them?", vbQuestion + vbYesNo + vbDefaultButton2)
        
        'If ExportConfirm <> vbYes Then
            'GoTo SkipExport
        'End If
    
    'ReDim SheetArray(1 To 4)
        'SheetArray(1) = "Cover"
        'SheetArray(2) = "Roster"
        'SheetArray(3) = "Simple"
        'SheetArray(4) = "Detailed"
    
    'Set NewBook = ExportMakeBook(RecordsDelRange, SheetArray) 'Figure out error handling here
    
    'Call ExportLocalSave(OldBook, NewBook)
    'OldBook.Activate
    
SkipExport:
    'Remove students from any open activity sheet
    For Each ActivitySheet In OldBook.Sheets
        Set c = ActivitySheet.Range("A1")
        
        If c.Value = "Practice" Then
            Call RemoveFromActivity(ActivitySheet, RecordsDelRange)
        End If
    Next ActivitySheet

    'Delete from Records
    Call UnprotectSheet(RecordsSheet)
    Call RemoveFromRecords(RecordsSheet, RecordsDelRange, "Yes") 'Will prompt for export
    
SkipRecords:
    'Delete from Roster
    Call UnprotectSheet(RosterSheet)
    Call RemoveRows(RosterSheet, RosterTable.DataBodyRange, RosterNameRange, RosterDelRange)
    
    'Parse the roster again and tabulate
    Call RosterParseButton
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
    Call TabulateAll

    RemoveFromRoster = 1

Footer:

End Function

Sub RemoveRows(TargetSheet As Worksheet, FullRange As Range, SearchRange As Range, DelRange As Range)
'Sorts the SearchRange, deletes every row in the intersection of DelRange and the FullRange
'Needs to be passed the full range to sort, i.e. a table DataBodyRange

    Dim SortDelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim TargetTable As ListObject
    Dim HasTable As Boolean
    
    Call UnprotectSheet(TargetSheet)

    'I don't think this is needed since I'm defining a number of cells to be deleted rather than the entire row. Need to test
    'Remove any table and formatting
    If TargetSheet.ListObjects.Count > 0 Then
        HasTable = True
        Call RemoveTable(TargetSheet)
    End If
    
    FullRange.FormatConditions.Delete
    
    'Flag each row to be deleted
    DelRange.Interior.Color = vbRed
    
    'Sort by color
    With TargetSheet.Sort
        .SortFields.Clear
        .SortFields.Add2(SearchRange, xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 0, 0)
        .SetRange FullRange
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Find the bounds of the red cells
    'Not looking at contents because the sub can be called to delete any row
    Set c = FullRange.Rows(1) 'First row
    
    For i = c.Row To FullRange.Rows(FullRange.Rows.Count + 1).Row 'In case every row is checked
        Set d = TargetSheet.Cells(i, SearchRange.Column)
        If d.Interior.Color <> vbRed Then
            Set d = d.Offset(-1, 0) 'Last row
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
        Set TargetTable = CreateReportTable
        Call FormatReportTable(TargetSheet, TargetTable)
    Else
        Set TargetTable = CreateTable(TargetSheet)
        Call FormatTable(TargetSheet, TargetTable)
    End If
    
Footer:

End Sub
