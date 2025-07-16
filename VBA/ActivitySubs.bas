Attribute VB_Name = "ActivitySubs"
Option Explicit

Function ActivityAddStudents(ActivitySheet As Worksheet) As Range
'Takes checked students on the RosterSheet and the ones not already on an ActivitySheet
'Call from the AddStudents user form
'Returns the range of new students

    Dim RosterSheet As Worksheet
    Dim RosterCheckRange As Range
    Dim ActivityNameRange As Range
    Dim UniqueRange As Range
    Dim c As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim ActivityTable As ListObject
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    Set ActivityTable = ActivitySheet.ListObjects(1)
    
    'Find checked students on the RosterTable
    Set c = RosterTable.ListColumns("Select").DataBodyRange.SpecialCells(xlCellTypeVisible)
    Set RosterCheckRange = FindChecks(c).Offset(0, 1)
    
    'If there are no students on the ActivitySheet, add all of them
    If Not ActivityTable.ListRows.Count > 0 Then
        Set UniqueRange = RosterCheckRange
    'Otherwise, find which students are unique
    Else
        Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
        Set UniqueRange = FindUnique(RosterCheckRange, ActivityNameRange)
        
        If UniqueRange Is Nothing Then
            GoTo Footer
        End If
    End If
    
    'Define the copy and paste ranges
    Set CopyRange = Intersect(UniqueRange.EntireRow, RosterTable.DataBodyRange)
    Set c = ActivityNameRange.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set PasteRange = c.Offset(1, -1)
    
    'Paste and return
    Set ActivityAddStudents = CopyRows(RosterSheet, CopyRange, ActivitySheet, PasteRange)

Footer:

End Function

Sub ActivityDelete(Optional LabelCell As Range)
'Deletes all attendance information and notes for the passed activity

    Dim ActivitySheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim RecordsAttendanceRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim i As Long
    Dim LabelString As String
    
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Check if there are students and activities. If there are no students, then clear everything
    i = CheckRecords(RecordsSheet)
    
    If i > 2 Then 'The records sheet has been broken. Write code to remake the page
        GoTo Footer
    End If
    
    'Define range
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
        If RecordsLabelRange Is Nothing Then
            GoTo Footer
        End If
        
    Set DelRange = FindRecordsAttendance(RecordsSheet, , LabelCell) 'This will be a single column if a range was passed
    
    'Delete the attendance and the notes from the RecordsSheet
    Call UnprotectSheet(RecordsSheet)
    
    DelRange.ClearContents
    RecordsLabelRange.Offset(1, 0).ClearContents
    
    'Clear from the ReportSheet and close any open sheets
    If LabelCell Is Nothing Then
        Call ReportClearAll
        
        For Each ActivitySheet In ThisWorkbook.Sheets
            If ActivitySheet.Range("A1").Value = "Practice" Then
                ActivitySheet.Delete
            End If
        Next ActivitySheet
        
    Else
        LabelString = LabelCell.Value
        
        For Each c In RecordsLabelRange
            Call RemoveFromReport(c)
            
            Set ActivitySheet = FindActivitySheet(LabelString)
            
            If Not ActivitySheet Is Nothing Then
                ActivitySheet.Delete
            End If
        Next c
    End If
    
Footer:

End Sub

Function ActivityNewSheet(InfoArray() As Variant, Optional OperationString As String) As Worksheet
'Called from the new activity form, returns a completed activity sheet
'Activates the sheet if it's already open and ends the subroutine
'The array is 2D and contains the information to be inserted and where it's to be inserted
'(1, 1) -> {What1}
'(1, 2) -> {Address1}
'(1, 3) -> {Value1}, etc.
'Passing "Load" grabs students from Records rather than the Roster

    Dim RecordsSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim NewSheet As Worksheet
    'Dim AttendanceRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim PracticeString As String
    Dim HeaderArray As Variant
    Dim RosterTable As ListObject
    Dim NewTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    'Grab the practice
    For i = LBound(InfoArray) To UBound(InfoArray)
        If InfoArray(i, 1) = "Practice" Then
            PracticeString = InfoArray(i, 3)
        End If
    Next i
    
    'First check if there's an open sheet for the same activity
    Set NewSheet = FindActivitySheet(PracticeString)
        If Not NewSheet Is Nothing Then
            Set ActivityNewSheet = NewSheet
            NewSheet.Activate
            
            GoTo Footer
        End If
        
    'If not, create a new sheet at the end of the workbook
    ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = PracticeString
    Set NewSheet = Worksheets(PracticeString)
    
    'Paste in the passed text
    Call UnprotectSheet(NewSheet)
    
    For i = LBound(InfoArray) To UBound(InfoArray)
        Set c = NewSheet.Range(InfoArray(i, 2))
        
        With c
            .Value = InfoArray(i, 1)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
            .Offset(0, 1).Value = InfoArray(i, 3)

            .Resize(1, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Resize(1, 2).Borders(xlEdgeBottom).Weight = xlMedium
            .Resize(1, 2).WrapText = False
        End With
    Next i
    
    'Autofit the first column
    NewSheet.Range("A1").EntireColumn.AutoFit
    
    'Add buttons
    Call ActivityNewButtons(NewSheet)
    
    'Add a table
    Set c = NewSheet.Range("A6")
    Set PasteRange = c.Offset(1, 0)
    
    HeaderArray = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.Transpose(RosterTable.HeaderRowRange)) 'Have to do it twice since it's in a row
    Call ResetTableHeaders(NewSheet, c, HeaderArray)
    Set c = Nothing
    
    'Don't pull in from the RosterSheet if we're loading
    If Not OperationString = "Load" Then
        Set c = FindChecks(RosterTable.ListColumns("Select").DataBodyRange.SpecialCells(xlCellTypeVisible)) 'Who is checked on the RosterSheet
        
        If Not c Is Nothing Then
            Set CopyRange = Intersect(c.EntireRow, RosterTable.DataBodyRange)
            
            Call CopyRows(RosterSheet, CopyRange, NewSheet, PasteRange)
        End If
    End If
    
    Set NewTable = CreateTable(NewSheet)
    Call FormatTable(NewSheet, NewTable)
    
    If OperationString = "Load" Then 'This is a messy logic flow
        Call ActivityPullAttendance(NewSheet, NewSheet.Range("B1"))
    End If
    
    'Clean up. This shouldn't be needed
    If Not c Is Nothing Then
        Call RemoveDupeBlank(NewSheet, NewTable.DataBodyRange, NewTable.ListColumns("First").DataBodyRange)
    End If
    
    Set ActivityNewSheet = NewSheet
    
Footer:
    Call ResetProtection

End Function

Sub ActivityNewButtons(ActivitySheet As Worksheet)
'Called when an activity is created or loaded

    Dim NewButton As Button
    Dim NewButtonRange As Range

    'Select All
    Set NewButtonRange = ActivitySheet.Range("A5:B5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SelectAllButton"
        .Caption = "Select All"
    End With

    'Delete Row
    Set NewButtonRange = ActivitySheet.Range("C5:D5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelectedButton"
        .Caption = "Delete Row"
    End With

    'Delete activity button
    Set NewButtonRange = ActivitySheet.Range("J1:K1")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ActivtyDeleteButton"
        .Caption = "Delete Activity"
    End With
    
    'Save Activity button
    Set NewButtonRange = ActivitySheet.Range("G1:H1")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ActivitySaveButton"
        .Caption = "Save Activity"
    End With
    
    'Close Activity button
    Set NewButtonRange = ActivitySheet.Range("G5:H5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ActivityCloseButton"
        .Caption = "Close Sheet"
    End With
    
    'Pull attendence button
    Set NewButtonRange = ActivitySheet.Range("E5:F5")
    Set NewButton = ActivitySheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ActivityPullAttendanceButton"
        .Caption = "Pull Attendence"
    End With

End Sub

Sub ActivityPullAttendance(ActivitySheet As Worksheet, LabelCell As Range)
'Pulls attendance for all students marked "present" in the Records sheet to an activity Sheet

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RecordsLabelCell As Range
    Dim RecordsNameRange As Range
    Dim ActivityNameRange As Range
    Dim AttendanceRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")

    'Check if there are both students and activities
    If CheckRecords(RecordsSheet) <> 1 Then
        GoTo Footer
    End If
    
    'Check that there are students
    i = CheckTable(ActivitySheet)
        If i > 3 Then
            GoTo Footer
        End If
    
    'Copy over any missing students
    Call CopyMissing(ActivitySheet, LabelCell, i)
    
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsLabelCell = FindRecordsLabel(RecordsSheet, LabelCell)
    
    'Clear out any checks on the ActivitySheet
    Set ActivityNameRange = ActivitySheet.ListObjects(1).ListColumns("First").DataBodyRange
        If ActivityNameRange Is Nothing Then
            GoTo Footer
        End If
    
    ActivityNameRange.Offset(0, -1).ClearContents
    
    'Loop through and find students marked present. See if any are missing on the ActivitySheet
    Set AttendanceRange = FindPresent(RecordsSheet, RecordsLabelCell, "All")

    For Each c In AttendanceRange
        Set d = FindName(c, ActivityNameRange)
        
        If Not d Is Nothing Then
            If RecordsSheet.Cells(c.Row, RecordsLabelCell.Column) = "1" Then
                d.Offset(0, -1).Value = "a"
            End If
        End If
    Next c
    
Footer:

End Sub

Sub ActivitySave(ActivitySheet As Worksheet, LabelCell As Range)
'Pushes the marked attendance from the activity sheet to the Records Page
'Separated out tabulation to avoid it being automatically called multiple times

    Dim RecordsSheet As Worksheet
    Dim RecordsNameRange As Range
    Dim RecordsLabelRange As Range
    Dim ActivityNameRange As Range
    Dim ActivityPresentRange As Range
    Dim c As Range
    Dim d As Range
    Dim ActivityTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange

    'Clear out any existing information
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell) 'This should always be present
    
    RecordsNameRange.Offset(0, RecordsLabelRange.Column - RecordsNameRange.Column).ClearContents
    
    'If there are no students anymore, skip to tabulating and close the sheet
    If ActivityNameRange Is Nothing Then
        GoTo TabulateActivity
    End If

    'Loop through the update attendance
    For Each c In ActivityNameRange
        Set d = FindName(c, RecordsNameRange)
        
        If Not d Is Nothing Then
            If c.Offset(0, -1) = "a" Then
                d.Offset(0, RecordsLabelRange.Column - d.Column).Value = 1
            Else
                d.Offset(0, RecordsLabelRange.Column - d.Column).Value = 0
            End If
        End If
    Next c

TabulateActivity:
    Call TabulateActivity(LabelCell)

    If ActivityNameRange Is Nothing Then
        ActivitySheet.Delete
    End If

Footer:

End Sub


