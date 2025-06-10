Attribute VB_Name = "RosterButtonSubs"
Option Explicit

Sub OpenAddStudentsButton()
'Adds selected students to a saved activity

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")

    'Make sure there is a roster table, that there's at least one student, and that there's at least one checked
    If CheckTable(RosterSheet) > 1 Then
        GoTo Footer
    End If
    
    'Make sure there are any saved activities
    If CheckRecords(RecordsSheet) > 1 Then
        MsgBox ("You have no saved activities.")
        
        GoTo Footer
    End If

    'Show form
    AddStudentsForm.Show

Footer:

End Sub

Sub OpenLoadActivityButton()
'Checks to see if there are any saved activities and opens the load activity form

    Dim RecordsSheet As Worksheet
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Check to make sure there are any saved activities
    If CheckRecords(RecordsSheet) > 1 Then
        MsgBox ("You have no saved activities")
        
        GoTo Footer
    End If
    
    'Show form
    LoadActivityForm.Show
    
Footer:

End Sub

Sub OpenNewActivityButton()
'Opens form to create a new activity. Does not require any selected students

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RosterCheckRange As Range
    Dim RecordsActivityRange As Range
    Dim c As Range
    Dim RosterTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Make sure there's a parsed table with at least one student selected
    If CheckTable(RosterSheet) > 1 Then
        GoTo Footer
    End If
    
    'Make sure the CoverSheet is filled out
    If CheckCover <> 1 Then
        MsgBox ("Please fill out your name, the date, and your center on the Cover Page")
        
        GoTo Footer
    End If
    
    'If every activity has been filled out, don't show
    Set RecordsActivityRange = FindRecordsLabel(RecordsSheet)
    
    'If there are no activities listed
    If RecordsActivityRange(1, 1).Value = "V BREAK" Then
        MsgBox ("Something has gone wrong. Please download a fresh copy of this file.")
        
        GoTo Footer
    End If
    
    For Each c In RecordsActivityRange
        'At least one remaining
        If IsChecked(FindRecordsAttendance(RecordsSheet, , c), "All") = False Then
            NewActivityForm.Show
    
            GoTo Footer
        End If
    Next c
    
    'All activities filled out
    MsgBox ("You have filled out all activities. To modify one, click 'Load Activity' or 'Add to Activity.'")
    
Footer:

End Sub

Sub RosterClearButton()
'Delete everything, reset columns, clear records, retabulate

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterDelRange As Range
    Dim RosterTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")

    'Skip if there are no rows
    If CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    End If
    
    'Pass to remove everything
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterDelRange = RosterTable.ListColumns("First").DataBodyRange
    
    If RemoveFromRoster(RosterSheet, RosterDelRange, RosterTable) <> 1 Then
        GoTo Footer
    End If
    
Footer:

End Sub

Sub RosterParseButton()
'Read in the roster, table with conditional formatting, Marlett boxes, push to the ReportSheet

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim NewStudentRange As Range
    Dim RosterTableRange As Range
    Dim MissingStudentRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim MessageString As String
    Dim HeaderArray() As Variant
    Dim RosterTable As ListObject
    Dim ReportTable As ListObject

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set c = RosterSheet.Range("A6") 'Where the table starts, in case there are no headers
    
    Call UnprotectSheet(RosterSheet)
    
    'If there is a table, remove filters, formatting, and unlist
    If CheckTable(RosterSheet) <> 4 Then
        RosterSheet.AutoFilterMode = False
        Call RemoveTable(RosterSheet)
    End If

    'Reset the headers. I think it'll be cleaner to do this every time
    HeaderArray = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
    
    Call ResetTableHeaders(RosterSheet, c, HeaderArray) 'This will not remove additional columns added to the right of the default ones

    'Find the range for the new table, break if there is nothing but the header
    Set RosterTableRange = FindTableRange(RosterSheet)
    
    If Not RosterTableRange.Rows.Count > 1 Then
        GoTo Footer
    End If
    
    'Make the new table, remove blanks and duplicates. Add formatting
    Set RosterTable = CreateTable(RosterSheet, "RosterTable", RosterTableRange)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set c = FindDuplicate(RosterNameRange)
    
    j = 0
    If Not c Is Nothing Then
        j = c.Cells.Count
    End If
    
    Call RemoveDupeBlank(RosterSheet, RosterTable.DataBodyRange, RosterNameRange)
    Set RosterTable = RosterSheet.ListObjects(1)
    Call FormatTable(RosterSheet, RosterTable)

    'Add Marlett boxes
    Set c = RosterTable.ListColumns("Select").DataBodyRange
    Call AddMarlettBox(c)
    
    'Push to the RecordsSheet and remove students no longer on the roster
    Set NewStudentRange = CopyToRecords(RosterSheet, RecordsSheet)
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set MissingStudentRange = FindUnique(RecordsNameRange, RosterNameRange)
        If Not MissingStudentRange Is Nothing Then
            Call RemoveFromRecords(RecordsSheet, MissingStudentRange, "Yes") 'Will prompt for exporting
        End If
        
    i = 0
    If Not NewStudentRange Is Nothing Then
        i = NewStudentRange.Cells.Count
    End If
    
    'Make sure there is a table on the ReportSheet
    If CheckTable(ReportSheet) > 2 Then
        Set ReportTable = CreateReportTable
    Else
        Set ReportTable = ReportSheet.ListObjects(1)
    End If
    
    'Push totals to the ReportSheet
    Call TabulateReportTotals
    
    'Show how many students were added and if any duplicates were removed
    If i > 0 Then
        MessageString = "Students added: " & i & vbCr
    End If

    If j > 0 Then
        MessageString = MessageString & "Duplicates removed: " & j
    End If

    If Len(MessageString) > 0 Then
        MsgBox MessageString
    End If
    
Footer:
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
