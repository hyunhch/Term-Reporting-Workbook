Attribute VB_Name = "RosterButtonSubs"
Option Explicit

Sub RosterClearButton()
'Delete everything, reset columns, clear records, retabulate

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim RosterDelRange As Range
    Dim DelConfirm As Long
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
    
    'Prompt for confirmation
    DelConfirm = MsgBox("You are about to permanently remove all students. This cannot be undone. Do you wish to continue?", vbQuestion + vbYesNo + vbDefaultButton2)
    
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If

    'Pass to remove everything
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterDelRange = RosterTable.ListColumns("First").DataBodyRange
    
    If RemoveFromRoster(RosterSheet, RosterDelRange, RosterTable) <> 1 Then
        GoTo Footer
    End If
    
    Call RemoveTable(RosterSheet)
    
Footer:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub RosterParseButton()
'Read in the roster, table with conditional formatting, Marlett boxes, push to the Records and ReportSheet

    Dim RecordsSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RosterNameRange As Range
    Dim c As Range
    Dim NumDuplicate As Long
    Dim NumAdded As Long
    Dim MessageString As String
    Dim RosterTable As ListObject

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")

    'Remake the Roster table to include new students, if any
    Set RosterTable = MakeRosterTable(RosterSheet)

    'Check if we have any students. Break if we don't
    If CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    End If
    
    'Remove duplicates and blank rows
    NumDuplicate = RemoveBadRows(RosterSheet, RosterTable.DataBodyRange, RosterTable.ListColumns("First").DataBodyRange)
        If NumDuplicate > 0 Then
            MessageString = NumDuplicate & " duplicates removed."
        End If

    'Remake the table
    Set RosterTable = MakeRosterTable(RosterSheet)

    'Check if we have any students. Break if we don't. This might happen when there's a table with all empty rows
    If CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    End If
    
    Call TableFormat(RosterSheet, RosterTable)

    'Pass to add new students to the RecordsSheet
    Set c = CopyToRecords(RosterSheet, RecordsSheet) 'Includes tabulation
        If Not c Is Nothing Then
            MessageString = c.Rows.Count & " students added." & vbCr + MessageString
        End If
        
    If Len(MessageString) > 0 Then
        MsgBox (MessageString)
    End If
        
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
