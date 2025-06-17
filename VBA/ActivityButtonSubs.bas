Attribute VB_Name = "ActivityButtonSubs"
Option Explicit

Sub ActivtyDeleteButton()
'Clears the attendance, notes, report, and closes the sheet

    Dim ActivitySheet As Worksheet
    Dim LabelCell As Range
    Dim DelConfirm As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ActivitySheet = ActiveSheet
    Set LabelCell = ActivitySheet.Range("A:A").Find("Practice", , xlValues, xlWhole).Offset(0, 1)
        If LabelCell Is Nothing Then
            ActivitySheet.Delete
            
            GoTo Footer
        End If
    
    'Confirm deletion
    DelConfirm = MsgBox("Do you wish to delete all saved atttendance for this activity? " & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    'Pass for deletion. Confirmation happens in child sub
    Call ActivityDelete(LabelCell)

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ActivitySaveButton()
'To call the SaveActivity() sub

    Dim ActivitySheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim LabelCell As Range
    Dim ActivityTable As ListObject

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ActivitySheet = ActiveSheet
    Set RecordsSheet = Worksheets("Records Page")

    'Look at the ActivityTable. Clear if there's no table or no rows
    If CheckTable(ActivitySheet) > 2 Then
        Call ActivityDelete(LabelCell)
        
        GoTo Footer
    End If

    'Check that the label is present. It always should be
    Set LabelCell = ActivitySheet.Range("1:1").Find("Practice", , xlValues, xlWhole).Offset(0, 1)
    If LabelCell Is Nothing Or Len(LabelCell.Value) < 1 Then
        MsgBox ("Something has gone wrong. Please close this activity and either load or recreate it.")
        GoTo Footer
    End If

    'Check that there are students on the Records Page. It's okay if there are no activities
    If CheckRecords(RecordsSheet) > 2 Then
        MsgBox ("Please parse the roster and try again.")
        GoTo Footer
    End If

    'Pass to save and close the sheet
    Call ActivitySave(ActivitySheet, LabelCell)

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ActivityCloseButton()
'Deletes the sheet
'Prompts for saving if what's on the ActivitySheet doesn't match the RecordsSheet

    Dim ActivitySheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ActivityNameRange As Range
    Dim ActivityCheckRange As Range
    Dim RecordsNameRange As Range
    Dim RecordsLabelRange As Range
    'Dim RecordsAttendanceRange As Range
    Dim LabelCell As Range
    Dim c As Range
    Dim d As Range
    Dim SaveConfirm As Long
    Dim IsDifferent As Boolean
    Dim ActivityTable As ListObject
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ActivitySheet = ActiveSheet
    Set RecordsSheet = Worksheets("Records Page")
    Set LabelCell = ActivitySheet.Range("A:A").Find("Practice", , xlValues, xlWhole).Offset(0, 1)
        If LabelCell Is Nothing Then
            ActivitySheet.Delete
            
            GoTo Footer
        End If
    
    'If there are no students on the RecordsSheet, just close the sheet
    If CheckRecords(RecordsSheet) > 1 Then
        ActivitySheet.Delete
    
        GoTo Footer
    End If
    
    'If there's no table or rows, delete the activity
    If CheckTable(ActivitySheet) > 2 Then
        Call ActivityDelete(LabelCell)
    End If
    
    'Compare the open sheet with what's on the RecordsSheet
    Set ActivityTable = ActivitySheet.ListObjects(1)
    Set ActivityNameRange = ActivityTable.ListColumns("First").DataBodyRange
    Set ActivityCheckRange = FindChecks(ActivityNameRange.Offset(0, -1))
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
    'Set RecordsAttendanceRange = FindRecordsAttendance(RecordsSheet, , LabelCell)
    
    IsDifferent = False
    For Each c In ActivityNameRange
        Set d = FindName(c, RecordsNameRange)
            If d Is Nothing Then
                GoTo NextRow
            End If
            
        If c.Offset(0, -1).Value = "a" And RecordsSheet.Cells(d.Row, RecordsLabelRange.Column) = "1" Then
            GoTo NextRow
        ElseIf c.Offset(0, -1).Value <> "a" And RecordsSheet.Cells(d.Row, RecordsLabelRange.Column) = "0" Then
            GoTo NextRow
        Else
            IsDifferent = True
        End If
NextRow:
    Next c

    'Prompt to save
    If IsDifferent = True Then
        SaveConfirm = MsgBox("There are unsaved changes on this activity. " & _
            "Would you like to save them before closing the sheet?", vbQuestion + vbYesNo + vbDefaultButton2)
        
        If SaveConfirm <> vbYes Then
            GoTo Footer
        End If
        
        Call ActivitySave(ActivitySheet, LabelCell)
    End If

    ActivitySheet.Delete

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ActivityPullAttendanceButton()
'To call ActivityPullAttendance()

    Dim ActivitySheet As Worksheet
    Dim LabelCell As Range

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ActivitySheet = ActiveSheet
    Set LabelCell = ActivitySheet.Range("A:A").Find("Practice", , xlValues, xlWhole).Offset(0, 1)
        If LabelCell Is Nothing Then
            ActivitySheet.Delete
            
            GoTo Footer
        End If

    Call ActivityPullAttendance(ActivitySheet, LabelCell)

Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
