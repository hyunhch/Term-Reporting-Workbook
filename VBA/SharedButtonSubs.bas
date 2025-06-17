Attribute VB_Name = "SharedButtonSubs"
Option Explicit

Sub RemoveSelectedButton()
'Can be called from the RosterSheet, ReportSheet, and ActivitySheet

    Dim DelSheet As Worksheet
    Dim CheckRange As Range
    Dim c As Range
    Dim SheetName As String
    Dim DelTable As ListObject

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set DelSheet = ActiveSheet
    
    'Make sure there's a table with at least one student checked
    If CheckTable(DelSheet) <> 1 Then
        GoTo Footer
    End If
    
    Set DelTable = DelSheet.ListObjects(1)
    Set c = DelTable.ListColumns("Select").DataBodyRange.SpecialCells(xlCellTypeVisible)
    Set CheckRange = FindChecks(c)
    
    'Different procedures depending on what sheet it's called from
    SheetName = DelSheet.Name
    
    Select Case SheetName
        Case "Roster Page"
            Call RemoveFromRoster(DelSheet, CheckRange.Offset(0, 1), DelTable)
            
        Case "Report Page"
            For Each c In CheckRange
                Call RemoveFromReport(c.Offset(0, 2)) 'Two columns over to the practice
            Next c
            
        Case Else
            Call RemoveFromActivity(DelSheet, CheckRange.Offset(0, 1))
    End Select
    
Footer:
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub SelectAllButton()
'Assigns value in Select column to "a" or ""
'The report sheet doesn't have a table, consider changing that in the future

    Dim SelectSheet As Worksheet
    Dim SelectRange As Range
    Dim c As Range
    Dim i As Long
    Dim SelectTable As ListObject
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set SelectSheet = ActiveSheet
    
    'Verify that there is a table with rows
    i = CheckTable(SelectSheet)
        If i > 2 Then
            GoTo Footer
        End If
        
    'Define the table and column where checks go
    Set SelectTable = SelectSheet.ListObjects(1)
    Set SelectRange = SelectTable.ListColumns("Select").DataBodyRange
    
    If SelectSheet.Name = "Report Page" Then 'For the Totals row
        If Not SelectTable.DataBodyRange.Rows.Count > 1 Then
            GoTo Footer
        End If
    
        Set c = SelectRange
        Set SelectRange = c.Offset(1, 0).Resize(c.Rows.Count - 1, 1)
    End If
    
    'Check all if any are blank, uncheck all if none are blank
    'Only apply to visible cells
    Call UnprotectSheet(SelectSheet)
    
    With SelectRange
        .Font.Name = "Marlett"
        i = .Cells.Count
        
        If Application.CountIf(SelectRange, "a") = i Then
            .SpecialCells(xlCellTypeVisible).ClearContents
        Else
            .SpecialCells(xlCellTypeVisible).Value = "a"
        End If
    End With

Footer:
    Call ResetProtection

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
