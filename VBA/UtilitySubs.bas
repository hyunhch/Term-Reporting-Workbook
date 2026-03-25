Attribute VB_Name = "UtilitySubs"
Option Explicit

Sub AddMarlettBox(BoxRange As Range)
'Doing this instead of actual checkboxes to deal with sorting issues
'This only changes the font of a range to Marlett
    
    Dim c As Range

    If BoxRange Is Nothing Then
        GoTo Footer
    End If

    With BoxRange
        .Font.Name = "Marlett"
        .HorizontalAlignment = xlRight
    End With
    
    'Preserve checks, but get rid of anything other than an "a"
    For Each c In BoxRange
        If c.Value <> "a" Then
            c.ClearContents
        End If
    Next c

Footer:

End Sub

Function BuildRange(NewCell As Range, Optional OldRange As Range) As Range
'A function for building ranges cell by cell
'This may be slower

    If OldRange Is Nothing Then
        Set BuildRange = NewCell
    Else
        Set BuildRange = Union(OldRange, NewCell)
    End If

Footer:

End Function

Sub CenterDropdown(TargetSheet As Worksheet, CenterRange As Range)
'Make a dropdown list with center names in the indicated cell

    Call UnprotectSheet(TargetSheet)

    With CenterRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CentersList"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .InputMessage = ""
        .ErrorMessage = "Please choose from the drop-down list"
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub ClearSheet(TargetSheet As Worksheet, Optional TargetRange As Range)
'Clears everything on a sheet and deletes tables
'Passing a range deletes everything in that range

    Dim DelRange As Range
    Dim DelTable As ListObject

     'If DelRange was passed, only delete within that range
     If Not TargetRange Is Nothing Then
        Set DelRange = TargetRange
    Else
        Set DelRange = TargetSheet.Cells
    End If
    
    'Remove any existing table and clear DelRange
    Call RemoveTable(TargetSheet)
    
    With DelRange
        .ClearContents
        .ClearFormats
        .Validation.Delete
    End With
    

End Sub

Sub DateValidation(TargetSheet As Worksheet, DateRange As Range)
'Date greater than 1990

    Call UnprotectSheet(TargetSheet)
    
    With DateRange.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="1/1/1990"
        .IgnoreBlank = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .ErrorMessage = "Please enter a date as mm/dd/yyyy"
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Function LettersOnly(str As String) As String
    Dim i As Long, letters As String, letter As String

    letters = vbNullString

    For i = 1 To Len(str)
        letter = VBA.Mid$(str, i, 1)

        If Asc(LCase(letter)) >= 97 And Asc(LCase(letter)) <= 122 Then
            letters = letters + letter
        End If
    Next
    LettersOnly = letters
End Function

Sub MakeDropdown(TargetSheet As Worksheet, TargetRange As Range, TargetList As String)
'General sub to make a dropdown list
'Pass a named range with TargetList

    Call UnprotectSheet(TargetSheet)

    With CenterRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & TargetList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Error"
        .InputMessage = ""
        .ErrorMessage = "Please choose from the drop-down list"
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Function NudgeToColumn(SourceSheet As Worksheet, SourceRange As Range, ColNum As Long) As Range
'Shifts a range to the passed column number
'Returns nudged range
'Returns nothing on error

    Dim NudgeRange As Range
    Dim c As Range
    
    'Validate passed variables
    If SourceRange Is Nothing Then
        GoTo Footer
    ElseIf Not ColNum > 0 Then
        GoTo Footer
    End If
    
    'If nothing needs to be nudged
    If SourceRange.Column = ColNum Then
        Set NudgeToColumn = SourceRange
    
        GoTo Footer
    End If

    'Intersect of the range rows and the target column
    Set c = SourceSheet.Cells(1, ColNum)
    Set NudgeRange = Intersect(SourceRange.EntireRow, c.EntireColumn)
    
    If Not NudgeRange Is Nothing Then
        Set NudgeToColumn = NudgeRange
    End If

Footer:

End Function

Function NudgeToHeader(SourceSheet As Worksheet, SourceRange As Range, HeaderName As String) As Range
'Shifts a range to a table column with the passed header
'Returns nudged range
'Returns nothing on error

    Dim TargetHeader As Range
    Dim TargetRange As Range
    Dim SourceTable As ListObject
    
    Set SourceTable = SourceSheet.ListObjects(1)

    Set TargetHeader = FindTableHeader(SourceSheet, HeaderName)
        If TargetHeader Is Nothing Then
            GoTo Footer
        ElseIf TargetHeader.Column = SourceRange.Column Then 'If it's already in the same column
            Set TargetRange = SourceRange
            
            GoTo ReturnRange
        End If
    
    Set TargetRange = Intersect(SourceRange.EntireRow, SourceTable.ListColumns(HeaderName).DataBodyRange) 'Not using offset to avoid swapping the sign of the # columns in .Offset()
        If TargetRange Is Nothing Then
            GoTo Footer
        End If
    
ReturnRange:
    Set NudgeToHeader = TargetRange
        
Footer:

End Function


Sub ResetProtection()
'Reset all sheet protections
    
    Dim ReportBook As Workbook
    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim ChangeSheet As Worksheet
    Dim ActivitySheet As Worksheet
    
    Set ReportBook = ActiveWorkbook
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    Set CoverSheet = Worksheets("Cover Page")
    Set ChangeSheet = Worksheets("Change Log")

    RosterSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    ReportSheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    CoverSheet.Protect , userinterfaceonly:=True
    ChangeSheet.Protect , userinterfaceonly:=True

    'Lock/Unlock areas
    CoverSheet.Range("B3:B5").Locked = False
    
    RosterSheet.Cells.Locked = False
    RosterSheet.Range("A1:A5").EntireRow.Locked = True
    
    'Lock the entire page besides the "Select: Column
    ReportSheet.Cells.Locked = True
    ReportSheet.Range("A:A").Locked = False
    ReportSheet.Range("A1:A5").EntireRow.Locked = True
    
    'All activity sheets
    For Each ActivitySheet In ReportBook.Sheets
        If ActivitySheet.Range("A1").Value = "Practice" Then
            ActivitySheet.Protect , userinterfaceonly:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
            ActivitySheet.Cells.Locked = False
            ActivitySheet.Range("A1:A5").EntireRow.Locked = True
            ActivitySheet.Range("B3").Locked = False 'Allow the notes to be editable
        End If
    Next ActivitySheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
End Sub

Sub UnprotectSheet(TargetSheet As Worksheet)
'Checks if a sheet is protected and unprotects
'Used to avoid trying to unprotect an already unprotected sheet

    If TargetSheet.ProtectContents = True Then
        TargetSheet.Unprotect
    End If

End Sub

Sub WipeSheet(TargetSheet As Worksheet)
'Takes everything off the passed sheet, including buttons

    On Error Resume Next
    
    Call UnprotectSheet(TargetSheet)
    With TargetSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With

    On Error GoTo 0

End Sub
