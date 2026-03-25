Attribute VB_Name = "TableSubs"
Option Explicit

Sub TableResetHeaders(TargetSheet As Worksheet, TargetCell As Range, TargetFields As Variant)
'Insert the default column names starting at the passed cell

    Dim FieldRange As Range
    
    Set FieldRange = TargetSheet.Range(TargetCell, TargetCell.Offset(0, UBound(TargetFields) - 1))
    FieldRange.Value = TargetFields

End Sub

Sub TableFormat(TargetSheet As Worksheet, TargetTable As ListObject)
'Flag blanks and bad entries on Roster Sheet
'Different formatting on Report Page

    Dim EthnicityRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim CreditsRange As Range
    Dim MajorRange As Range
    Dim c As Range
    Dim FormulaString1 As String
    Dim FormulaString2 As String
    
    'If there are no rows, skip
    If Not TargetTable.ListRows.Count > 0 Then
        GoTo Footer
    End If
    
    'Blank cells flagged yellow, except in the first column
    With TargetTable
        .DataBodyRange.FormatConditions.Delete
        .DataBodyRange.FormatConditions.Add Type:=xlBlanksCondition
        .ListColumns("Select").DataBodyRange.FormatConditions.Delete
        
        With .DataBodyRange.FormatConditions(1)
        .StopIfTrue = False
        .Interior.ColorIndex = 36
        End With
    End With
    
    'Validate demographics using tables on Ref Tables sheet
    Set EthnicityRange = TargetTable.ListColumns("Race").DataBodyRange
    Set GenderRange = TargetTable.ListColumns("Gender").DataBodyRange

    For Each c In EthnicityRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "RaceList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    For Each c In GenderRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "GenderList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    'Check if this is for College Prep
    If IsCollege = True Then
        GoTo CollegePrep
    End If

    'Credit hours don't need the reference table, add majors to validation
    Set CreditsRange = TargetTable.ListColumns("Credits").DataBodyRange
    Set MajorRange = TargetTable.ListColumns("Major").DataBodyRange
    
    For Each c In CreditsRange
        FormulaString1 = "=AND(NOT(ISNUMBER(" + c.Address + ")),"
        FormulaString2 = " NOT(ISBLANK(" + c.Address + ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c
    
    'Flag majors orange instead of red
    For Each c In MajorRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "MajorList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.ColorIndex = 45
        End With
    Next c

    GoTo Footer

CollegePrep:
    'Grades work as both strings and numbers
    Set GradeRange = TargetTable.ListColumns("Grade").DataBodyRange
    
    For Each c In GradeRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "GradeList" & FormulaString2
        With c.FormatConditions(2)
            .StopIfTrue = False
            .Interior.Color = vbRed
        End With
    Next c

Footer:

End Sub

Sub TableFormatReport(ReportSheet As Worksheet, ReportTable As ListObject)
'Adjust font, color, number formats for the table on the Report Page

    Dim RefRange As Range
    Dim c As Range
    Dim i As Long
    Dim RedValue As String
    Dim GreenValue As String
    Dim BlueValue As String
    Dim TempArray() As String
    Dim ColorArray() As Variant

    'I'm not sure how to prevent the totals row from sorting, so for now I'll disable the autofilter
    ReportTable.ShowAutoFilterDropDown = False
    
    'Remove all colors, then add them in the correct columns. Add vertical lines
    Set RefRange = Range("ReportRGBList")
    
    ReDim ColorArray(1 To RefRange.Cells.Count)
    
    i = 1
    For Each c In RefRange
        ColorArray(i) = c.Value
    
        i = i + 1
    Next c
    
    For i = 1 To ReportTable.ListColumns.Count
        TempArray = Split(ColorArray(i), ",")
        RedValue = TempArray(0)
        GreenValue = TempArray(1)
        BlueValue = TempArray(2)
        With ReportTable.ListColumns(i)
            .Range.Interior.Color = RGB(RedValue, GreenValue, BlueValue)
            .Range.Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range.Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range.Columns.HorizontalAlignment = xlCenter
        End With
    Next i

    'Header row: text black, bold
    With ReportTable.HeaderRowRange
        .Font.Color = vbBlack
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    'Cell under "Total" bold as well
    With ReportTable.ListColumns("Total").DataBodyRange
        .Font.Color = vbBlack
        .Font.Bold = True
    End With
    
    'Center, Name, Date left aligned and autofitted. Date formatted
    With ReportTable.ListColumns("Date").DataBodyRange
        .NumberFormat = "mm/dd/yyyy"
    End With
    
    With ReportSheet.Range(ReportTable.ListColumns("Center").DataBodyRange, ReportTable.ListColumns("Date").DataBodyRange)
        .HorizontalAlignment = xlLeft
        .EntireColumn.AutoFit
    End With

Footer:

End Sub

Sub ResetTableHeaders(TargetSheet As Worksheet, TargetCell As Range, TargetFields As Variant)
'Insert the default column names starting at the passed cell

    Dim FieldRange As Range
    Dim i As Long
    
    'Trim off whitespace
    For i = LBound(TargetFields) To UBound(TargetFields)
        TargetFields(i) = Trim(TargetFields(i))
    Next i
    
    Set FieldRange = TargetSheet.Range(TargetCell, TargetCell.Offset(0, UBound(TargetFields) - 1))
    FieldRange.Value = TargetFields

End Sub

