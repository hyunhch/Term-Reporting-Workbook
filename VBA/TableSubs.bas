Attribute VB_Name = "TableSubs"
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

Function CreateTable(TargetSheet As Worksheet, Optional NewTableName As String, Optional OldTableRange As Range) As ListObject
'Create a table on an existing or new sheet
'Can pass a name or a range for the table

    Dim NewTableRange As Range
    Dim NameRange As Range
    Dim BoxRange As Range
    Dim DelRange As Range
    Dim c As Range
    Dim NewTable As ListObject
    
    'Unprotect
    Call UnprotectSheet(TargetSheet)
    
    'Find the range to use
    If OldTableRange Is Nothing Then
        Set NewTableRange = FindTableRange(TargetSheet)
    Else
        Set NewTableRange = OldTableRange
    End If

    If NewTableRange Is Nothing Then
        MsgBox ("There was a problem creating a table on sheet " & TargetSheet.Name)
        GoTo Footer
    End If
    
    'Unlist any existing table
    Call RemoveTable(TargetSheet)

    'Clear formats and create a table
    NewTableRange.ClearFormats
    
    Set NewTable = TargetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=NewTableRange, _
        xlListObjectHasHeaders:=xlYes)
        
    NewTable.ShowTableStyleRowStripes = False
    
    'Assign a name if passed
    If Len(NewTableName) > 0 Then
        NewTable.Name = NewTableName
    End If
    
    'Removed getting rid of blank rows
    'Better to use a RemoveBlanks function that can be pointed to any column, rather than just the first name column of a table
    
    'Put in Marlett boxes. I had taken this out but can 't remember why
    If NewTable.HeaderRowRange.Find("Select") Is Nothing Then
        GoTo ReturnTable
    End If
    
    If Not NewTable.DataBodyRange Is Nothing Then
        Set BoxRange = NewTable.ListColumns("Select").DataBodyRange

        Call AddMarlettBox(BoxRange)
    End If
    
ReturnTable:
    Set CreateTable = NewTable

Footer:

End Function

Function CreateReportTable() As ListObject
'Grabs headers from reference page, unmakes and remakes the table
'Called when adding or deleting rows, tabulating totals

    Dim ReportSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim ReportTableStart As Range
    Dim ReportTableRange As Range
    Dim HeaderRange As Range
    Dim BoxRange As Range
    Dim c As Range
    Dim i As Long
    Dim HeaderArray() As Variant
    Dim TotalsArray() As Variant
    Dim ActivityArray() As Variant
    Dim ReportTable As ListObject
    
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set ReportTableStart = ReportSheet.Range("A:A").Find("Select", , xlValues, xlWhole)
    
    If ReportTableStart Is Nothing Then 'If the table headers got messed up
        Set ReportTableStart = ReportSheet.Range("A6")
    End If
       
    Call UnprotectSheet(ReportSheet)
    
    'Remove any existing filters, unlist the table and remove formatting
    If ReportSheet.AutoFilterMode = True Then
        ReportSheet.AutoFilterMode = False
    End If
    
    Call RemoveTable(ReportSheet)
    
    'Reset headers. This creates two rows
    HeaderArray = Application.Transpose(ActiveWorkbook.Names("ReportColumnNamesList").RefersToRange.Value)
    TotalsArray = Application.Transpose(ActiveWorkbook.Names("ReportTotalsRowList").RefersToRange.Value)
    
    Call ResetTableHeaders(ReportSheet, ReportTableStart, HeaderArray)
    Call ResetTableHeaders(ReportSheet, ReportTableStart.Offset(1, 0), TotalsArray) 'Only three columns
    
    'Put in the activity category and practice. Category is one column to the left of Practice
    i = Range("ActivitiesList").Cells.Count
    ReDim ActivityArray(1 To i, 1 To 2)
    
    'Grab the category and practice reference lists
    i = 1
    
    For Each c In Range("ActivitiesList")
        ActivityArray(i, 1) = Trim(c.Offset(0, -1).Value)
        ActivityArray(i, 2) = Trim(c.Value)
    
        i = i + 1
    Next c
        
    'Copy over to the ReportSheet
    Set c = ReportTableStart.Offset(2, 1) 'Below the Total row
    
    For i = 1 To UBound(ActivityArray)
        c.Offset(i - 1, 0).Value = ActivityArray(i, 1)
        c.Offset(i - 1, 1).Value = ActivityArray(i, 2)
    Next i
    
    'Define table range and clear formats
    Set ReportTableRange = FindTableRange(ReportSheet)
    ReportTableRange.ClearFormats
    
    'Make a new table
    Set ReportTable = ReportSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=ReportTableRange, _
        xlListObjectHasHeaders:=xlYes)
    ReportTable.Name = "ReportTable"
    
FormatTable:
    'Format
    ReportTable.ShowTableStyleRowStripes = False
    Call FormatReportTable(ReportSheet, ReportTable)
    
    'Add Marlett Boxes to everything but the Totals row
    Set BoxRange = ReportTable.ListColumns("Select").DataBodyRange
    
    Call AddMarlettBox(BoxRange)
    ReportTable.ListColumns("Select").DataBodyRange(1, 1).Font.Name = "Aptos Narrow" 'This can be anything except Marlett to prevent the cell from being checked
    
    'Autofit Category and Practice columns
    ReportTable.ListColumns("Category").Range.EntireColumn.AutoFit
    ReportTable.ListColumns("Practice").Range.EntireColumn.AutoFit

    'Return
    Set CreateReportTable = ReportTable

Footer:

End Function

Sub FormatTable(TargetSheet As Worksheet, NewTable As ListObject)
'Flag blanks and bad entries on Roster and Activity Sheets
'Different formatting on Report Page

    Dim CoverSheet As Worksheet
    Dim EthnicityRange As Range
    Dim GenderRange As Range
    Dim GradeRange As Range
    Dim CreditsRange As Range
    Dim MajorRange As Range
    Dim c As Range
    Dim FormulaString1 As String
    Dim FormulaString2 As String
    Dim IsCollegePrep As Boolean
    
    'If there are no rows, skip
    If CheckTable(TargetSheet) > 2 Then
        GoTo Footer
    End If

    'Check if this is for College Prep
    Set CoverSheet = Worksheets("Cover Page")
    
    IsCollegePrep = IsCollege

    'Blank cells flagged yellow, except in the first column
    With NewTable
        .DataBodyRange.FormatConditions.Delete
        .DataBodyRange.FormatConditions.Add Type:=xlBlanksCondition
        .ListColumns("Select").DataBodyRange.FormatConditions.Delete
        
        With .DataBodyRange.FormatConditions(1)
        .StopIfTrue = False
        .Interior.ColorIndex = 36
        End With
    End With

    'Validate demographics using tables on Ref Tables sheet
    Set EthnicityRange = NewTable.ListColumns("Ethnicity").DataBodyRange
    Set GenderRange = NewTable.ListColumns("Gender").DataBodyRange

    For Each c In EthnicityRange
        FormulaString1 = "=AND(COUNTIFS("
        FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
        c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "EthnicityList" & FormulaString2
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

    If IsCollegePrep = True Then
        'Grades work as both strings and numbers
        Set GradeRange = NewTable.ListColumns("Grade").DataBodyRange
        
        For Each c In GradeRange
            FormulaString1 = "=AND(COUNTIFS("
            FormulaString2 = "," & "Trim(" & c.Address & ")) < 1, NOT(ISBLANK(" & c.Address & ")))"
            c.FormatConditions.Add Type:=xlExpression, Formula1:=FormulaString1 & "GradeList" & FormulaString2
            With c.FormatConditions(2)
                .StopIfTrue = False
                .Interior.Color = vbRed
            End With
        Next c
    Else
        'Credit hours don't need the reference table, add majors to validation
        Set CreditsRange = NewTable.ListColumns("Credits").DataBodyRange
        Set MajorRange = NewTable.ListColumns("Major").DataBodyRange
        
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
    End If

Footer:

End Sub

Sub FormatReportTable(ReportSheet As Worksheet, ReportTable As ListObject)
'Adjust font, color, number formats for the table on the Report Page

    Dim i As Long
    Dim RedValue As String
    Dim GreenValue As String
    Dim BlueValue As String
    Dim TempArray() As String
    Dim ColorArray() As Variant

    'I'm not sure how to prevent the totals row from sorting, so for now I'll disable the autofilter
    ReportTable.ShowAutoFilterDropDown = False
    
    'Add color and vertical borders
    ColorArray = Application.Transpose(ActiveWorkbook.Names("ReportColumnRGBList").RefersToRange.Value)
    
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

    'Grey out the cell under "Select"
    TempArray = Split(ColorArray(2), ",")
        RedValue = TempArray(0)
        GreenValue = TempArray(1)
        BlueValue = TempArray(2)
    ReportTable.ListColumns("Select").Range.Resize(1, 1).Offset(1, 0).Interior.Color = RGB(RedValue, GreenValue, BlueValue)

    'Header and Total row, Total column text black, bold
    With ReportTable.HeaderRowRange.Resize(2, ReportTable.ListColumns.Count)
        .Font.Color = vbBlack
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    With ReportTable.ListColumns("Total").DataBodyRange
        .Font.Color = vbBlack
        .Font.Bold = True
    End With
    
    'Category, Pratice left aligned. Practice bold
    With ReportSheet.Range(ReportTable.ListColumns("Category").DataBodyRange, ReportTable.ListColumns("Practice").DataBodyRange)
        .HorizontalAlignment = xlLeft
    End With
    
    With ReportTable.ListColumns("Practice").DataBodyRange
        .Font.Bold = True
    End With
    
    '2nd row of both back to center aligned
    ReportTable.ListRows(1).Range.HorizontalAlignment = xlCenter

End Sub

Sub RemoveTable(TargetSheet As Worksheet)
'Unlists all table objects and removes formatting

    Dim DelTableRange As Range
    Dim DelTable As ListObject
    
    Call UnprotectSheet(TargetSheet)
    
    For Each DelTable In TargetSheet.ListObjects
        Set DelTableRange = DelTable.Range
        
        DelTable.Unlist
        DelTableRange.FormatConditions.Delete
        DelTableRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    Next DelTable

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

