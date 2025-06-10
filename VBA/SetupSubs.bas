Attribute VB_Name = "SetupSubs"
Option Explicit

Sub Tester()

'Call ChooseProgram("University Ref")
'Call ChooseProgram("College Ref")
Call ChooseProgram("Transfer Ref")

End Sub

Sub TesterClearTables()

    Dim RefSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim ClearTable As ListObject
    Dim btn As Button
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    
    On Error Resume Next
    
    Call UnprotectSheet(CoverSheet)
    With CoverSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    Call UnprotectSheet(RecordsSheet)
    With RecordsSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    Call UnprotectSheet(ReportSheet)
    With ReportSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    Call UnprotectSheet(RosterSheet)
    With RosterSheet
        .Cells.ClearContents
        .Cells.ClearFormats
        .Buttons.Delete
        .Columns.UseStandardWidth = True
    End With
    
    If Not Worksheets(1).Name = "University Ref" Then
        Worksheets(1).Name = "University Ref"
    ElseIf Not Worksheets(2).Name = "Transfer Ref" Then
        Worksheets(2).Name = "Transfer Ref"
    ElseIf Not Worksheets(3).Name = "College Ref" Then
        Worksheets(3).Name = "College Ref"
    End If
    
    Set RefSheet = Worksheets("University Ref")

    For Each ClearTable In RefSheet.ListObjects
        If Not ClearTable.Name = "UniversityTableGen" And Not ClearTable.Name = "UniversityRangeGen" Then
            ClearTable.Unlist
        End If
    Next ClearTable
    
    
    Set RefSheet = Worksheets("Transfer Ref")
        
    For Each ClearTable In RefSheet.ListObjects
        If Not ClearTable.Name = "TransferTableGen" And Not ClearTable.Name = "TransferRangeGen" Then
            ClearTable.Unlist
        End If
    Next ClearTable
    
    
    Set RefSheet = Worksheets("College Ref")
        
    For Each ClearTable In RefSheet.ListObjects
        If Not ClearTable.Name = "CollegeTableGen" And Not ClearTable.Name = "CollegeRangeGen" Then
            ClearTable.Unlist
        End If
    Next ClearTable

End Sub

Sub ChooseProgram(ProgramString As String)
'User selects the program from a dropdown list
'Set up table, ranges, and references specific to that program, then disable the ability to select

        Dim RefSheet As Worksheet
        Dim RecordsSheet As Worksheet
        Dim ReportSheet As Worksheet
        Dim RosterSheet As Worksheet
        Dim CoverSheet As Worksheet
        Dim sh As Worksheet
        Dim StartCell As Range
        Dim StopCell As Range
        Dim BotCell As Range
        Dim TableRange As Range
        Dim SearchRange As Range
        Dim CoverTitleRange As Range
        Dim CoverRefRange As Range
        Dim CoverCenterRange As Range
        Dim c As Range
        Dim TableString As String
        Dim HeaderArray() As Variant
        Dim TotalsArray() As Variant
        Dim TableGenTable As ListObject
        Dim RangeGenTable As ListObject
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        
        'Find the refence sheet for the selected program
        Set RefSheet = Worksheets(ProgramString)
        RefSheet.Name = "Ref Tables"
        
        If RefSheet Is Nothing Then 'This shouldn't happen
            GoTo Footer
        End If

        'Make and name reference tables. Each table has an empty column between it and the next
        'A table for table names and for range names/references already exist
        TableString = Left(ProgramString, InStr(ProgramString, " ") - 1) 'Snip off the " Ref" at the end
        
        With RefSheet
            Set TableGenTable = .ListObjects(TableString & "TableGen")
            Set RangeGenTable = .ListObjects(TableString & "RangeGen")
        
            'If ProgramString = "University Ref" Then
                'Set TableGenTable = .ListObjects("UniversityTableGen")
                'Set RangeGenTable = .ListObjects("UniversityRangeGen")
            'ElseIf ProgramString = "Transfer Ref" Then
                'Set TableGenTable = .ListObjects("TransferTableGen")
                'Set RangeGenTable = .ListObjects("TransferRangeGen")
            'ElseIf ProgramString = "College Ref" Then
                'Set TableGenTable = .ListObjects("CollegeTableGen")
                'Set RangeGenTable = .ListObjects("CollegeRangeGen")
            'End If
            
            Set SearchRange = TableGenTable.ListColumns("First Header").DataBodyRange
            
            'The TableGenTable as the names of each header in the 1st column. Find the header, first blank column after, and last row
            For Each c In SearchRange
                Set StartCell = .Range("1:1").Find(c.Value, , xlValues, xlWhole)
                If Not StartCell Is Nothing Then
                    'Define table range
                    Set StopCell = .Range(StartCell, Cells(1, Columns.Count).Address).Find("", , xlValues, xlWhole) 'This is a blank cell one past the last column
                    Set BotCell = StartCell.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
                    Set TableRange = StartCell.Resize(BotCell.row, StopCell.Column - StartCell.Column)
                    
                    'Make and name table
                    .ListObjects.Add(xlSrcRange, TableRange, , xlYes).Name = c.Offset(0, -1).Value 'Names of tables are stored one to the left
                End If
            Next c
            
            'Define named ranges
            Set SearchRange = RangeGenTable.ListColumns("Range Name").DataBodyRange
            
            For Each c In SearchRange
                ThisWorkbook.Names.Add Name:=c.Value, RefersTo:=.Range("=" & c.Offset(0, 1).Value)
            Next c

        End With

    'Populate the Cover Page
    Set CoverSheet = Worksheets("Cover Page")
    
    Call UnprotectSheet(CoverSheet)
    Call CoverSheetText(RefSheet, CoverSheet, ProgramString)
    Call CoverSheetButtons(ProgramString)

    'Make table on Report Page and add buttons
    Set ReportSheet = Worksheets("Report Page")
    Set c = ReportSheet.Range("A6") 'Where the headers begin

    Call UnprotectSheet(ReportSheet)
    c.Value = "Select" 'This is what the following sub looks for
    Call CreateReportTable
    Call ReportSheetButtons
    
    'Put in headers for Roster Page and add buttons. Making the table will happen when it's parsed
    Set RosterSheet = Worksheets("Roster Page")
    Set c = RosterSheet.Range("A6") 'Where the headers begin
    
    Call UnprotectSheet(RosterSheet)
    HeaderArray = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
    Call ResetTableHeaders(RosterSheet, c, HeaderArray)
    Call RosterSheetButtons
    
    'Populate the RecordsSheet
    Set RecordsSheet = Worksheets("Records Page")
    
    Call UnprotectSheet(RecordsSheet)
    Call RecordsSheetText
    
    'Make sure the workbook can be edited
    Call ResetProtection
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverSheetText(RefSheet As Worksheet, CoverSheet As Worksheet, ProgramString As String)
'Text, formatting, tables for CoverSheet

    Dim TextRange As Range
    Dim DateRange As Range
    Dim CenterRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim i As Long
    Dim BookTitle As String
    Dim BookVersion As String
    Dim TextString As String
    Dim TextArray() As String
    Dim TableString As String
    Dim TableNameArray() As Variant
    Dim TempTable As ListObject
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Unprotect. This shouldn't ever be needed
    Call UnprotectSheet(CoverSheet)
    
    'Define the title and edition
    Select Case ProgramString
        Case "University Ref"
            BookTitle = "MESA University Weekly Report"
            
        Case "Transfer Ref"
            BookTitle = "Transfer Prep Weekly Report"
            
        Case "College Ref"
            BookTitle = "College Prep Weekly Report"
    End Select
    
    'If ProgramString = "University Ref" Then
        'BookTitle = "MESA University Weekly Report"
    'ElseIf ProgramString = "Transfer Ref" Then
        'BookTitle = "Transfer Prep Weekly Report"
    'ElseIf ProgramString = "College Ref" Then
        'BookTitle = "College Prep Weekly Report"
    'End If
    
    BookVersion = GetVersion()

    'Insert text
    With CoverSheet
        Set TextRange = .Range("A1:A5")
        
        TextString = BookTitle & ";" & "Version " & BookVersion & ";Name;Date;Center"
        TextArray = Split(TextString, ";")
        TextRange.Value = Application.Transpose(TextArray)
    
        'Date validation and a dropdown menu for the center
        Set DateRange = .Range("A:A").Find("Date", , xlValues, xlWhole)
        Set CenterRange = .Range("A5").Find("Center", , xlValues, xlWhole)

        Call DateValidation(CoverSheet, DateRange.Offset(0, 1))
        Call CenterDropdown(CoverSheet, CenterRange.Offset(0, 1))
    End With
    
    'Add formatting. No lines under the first two rows
    i = 1
    For Each c In TextRange
        c.Font.Bold = True
        
        If i > 2 Then
            c.HorizontalAlignment = xlRight
            Set c = Union(c, c.Offset(0, 1))
        
            c.Borders(xlEdgeBottom).LineStyle = xlContinuous
            c.Borders(xlEdgeBottom).Weight = xlMedium
        Else
            Set c = Union(c, c.Offset(0, 1))
        End If
       
        c.WrapText = False
        
        i = i + 1
    Next c
    
    'Add reference tables
    Set c = CoverSheet.Range("H1")
    
    If ProgramString = "College Ref" Then
        ReDim TableNameArray(1 To 3)
        
        TableNameArray(1) = "EthnicityTable"
        TableNameArray(2) = "GenderTable"
        TableNameArray(3) = "GradeTable"
    Else
        ReDim TableNameArray(1 To 5)
        
        TableNameArray(1) = "EthnicityTable"
        TableNameArray(2) = "GenderTable"
        TableNameArray(3) = "MajorTable"
        TableNameArray(4) = "FirstGenerationTable"
        TableNameArray(5) = "LowIncomeTable"
    End If
    
    For i = 1 To UBound(TableNameArray)
        TableString = TableNameArray(i)
        
        Set TempTable = RefSheet.ListObjects(TableString)
        Set CopyRange = TempTable.Range
        Set PasteRange = c.Resize(TempTable.Range.Rows.Count, 1).Offset(0, i - 1)
        
        PasteRange.Value(11) = CopyRange.Value(11)
        PasteRange.HorizontalAlignment = xlLeft
        PasteRange.BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    Next i
    
    'Autofitting
    Set PasteRange = Range(c, c.Offset(0, i - 1)).EntireColumn
    
    PasteRange.Columns.AutoFit
    
    
    
End Sub

Sub CoverSheetButtons(ProgramString)
'Called when the program is chosen

    Dim CoverSheet As Worksheet
    Dim NewButton As Button
    Dim NewButtonRange As Range
      
    Set CoverSheet = Worksheets("Cover Page")
  
    'Submit button
    Set NewButtonRange = CoverSheet.Range("D1:F2")
    Set NewButton = CoverSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "CoverSharePointButton"
        .Caption = "Submit to SharePoint"
        .Name = "CoverSharePointExportButton"
    End With
        
    'Save button
    Set NewButtonRange = CoverSheet.Range("D4:F5")
    Set NewButton = CoverSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "CoverLocalSaveButton"
        .Caption = "Save a Copy"
        .Name = "CoverSaveCopyButton"
    End With
        
    'Import button
    'Set NewButtonRange = CoverSheet.Range("L1:M2")
    
    'Nudge for extra columns
    'If ProgramString <> "College Ref" Then
        'Set NewButtonRange = NewButtonRange.Offset(0, 2)
    'End If
    
    'Set NewButton = CoverSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    'With NewButton
        '.OnAction = "CoverImportButton"
        '.Caption = "Import Records"
        '.Name = "CoverImportButton"
    'End With
        
End Sub

Sub RosterSheetButtons()
'Called when the program is chosen

    Dim RosterSheet As Worksheet
    Dim NewButton As Button
    Dim NewButtonRange As Range
    
    Set RosterSheet = Worksheets("Roster Page")
    
    'Select All
    Set NewButtonRange = RosterSheet.Range("A5:B5")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SelectAllButton"
        .Caption = "Select All"
        .Name = "RosterSelectAllButton"
    End With

    'Delete Row
    Set NewButtonRange = RosterSheet.Range("D5:E5")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelectedButton"
        .Caption = "Delete Row"
        .Name = "RosterRemoveSelectedButton"
    End With
    
    'Select activity
    Set NewButtonRange = RosterSheet.Range("G4:H5")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "OpenNewActivityButton"
        .Caption = "New Activity"
        .Name = "RosterNewActivityButton"
    End With
    
    'Load activity
    Set NewButtonRange = RosterSheet.Range("G2:H2")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "OpenLoadActivityButton"
        .Caption = "Load Activity"
        .Name = "RosterLoadActivityButton"
    End With
    
    'Add students
    Set NewButtonRange = RosterSheet.Range("G1:H1")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "OpenAddStudentsButton"
        .Caption = "Add to Activity"
        .Name = "RosterAddSelectedButton"
    End With
    
    'Read roster
    Set NewButtonRange = RosterSheet.Range("A1:B2")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RosterParseButton"
        .Caption = "Parse Roster"
        .Name = "RosterParseButton"
    End With
    
    'Clear roster
    Set NewButtonRange = RosterSheet.Range("D1:E1")
    Set NewButton = RosterSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RosterClearButton"
        .Caption = "Clear Roster"
        .Name = "RosterClearButton"
    End With

End Sub

Sub ReportSheetButtons()
'Called when the program is chosen

    Dim ReportSheet As Worksheet
    Dim NewButton As Button
    Dim NewButtonRange As Range
    
    Set ReportSheet = Worksheets("Report Page")
    
    'Select All
    Set NewButtonRange = ReportSheet.Range("A5:B5")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "SelectAllButton"
        .Caption = "Select All"
        .Name = "ReportSelectAllButton"
    End With
    
    'Pull Totals
    Set NewButtonRange = ReportSheet.Range("A1:B2")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ReportTabulateTotalsButton"
        .Caption = "Tabulate Totals"
        .Name = "ReportTabTotalsButton"
    End With
    
    'Clear the report
    Set NewButtonRange = ReportSheet.Range("D1:E2")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "ReportClearAllButton"
        .Caption = "Clear Report"
        .Name = "ReportClearButton"
    End With
    
    'Tabulate activities
    Set NewButtonRange = ReportSheet.Range("C1:C2")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "OpenTabulateActivityButton"
        .Caption = "Tabulate Activities"
        .Name = "ReportTabActivitiesButton"
    End With
    
    'Remove row
    Set NewButtonRange = ReportSheet.Range("D4:E5")
    Set NewButton = ReportSheet.Buttons.Add(NewButtonRange.Left, NewButtonRange.Top, _
        NewButtonRange.Width, NewButtonRange.Height)
    
    With NewButton
        .OnAction = "RemoveSelectedButton"
        .Caption = "Delete Row"
        .Name = "ReportRemoveSelectedButton"
    End With

End Sub

Sub RecordsSheetText()
'Put in the corresponding activities for the program
'Make this programatic in the future
    
    Dim RecordsSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim HeaderArray() As Variant
    Dim ActivityArray() As Variant
    
    Set RecordsSheet = Worksheets("Records Page")
    
    'Headers for students and activities
    ReDim HeaderArray(1 To 6, 1 To 2)
        HeaderArray(1, 1) = "A3"
        HeaderArray(2, 1) = "A4"
        HeaderArray(3, 1) = "B3"
        HeaderArray(4, 1) = "C1"
        HeaderArray(5, 1) = "C2"
        HeaderArray(6, 1) = "D1"
        
        HeaderArray(1, 2) = "First"
        HeaderArray(2, 2) = "H BREAK"
        HeaderArray(3, 2) = "Last"
        HeaderArray(4, 2) = "Activity"
        HeaderArray(5, 2) = "Notes"
        HeaderArray(6, 2) = "V BREAK"
    
    For i = LBound(HeaderArray) To UBound(HeaderArray)
        Set c = RecordsSheet.Range(HeaderArray(i, 1))
        
        c.Value = HeaderArray(i, 2)
    Next i
    
    'Grab the list of activities and insert them one cell right of the "V BREAK" padding cell
     ActivityArray = Application.Transpose(ActiveWorkbook.Names("ActivitiesList").RefersToRange.Value)
     
     Set c = RecordsSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole).Offset(0, 1)
     
     Call ResetTableHeaders(RecordsSheet, c, ActivityArray)
    
Footer:
    
End Sub
