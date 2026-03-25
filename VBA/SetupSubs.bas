Attribute VB_Name = "SetupSubs"
Option Explicit

Sub Tester()

'Call ChooseProgram("University Ref")
Call ChooseProgram("College Ref")
'Call ChooseProgram("Transfer Ref")

End Sub

Sub dimtbltest()

    Dim DirectorySheet As Worksheet
    Dim ProgramString As String
    
    Set DirectorySheet = Worksheets("Narrative Page")
    'ProgramString = "College"
    'ProgramString = "Transfer"
    ProgramString = "University"
    
    'Call SetupNarrativeTables(ProgramString)
    'Call SetupDirectoryTables(ProgramString)
    Call SetupOtherTables(ProgramString)

End Sub

Sub listttest()

    Dim ProgramString As String
    
    ProgramString = "College"
    
    Call SetupRanges(ProgramString)

End Sub

Sub trimtest()

    Dim TempString As String
    Dim TrimString As String
    
    TempString = "FirstGenerationTable[First Generation]"
    TrimString = Left(TempString, InStr(1, TempString, "[") - 1)
    Debug.Print TrimString

End Sub

Sub newlisttest()

    Dim RefSheet As Worksheet
    Dim TestSheet As Worksheet
    Dim ListRange As Range
    Dim c As Range
    Dim i As Long
    Dim ListArray As Variant
    
    Set RefSheet = Worksheets("Ref Tables")
    Set TestSheet = Worksheets("Test")
    
    'Set ListRange = Range("RosterHeadersList")
    
    For Each c In ListRange
        Debug.Print c.Value
    Next c

End Sub

Sub SetupCoverButtons(ProgramString)
'Called when the program is chosen

    Dim CoverSheet As Worksheet
    Dim i As Long
    Dim j As Long
    Dim ButtonArray As Variant
    Dim TempArray As Variant
      
    Set CoverSheet = Worksheets("Cover Page")
    
    'Deleting Select Program button, if it's still there
    CoverSheet.Buttons.Delete
  
    ReDim ButtonArray(1 To 1)
    i = 1
  
    'SharePoint
    ReDim TempArray(1 To 4)
        TempArray(1) = "A7:C8"
        TempArray(2) = "CoverSharePointExportButton"
        TempArray(3) = "Submit to SharePoint"
        TempArray(4) = "ButtonCoverSharePointExport"
        
    ReDim ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
    
    'Local save
    ReDim TempArray(1 To 4)
        TempArray(1) = "A10:C11"
        TempArray(2) = "CoverSaveCopyButton"
        TempArray(3) = "Save a Copy"
        TempArray(4) = "ButtonCoverSaveCopy"

    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    For i = 1 To UBound(ButtonArray)
        TempArray = ButtonArray(i)
        Call MakeButton(CoverSheet, TempArray)
    Next i
        
End Sub

Sub SetupCoverInitialize()
'Puts a button to choose the program

    Dim CoverSheet As Worksheet
    Dim ButtonArray As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    
    ReDim ButtonArray(1 To 4)
        ButtonArray(1) = "A1:C3"
        ButtonArray(2) = "CoverChooseProgramButton"
        ButtonArray(3) = "Choose Program"
        ButtonArray(4) = "ButtonCoverChooseProgram"
    
    Call MakeButton(CoverSheet, ButtonArray)
    
End Sub

Sub SetupCoverText(RefSheet As Worksheet, CoverSheet As Worksheet, ProgramString As String)
'Text, formatting, tables for CoverSheet

    Dim TextRange As Range
    Dim RefRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim AddressString As String
    Dim BookTitle As String
    Dim BookVersion As String
    Dim TableName As String
    Dim PasteTable As ListObject
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Unprotect. This shouldn't ever be needed
    Call UnprotectSheet(CoverSheet)
    
    'Define the title and edition
    Select Case ProgramString
        Case "University"
            BookTitle = "MESA University Term Report"
            
        Case "Transfer"
            BookTitle = "Transfer Prep Term Report"
            
        Case "College"
            BookTitle = "College Prep Term Report"
    End Select
    
    BookVersion = GetEdition()

    'Insert text
    Set TextRange = Range("CoverTextList")
    
    i = 1
    For Each c In TextRange 'The first two headers will be replaced
        AddressString = c.Offset(0, 1).Value 'One to the right
    
        Set PasteRange = CoverSheet.Range(AddressString)
        
        'Put in header
        Select Case c.Value
            Case "Title"
                PasteRange.Value = BookTitle
            
            Case "Version"
                PasteRange.Value = BookVersion
            
            Case Else
                PasteRange.Value = c.Value
                
                 If c.Value = "Date" Then
                    Call DateValidation(CoverSheet, c.Offset(0, 1))
                ElseIf c.Value = "Center" Then
                    Call CenterDropdown(CoverSheet, c.Offset(0, 1))
                End If
        End Select
        
        'Bold and underline
        Set d = PasteRange.Resize(1, 2)
        PasteRange.Font.Bold = True
        d.WrapText = False
        
        'No underline for the first two rows
        If c.Value = "Title" Or c.Value = "Version" Then
            GoTo NextRow
        End If
        
        PasteRange.HorizontalAlignment = xlRight
        d.Borders(xlEdgeBottom).LineStyle = xlContinuous
        d.Borders(xlEdgeBottom).Weight = xlMedium
NextRow:
    Next c
    
    'Add reference tables
    Set d = CoverSheet.Range("H1")
    Set RefRange = Range("CoverReferenceList")
    
    i = 0
    For Each c In RefRange
        TableName = c.Value
        
        Set CopyRange = RefSheet.ListObjects(TableName).Range 'None of these have helper columns
        Set PasteRange = d.Resize(CopyRange.Rows.Count, 1).Offset(0, i)
        
        PasteRange.Value(11) = CopyRange.Value(11)
        Set PasteTable = CoverSheet.ListObjects.Add(xlSrcRange, PasteRange, , xlYes)
        PasteRange.HorizontalAlignment = xlLeft
        PasteRange.BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        
        i = i + 1
    Next c
    
    'Autofitting
    Set PasteRange = Range(d, d.Offset(0, i)).EntireColumn
    
    PasteRange.Columns.AutoFit
    
End Sub

Sub SetupDirectoryButtons(ProgramString As String)
'Puts in a button to tabulate schools

    Dim DirectorySheet As Worksheet
    Dim ButtonArray As Variant
    
    If ProgramString <> "College" Then
        GoTo Footer
    End If
    
    Set DirectorySheet = Worksheets("Directory Page")
    
    ReDim ButtonArray(1 To 4)
        ButtonArray(1) = "Q1:R1"
        ButtonArray(2) = "DirectoryTabulateSchoolsButton"
        ButtonArray(3) = "Tabulate Schools"
        ButtonArray(4) = "ButtonDirectoryTabulateSchools"
    
    Call MakeButton(DirectorySheet, ButtonArray)

Footer:

End Sub

Sub SetupDirectoryTables(ProgramString As String)
'College Prep gets the Staff, School, and Teacher tables
'Transfer Prep gets the Staff and Mentors table
    
    Dim DirectorySheet As Worksheet
    Dim RefSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim PasteOffset As Long
    Dim i As Long
    Dim TableName As String
    Dim TableNameArray() As Variant

    Set DirectorySheet = Worksheets("Directory Page")
    Set RefSheet = Worksheets("Ref Tables")
    Set PasteRange = DirectorySheet.Range("A1")
 
    Call UnprotectSheet(DirectorySheet)
    
    'Create an array with the tables we need
    If ProgramString = "College" Then
        ReDim TableNameArray(1 To 3)
            TableNameArray(1) = "DirectoryTable"
            TableNameArray(2) = "TeachersTable"
            TableNameArray(3) = "SchoolsTable"
    Else
        ReDim TableNameArray(1) '(1 To 2)
            TableNameArray(1) = "DirectoryTable"
            'TableNameArray(2) = "MentorsTable"
    End If
    
    'Loop through, offsetting the PasteRange to put a column between each table
    PasteOffset = 0
    
    For i = 1 To UBound(TableNameArray)
        TableName = TableNameArray(i)
    
        If Not Len(TableName) > 0 Then 'If there's a blank or bad table name
            'GoTo NextTable
        End If
        
        Set c = SetupResponseTables(DirectorySheet, PasteRange.Offset(0, PasteOffset), TableName)
        
        If c Is Nothing Then
            'GoTo NextTable
        End If
        
        PasteOffset = PasteOffset + c.Columns.Count + 1
        
NextTable:
    Next i
    
    'Put "Director" in for all programs
    PasteRange.Offset(3, 0).Value = "Director"
    
    'Insert "RA" and "Faculty Sponsor" for Transfer Prep and MESA University
    If ProgramString = "Transfer" Or ProgramString = "University" Then 'Make this programmatic in the future
        PasteRange.Offset(4, 0).Value = "RA"
        PasteRange.Offset(5, 0).Value = "Faculty Sponsor"
    End If

End Sub

Sub SetupPopulateWorkbook(ProgramString As String)
'User selects the program from a dropdown list
'Set up table, ranges, and references specific to that program, then disable the ability to select

    Dim RefSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim OtherSheet As Worksheet

    'Application.ScreenUpdating = False
    'Application.DisplayAlerts = False
    'Application.EnableEvents = False

    'Reference tables for all three programs are on the same sheet
    Set RefSheet = Worksheets("Ref Tables")
    Set CoverSheet = Worksheets("Cover Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set OtherSheet = Worksheets("Other Page")
    
    'Create reference tables and named ranges, passing the selected program
    Call SetupReferenceTables
    Call SetupRanges(ProgramString)
    
    'Populate the Cover Page
    Call UnprotectSheet(CoverSheet)
    Call SetupCoverText(RefSheet, CoverSheet, ProgramString)
    Call SetupCoverButtons(ProgramString)

    'Report Page
    Call UnprotectSheet(ReportSheet)
    Call MakeReportTable
    Call SetupReportButtons
    
    'Roster Page
    Call UnprotectSheet(RosterSheet)
    Call MakeRosterTable(RosterSheet)
    Call SetupRosterButtons
    
    'Records Page
    Call UnprotectSheet(RecordsSheet)
    Call SetupRecordsText
    
    'Directory, Narrative, and Other Pages
    Call SetupDirectoryTables(ProgramString)
    Call SetupDirectoryButtons(ProgramString)
    Call SetupNarrativeTables(ProgramString)
    Call SetupOtherTables(ProgramString) 'No difference between programs

    'Break external links
    Call BreakExternalLinks

    'Make sure the workbook can be edited
    Call ResetProtection
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SetupRanges(ProgramString As String)
'Goes through the RangeGenTable and dynamically creates named lists using helper columns
'The passed string determines which helper column is used

    Dim RefSheet As Worksheet
    Dim RangeNameRange As Range
    Dim SearchRange As Range
    Dim IncludeRange As Range
    Dim TempRange As Range
    Dim c As Range
    Dim d As Range
    Dim RangeName As String
    Dim RangeAddress As String
    Dim TableName As String
    Dim RangeGenTable As ListObject
    Dim TempTable As ListObject

    Set RefSheet = Worksheets("Ref Tables")
        If RefSheet Is Nothing Then 'This shouldn't happen
            GoTo Footer
        End If
    
    Set RangeGenTable = RefSheet.ListObjects("RangeGenTable")
    Set RangeNameRange = RangeGenTable.ListColumns("Range Name").DataBodyRange
       
    'Loop through the names of ranges and grab the table column they refer to
    For Each c In RangeNameRange
        'Grab the name we'll give the named range, its address reference, and the name of the table the address lives in
        Set d = RangeGenTable.HeaderRowRange.Find("Range Ref", , xlValues, xlWhole)
        
        RangeName = c.Value
        RangeAddress = RefSheet.Cells(c.Row, d.Column)
        TableName = Left(RangeAddress, InStr(1, RangeAddress, "[") - 1)
        
        'Grab the value in the "Filter" column, a "1" tells us that only part of the reference table will be used
        Set d = RangeGenTable.HeaderRowRange.Find("Filter", , xlValues, xlWhole)
            If RefSheet.Cells(c.Row, d.Column) = 0 Then
                'Skip filtering and use the whole column
                GoTo DefineList
            End If

        'If we're filtering, go through and build a range using the helper columns. Each program has its own column
        Set TempRange = RefSheet.Range(RangeAddress) '("=" & RangeAddress)
        Set TempTable = RefSheet.ListObjects(TableName)
        Set SearchRange = TempTable.ListColumns(ProgramString).DataBodyRange
        
        'Loop through the helper column
        For Each d In SearchRange
            If d.Value <> 1 Then
                GoTo NextRow
            End If
            
            Set IncludeRange = BuildRange(RefSheet.Cells(d.Row, TempRange.Column), IncludeRange)
NextRow:
        Next d
        
        RangeAddress = IncludeRange.Address
        Set IncludeRange = Nothing
        
DefineList:
        'If filter is False, assign the entire column in the reference table
        ThisWorkbook.Names.Add Name:=c.Value, RefersTo:=RefSheet.Range(RangeAddress)
        
NextList:
    Next c
    
Footer:

End Sub

Sub SetupRemoveTables()

    Dim RefSheet As Worksheet
    Dim i As Long
    Dim ClearTable As ListObject
    
    Set RefSheet = Worksheets("Ref Tables")

    'Remove all but the first two tables
    For Each ClearTable In RefSheet.ListObjects
        If Not ClearTable.Name = "TableGenTable" And Not ClearTable.Name = "RangeGenTable" Then
            ClearTable.Unlist
        End If
    Next ClearTable

    'Remove all named ranges
    Dim NameString
    
    For i = ThisWorkbook.Names.Count To 1 Step -1
     
        NameString = ThisWorkbook.Names(i)

        If InStr(NameString, "=#NAME?") > 0 Then 'For a null we're getting for some reason
            GoTo NextName
        End If

        ThisWorkbook.Names(i).Delete
NextName:
    Next i

End Sub

Sub SetupNarrativeTables(ProgramString As String)
'Tables for highlights, goals, educatorPD, parent development

    Dim NarrativeSheet As Worksheet
    Dim RefSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim PasteOffset As Long
    Dim SubtextString As String
    Dim TableName As String
    Dim TableNameArray() As Variant

    Set NarrativeSheet = Worksheets("Narrative Page")
    Set RefSheet = Worksheets("Ref Tables")
    Set PasteRange = NarrativeSheet.Range("A1")
 
    Call UnprotectSheet(NarrativeSheet)

    'Create an array with the tables we need
    ReDim TableNameArray(1 To 3)
        TableNameArray(1) = "HighlightTable"
        TableNameArray(2) = "GoalsTable"
        TableNameArray(3) = "EducatorPDTable"
    
    If ProgramString = "College" Then
        ReDim Preserve TableNameArray(1 To 4)
        TableNameArray(4) = "ParentDevelopmentTable"
    End If
    
    'Loop through, offsetting the PasteRange to put a column between each table
    PasteOffset = 0
    
    For i = 1 To UBound(TableNameArray)
        TableName = TableNameArray(i)
    
        If Not Len(TableName) > 0 Then 'If there's a blank or bad table name
            'GoTo NextTable
        End If
        
        Set c = SetupResponseTables(NarrativeSheet, PasteRange.Offset(0, PasteOffset), TableName)
        
        If c Is Nothing Then
            'GoTo NextTable
        End If
        
        'College Prep has an extra column in the "Highlights" table. Remove for the other programs
        If ProgramString <> "College" Then
            If TableName = "HighlightTable" Then
                NarrativeSheet.ListObjects(1).ListColumns(1).Delete
                
                Set d = c.Resize(1, 1).Offset(1, 0) 'For resizing the columns
         
                SubtextString = d.Value
                d.ClearContents
                d.EntireColumn.AutoFit
                d.Value = SubtextString
                
                PasteOffset = PasteOffset + c.Columns.Count
                GoTo NextTable
            End If
        End If
        
        PasteOffset = PasteOffset + c.Columns.Count + 1
NextTable:
    Next i
    
End Sub

Sub SetupOtherTables(ProgramString As String)
'Catch-all for anything that didn't belong in the Report or Narrative
'Just the headers and blank rows

    Dim OtherSheet As Worksheet
    Dim RefSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim PasteOffset As Long
    Dim i As Long
    Dim TableName As String
    Dim TableNameArray() As Variant
    
    Set OtherSheet = Worksheets("Other Page")
    Set RefSheet = Worksheets("Ref Tables")
    Set PasteRange = OtherSheet.Range("A1")
    
    Call UnprotectSheet(OtherSheet)

    'Create an array with the tables we need. It's just one in this case
    ReDim TableNameArray(1)
        TableNameArray(1) = "OtherSheetTable"

    'Loop through, offsetting the PasteRange to put a column between each table
    PasteOffset = 0
    
    For i = 1 To UBound(TableNameArray)
        TableName = TableNameArray(i)
    
        If Not Len(TableName) > 0 Then 'If there's a blank or bad table name
            'GoTo NextTable
        End If
        
        Set c = SetupResponseTables(OtherSheet, PasteRange.Offset(0, PasteOffset), TableName)
        
        If c Is Nothing Then
            'GoTo NextTable
        End If
        
        PasteOffset = PasteOffset + c.Columns.Count + 1
        
NextTable:
    Next i


End Sub

Sub SetupRecordsText()
'Put in the corresponding activities for the program
'Make this programatic in the future
    
    Dim RecordsSheet As Worksheet
    Dim RefRange As Range
    Dim c As Range
    Dim AddressString As String
    
    Set RecordsSheet = Worksheets("Records Page")
    Set RefRange = Range("RecordsHeadersList")
    
    'Paste in the headers. In this case, we only have students, no activities
    For Each c In RefRange
        AddressString = c.Offset(0, 1).Value
        RecordsSheet.Range(AddressString).Value = c.Value
    Next c
    
Footer:
    
End Sub

Sub SetupReportButtons()
'Called when the program is chosen

    Dim ReportSheet As Worksheet
    Dim i As Long
    Dim TempArray As Variant
    Dim ButtonArray As Variant
    
    Set ReportSheet = Worksheets("Report Page")
    
    ReDim ButtonArray(1 To 1)
    i = 1
    
    'Pull Totals
    ReDim TempArray(1 To 4)
        TempArray(1) = "A1:B2"
        TempArray(2) = "ReportTabulateTotalsButton"
        TempArray(3) = "Tabulate Totals"
        TempArray(4) = "ButtonReportTabTotals"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
    
    'Clear the Report
    ReDim TempArray(1 To 4)
        TempArray(1) = "D1:E2"
        TempArray(2) = "ReportClearTotalsButton"
        TempArray(3) = "Clear Report"
        TempArray(4) = "ButtonReportClearTotals"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    For i = 1 To UBound(ButtonArray)
        TempArray = ButtonArray(i)
        Call MakeButton(ReportSheet, TempArray)
    Next i

Footer:

End Sub

Sub SetupResetWorkbook()

    Dim RefSheet As Worksheet
    Dim CoverSheet As Worksheet
    Dim TempSheet As Worksheet
    Dim i As Long
    Dim SheetArray As Variant
    Dim ButtonArray As Variant
    
    Set RefSheet = Worksheets("Ref Tables")
    Set CoverSheet = Worksheets("Cover Page")
    
    'Wipe all sheets except the RefSheet
    SheetArray = Split("Cover Page, Records Page, Report Page, Roster Page, Directory Page, Narrative Page, Other Page", ",")
    ReDim Preserve SheetArray(1 To UBound(SheetArray) + 1) 'Make base 1
    
    For i = 1 To UBound(SheetArray)
        Set TempSheet = Worksheets(Trim(SheetArray(i)))
        Call WipeSheet(TempSheet)
        Set TempSheet = Nothing
    Next i

    'Remove tables and named ranges from the RefSheet
    Call SetupRemoveTables

    'Put the Choose Program button back on the Cover Sheet
    ReDim ButtonArray(1 To 4)
        ButtonArray(1) = "A1:C3"
        ButtonArray(2) = "CoverChooseProgramButton"
        ButtonArray(3) = "Choose Program"
        ButtonArray(4) = "ButtonCoverChooseProgram"
    
    Call MakeButton(CoverSheet, ButtonArray)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SetupRosterButtons()
'Called when the program is chosen

    Dim RosterSheet As Worksheet
    Dim i As Long
    Dim ButtonArray As Variant
    Dim TempArray As Variant

    Set RosterSheet = Worksheets("Roster Page")
    
    ReDim ButtonArray(1 To 1)
    i = 1
    
    'Select All
    ReDim TempArray(1 To 4)
        TempArray(1) = "A5:B5"
        TempArray(2) = "SelectAllButton"
        TempArray(3) = "Select All"
        TempArray(4) = "ButtonRosterSelectAll"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1
    
    'Delete Row
    ReDim TempArray(1 To 4)
        TempArray(1) = "D5:E5"
        TempArray(2) = "RemoveSelectedButton"
        TempArray(3) = "Delete Row"
        TempArray(4) = "ButtonRosterRemoveSelected"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    'Parse Roster
    ReDim TempArray(1 To 4)
        TempArray(1) = "A1:B2"
        TempArray(2) = "RosterParseButton"
        TempArray(3) = "Parse Roster"
        TempArray(4) = "ButtonRosterParse"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    'Clear Roster
    ReDim TempArray(1 To 4)
        TempArray(1) = "D1:E1"
        TempArray(2) = "RosterClearButton"
        TempArray(3) = "Clear Roster"
        TempArray(4) = "ButtonRosterClear"
        
    ReDim Preserve ButtonArray(1 To i)
    ButtonArray(i) = TempArray
    i = i + 1

    For i = 1 To UBound(ButtonArray)
        TempArray = ButtonArray(i)
        Call MakeButton(RosterSheet, TempArray)
    Next i
    
Footer:

End Sub

Sub SetupReferenceTables()
'Searches through two established tables and generates the rest on the RefSheet
'This is a little redundant, except for culling old tables when the workbook is reset

    Dim RefSheet As Worksheet
    Dim StartCell As Range
    Dim StopCell As Range
    Dim BotCell As Range
    Dim SearchRange As Range
    Dim TableRange As Range
    Dim c As Range
    Dim TableGenTable As ListObject
    
    Set RefSheet = Worksheets("Ref Tables")
        If RefSheet Is Nothing Then 'This shouldn't happen
            GoTo Footer
        End If
    
    'Make and name reference tables. Each table has an empty column between it and the next
    'A table for table names and for range names/references already exist
    Set TableGenTable = RefSheet.ListObjects("TableGenTable")
    Set SearchRange = TableGenTable.ListColumns("First Header").DataBodyRange
    
    With RefSheet
        'The TableGenTable has the names of each header in the 1st column. Find the header, first blank column after, and last row
        For Each c In SearchRange
            Set StartCell = .Range("1:1").Find(c.Value, , xlValues, xlWhole)
            
            If StartCell Is Nothing Then
                GoTo NextTable
            End If
                             
            'Define table range
            Set StopCell = .Range(StartCell, Cells(1, Columns.Count).Address).Find("", , xlValues, xlWhole) 'This is a blank cell one past the last column
            Set BotCell = StartCell.EntireColumn.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            Set TableRange = StartCell.Resize(BotCell.Row, StopCell.Column - StartCell.Column)
            
            'Make and name table
            .ListObjects.Add(xlSrcRange, TableRange, , xlYes).Name = c.Offset(0, -1).Value 'Names of tables are stored one to the left
        Next c
NextTable:
    End With

Footer:

End Sub

Function SetupResponseTables(TargetSheet As Worksheet, PasteRange As Range, TableName As String) As Range
'Takes a reference table on the RefSheet and transforms it at the passed range
'The first two rows of the reference table are a title and subtext
'The rest of the rows are table columns
'Puts five empty rows below
'Returns the range of the created table up to Row 1
'Returns nothing on error

    Dim RefSheet As Worksheet
    Dim CopyRange As Range
    Dim HeaderRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim SubtextString As String
    Dim TargetTable As ListObject
    Dim CopyArray As Variant
    
    Set RefSheet = Worksheets("Ref Tables")
    Set CopyRange = RefSheet.ListObjects(TableName).Range
        If CopyRange Is Nothing Then
            GoTo Footer
        ElseIf Not CopyRange.Cells.Count > 2 Then 'Need at least one column
            GoTo Footer
        End If
    
    'Read into an array
    ReDim CopyArray(1 To CopyRange.Cells.Count)
    
    i = 1
    For Each c In CopyRange
        CopyArray(i) = c.Value
        
        i = i + 1
    Next c
    
    'The first two rows are a title and subtext
    PasteRange.Value = CopyArray(1)
    PasteRange.Offset(1, 0).Value = CopyArray(2)
    
    'The rest are columns
    j = 0
    For i = 3 To UBound(CopyArray)
        PasteRange.Offset(2, j).Value = CopyArray(i)
    
        j = j + 1
    Next i

    'Make a table object with five extra rows
    Set HeaderRange = TargetSheet.Range(PasteRange.Offset(2, 0), PasteRange.Offset(2, j - 1))
    Set TargetTable = TargetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=HeaderRange, xlListObjectHasHeaders:=xlYes)
    
    For i = 1 To 5
        TargetTable.ListRows.Add
    Next i
    
    'Format the title
    With PasteRange
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        '.Columns.AutoFit
    End With
    
    'Autofit based on the title and 1st table column, but not the subtext
    Set c = PasteRange.Offset(1, 0)
    
    SubtextString = c.Value
    c.ClearContents
    c.EntireColumn.AutoFit
    c.Value = SubtextString
    
    'If there's a "Date" column, format
    Set c = TargetTable.HeaderRowRange.Find("Date", , xlValues, xlWhole)
    
    If Not c Is Nothing Then
        TargetTable.ListColumns("Date").DataBodyRange.NumberFormat = "mm/dd/yyyy"
    End If
     
    
    'Return the full used range
    Set c = TargetTable.Range
    Set ReturnRange = TargetSheet.Range(PasteRange, c)
        If ReturnRange Is Nothing Then
            GoTo Footer
        End If
        
    Set SetupResponseTables = ReturnRange
    
Footer:
    
End Function

Sub responsetabletest()

    Dim TestSheet As Worksheet
    Dim PasteRange As Range
    Dim c As Range
    Dim i As Long
    Dim TestTable As ListObject
    
    Set TestSheet = Worksheets("Test")
    Set PasteRange = TestSheet.Range("A1")
    Set c = SetupResponseTables(TestSheet, PasteRange, "TeachersTable")
    
    'Debug.Print c.Address
        

End Sub


Sub emptyrowtest()

    Dim TestSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim TestTable As ListObject
    
    Set TestSheet = Worksheets("Test")
    Set c = TestSheet.Range("A6:G6")
    
    For Each TestTable In TestSheet.ListObjects
        TestTable.Unlist
    Next TestTable
    
    TestSheet.Cells.ClearContents
    TestSheet.Cells.ClearFormats
    
    Set TestTable = TestSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=c, xlListObjectHasHeaders:=xlYes)
    
    For i = 1 To 5
    TestTable.ListRows.Add
    Next i
    

End Sub
