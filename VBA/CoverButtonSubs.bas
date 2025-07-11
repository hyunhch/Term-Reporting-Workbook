Attribute VB_Name = "CoverButtonSubs"
Option Explicit

Sub CoverChooseProgramButton()
'Displays on the cover sheet on a fresh copy of the file, deletes after a program is chosen
'Using a button rather than an OnActivate event to avoid issues with macros being disabled and frustration with the popup

    Dim CoverSheet As Worksheet
    Dim RefSheet As Worksheet
    Dim ProgramString As String
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'Check if one of the reference sheets has been renamed. This shouldn't happen
    For Each RefSheet In ThisWorkbook.Worksheets
        If RefSheet.Name = "Ref Tables" Then
            Call TesterClearTables
            
            Exit For
        End If
    Next RefSheet

    ChooseProgramForm.Show

Footer:

End Sub

Sub CoverImportButton()
'Imports data from a different workbook, intended to be used when updates happen

    Dim CopyBook As Workbook
    Dim PasteBook As Workbook
    Dim ReportSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim CopySheet As Worksheet
    'Dim CopyRange As Range
    'Dim PasteRange As Range
    Dim ImportConfirm As Long
    Dim i As Long
    Dim j As Long
    Dim SheetName As String
    Dim CopyFilePath As String
    Dim FinishMessage As String
    Dim CheckArray As Variant
    Dim CopyArray As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set PasteBook = ThisWorkbook
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    
    'Check if there's any content on the three sheets
    CheckArray = GetReadyToExport
        
    'If a number > 2 is returned, there's no content  on the sheet
    For i = LBound(CheckArray, 1) To UBound(CheckArray, 1)
        SheetName = CheckArray(i, 1)
        j = CheckArray(i, 2)
        
        Select Case SheetName
            Case "Report Page", "Records Page", "Roster Page"
                If j < 3 Then
                    GoTo ShowPrompt
                End If
        End Select
    Next i
    
    'If all three pages were empty
    GoTo SkipPrompt
        
ShowPrompt:
    'Prompt for confirmation since this will erase everything currently in the workbook
    ImportConfirm = MsgBox("Importing will delete all data currently in this workbook. " & _
        "Do you want to proceed?", vbQuestion + vbYesNo + vbDefaultButton2) 'Remove if this is too annoying
        
        If ImportConfirm <> vbYes Then
            GoTo Footer
        End If
    
SkipPrompt:
    'Choose the file to import
    CopyFilePath = Application.GetOpenFilename("Excel Files (*.xlsm*), *xlsm*", , "Select the file to import")
        'Clicking "cancel" or otherwise closing the selection window
        If CopyFilePath = "False" Then
            GoTo Footer
        End If

    Set CopyBook = Workbooks.Open(CopyFilePath)
    
    'Look for the Roster, Report, and Records sheets in the selected file
    CopyArray = GetImportSheets(CopyBook)
        'Returns empty on an error
        'Will return empty elements in the 2nd dimension if the sheets weren't found
        If IsEmpty(CopyArray) Then
            MsgBox ("This does not appear to be a valid workbook to import")
        
            GoTo Footer
        End If
        
    'Clear out the existing sheets, without any export prompts, then replace text
    PasteBook.Activate
    
    'Roster sheet
    Call UnprotectSheet(RosterSheet)
    Call ClearSheet(RosterSheet)
    Call RosterParseButton
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

    'Records page
    Call UnprotectSheet(RecordsSheet)
    Call ClearSheet(RecordsSheet)
    Call RecordsSheetText
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
    'Report page
    Call ReportClearAll 'Nothing needs to be put back in
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

    'Loop through CopyArray and copy over the ones that have something on them
    For i = LBound(CopyArray, 1) To UBound(CopyArray, 1)
        If IsEmpty(CopyArray(i, 2)) Then
            GoTo NextSheet
        End If
            
        SheetName = CopyArray(i, 1)
        Set CopySheet = CopyArray(i, 2)
        
        Select Case SheetName
            Case "Records Page"
                Call UnprotectSheet(RecordsSheet)
                If ImportRecords(CopySheet, RecordsSheet) = 1 Then
                    CopyArray(i, 3) = 1
                End If
                
                Application.EnableEvents = False
                Application.ScreenUpdating = False
                Application.DisplayAlerts = False
            
            Case "Report Page" 'This is temporary
                CopyArray(i, 3) = 1
                
            'Case "Report Page" **Omitting this for now. The report can simply be retabulated. Add better functionality later
                'Set CopyRange = CopySheet.ListObjects(1).Range
                'Set PasteRange = ReportSheet.ListObjects(1).Range
                    'If CopyRange.Cells.Count <> PasteRange.Cells.Count Then
                        'MsgBox ("The old and new Report Pages did not match. You may need to retabulate manually")
                    'End If
                    
                'Call UnprotectSheet(ReportSheet)
                'PasteRange.Value = CopyRange.Value 'This gets rid of the table for some reason
                'Call CreateReportTable
    
                'Application.EnableEvents = False
                'Application.ScreenUpdating = False
                'Application.DisplayAlerts = False
            
            Case "Roster Page"
                Call UnprotectSheet(RosterSheet)
                If ImportRoster(CopySheet, RosterSheet) = 1 Then
                    CopyArray(i, 3) = 1
                End If
                
                Application.EnableEvents = False
                Application.ScreenUpdating = False
                Application.DisplayAlerts = False
        End Select
    
NextSheet:
    Next i

    'Tabulate everything
    Call TabulateAll

    'Close the selected file
    CopyBook.Close savechanges:=False
    
    'Confirm that everything copied over
    For i = LBound(CopyArray, 2) To UBound(CopyArray, 2)
        If CopyArray(i, 3) = 0 Then
            FinishMessage = "One or more sheets did not import. Please ensure the selected file is not missing any sheets and try again."
            
            GoTo PopupMessage
        End If
    Next i

    FinishMessage = "Import successful"
    
PopupMessage:
    MsgBox (FinishMessage)

Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverLocalSaveButton()

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim i As Long
    Dim SheetNameArray() As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set OldBook = ThisWorkbook
    
    'Define which pages to export. For a local save, it's all of them
    ReDim SheetNameArray(1 To 7) 'Make programmatic
        SheetNameArray(1) = "Roster"
        SheetNameArray(2) = "Simple"
        SheetNameArray(3) = "Detailed"
        SheetNameArray(4) = "Report"
        SheetNameArray(5) = "Narrative"
        SheetNameArray(6) = "Directory"
        SheetNameArray(7) = "Other"

    'Pass for making the workbook and saving
    Set NewBook = ExportMakeBook(, SheetNameArray)
    
    If Not NewBook Is Nothing Then
        i = ExportLocalSave(OldBook, NewBook)
    End If
    
    If i <> 1 Then
        MsgBox ("Something has gone wrong. Please restart this workbook and try again.")
    Else
        MsgBox ("Save complete.")
    End If
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverSharePointButton()

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim i As Long
    Dim SheetNameArray() As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set OldBook = ThisWorkbook
    
    'Define which pages to export. For SharePoint, it's the report, narrative, directory, and other
    ReDim SheetNameArray(1 To 4) 'Make programmatic
        SheetNameArray(1) = "Report"
        SheetNameArray(2) = "Narrative"
        SheetNameArray(3) = "Directory"
        SheetNameArray(4) = "Other"

    'Pass for making the workbook and saving
    Set NewBook = ExportMakeBook(, SheetNameArray)
    
    If Not NewBook Is Nothing Then
        i = ExportSharePoint(OldBook, NewBook)
    End If
    
    If i <> 1 Then
        MsgBox ("Something has gone wrong. Please restart this workbook and try again.")
    Else
        MsgBox ("Exported to SharePoint.")
    End If
  
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
