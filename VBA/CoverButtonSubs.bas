Attribute VB_Name = "CoverButtonSubs"
Option Explicit

Sub CoverChooseProgramButton()
'Displays on the cover sheet on a fresh copy of the file, deletes after a program is chosen
'Using a button rather than an OnActivate event to avoid issues with macros being disabled and frustration with the popup

    ChooseProgramForm.Show

Footer:

End Sub

Sub CoverSaveCopyButton()

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim i As Long
    Dim SheetNameArray() As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set OldBook = ThisWorkbook
    
    'Define which pages to export. For a local save, it's all of them
    ReDim SheetNameArray(1 To 6) 'Make programmatic
        SheetNameArray(1) = "Cover"
        SheetNameArray(2) = "Roster"
        SheetNameArray(3) = "Report"
        SheetNameArray(4) = "Narrative"
        SheetNameArray(5) = "Directory"
        SheetNameArray(6) = "Other"

    'Pass for making the workbook and saving
    Set NewBook = ExportMakeBook(SheetNameArray)
    
    If Not NewBook Is Nothing Then
        i = ExportLocalSave(OldBook, NewBook)
    End If
    
    If i = 0 Then
        MsgBox ("Something has gone wrong. Please restart this workbook and try again.")
    ElseIf i = 1 Then
        MsgBox ("Save complete.")
    End If
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub CoverSharePointExportButton()

    Dim OldBook As Workbook
    Dim NewBook As Workbook
    Dim i As Long
    Dim SheetNameArray() As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set OldBook = ThisWorkbook
    
    'Define which pages to export. For SharePoint, it's the report, narrative, directory, and other
    ReDim SheetNameArray(1 To 5) 'Make programmatic
        SheetNameArray(1) = "Cover"
        SheetNameArray(2) = "Report"
        SheetNameArray(3) = "Narrative"
        SheetNameArray(4) = "Directory"
        SheetNameArray(5) = "Other"

    'Pass for making the workbook and saving
    Set NewBook = ExportMakeBook(SheetNameArray)
    
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
