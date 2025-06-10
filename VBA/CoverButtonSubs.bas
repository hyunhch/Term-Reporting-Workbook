Attribute VB_Name = "CoverButtonSubs"
Option Explicit

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
