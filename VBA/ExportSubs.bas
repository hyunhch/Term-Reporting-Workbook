Attribute VB_Name = "ExportSubs"
Option Explicit

Function ExportCoverSheet(OldBook As Workbook, NewBook As Workbook) As Long
'Makes a simple cover sheet for the new book
'Passing OldBook to make sure we are pulling from the correct workbook, even if the NewBook is active
'Returns 1 if successful
    
    Dim OldCoverSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim RefRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    
    Set OldCoverSheet = OldBook.Worksheets("Cover Page")
    
    ExportCoverSheet = 0
    OldBook.Activate
    
    'Make a new sheet and insert information
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Cover Page"
    End With
    
    'Find where the needed text in. It's stored one column right of the CoverTextList range
    Set RefRange = Range("CoverTextList")
        If RefRange Is Nothing Then
            GoTo Footer
        End If
    
    For Each c In RefRange.Offset(0, 1)
        Set CopyRange = BuildRange(OldCoverSheet.Range(c.Value), CopyRange)
        Set PasteRange = BuildRange(NewSheet.Range(c.Value), PasteRange)
    Next c
    
    'Copy values and formats
    CopyRange.Resize(CopyRange.Rows.Count, 2).Copy
    PasteRange.Resize(PasteRange.Rows.Count, 2).PasteSpecial Paste:=xlPasteValues
    PasteRange.Resize(PasteRange.Rows.Count, 2).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    NewSheet.Range("B:B").EntireColumn.AutoFit
    
    'Return
    ExportCoverSheet = 1
    
Footer:

End Function

Function ExportLocalSave(OldBook As Workbook, NewBook As Workbook) As Long
'For making a local save
'Returns 1 if successful, 2 if canceled
'Returns 0 on error

    Dim CoverSheet As Worksheet
    Dim CenterString As String
    Dim FileName As String
    Dim LocalPath As String
    Dim SaveName As String
    Dim SubDate As String
    Dim SubTime As String

    ExportLocalSave = 0

    Set CoverSheet = Worksheets("Cover Page")

    'Pull in center from the CoverSheet, date, and time
    CenterString = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1).Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Make a file name using the center, date, and time of submission
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"
    
    'Where the OldBook is stored
    OldBook.Activate
    LocalPath = GetLocalPath(ThisWorkbook.path)

    'For Win and Mac
    If Application.OperatingSystem Like "*Mac*" Then
        SaveName = Application.GetSaveAsFilename(LocalPath & "/" & FileName) ', "Excel Files (*.xlsm), *.xlsm")  MacOS sandboxing can't use file filters
        If SaveName = "False" Then
            NewBook.Close savechanges:=False
            ExportLocalSave = 2
            
            GoTo Footer
        End If
        NewBook.SaveAs FileName:=LocalPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    Else
        SaveName = Application.GetSaveAsFilename(LocalPath & "\" & FileName, "Excel Files (*.xlsm), *.xlsm")
        If SaveName = "False" Then
            NewBook.Close savechanges:=False
            ExportLocalSave = 2
            
            GoTo Footer
        End If
        NewBook.SaveAs FileName:=SaveName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        'ActiveWorkbook.Close SaveChanges:=False
    End If
    
    'Everything worked
    ExportLocalSave = 1

Footer:

End Function

Function ExportMakeBook(SheetArray As Variant, Optional ExportRange As Range) As Workbook
'Container function for each section of exporting
'Every sheet passed gets included in the returned book

    Dim NewBook As Workbook
    Dim OldBook As Workbook
    Dim NewSheet As Worksheet
    Dim i As Long
    Dim SheetValue As Long
    Dim ErrorString As String
    Dim SheetName As String
    Dim ReadyArray() As Variant
    
    Set OldBook = ThisWorkbook
    
    'If no array was passed, break
    If IsArray(SheetArray) = False Or IsEmpty(SheetArray) = True Then
        GoTo ReturnBook
    End If
    
    'Determine which sheets have been filled out
    ReadyArray = GetReadyToExport
    
    For i = 1 To UBound(ReadyArray, 2)
        SheetName = ReadyArray(1, i)
        SheetValue = ReadyArray(2, i)
        
        If Not SheetValue < 2 Then
            Select Case SheetName
                Case "Cover Page"
                    ErrorString = ErrorString & "Please enter your name, date, and center on the Cover Page" & vbCr
                
                Case "Roster Page"
                    ErrorString = ErrorString & "You have no students on your roster" & vbCr
                
                Case "Records Page"
                    ErrorString = ErrorString & "Please parse your roster" & vbCr
                
                Case "Report Page"
                    ErrorString = ErrorString & "Please tabulate the report" & vbCr
                    
                Case Else
                    
            End Select
        End If
    Next i
    
    'If there were unfinished sheets
    If Len(ErrorString) > 0 Then
        MsgBox (ErrorString)
        
        GoTo Footer
    End If
    
    'Create new book and add sheets
    Set NewBook = Workbooks.Add

    'Loop through
    For i = LBound(SheetArray) To UBound(SheetArray)
        SheetName = SheetArray(i)
    
        Select Case SheetName
            Case "Cover"
                Call ExportCoverSheet(OldBook, NewBook)
                
            Case "Roster"
                Call ExportRosterSheet(OldBook, NewBook, ExportRange)
                
            Case "Report"
                Call ExportReportSheet(OldBook, NewBook)
                
            Case "Narrative", "Directory", "Other"
                Call ExportTableSheet(OldBook, NewBook, SheetName)
                
            'Case "Directory"
                'Call ExportTableSheet(OldBook, NewBook, SheetName)
                
            'Case "Other"
                'Call ExportTableSheet(OldBook, NewBook, SheetName)
        End Select
    Next i

ReturnBook:
    'Delete "Sheet1"
    Set NewSheet = NewBook.Worksheets("Sheet1")
    NewSheet.Delete

    Set ExportMakeBook = NewBook
    
    GoTo Footer

ErrorMessage:
    MsgBox ("Something has gone wrong, please close and reopen this file, then try again." & vbCr _
        & "If the problem persists, please contact the state office.")
            
    NewBook.Close savechanges:=False

Footer:

End Function

Function ExportReportSheet(OldBook As Workbook, NewBook As Workbook) As Long
'Grabs the entire Report sheet, only done when exporting the entire roster
'Returns 1 if successful, 0 otherwise

    Dim OldReportSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim OldReportTable As ListObject
    Dim NewTable As ListObject
    
    Set OldReportSheet = OldBook.Worksheets("Report Page")
    Set OldReportTable = OldReportSheet.ListObjects(1)
    
    ExportReportSheet = 0
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Report Page"
    End With
    
    'Copy and paste the entire table data
    Set CopyRange = OldReportTable.Range
    Set c = NewSheet.Range("A1")
    Set PasteRange = c.Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
    
    CopyRange.Copy Destination:=PasteRange
    
    'Autofit
    PasteRange.EntireColumn.AutoFit
    
    ExportReportSheet = 1

Footer:

End Function

Function ExportRosterSheet(OldBook As Workbook, NewBook As Workbook, Optional ExportRange As Range) As Long
'Reproduces the roster in the NewBook. Grabs all students by default
'Passing a range restricts to only those students

    Dim OldRosterSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim OldRosterTable As ListObject
    
    Set OldRosterSheet = OldBook.Worksheets("Roster Page")
    Set OldRosterTable = OldRosterSheet.ListObjects(1)
    
    ExportRosterSheet = 0
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Roster Page"
    End With

    'Reproduce the entire table
    Set CopyRange = OldRosterTable.Range
    Set PasteRange = NewSheet.Range("A1").Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
    
    CopyRange.Copy Destination:=PasteRange
    
    'Make a new table
    'Call CreateTable(NewSheet)
    
    'Chop off the first column and return
    NewSheet.Range("A1").EntireColumn.Delete
    
    ExportRosterSheet = 1

Footer:

End Function

Function ExportSharePoint(OldBook As Workbook, NewBook As Workbook) As Long
'Sends the cover sheet and report to SharePoint
'Returns 1 if successful
'Returns 0 on error
    Dim CoverSheet As Worksheet
    Dim CenterString As String
    Dim FileName As String
    Dim SaveName As String
    Dim SubDate As String
    Dim SubTime As String
    Dim SpPath As String
    Dim TempArray() As Variant

    ExportSharePoint = 0

    Set CoverSheet = Worksheets("Cover Page")

    'Pull in center from the CoverSheet, date, and time
    CenterString = CoverSheet.Range("A:A").Find("Center", , xlValues, xlWhole).Offset(0, 1).Value
    SubDate = Format(Date, "yyyy-mm-dd")
    SubTime = Format(Time, "hh-nn AM/PM")

    'Make a file name using the center, date, and time of submission
    FileName = CenterString & " " & SubDate & "." & SubTime & ".xlsm"

    'The address where the new book will be save in SharePoint
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Center%20Files%20Data%20and%20Reports/Report%20Submissions/"

    'Upload
    NewBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    NewBook.Close savechanges:=False
    
    'Everything worked
    ExportSharePoint = 1
    
Footer:

End Function

Function ExportTableSheet(OldBook As Workbook, NewBook As Workbook, SheetName As String) As Long
'Reproduces the NarrativeSheet, DirectorySheet, or OtherSheet
'Copy paste and tables and header rows separately
'Returns 1 if successful, 0 otherwise

    Dim OldSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim CopyTable As ListObject
    
    ExportTableSheet = 0
    
    Set OldSheet = OldBook.Worksheets(SheetName & " Page")
        If OldSheet Is Nothing Then
            GoTo Footer
        End If
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = OldSheet.Name '& " Test"
    End With
    
    'Find the entire used range and copy over
    Set c = FindUsedRange(OldSheet)
        If c Is Nothing Then
            GoTo Footer
        End If
        
    'We only needs the first two row
    Set CopyRange = c.Resize(2, c.Columns.Count)
    Set PasteRange = NewSheet.Range(CopyRange.Address)
    
    CopyRange.Copy Destination:=PasteRange
    
    'Grab all the tables
    For Each CopyTable In OldSheet.ListObjects
        CopyTable.Range.Copy Destination:=NewSheet.Range(CopyTable.Range.Address)
    Next CopyTable

    NewSheet.Range("3:3").Columns.AutoFit
    
    ExportTableSheet = 1
    
Footer:

End Function

'****Not finished****

Function ExportToRA() As Long

    Dim DirectorySheet As Worksheet
    Dim c As Range
    Dim DirectorNameString As String
    Dim DirectorAddressString As String
    Dim RANameString As String
    Dim RAAddressString As String
    Dim FilePath As String
    Dim FileName As String
    Dim DirectoryTable As ListObject
    
    'Check the table
    Set DirectorySheet = Worksheets("Directory Page")
        If CheckTable(DirectorySheet) > 2 Then
            ExportToRA = 0
            
            GoTo Footer
        End If

    Set DirectoryTable = DirectorySheet.ListObjects(1) 'Prompt to fill these out
    
    'Grab RA information
    Set c = DirectoryTable.ListColumns("Position").DataBodyRange.Find("RA", , xlValues, xlWhole)
        If c Is Nothing Then
            ExportToRA = 0
        
            GoTo Footer
        End If
    
    RANameString = DirectorySheet.Cells(c.Row, DirectoryTable.ListColumns("Name").Range.Column)
    RAAddressString = DirectorySheet.Cells(c.Row, DirectoryTable.ListColumns("Email").Range.Column)
    
    'Director information
    Set c = DirectoryTable.ListColumns("Position").DataBodyRange.Find("Director", , xlValues, xlWhole)
        If c Is Nothing Then
            ExportToRA = 0
        
            GoTo Footer
        End If
    
    DirectorNameString = DirectorySheet.Cells(c.Row, DirectoryTable.ListColumns("Name").Range.Column)
    DirectorAddressString = DirectorySheet.Cells(c.Row, DirectoryTable.ListColumns("Email").Range.Column)

    'Verify they all have something in them. In the future, validate email addresses
    If Not Len(RANameString) > 0 Or Not Len(RAAddressString) > 0 Or Not Len(DirectorNameString) > 0 Or Not Len(DirectorNameString) > 0 Then
        ExportToRA = 0
    
        GoTo Footer
    End If

    'Where it goes
    FilePath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Internal%20Documentation/Word Test.Docx"
    FileName = "Email to " & RANameString
    
    'Make the body of the email
    
    'Pass
    Call MakeDoc(RANameString, DirectorNameString, FilePath, FileName)



    ExportToRA = 1
    'Send error email if not 1


Footer:

    'Debug.Print ExportToRA

End Function

Sub MakeDoc(RANameString As String, DirectorNameString As String, FilePath As String, FileName As String)

    Dim WordApp As Object
    Dim WordDoc As Object
    Dim ContentString As String
    Dim LinkString As String
    
    Set WordApp = CreateObject("Word.Application")
    Set WordDoc = WordApp.Documents.Add
    
    WordDoc.SaveAs FilePath
    'LinkString = "https://uwnetid.sharepoint.com/:f:/r/sites/partner_university_portal/Data%20Portal/Term%20Reports?csf=1&amp;web=1&amp;e=6iOr7A"
    LinkString = "https:\\uwnetid.sharepoint.com\:f:\r\sites\partner_university_portal\Data%20Portal\Term%20Reports?csf=1&amp;web=1&amp;e=6iOr7A"
    ContentString = "<p>" & "Dear " & RANameString & "," & "</p>" & _
        "<p>" & DirectorNameString & ", your " & _
        "local Washington MESA director, has submitted their bi-annual data report to the state office and a copy is attached. " & _
        "It contains several items. First is a demographic breakdown of the center&rsquo;s student population, tabulated by which " & _
        "interventions or activities they participated in. In addition, it shows the focal areas and goals from the previous six " & _
        "months; a directory of their staff, faculty sponsors, and local educators; and a page to document " & _
        "any work that did not fall into the above categories." & "</p><p>&nbsp;</p><p>" & _
        "We ask that you review this report and e-sign this message to verify that you have received it." & "</p>" & _
        "<p>&nbsp;</p><p>" & "You should have access to this and all previous reports since 2021 at " & _
        "<a><hrf =" & LinkString & ">this link. </a>" & _
        "If you cannot access it, please contact the state office at " & "<a><href=" & "mailto:wamesa@uw.edu" & ">wamesa@uw.edu</a>" & ". " & _
        ". " & DirectorNameString & " has access to more detailed records, should more granular data be needed." & "</p><p>&nbsp;</p>" & _
        "<p>" & "Thank you for your time and support," & "</p>" & _
        "<p>" & "Washington MESA, State Office" & "</p>"
    
    Call htmltext(ContentString)
    
    'WordDoc.Content.Text = ContentString
    'Set WordDoc = WordApp.Documents.Open(FilePath)
    
    'WordDoc.Close
    'Set WordDoc = Nothing
    'WordApp.Quit
    'Set WordApp = Nothing
End Sub


'Function HtmlToText(sHTML) As String
Sub htmltext(ContentString As String)
    'Dim oDoc As Object
    'Dim oDoc As HTMLDocument
    Dim html As Object
    Dim FilePath As String
    'Set oDoc = New HTMLDocument
    'Set oDoc = CreateObject("HTMLDocument")
    Set html = CreateObject("htmlfile")
    'FilePath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/_layouts/15/download.aspx?UniqueId=f75271ad%2D055c%2D4a11%2D844a%2D6b48a5d0876d"
    FilePath = "C:\Users\hyunhch\OneDrive - UW\Desktop\Dear (2).html"
    
    html.body.innerHTML = ContentString
    'html.SaveAs FileName:="C:\Users\hyunhch\OneDrive - UW\Desktop\Dear (2).htm"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(FilePath, True, True)
    Fileout.Write ContentString
    Fileout.Close
    
    Dim TestVar As String
    TestVar = html.getElementsByTagName("p").Item(0).innerText
    
    Debug.Print TestVar
    
    'oDoc.body.innerHTML = FilePath
    'HtmlToText = oDoc.body.innerText
End Sub



