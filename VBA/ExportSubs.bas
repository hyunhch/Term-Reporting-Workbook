Attribute VB_Name = "ExportSubs"
Option Explicit

Function ExportCoverSheet(OldBook As Workbook, NewBook As Workbook) As Long
'Makes a simple cover sheet for the new book
'Passing OldBook to make sure we are pulling from the correct workbook, even if the NewBook is active
'Returns 1 if successful
    
    Dim OldCoverSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim i As Long
    Dim CoverInfoArray() As Variant
    
    Set OldCoverSheet = OldBook.Worksheets("Cover Page")
    
    ExportCoverSheet = 0
    OldBook.Activate
    
    'Grab the needed info
    CoverInfoArray = GetCoverInfo
    
    If IsEmpty(CoverInfoArray) = True Then 'This shouldn't happen
        GoTo Footer
    End If
    
    'Make a new sheet and insert information
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Cover Page"
    End With
    
    For i = 1 To UBound(CoverInfoArray)
        NewSheet.Cells(i, 1).Value = CoverInfoArray(i, 1)
        NewSheet.Cells(i, 2).Value = CoverInfoArray(i, 2)
        
        'Format the date
        If CoverInfoArray(i, 1) = "Date" Then
            NewSheet.Cells(i, 2).NumberFormat = "mm/dd/yyyy"
        End If
    Next i
    
    'Return
    ExportCoverSheet = 1
    
Footer:

End Function

Function ExportDetailedAttendance(OldBook As Workbook, NewBook As Workbook, Optional ExportRange As Range) As Long
'Exports a line for each time a student was marked present for an activity
'Passing a range only exports for those students. Range should be from the RecordsSheet
'Returns 1 if successful, 0 otherwise

    Dim OldRecordsSheet As Worksheet
    Dim OldRosterSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim OldRosterNameRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim PresentRange As Range
    Dim SearchRange As Range
    Dim NameCell As Range
    Dim c As Range
    Dim d As Range
    Dim OldRosterTable As ListObject
    
    ExportDetailedAttendance = 0

    Set OldRecordsSheet = OldBook.Worksheets("Records Page")
    Set OldRosterSheet = OldBook.Worksheets("Roster Page")
    Set OldRosterTable = OldRosterSheet.ListObjects(1)
    Set OldRosterNameRange = OldRosterTable.ListColumns("First").DataBodyRange

    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Detailed Attendance"
    End With
    
    'Put in the headers, which will be the headers on the RosterTable except the first row, then the activity name and notes
    Set c = OldRosterTable.HeaderRowRange
    Set CopyRange = c.Resize(1, c.Columns.Count - 1).Offset(0, 1)
    
    Set c = NewSheet.Range("A1")
    Set PasteRange = c.Resize(1, CopyRange.Columns.Count)
        PasteRange.Value = CopyRange.Value
    
    Set d = c.Offset(0, PasteRange.Columns.Count)
        d.Value = "Activity"
        d.Offset(0, 1).Value = "Notes"
    
    'Define the range to search
    If ExportRange Is Nothing Then
        Set SearchRange = FindRecordsName(OldRecordsSheet)
    Else
        Set SearchRange = ExportRange
    End If
    
    'Loop through range, looking for any instances where the student is marked present
    For Each c In SearchRange
        Set d = FindRecordsAttendance(OldRecordsSheet, c)
            If d Is Nothing Then
               GoTo NextStudent
            End If

        Set PresentRange = FindChecks(d)
            If PresentRange Is Nothing Then
                GoTo NextStudent
            End If
        
        'Find the student on the RosterTable
        Set NameCell = FindName(c, OldRosterNameRange)
            If NameCell Is Nothing Then
                GoTo NextStudent
            End If
            
        'Copy over student information
        Call ExportDetailedHelper(NewSheet, OldRecordsSheet, OldRosterSheet, PresentRange, NameCell)
        
NextStudent:
    Next c

    'Make a table
    Set c = NewSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set d = NewSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlColumns)
    Set PasteRange = NewSheet.Range("A1", Cells(c.Row, d.Column).Address)
    
    Call CreateTable(NewSheet, , PasteRange)
    
    ExportDetailedAttendance = 1

Footer:

End Function

Function ExportDetailedHelper(NewSheet As Worksheet, OldRecordsSheet As Worksheet, OldRosterSheet As Worksheet, PresentRange As Range, NameCell As Range) As Long
'Given a student and their attendance record on the RecordsSheet, creates a new row on the NewSheet each time they are marked present
'Returns 1 if successful

    Dim OldRosterNameRange As Range
    Dim StudentInfoRange As Range
    Dim ActivityInfoRange As Range
    Dim c As Range
    Dim d As Range
    Dim PasteRange As Range
    Dim i As Long
    Dim OldRosterTable As ListObject
    
    ExportDetailedHelper = 0
    
    Set OldRosterTable = OldRosterSheet.ListObjects(1)
    Set OldRosterNameRange = OldRosterTable.ListColumns("First").DataBodyRange

    'Make sure the student is on the RosterSheet, if not then break
    Set c = FindName(NameCell, OldRosterNameRange)
    
    If c Is Nothing Then
        GoTo Footer
    End If
    
    'Grab the student's info, so the entire table except the first column
    Set StudentInfoRange = c.Resize(1, OldRosterTable.ListColumns.Count - 1) '.Offset(0, 1)
    
    'Find where to start pasting
    Set PasteRange = NewSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    'Loop through the PresentRange, pasting the same student info and the corresponding activity info
    i = 1
    For Each c In PresentRange
        Set d = PasteRange.Resize(1, StudentInfoRange.Columns.Count).Offset(i, 0)
            d.Value = StudentInfoRange.Value
    
        Set ActivityInfoRange = OldRecordsSheet.Cells(1, c.Column)
            PasteRange.Offset(i, StudentInfoRange.Columns.Count).Value = ActivityInfoRange.Value
            PasteRange.Offset(i, StudentInfoRange.Columns.Count + 1).Value = ActivityInfoRange.Offset(1, 0).Value
    
        i = i + 1
    Next c

    ExportDetailedHelper = 1
    
Footer:

End Function

Function ExportGenericSheet(OldBook As Workbook, NewBook As Workbook, OldSheet As Worksheet) As Long
'Simply reproduces the data and formatting of a sheet

    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim LCol As Long
    Dim LRow As Long

    ExportGenericSheet = 0

    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = OldSheet.Name
    End With
    
    'Find the used area, copy over
    LCol = OldSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    LRow = OldSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    Set c = OldSheet.Cells(LRow, LCol)
    Set CopyRange = OldSheet.Range("A1", c)
    Set PasteRange = NewSheet.Range(CopyRange.Address)
    
    'Use copy/paste to preserve formatting
    CopyRange.Copy Destination:=PasteRange

    ExportGenericSheet = 1

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

Function ExportMakeBook(Optional ExportRange As Range, Optional SheetArray As Variant) As Workbook
'Container function for each section of exporting
'By default, only exports the cover page
'Every sheet passed gets included in the returned book

    Dim NewBook As Workbook
    Dim OldBook As Workbook
    Dim NewSheet As Worksheet
    Dim OldNarrativeSheet As Worksheet
    Dim OldDirectorySheet As Worksheet
    Dim OldOtherSheet As Worksheet
    Dim i As Long
    Dim SheetValue As Long
    Dim SheetName As String
    Dim ReadyArray() As Variant
    
    Set OldBook = ThisWorkbook
    Set OldNarrativeSheet = Worksheets("Narrative Page")
    Set OldDirectorySheet = Worksheets("Directory Page")
    Set OldOtherSheet = Worksheets("Other Page")
    
    'Determine which sheets have been filled out
    ReadyArray = GetReadyToExport
    
    For i = 1 To UBound(ReadyArray)
        SheetName = ReadyArray(i, 1)
        SheetValue = ReadyArray(i, 2)
 
        'Break if we're missing needed information
        If SheetName = "Cover Page" And SheetValue <> 1 Then
            MsgBox ("Please enter your name, date, and center on the Cover Page")
            GoTo Footer
        ElseIf SheetName = "Roster Page" And SheetValue > 2 Then 'This shouldn't happen
            MsgBox ("You have no students on your roster.")
            GoTo Footer
        ElseIf SheetName = "Roster Page" And SheetValue > 2 Then
            MsgBox ("You have no saved attendance information. Please parse your roster.") 'This shouldn't happen
            GoTo Footer
        ElseIf SheetName = "Report Page" And SheetValue > 2 Then 'We only need the report in some cases
            MsgBox ("You have no activities on the Report Page.")
            GoTo Footer
        End If
    Next i
    
    'Create new book and add sheets
    Set NewBook = Workbooks.Add

    'CoverSheet will always be added
    If ExportCoverSheet(OldBook, NewBook) <> 1 Then
        GoTo ErrorMessage
    End If
    
    'If no array was passed, we're done
    If IsArray(SheetArray) = False Then
        GoTo ReturnBook
    End If
    
    'Loop through
    For i = LBound(SheetArray) To UBound(SheetArray)
        SheetName = SheetArray(i)
    
        Select Case SheetName
            Case "Roster"
                Call ExportRosterSheet(OldBook, NewBook, ExportRange)
            Case "Simple"
                Call ExportSimpleAttendance(OldBook, NewBook, ExportRange)
            Case "Detailed"
                Call ExportDetailedAttendance(OldBook, NewBook, ExportRange)
            Case "Report"
                Call ExportReportSheet(OldBook, NewBook)
            Case "Narrative"
                Call ExportGenericSheet(OldBook, NewBook, OldNarrativeSheet)
            Case "Directory"
                Call ExportGenericSheet(OldBook, NewBook, OldDirectorySheet)
            Case "Other"
                Call ExportGenericSheet(OldBook, NewBook, OldOtherSheet)
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
    
    'Chop off the first row, autofit
    c.EntireColumn.Delete
    PasteRange.EntireColumn.AutoFit
    
    ExportReportSheet = 1

Footer:

End Function

Function ExportRosterSheet(OldBook As Workbook, NewBook As Workbook, Optional ExportRange As Range) As Long
'Reproduces the roster in the NewBook. Grabs all students by default
'Passing a range restricts to only those students

    Dim OldRosterSheet As Worksheet
    Dim OldRecordsSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim OldRosterTable As ListObject
    Dim NewTable As ListObject
    
    Set OldRosterSheet = OldBook.Worksheets("Roster Page")
    Set OldRecordsSheet = OldBook.Worksheets("Records Page")
    Set OldRosterTable = OldRosterSheet.ListObjects(1)
    
    ExportRosterSheet = 0
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Roster Page"
    End With

    'If there is no passed ranged, reproduce the entire table
    If ExportRange Is Nothing Then
        Set CopyRange = OldRosterTable.Range
        Set PasteRange = NewSheet.Range("A1").Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
        
        PasteRange.Value = CopyRange.Value
    Else
        'Make sure we're on the RosterSheet. If not, search there for the names to be exported
        If ExportRange.Worksheet.Name = "Records Page" Then
            Set c = OldRosterTable.ListColumns("First").DataBodyRange
            Set d = FindName(ExportRange, c)
                If d Is Nothing Then
                    GoTo Footer
                End If
        Else
            Set d = ExportRange
        End If
    
        'Resize to get the entire row
        Set CopyRange = Intersect(d.EntireRow, OldRosterTable.DataBodyRange)
    
        'Include the header
        Set CopyRange = Union(OldRosterTable.HeaderRowRange, CopyRange)
        Set PasteRange = NewSheet.Range("A1")
        
        Call CopyRows(OldRosterSheet, CopyRange, NewSheet, PasteRange)
    End If
    
    'Make a new table
    Call CreateTable(NewSheet)
    
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
    SpPath = "https://uwnetid.sharepoint.com/sites/partner_university_portal/Data%20Portal/Report%20Submissions/"

    'Upload
    NewBook.SaveAs FileName:=SpPath & "/" & FileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    NewBook.Close savechanges:=False
    
    'Everything worked
    ExportSharePoint = 1
    
Footer:

End Function

Function ExportSimpleAttendance(OldBook As Workbook, NewBook As Workbook, Optional ExportRange As Range) As Long
'Exports if a student was present, absent, or unrecorded for every activity
    '1 - present
    '0 - absent
    '[nothing] - N/A
'Passing a range only exports for those students. Range should be from the RecordsSheet
'Returns 1 if successful, 0 otherwise

    Dim OldRecordsSheet As Worksheet
    Dim NewSheet As Worksheet
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim OldRecordsHeaderRange As Range
    Dim OldRecordsFoundNames As Range
    Dim c As Range
    Dim d As Range
    Dim LRow As Long
    Dim LCol As Long
    
    ExportSimpleAttendance = 0
    
    Set OldRecordsSheet = OldBook.Worksheets("Records Page")
    
    'Make a new sheet
    With NewBook
        Set NewSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        NewSheet.Name = "Simple Attendance"
    End With

    'If there is no passed range, we can copy and paste the entire page
    If ExportRange Is Nothing Then
        LRow = OldRecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        LCol = OldRecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
        Set c = OldRecordsSheet.Cells(LRow, LCol)
        Set CopyRange = OldRecordsSheet.Range("A1", c)
        
        Set c = NewSheet.Range("A1")
        Set PasteRange = c.Resize(CopyRange.Rows.Count, CopyRange.Columns.Count)
    
        PasteRange.Value = CopyRange.Value
    Else
        'Grab all of the activities and the headers
        Set c = OldRecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
        Set d = OldRecordsSheet.Range("1:1").Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        Set OldRecordsHeaderRange = OldRecordsSheet.Range("A1", Cells(c.Row, d.Column).Address)
        
        'Make sure we're on the RecordsSheet. If not, search there for the names to be exported
        If Not ExportRange.Worksheet.Name = "Records Page" Then
            Set c = FindRecordsName(OldRecordsSheet)
                If c Is Nothing Then
                    GoTo Footer
                End If
            
            Set OldRecordsFoundNames = FindName(ExportRange, c)
                If OldRecordsFoundNames Is Nothing Then
                    GoTo Footer
                End If
        Else
            Set OldRecordsFoundNames = ExportRange
        End If
        
        'Define the range to copy over
        Set c = Intersect(OldRecordsFoundNames.EntireRow, OldRecordsHeaderRange.EntireColumn)
        Set CopyRange = Union(OldRecordsHeaderRange, c)
        Set PasteRange = NewSheet.Range("A1")
        
        Call CopyRows(OldRecordsSheet, CopyRange, NewSheet, PasteRange)
    End If
    
    'Format the headers bold
    Set c = NewSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole)
    
    NewSheet.Range("A1", c).EntireRow.Font.Bold = True
    
    'Delete the padding cells
    c.EntireRow.Delete
    
    Set c = NewSheet.Range("1:1").Find("V BREAK", , xlValues, xlWhole)
    c.EntireColumn.Delete
    
    ExportSimpleAttendance = 1
    
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



