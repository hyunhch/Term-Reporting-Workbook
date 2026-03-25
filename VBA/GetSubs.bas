Attribute VB_Name = "GetSubs"
Option Explicit

Function GetCoverInfo() As Variant
'Grabs the name, date, center, program, and version from the CoverSheet
'Returns an array with each value
'Returns nothing if some fields are missing
    '(1, i) - Header
    '(2, i) - Value

    Dim CoverSheet As Worksheet
    Dim RefRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim ReturnArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RefRange = Range("CoverTextList")
        If RefRange Is Nothing Then
            GoTo Footer
        End If
    
    'Check that everything has been filled out
    If CheckCover <> 1 Then
        GoTo Footer
    End If
    
    'Grab the items we want
    ReDim ReturnArray(1 To 2, 1 To RefRange.Cells.Count)
    
    i = 1
    For Each c In RefRange 'They will be in the same order
        Set d = CoverSheet.Range(c.Offset(0, 1).Value)
        
        ReturnArray(1, i) = c.Value
        
        If i < 3 Then
            ReturnArray(2, i) = d.Value
        Else
            ReturnArray(2, i) = d.Offset(0, 1).Value
        End If
    
        i = i + 1
    Next c
        
    'Return
    GetCoverInfo = ReturnArray

Footer:

End Function

Function GetEdition() As String

    Dim ChangeSheet As Worksheet
    
    Set ChangeSheet = Worksheets("Change Log")
    
    GetEdition = ChangeSheet.Range("A1").Value
    
End Function

Function GetImportSheets(CopyBook As Workbook) As Variant
'Checks a selected file if the Roster, Report, and Records sheets exist
'Returns each sheet it finds
'Returns empty elements if the sheet isn't found
    '(1, i) - Sheet name
    '(2, i) - Sheet object
    '(3, i) - 0, to be used later. Set to 1 if importing the sheet is successful

    Dim CopySheet As Worksheet
    Dim i As Long
    Dim CopyFilePath As String
    Dim SheetArray() As Variant
    
    'Define the names of the sheet to look for
    ReDim SheetArray(1 To 3, 1 To 3)
        SheetArray(1, 1) = "Report Page"
        SheetArray(2, 1) = "Roster Page"
        SheetArray(3, 1) = "Records Page"
    
    'Loop through sheets. If there are fewer than 3, we can break
    If CopyBook.Sheets.Count < 3 Then
        GoTo Footer
    End If

    For Each CopySheet In CopyBook.Sheets
        For i = 1 To UBound(SheetArray, 1)
            If CopySheet.Name = SheetArray(i, 1) Then
                Set SheetArray(i, 2) = CopyBook.Sheets(CopySheet.Name)
            End If
        Next i
    Next CopySheet

    'Return
    GetImportSheets = SheetArray

Footer:

End Function

Function GetProgram() As String
'Returns the program listed on the Cover Page
'Return nothing if there's no match

    Dim CoverSheet As Worksheet
    Dim TitleString As String
    Dim ReturnString As String
    
    Set CoverSheet = Worksheets("Cover Page")
    
    'See which program is in the workbook title
    TitleString = CoverSheet.Range("A1").Value

    If TitleString Like "*College*" Then
        ReturnString = "College Prep"
    ElseIf TitleString Like "*Transfer*" Then
        ReturnString = "Transfer Prep"
    ElseIf TitleString Like "*University*" Then
        ReturnString = "MESA University"
    Else
        GoTo Footer
    End If
    
    GetProgram = ReturnString

Footer:

End Function

Function GetReadyToExport() As Variant
'Checks the Cover, Report, Roster, Records, Narrative, and Directory
'Returns an array that shows if they're filled out or not

    Dim CoverSheet As Worksheet
    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim NarrativeSheet As Worksheet
    Dim DirectorySheet As Worksheet
    Dim OtherSheet As Worksheet
    Dim ReadyArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    Set RosterSheet = Worksheets("Roster Page")
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set NarrativeSheet = Worksheets("Narrative Page")
    Set DirectorySheet = Worksheets("Directory Page")
    Set OtherSheet = Worksheets("Other Page")
    
    'Read in the names of the sheets to check. Make this programmatic in the future
    ReDim ReadyArray(1 To 2, 1 To 7)
        ReadyArray(1, 1) = "Cover Page"
        ReadyArray(1, 2) = "Roster Page"
        ReadyArray(1, 3) = "Records Page"
        ReadyArray(1, 4) = "Report Page"
        ReadyArray(1, 5) = "Narrative Page"
        ReadyArray(1, 6) = "Directory Page"
        ReadyArray(1, 7) = "Other Page"
         
    'Go through each sheet
        ReadyArray(2, 1) = CheckCover
        ReadyArray(2, 2) = CheckTable(RosterSheet)
            If ReadyArray(2, 2) = 2 Then 'To make the number consistent with other check functions
                ReadyArray(2, 2) = 1
            End If
        ReadyArray(2, 3) = CheckRecords(RecordsSheet)
        ReadyArray(2, 4) = CheckReport(ReportSheet)
        ReadyArray(2, 5) = 0 'Figure out how to verify these
        ReadyArray(2, 6) = 0
        ReadyArray(2, 7) = 0
        
    GetReadyToExport = ReadyArray
        
Footer:

End Function

Function GetRosterSchools(RosterSheet As Worksheet, RosterTable As ListObject) As Variant
'Returns each school, the number of students, number of teachers, and district
'Returns nothing on error
'Blank lines are put into an "Other" category
'Errors on the roster, i.e. two different districts listed for one school, are NOT corrected
    '(1, i) - School Name
    '(2, i) - # Students
    '(3, i) - # Teachers
    '(4, i) - District Name

    Dim SchoolCol As Range
    Dim DistrictCol As Range
    Dim TempRange As Range
    Dim SearchRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim SchoolName As String
    Dim DistrictName As String
    Dim SchoolArray As Variant
    Dim TempArray As Variant
    Dim CountArray As Variant
    
    'Validating tables happens in parent sub
    Set SchoolCol = RosterTable.ListColumns("School").DataBodyRange
    Set DistrictCol = RosterTable.ListColumns("District").DataBodyRange
    
    'Grab unique schools
    SchoolArray = GetUniqueValues(SchoolCol)
        If IsEmpty(SchoolArray) Or Not IsArray(SchoolArray) Then
            GoTo Footer
        ElseIf SchoolArray(1, 1) = "Other" Then 'Only the "other" category
            GoTo Footer
        End If

    'Make array to store district and # teachers as well
    ReDim CountArray(1 To 4, 1 To UBound(SchoolArray, 2))
    
    For i = 1 To UBound(SchoolArray, 2)
        CountArray(1, i) = SchoolArray(1, i)
        CountArray(2, i) = SchoolArray(2, i)
        
        'Find each cell belonging to a school
        SchoolName = CountArray(1, i)
            'Skip "other"
            If SchoolName = "Other" Or Not Len(SchoolName) > 0 Then
                GoTo NextSchool
            End If
        
        Set TempRange = FindAll(SchoolCol, SchoolName)
            'If a range isn't returned. This shouldn't happen
            If TempRange Is Nothing Then
                GoTo NextSchool
            End If
        
        'Nudge to Teachers and find unqiue
        Set SearchRange = NudgeToHeader(RosterSheet, TempRange, "Teacher")
        
        TempArray = GetUniqueValues(SearchRange)
            If IsEmpty(TempArray) Or Not IsArray(TempArray) Then
                GoTo FindDistrict
            End If
            
        'If there's an "Other" element, remove it
        j = UBound(TempArray, 2)
            If TempArray(1, j) = "Other" Then
                j = j - 1
            End If
            
            If j = 0 Then
                GoTo FindDistrict
            End If
            
        'Count of teachers per school
        CountArray(3, i) = j
        
FindDistrict:
        For Each c In TempRange
            DistrictName = RosterSheet.Cells(c.Row, DistrictCol.Column)
            
            If Len(DistrictName) > 0 Then
                CountArray(4, i) = DistrictName
                
                GoTo NextSchool
            End If
        Next c
        
NextSchool:
    Next i

    'Return
    If IsEmpty(CountArray) Or Not IsArray(CountArray) Then
        GoTo Footer
    End If
    
    GetRosterSchools = CountArray
    
Footer:

End Function

Function GetUniqueValues(SearchRange As Range) As Variant
'Loops through a passed range, finds each unique valaue and the number of instances
'This isn't much different from the DemoTabulate functions, but is more general. Consider merging later
'Returns a 1 x 2 array
    '(1, i) - Value
    '(2, i) - Frequency
    
    Dim c As Range
    Dim OtherCounter As Long
    Dim i As Long
    Dim j As Long
    Dim TempValue As String
    Dim CountArray As Variant
    
    'Make sure there's a passed range
    If SearchRange Is Nothing Then
        GoTo Footer
    End If

    'Initialize the array
    ReDim CountArray(1 To 2, 1)
    OtherCounter = 0
    
    i = 1
    For Each c In SearchRange
        TempValue = c.Value
        
        'Blanks will go in an "Other"
        If Not Len(c.Value) > 0 Then
            OtherCounter = OtherCounter + 1
        
            GoTo NextCell
        End If
        
        'Loop through terms and iterate
        For j = 1 To UBound(CountArray, 2)
            If CountArray(1, j) = TempValue Then
                CountArray(2, j) = CountArray(2, j) + 1
                
                GoTo NextCell
            End If
        Next j

        'If it's not already in the array, add it
        ReDim Preserve CountArray(1 To 2, 1 To i) 'This is erasing the "Other" assigned above. The numbers work out, but the text is lost
            CountArray(1, i) = TempValue
            CountArray(2, i) = 1
            
        i = i + 1
NextCell:
    Next c

    'This shouldn't happen
    If Not IsArray(CountArray) Or IsEmpty(CountArray) Then
        GoTo Footer
    End If

    'If there are blanks
    If OtherCounter > 0 Then
        If OtherCounter = SearchRange.Cells.Count Then 'All empty values
            i = 1
        Else
            i = UBound(CountArray, 2) + 1
        End If
        
        ReDim Preserve CountArray(1 To 2, 1 To i)
            CountArray(1, i) = "Other"
            CountArray(2, i) = OtherCounter
    End If
    
    'Return
    GetUniqueValues = CountArray

Footer:

End Function

Function GetVersion() As String
'Returns the version listed in the change log

    Dim ChangeSheet As Worksheet
    Dim c As Range
    
    Set ChangeSheet = Worksheets("Change Log")
    Set c = ChangeSheet.Range("A1")
    
    'If, for some reason, there's nothing there
    If Not InStr(c.Value, "Version") > 0 Then
        Set c = ChangeSheet.Range("A:A").Find("Version", , xlValues, xlPart)
        GetVersion = "Unknown version - " & c.Value & "+"
        
        GoTo Footer
    End If

    GetVersion = c.Value
    
Footer:

End Function
