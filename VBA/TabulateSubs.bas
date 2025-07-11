Attribute VB_Name = "TabulateSubs"
Option Explicit

Function DemoTabulate(SearchRange As Range, SearchType As String) As Variant
'This is an intermediate function that will call either a function for Windows or for MacOS
'MacOS doesn't support dictionaries, but that method is faster
'Returns an array with the count of each category being tabulated

    If Application.OperatingSystem Like "*Mac*" Then
        DemoTabulate = DemoTabulateMac(SearchRange, SearchType)
    Else
        DemoTabulate = DemoTabulateWin(SearchRange, SearchType)
    End If

Footer:

End Function

Function DemoTabulateWin(SearchRange As Range, SearchType As String) As Variant
'Returns an array with the values in the passed range tabulated
'Uses a dictionary so no reference to the columns needs to be done
'SearchType is for renaming "Other" values. This is why this is done piecemeal instead of all at once

    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim RenameString As String
    Dim TypeArray As Variant
    Dim DemoElement As Variant
    Dim SearchArray As Variant
    Dim CountArray As Variant
    Dim DemoDict As Object
    'Dim DemoDict As Scripting.Dictionary

    Set DemoDict = CreateObject("Scripting.Dictionary")
    'Set DemoDict = New Scripting.Dictionary
    
    'Ignore case
    DemoDict.CompareMode = vbTextCompare
    
    'Read values into an array. This should always be 1 dimensional since only one column is passed
    'Need to loop for any non-contiguous range
    ReDim SearchArray(1 To SearchRange.Cells.Count)
    
    i = 1
    For Each c In SearchRange
        SearchArray(i) = Trim(c.Value)
        i = i + 1
    Next c
    
    'Tabulating credits is done differently since we're putting integers into buckets
    If SearchType = "Credits" Then
        ReDim CountArray(1 To 4, 1 To 2) 'Should probably make this programmatic
            CountArray(1, 1) = "<45"
            CountArray(2, 1) = "45-90"
            CountArray(3, 1) = ">90"
            CountArray(4, 1) = "Other Credits"
            
        For i = 1 To UBound(SearchArray)
            If SearchArray(i) = "" Then
                CountArray(4, 2) = CountArray(4, 2) + 1
            Else
                j = SearchArray(i)
            End If
            
            If IsEmpty(j) Or j = 0 Then 'VBA will return true for <45 on empty cells
                CountArray(4, 2) = CountArray(4, 2) + 1
            ElseIf Not (IsNumeric(j)) Then
                CountArray(4, 2) = CountArray(4, 2) + 1
            ElseIf j < 45 Then
                CountArray(1, 2) = CountArray(1, 2) + 1
            ElseIf j <= 90 Then
                CountArray(2, 2) = CountArray(2, 2) + 1
            ElseIf j > 90 Then
                CountArray(3, 2) = CountArray(3, 2) + 1
            Else
                CountArray(4, 2) = CountArray(4, 2) + 1 'To catch anything else
            End If
        Next i

        GoTo ReturnArray
    End If

    'Rename the SearchType if it's ethnicity. Resolve this problem in the future
    If SearchType = "Ethnicity" Then
        SearchType = "Race"
    End If

    'Read into the dictionary
    For i = 1 To UBound(SearchArray)
        DemoElement = SearchArray(i)
        
        If Not DemoDict.Exists(DemoElement) Then
            DemoDict.Add DemoElement, 1
        Else
            DemoDict(DemoElement) = DemoDict(DemoElement) + 1
        End If
    Next i

    'First Generation and Low Income don't have an other category
    If SearchType = "First Generation" Or SearchType = "Low Income" Then
        If DemoDict.Exists("Yes") Then
            DemoDict.Key("Yes") = SearchType
        End If
        
        GoTo SkipOther
    End If

    'Rename the "Other" key. Insert one if it doesn't exist
    If Not DemoDict.Exists("Other") Then
        DemoDict.Add "Other", 0
    End If
    
    'If DemoDict.Exists("Other") Then
        RenameString = "Other " & SearchType
        DemoDict.Key("Other") = RenameString
    'End If
    
SkipOther:
    'Read into an array for counting
    ReDim CountArray(1 To DemoDict.Count, 1 To 2)
    
    i = 0
    For Each DemoElement In DemoDict.Keys
        i = i + 1
        
        CountArray(i, 1) = DemoElement
        CountArray(i, 2) = DemoDict(DemoElement)
    Next DemoElement

ReturnArray:
    DemoTabulateWin = CountArray

Footer:

End Function

Function DemoTabulateMac(SearchRange As Range, SearchType As String) As Variant
'Returns an array with the values in the passed range tabulated
'No dictionary for Mac
'SearchType is for renaming "Other" values. This is why this is done piecemeal instead of all at once

    Dim HeaderRange As Range
    Dim c As Range
    Dim i As Long
    Dim j As Long
    Dim RenameString As String
    Dim ListName As String
    Dim TypeArray As Variant
    Dim DemoElement As Variant
    Dim SearchArray As Variant
    Dim CountArray As Variant
    
    'Read values into an array. This should always be 1 dimensional since only one column is passed
    'Need to loop for any non-contiguous range
    ReDim SearchArray(1 To SearchRange.Cells.Count)
    
    i = 1
    For Each c In SearchRange
        SearchArray(i) = c.Value
        i = i + 1
    Next c
    
    'Define the list of values to pull from
    ListName = SearchType & "List"
    Set HeaderRange = Range(ListName)
     
     'Rename the SearchType if it's ethnicity. Resolve this problem in the future
    If SearchType = "Ethnicity" Then
        SearchType = "Race"
    End If
    
    'Create the array for counting the values
    ReDim CountArray(1 To HeaderRange.Cells.Count, 1 To 2)
    
    i = 1
    For Each c In HeaderRange
        CountArray(i, 1) = c.Value
        
        'Rename the "Other" category, if present
        If c.Value = "Other" Then
            CountArray(i, 1) = "Other " & SearchType
        End If
        
        i = i + 1
    Next c
    
    'Different procedure for credits since we're working with buckets
    If SearchType = "Credits" Then
        GoTo TabulateCredits
    End If
    
    'Loop through the Search array and count instances
    For i = LBound(SearchArray) To UBound(SearchArray)
        For j = LBound(CountArray, 1) To UBound(CountArray, 1)
            If SearchArray(i) = CountArray(j, 1) Then
                CountArray(j, 2) = CountArray(j, 2) + 1
                
                GoTo MatchFound
            End If
        Next j
        
        'If not found, put in the "Other" category, if any
        If InStr(CountArray(j - 1, 1), "Other") > 0 Then
            CountArray(j - 1, 2) = CountArray(j - 1, 2) + 1
        End If
MatchFound:
    Next i
    
    GoTo ReturnArray
    
TabulateCredits:
    'Tabulating credits is done differently since we're putting integers into buckets
    For i = LBound(SearchArray) To UBound(SearchArray)
        j = SearchArray(i)
        
        If IsEmpty(j) Or j = 0 Then 'VBA will return true for <45 on empty cells
            CountArray(4, 2) = CountArray(4, 2) + 1
        ElseIf Not (IsNumeric(j)) Then
            CountArray(4, 2) = CountArray(4, 2) + 1
        ElseIf j < 45 Then
            CountArray(1, 2) = CountArray(1, 2) + 1
        ElseIf j <= 90 Then
            CountArray(2, 2) = CountArray(2, 2) + 1
        ElseIf j > 90 Then
            CountArray(3, 2) = CountArray(3, 2) + 1
        Else
            CountArray(4, 2) = CountArray(4, 2) + 1 'To catch anything else
        End If
    Next i
    
ReturnArray:
    DemoTabulateMac = CountArray
    
Footer:

End Function

Sub TabulateActivity(LabelCell As Range)
'Tabulates a given activity and pushes it to the ReportSheet
'Anything in the LabelCell row is deleted before retabulation
'We don't tabulate from the activity sheet so that we don't tabulate unsaved changes. Everything comes from the RecordsSheet

    Dim RosterSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim ReportLabelRange As Range
    Dim ReportTotalRange As Range
    Dim TabulateRange As Range
    Dim c As Range
    Dim RosterTable As ListObject
    Dim ReportTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RosterSheet = Worksheets("Roster Page")

    'Make sure we have students and activities in records
    If CheckRecords(RecordsSheet) > 1 Then
        Call ReportClearAll
    
        GoTo Footer
    'Make sure the RosterSheet has a table with students
    ElseIf CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    'Make sure there's a table on the ReportSheet
    ElseIf CheckTable(ReportSheet) > 2 Then
        Call CreateReportTable
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set ReportTable = ReportSheet.ListObjects(1)
    
    'Make sure the activity is in the RecordsSheet
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet, LabelCell)
        If RecordsLabelRange Is Nothing Then 'This shouldn't happen
            GoTo Footer
        End If

    'Find on the ReportSheet
    Set ReportLabelRange = FindReportLabel(ReportSheet, LabelCell)
        If ReportLabelRange Is Nothing Then 'This shouldn't happen
            GoTo Footer
        End If
        
    'Clear out everything currently in the row
    If RemoveFromReport(ReportLabelRange) <> 1 Then
        GoTo Footer
    End If

    'Define the range to tabulate
    Set TabulateRange = FindTabulateRange(RosterSheet, RecordsSheet, RecordsLabelRange)
        If TabulateRange Is Nothing Then 'This happens when there are no students, i.e. after clearing the roster
            GoTo Footer
        End If
   
   'Pass for tabulation add Total and Notes
   Set c = ReportTable.HeaderRowRange.Find("Total", , xlValues, xlWhole)
   Set ReportTotalRange = ReportSheet.Cells(ReportLabelRange.Row, c.Column)
   
   ReportTotalRange.Value = TabulateRange.Cells.Count
   ReportLabelRange.Offset(0, 1).Value = RecordsLabelRange.Offset(1, 0).Value 'Both are one cell away
   Call TabulateHelper(ReportSheet, RosterSheet, ReportLabelRange, TabulateRange)
   
Footer:
    Call ResetProtection

End Sub

Sub TabulateAll()
'Tabulates all practices
'Called when the roster changes or from the ReportSheet

    Dim RecordsSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim ReportLabelRange As Range
    Dim c As Range
    Dim ReportTable As ListObject
    
    Set RecordsSheet = Worksheets("Records Page")
    Set ReportSheet = Worksheets("Report Page")
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportLabelRange = ReportTable.ListColumns("Practice").DataBodyRange
    
    'If there aren't students and activities, then clear the Report instead
    If CheckRecords(RecordsSheet) <> 1 Then
        Call ReportClearAll
        
        GoTo Footer
    End If
    
    For Each c In ReportLabelRange
        Call TabulateActivity(c)
    Next c
    
Footer:
    Call ResetProtection
    
End Sub

Sub TabulateHelper(ReportSheet As Worksheet, RosterSheet As Worksheet, PasteCell As Range, Optional NameRange As Range)
'Tabulates every category since this needs to be done both for the totals row and tabulating activities
'Tabulates for all students on the RosterSheet by default
'Passing NameRange limits tabulation to only those names. This should be a range on the RosterSheet

    Dim SearchTermRange As Range
    Dim TempRange As Range
    Dim c As Range
    Dim i As Long
    Dim SearchTerm As String
    Dim SearchTermArray() As Variant
    Dim TempArray() As Variant
    Dim RosterTable As ListObject
    
    Set SearchTermRange = Range("TabulateTermList")
    Set RosterTable = RosterSheet.ListObjects(1)
    
    'Populate an array for the terms to tabulate and the reference lists to use
        '(i, 1) Terms
        '(i, 2) Lists
    ReDim SearchTermArray(1 To SearchTermRange.Cells.Count, 1 To 2)
    
    i = 1
    For Each c In SearchTermRange
        SearchTermArray(i, 1) = c.Value
        SearchTermArray(i, 2) = c.Offset(0, 1).Value
        
        i = i + 1
    Next c
    
    'Loop through and pass each term for tabulation
    For i = 1 To UBound(SearchTermArray)
        SearchTerm = SearchTermArray(i, 1)
        
        Set TempRange = RosterTable.ListColumns(SearchTerm).DataBodyRange
        
        'Only grab the passed students if a range was passed
        If Not NameRange Is Nothing Then
            Set c = NameRange.Offset(0, TempRange.Column - NameRange.Column)
            Set TempRange = c
        End If
    
        'Push to the ReportSheet
        Erase TempArray
        TempArray = DemoTabulate(TempRange, SearchTerm)
        
        Call CopyToReport(ReportSheet, PasteCell, TempArray)
    Next i
    
End Sub

Sub TabulateReportTotals()
'Called from a button and when the roster is parsed
'Not entirely programmatic yet. Adding a tabulation table on the RefSheet might work

    Dim RosterSheet As Worksheet
    Dim ReportSheet As Worksheet
    Dim RecordsSheet As Worksheet
    Dim ReportHeaderRange As Range
    Dim c As Range
    Dim RosterTable As ListObject
    Dim ReportTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set ReportSheet = Worksheets("Report Page")
    Set RecordsSheet = Worksheets("Records Page")

    'Make sure we have students in records
    If CheckRecords(RecordsSheet) > 2 Then
        Call ReportClearAll
    
        GoTo Footer
    'Make sure the RosterSheet has a table with students
    ElseIf CheckTable(RosterSheet) > 2 Then
        GoTo Footer
    'Make sure there's a table on the ReportSheet with rows
    ElseIf CheckTable(ReportSheet) > 2 Then
        Call CreateReportTable
    End If

    Set RosterTable = RosterSheet.ListObjects(1)
    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportHeaderRange = ReportTable.HeaderRowRange
    
    Call UnprotectSheet(ReportSheet)
    
    'Clear the totals row
    Call ReportClearTotals
    
    'Find the cell under the "Total" header and pass for tabulating
    Set c = ReportHeaderRange.Find("Total", , xlValues, xlWhole).Offset(1, 0)
    
    Call TabulateHelper(ReportSheet, RosterSheet, c)
    
    'Insert total students
    c.Value = RosterTable.ListRows.Count
    
Footer:

End Sub
