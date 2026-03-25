Attribute VB_Name = "CopySubs"
Option Explicit

Function CopyToRecords(RosterSheet As Worksheet, RecordsSheet As Worksheet) As Range
'Copies new students from the RosterSheet to the RecordsSheet
'Duplicates are ignored, blank rows and duplicates are deleted
'Returns the range of copied names

    Dim FullNameRange As Range
    Dim DelRange As Range
    Dim RosterNameRange As Range
    Dim RecordsNameRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim d As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim i As Long
    Dim RosterTable As ListObject
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set RosterNameRange = RosterTable.ListColumns("First").DataBodyRange
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    
    If RosterNameRange Is Nothing Then
        GoTo Footer
    End If
    
    'Find the bottom of the name list
    Set c = RecordsSheet.Range("A:A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set PasteRange = c.Offset(1, 0).Resize(1, 2) 'For first and last name
    
    'If there are no names, simply copy everything
    If c.Value = "H BREAK" Then
        Set d = RosterNameRange
        
        GoTo CopyUnique
    End If
    
    'Find any students on the RecordsSheet but not the RosterSheet
    Set c = FindUnique(RecordsNameRange, RosterNameRange)
        If Not c Is Nothing Then
            Set DelRange = c
            Set FullNameRange = RecordsNameRange.Resize(RecordsNameRange.Rows.Count, 2) 'first and last names
        
            Call RemoveRows(RecordsSheet, FullNameRange, DelRange)
        End If
    
    'Find all non-duplicative students and copy over
    Set d = FindUnique(RosterNameRange, RecordsNameRange)
        If d Is Nothing Then
            GoTo CleanUp
        End If

CopyUnique:
    
    i = 0
    For Each c In d
        Set CopyRange = Union(c, c.Offset(0, 1))
    
        PasteRange.Offset(i, 0).Value = CopyRange.Value
        
        Set ReturnRange = BuildRange(PasteRange.Offset(i, 0), ReturnRange)
        i = i + 1
    Next c

    'Return
    Set CopyToRecords = ReturnRange

CleanUp:
    'Delete blanks and duplicates
    Set RecordsNameRange = FindRecordsName(RecordsSheet)
    Set FullNameRange = RecordsNameRange.Resize(RecordsNameRange.Rows.Count, 2) 'Both columns
    
    Call RemoveBadRows(RecordsSheet, FullNameRange, RecordsNameRange)
    
    'Tabulate student totals
    Call TabulateReportTotals
       
Footer:

End Function

Function CopyToReport(ReportSheet As Worksheet, PasteCell As Range, PasteArray As Variant) As Range
'Copies values passed in the array to the row of the PasteCell
'PasteArray has the header names in (i, 1), the values in (i, 2)
'Can pass any column in the report header

    Dim ReportHeaderRange As Range
    Dim ReturnRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim OtherIndex As Long
    Dim OtherString As String
    Dim HeaderString As String
    Dim ReportTable As ListObject

    Set ReportTable = ReportSheet.ListObjects(1)
    Set ReportHeaderRange = ReportTable.HeaderRowRange
    
    Call UnprotectSheet(ReportSheet)
    
    'Loop through the paste array and find corresponding headers
    For i = LBound(PasteArray) To UBound(PasteArray)
        HeaderString = PasteArray(i, 1)
        
        'Grab the index of the "Other" category, if any
        If InStr(1, HeaderString, "Other") > 0 Then
            OtherIndex = i
            OtherString = PasteArray(OtherIndex, 1)
        End If
        
        'Find the header, insert value if found
        Set c = ReportHeaderRange.Find(HeaderString, , xlValues, xlWhole)
        
        If Not c Is Nothing Then
            'Paste under the matching header
            Set d = ReportSheet.Cells(PasteCell.Row, c.Column)
            d.Value = PasteArray(i, 2)
            Set ReturnRange = BuildRange(d, ReturnRange)
            
            'Get rid of zeroes
            If d.Value = 0 Then
                d.ClearContents
            End If
        Else
            'Sum up elements that aren't found in the header
            If IsNumeric(PasteArray(i, 2)) = True Then
                j = j + PasteArray(i, 2) 'Using this so we don't run into problems with strings
            End If
        End If
    Next i

    'Skip adding up "others" if no other category was found, such as with Low Income
    If OtherIndex < 1 Then
        GoTo ReturnRange
    End If

    'Push all leftover elements into the "Other" category
    'This will allow the list of categories to change in the future
    If j > 0 Then
        PasteArray(OtherIndex, 2) = PasteArray(OtherIndex, 2) + j
        Set c = ReportHeaderRange.Find(OtherString, , xlValues, xlWhole)
        Set d = ReportSheet.Cells(PasteCell.Row, c.Column)
        
        d.Value = PasteArray(OtherIndex, 2)
        Set ReturnRange = BuildRange(d, ReturnRange)
    End If

ReturnRange:
    If Not ReturnRange Is Nothing Then
        Set CopyToReport = ReturnRange
    End If

Footer:

End Function
