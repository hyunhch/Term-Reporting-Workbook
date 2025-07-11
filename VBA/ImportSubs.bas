Attribute VB_Name = "ImportSubs"
Option Explicit

Public OldMatchArray As Variant
Public NewMatchArray As Variant

Function ImportCompareHeaders(ReturnArray As Variant, CompareArray As Variant) As Variant
'Compares two arrays created from table headers and matches the header strings
'If a match it found, element (3, i) takes the address of the matched cell
'If no match found, it retains value of 0
'Returns nothing on error
    
    Dim i As Long
    Dim j As Long
    Dim CompareString As String
    Dim MatchString As String
   
    'Loop through and compare each element in the 1st array against the 2nd
    For i = 1 To UBound(ReturnArray, 2)
        MatchString = Trim(ReturnArray(1, i)) 'Remove whitespace
                
        For j = 1 To UBound(CompareArray, 2)
            CompareString = Trim(CompareArray(1, j))
            
            If CompareString = MatchString Then
                ReturnArray(3, i) = CompareArray(2, j) 'Record the address
                
                GoTo NextElement
            End If
        Next j
NextElement:
    Next i

    'Return
    ImportCompareHeaders = ReturnArray

Footer:

End Function

Function ImportCopyColumns(CopySheet As Worksheet, PasteSheet As Worksheet) As Long
'Copies over columns, matching headers
'Adds new columns if needed
'Returns 1 if successful, 0 if not, nothing on error

    Dim c As Range
    Dim d As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim i As Long
    Dim LRow As Long
    
    ImportCopyColumns = 0
    
    'Find the last used row. Crude, but will work. Checking that there are rows to copy should have already happened
    LRow = CopySheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    'Each array already has the address information of where it needs to go. Loop through and paste
    For i = 1 To UBound(OldMatchArray, 2)
        If OldMatchArray(3, i) = 0 Or OldMatchArray(3, i) = 1 Then
            GoTo NextPaste
        End If
    
        Set c = CopySheet.Range(OldMatchArray(2, i))
        Set d = PasteSheet.Range(OldMatchArray(3, i))
            'Rename in case of matching
            c.Value = OldMatchArray(1, i)
        
        Set CopyRange = CopySheet.Range(c, c.Offset(LRow - 1, 0))
        Set PasteRange = d.Resize(CopyRange.Rows.Count, 1)
            PasteRange.Value = CopyRange.Value
NextPaste:
    Next i

    ImportCopyColumns = 1

Footer:

End Function

Function ImportMakeHeaderArray(HeaderRange As Range) As Variant
'Takes a header range and returns a 2D array
'Returns nothing on error
        '(1, i) - Header string
        '(2, i) - Header cell address
        '(3, i) - 0, will take the address of a matched cell
        
    Dim c As Range
    Dim i As Long
    Dim ReturnArray As Variant
    
    'Make sure the header range exists
    If HeaderRange Is Nothing Then
        GoTo Footer
    End If
    
    'Build the array and return
    ReDim ReturnArray(1 To 3, 1 To HeaderRange.Cells.Count)

    i = 1
    For Each c In HeaderRange
        ReturnArray(1, i) = c.Value
        ReturnArray(2, i) = c.Address
        ReturnArray(3, i) = 0
        
        i = i + 1
    Next c

    ImportMakeHeaderArray = ReturnArray

Footer:

End Function

Function ImportRecords(OldRecordsSheet As Worksheet, NewRecordsSheet As Worksheet) As Long
'Compares the activities given on the RecordsSheet in the two workbooks
'Identifies any activities that are different and prompt the user to match them
'Returns 1 if successful, 0 if there's an issue, nothing on error

    Dim NewRecordsNameRange As Range
    Dim OldRecordsNameRange As Range
    Dim NewRecordsLabelRange As Range
    Dim OldRecordsLabelRange As Range
    Dim CopyRange As Range
    Dim PasteRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    Dim IsMissing As Boolean
    
    'First make sure there are activity labels in both books
    If CheckRecords(OldRecordsSheet) > 2 Then
        'Skip if there are no activites to copy over
        GoTo ImportFailed
    ElseIf CheckRecords(NewRecordsSheet) > 2 Then
        'This should already be there
        Call RecordsSheetText
    End If
    
    'Define names to copy
    Set OldRecordsNameRange = FindRecordsName(OldRecordsSheet)
        If OldRecordsNameRange.Cells.Count = 1 Then
            If OldRecordsNameRange.Value = "H BREAK" Then 'Nothing to copy over
                GoTo ImportFailed
            End If
        End If
    
    'Paste over
    Set c = NewRecordsSheet.Range("A:A").Find("H BREAK", , xlValues, xlWhole).Offset(1, 0)
    Set CopyRange = OldRecordsNameRange.Resize(OldRecordsNameRange.Rows.Count, 2)
    Set PasteRange = c.Resize(CopyRange.Rows.Count, 2)
    Set NewRecordsNameRange = PasteRange
        PasteRange.Value = CopyRange.Value
    
    'Grab the activities in the both workbooks
    Set OldRecordsLabelRange = FindRecordsLabel(OldRecordsSheet)
        If OldRecordsLabelRange Is Nothing Then
            GoTo ImportFailed
        End If
    
    Set NewRecordsLabelRange = FindRecordsLabel(NewRecordsSheet)
        If NewRecordsLabelRange Is Nothing Then 'This is redundant
            GoTo ImportFailed
        End If
    
    'Read into arrays
    'Erase OldMatchArray
    'Erase NewMatchArray
    
    ReDim OldMatchArray(1 To 3, 1 To OldRecordsLabelRange.Cells.Count)
    ReDim NewMatchArray(1 To 3, 1 To NewRecordsLabelRange.Cells.Count)
        '1: Activity name
        '2: Address of the cell
        '3: Address of matched cell, 0 if no match found
    
    OldMatchArray = ImportMakeHeaderArray(OldRecordsLabelRange)
    NewMatchArray = ImportMakeHeaderArray(NewRecordsLabelRange)
    
    'Compare the two arrays
    OldMatchArray = ImportCompareHeaders(OldMatchArray, NewMatchArray)
    NewMatchArray = ImportCompareHeaders(NewMatchArray, OldMatchArray)
    
    'Loop through the arrays and see if there's a 0 in the 3rd column
    IsMissing = False
    
    For i = 1 To UBound(OldMatchArray, 2)
        If OldMatchArray(3, i) = 0 Then
            IsMissing = True
            
            Exit For
        End If
    Next i
    
    'Both arrays need to have unmatched values
    If IsMissing = False Then
        GoTo SkipMatching
    End If
    
    IsMissing = False
    
    For j = 1 To UBound(NewMatchArray, 2)
        If NewMatchArray(3, j) = 0 Then
            IsMissing = True
            
            Exit For
        End If
    Next j
    
    'Bring up userform to match activities
     If IsMissing = True Then
        ImportMatchForm.Show
    End If
     
SkipMatching:
    'Copy over each column that has a destination in the array. This may reorder the columns or not copy some over, which is important if the reference list changes
    For i = 1 To UBound(OldMatchArray, 2)
        If OldMatchArray(3, i) = 0 Or OldMatchArray(3, i) = 1 Then
            GoTo NextPaste
        End If
        
        'The 2nd y element is the original cell, the 3rd is the desination cell
        Set c = OldRecordsSheet.Range(OldMatchArray(2, i))
        Set d = NewRecordsSheet.Range(OldMatchArray(3, i))
        Set CopyRange = Intersect(c.EntireColumn, OldRecordsNameRange.EntireRow)
        Set PasteRange = Intersect(d.EntireColumn, NewRecordsNameRange.EntireRow)
            PasteRange.Value = CopyRange.Value
NextPaste:
    Next i
    
    ImportRecords = 1
    GoTo Footer

ImportFailed:
    ImportRecords = 0

Footer:

End Function

Function ImportRoster(OldRosterSheet As Worksheet, NewRosterSheet As Worksheet) As Long
'Copies over all students on the old Roster
'Returns 1 if successful, 0 if there's an issue, nothing on error

    Dim OldRosterRange As Range
    Dim NewRosterRange As Range
    Dim OldHeaderRange As Range
    Dim NewHeaderRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    Dim j As Long
    
    Dim OldRosterTable As ListObject
    Dim NewRosterTable As ListObject
    Dim IsMissing As Boolean
    
    'There should be a roster on the old workbook. Break if there isn't
    If Not OldRosterSheet.ListObjects.Count > 0 Then
        GoTo ImportFailed
    End If
    
    'Make sure there are rows. This should have already happened
    Set OldRosterTable = OldRosterSheet.ListObjects(1)
        If Not OldRosterTable.ListRows.Count > 0 Then
            GoTo ImportFailed
        End If
    
    Set OldRosterRange = OldRosterTable.DataBodyRange
    Set OldHeaderRange = OldRosterTable.HeaderRowRange

    'If there's no table on the new roster, create one. This shouldn't happen
    If Not NewRosterSheet.ListObjects.Count > 0 Then
        Set c = NewRosterSheet.Range("A6")
        NewMatchArray = Application.Transpose(ActiveWorkbook.Names("ColumnNamesList").RefersToRange.Value)
        
        Call ResetTableHeaders(NewRosterSheet, c, NewMatchArray)
        Set NewRosterTable = CreateTable(NewRosterSheet)
        
        Erase NewMatchArray
    Else
        Set NewRosterTable = NewRosterSheet.ListObjects(1)
    End If
    
    Set NewHeaderRange = NewRosterTable.HeaderRowRange

    'Grab the headers from both sheets and their addresses
        '(1, i) - Header string
        '(2, i) - Header cell address
        '(3, i) - Address of header cell of other table, or 1 if no match

    OldMatchArray = ImportMakeHeaderArray(OldHeaderRange)
    NewMatchArray = ImportMakeHeaderArray(NewHeaderRange)
    
    'Compare the two. Unlike with the RecordSheet, we can add additional columns
    OldMatchArray = ImportCompareHeaders(OldMatchArray, NewMatchArray)
    NewMatchArray = ImportCompareHeaders(NewMatchArray, OldMatchArray)
    
    'Loop through the Old array to see if any element (3, i) = 0
    IsMissing = False
    
    For i = 1 To UBound(OldMatchArray, 2)
        If OldMatchArray(3, i) = 0 Then
            IsMissing = True
            
            Exit For
        End If
    Next i
    
    'If there are missing matches, bring up the userform and insert "Add new column" as an option
    If Not IsMissing = True Then
        GoTo CopyOver
    End If
    
    ImportMatchForm.ImportMatchNewListbox.AddItem ("Add new column")
    ImportMatchForm.Show
    
CopyOver:
    Call ImportCopyColumns(OldRosterSheet, NewRosterSheet)
    
    'Remake the table in case new columns were added
    Call CreateTable(NewRosterSheet)
    
    ImportRoster = 1
    GoTo Footer

ImportFailed:
    ImportRoster = 0

Footer:

End Function
