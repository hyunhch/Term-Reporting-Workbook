Attribute VB_Name = "MakeSubs"
Option Explicit

Function MakeButton(TargetSheet As Worksheet, TargetArray As Variant) As Long
'Makes a button on the passed sheet, an array contains the arguments
'Returns 1 on sucess
    '(1) - Range
    '(2) - OnAction
    '(3) - Caption
    '(4) - Name
    
    Dim TargetRange As Range
    Dim RangeString As String
    Dim OnActionString As String
    Dim CaptionString As String
    Dim NameString As String
    Dim TargetButton As Button
    
    RangeString = TargetArray(1)
    OnActionString = TargetArray(2)
    CaptionString = TargetArray(3)
    NameString = TargetArray(4)
    
    Set TargetRange = TargetSheet.Range(RangeString)
    Set TargetButton = TargetSheet.Buttons.Add(TargetRange.Left, TargetRange.Top, _
        TargetRange.Width, TargetRange.Height)

    With TargetButton
        .OnAction = OnActionString
        .Caption = CaptionString
        .Name = NameString
    End With

    MakeButton = 1

Footer:

End Function

Function MakeReportTable() As ListObject
'Only has a single row

    Dim ReportSheet As Worksheet
    Dim ReportTableRange As Range
    Dim ReportTableStart As Range
    Dim TempRange As Range
    Dim RefRange As Range
    Dim c As Range
    Dim i As Long
    Dim ReportTable As ListObject
    Dim HeadersArray() As Variant
    Dim CoverArray() As Variant

    Set ReportSheet = Worksheets("Report Page")
    Set ReportTableStart = ReportSheet.Range("A:A").Find("Center", , xlValues, xlWhole) 'There is no "Select" column
        If ReportTableStart Is Nothing Then 'If the table headers got messed up
            Set ReportTableStart = ReportSheet.Range("A6")
        End If
      
    'Remove the table and formatting
    Call UnprotectSheet(ReportSheet)
    Call RemoveTable(ReportSheet)
    
    'Read in the headers. It can be discontiguous, so we can't use the Transpose function
    Set RefRange = Range("ReportHeadersList")
    
    ReDim HeadersArray(1 To RefRange.Cells.Count)
    
    i = 1
    For Each c In RefRange
        HeadersArray(i) = c.Value
        
        i = i + 1
    Next c

    Call TableResetHeaders(ReportSheet, ReportTableStart, HeadersArray)
    
    'Define table range and clear formats
    Set ReportTableRange = FindTableRange(ReportSheet)
        ReportSheet.Cells.ClearFormats
        
    'Make a new table and format
    Set ReportTable = ReportSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=ReportTableRange, _
        xlListObjectHasHeaders:=xlYes)
        
    With ReportTable
        .Name = "ReportTable"
        .ShowTableStyleRowStripes = False
        'Add a ListRow if there isn't one
        If Not .ListRows.Count > 0 Then
            .ListRows.Add
        End If
    End With
        
    'Format and return
    Call TableFormatReport(ReportSheet, ReportTable)
    
    Set MakeReportTable = ReportTable
    
Footer:

End Function

Function MakeRosterTable(RosterSheet As Worksheet) As ListObject
'Called when parsing the roster
'Returns the RosterTable if successful
'Returns nothing on error

    Dim RefRange As Range
    Dim c As Range
    Dim i As Long
    Dim RosterTable As ListObject
    Dim HeaderArray As Variant
    
    'Read in the headers. It can be discontiguous, so we can't use the Transpose function
    Set RefRange = Range("RosterHeadersList")
    
    ReDim HeaderArray(1 To RefRange.Cells.Count)
    
    i = 1
    For Each c In RefRange
        HeaderArray(i) = c.Value
        
        i = i + 1
    Next c
    
    'Make new table
    Set c = RosterSheet.Range("A6")
    Set RosterTable = MakeTable(RosterSheet, HeaderArray, "RosterTable", c)
        If RosterTable.DataBodyRange Is Nothing Then
            GoTo Footer
        End If
    
    Call TableFormat(RosterSheet, RosterTable)
    
    If Not RosterTable.ListColumns("Select").DataBodyRange Is Nothing Then
        Call AddMarlettBox(RosterTable.ListColumns("Select").DataBodyRange)
    End If
    
    Set MakeRosterTable = RosterTable
    
Footer:
    
End Function

Function MakeTable(TargetSheet As Worksheet, HeaderArray As Variant, Optional TableName As String, Optional TargetRange As Range) As ListObject
'Creates a table on the passed sheet with the passed headers
'Optionally can pass the name for the new table or a custom range
'If not range is passed, it will default to where FindTableRange() returns, which should be "A6"
'Returns a table object on success, returns nothing on error

    Dim TableRange As Range
    Dim c As Range
    Dim NewTable As ListObject
    
    Call UnprotectSheet(TargetSheet)
    
    'Unlist any existing table
    TargetSheet.AutoFilterMode = False
    Call RemoveTable(TargetSheet)
    
    'Insert the passed headers
    Call TableResetHeaders(TargetSheet, TargetRange(1, 1), HeaderArray)
    
    'Define range to use
    Set TableRange = FindTableRange(TargetSheet)
    
    If TableRange Is Nothing Then
        MsgBox ("There was a problem creating a table on sheet " & TargetSheet.Name)
        
        GoTo Footer
    End If
    
    'If there were more passed headers than the original range, resize
    If UBound(HeaderArray) > TableRange.Columns.Count Then
        Set c = TableRange.Resize(TableRange.Rows.Count, UBound(HeaderArray, 2))
        Set TableRange = c
    End If
    
    TableRange.ClearFormats
    
    'Create a table
    Set NewTable = TargetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=TableRange, _
        xlListObjectHasHeaders:=xlYes)
        
    NewTable.ShowTableStyleRowStripes = False

    'Assign a name if passed
    If Len(TableName) > 0 Then
        NewTable.Name = TableName
    End If

    'Removed getting rid of blank rows
    'Better to use a RemoveBlanks function that can be pointed to any column, rather than just the first name column of a table
    'Put in Marlett boxes. I had taken this out but can't remember why
    
    'ReturnTable
    Set MakeTable = NewTable
    
Footer:

End Function
