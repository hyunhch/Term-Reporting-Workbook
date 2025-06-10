Attribute VB_Name = "GetSubs"
Option Explicit

Function GetVersion() As String
'Returns the version listed at the end of the file name

    Dim FileName As String
    Dim TempName As String

    FileName = ThisWorkbook.Name
    TempName = Left(FileName, InStrRev(FileName, ".") - 1)
    GetVersion = Right(TempName, Len(TempName) - InStrRev(TempName, " "))

End Function

Function GetCoverInfo() As Variant
'Grabs the name, date, center, program, and version from the CoverSheet
'Returns an array with each value
'Returns nothing if some fields are missing

    Dim CoverSheet As Worksheet
    Dim c As Range
    Dim i As Long
    Dim TempArray() As Variant
    
    Set CoverSheet = Worksheets("Cover Page")
    Set c = CoverSheet.Range("A1")
    
    'The first five rows contain information in this order
        'Program/report type
        'Version
        'Name
        'Date
        'Center
    'Make this programmatic in the future
        
    'Check that everything has been filled out
    If CheckCover <> 1 Then
        GoTo Footer
    End If
    
    ReDim TempArray(1 To 5, 1 To 2)
        TempArray(1, 1) = "Program"
        TempArray(1, 2) = c.Offset(0, 0).Value
        
        TempArray(2, 1) = "Version"
        TempArray(2, 2) = c.Offset(1, 0).Value
        
        TempArray(3, 1) = "Name"
        TempArray(3, 2) = c.Offset(2, 1).Value
        
        TempArray(4, 1) = "Date"
        TempArray(4, 2) = c.Offset(3, 1).Value
        
        TempArray(5, 1) = "Center"
        TempArray(5, 2) = c.Offset(4, 1).Value

    'Return
    GetCoverInfo = TempArray

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
    ReDim ReadyArray(1 To 7, 1 To 2)
        ReadyArray(1, 1) = "Cover Page"
        ReadyArray(2, 1) = "Roster Page"
        ReadyArray(3, 1) = "Records Page"
        ReadyArray(4, 1) = "Report Page"
        ReadyArray(5, 1) = "Narrative Page"
        ReadyArray(6, 1) = "Directory Page"
        ReadyArray(7, 1) = "Other Page"
         
    'Go through each sheet
        ReadyArray(1, 2) = CheckCover
        ReadyArray(2, 2) = CheckTable(RosterSheet)
        ReadyArray(3, 2) = CheckRecords(RecordsSheet)
        ReadyArray(4, 2) = CheckReport(ReportSheet)
        ReadyArray(5, 2) = 0 'Figure out how to verify these
        ReadyArray(6, 2) = 0
        ReadyArray(7, 2) = 0
        
    GetReadyToExport = ReadyArray
        
Footer:

End Function
