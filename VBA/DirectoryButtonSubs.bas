Attribute VB_Name = "DirectoryButtonSubs"
Option Explicit

Sub DirectoryTabulateSchoolsButton()
'Autofills the "School Directory" table on the Directory Sheet
'Pulls from the Roster Page, doesn't automatically call

    Dim RosterSheet As Worksheet
    Dim DirectorySheet As Worksheet
    Dim c As Range
    Dim RosterTable As ListObject
    Dim SchoolTable As ListObject
    
    Set RosterSheet = Worksheets("Roster Page")
    Set DirectorySheet = Worksheets("Directory Page")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Make sure it's College Prep. The button shouldn't appear unless it is
    If GetProgram <> "College Prep" Then
        GoTo Footer
    End If
    
    'Make sure there's a RosterTable with rows and something in the School column
    If CheckTable(RosterSheet) > 2 Then
        MsgBox ("Please parse the roster and try again")
    
        GoTo Footer
    End If
    
    Set RosterTable = RosterSheet.ListObjects(1)
    Set c = RosterTable.ListColumns("School").DataBodyRange.Find("*", , xlValues, xlWhole)
        If c Is Nothing Then
            GoTo Footer
        End If
        
    'Make sure there's a table on the DirectorySheet. It will be the third table. Make this programmatic in the future
    If Not DirectorySheet.ListObjects.Count = 3 Then
        GoTo Footer
    End If
    
    Set SchoolTable = DirectorySheet.ListObjects(3)
    
    'Call the needed functions
    Call TabulateSchools(RosterSheet, DirectorySheet, RosterTable, SchoolTable)
    
Footer:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True


End Sub

