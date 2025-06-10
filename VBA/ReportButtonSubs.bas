Attribute VB_Name = "ReportButtonSubs"
Option Explicit

Sub ReportClearTotalsButton()
'Container for the daughter sub

    Dim ReportSheet As Worksheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    Call UnprotectSheet(ReportSheet)
    
    'Verify that the table is there
    If CheckTable(ReportSheet) > 2 Then
        Call CreateReportTable
        
        GoTo Footer
    End If
    
    Call ReportClearTotals
    
Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ReportClearAllButton()
'Container for the daughter sub

    Dim ReportSheet As Worksheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    Call UnprotectSheet(ReportSheet)
    
    'Verify that the table is there
    If CheckTable(ReportSheet) > 2 Then
        Call CreateReportTable
        
        GoTo Footer
    End If
    
    Call ReportClearAll
    
Footer:
    Call ResetProtection

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub ReportTabulateTotalsButton()
'Calls the sub, here to control when screen updating happens

    Dim ReportSheet As Worksheet

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ReportSheet = Worksheets("Report Page")
    
    Call UnprotectSheet(ReportSheet)
    Call TabulateReportTotals
    
Footer:
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


Sub OpenTabulateActivityButton()
'Checks that there is anything to tabulate first

    Dim RecordsSheet As Worksheet
    
    Set RecordsSheet = Worksheets("Records Page")
        If CheckRecords(RecordsSheet) > 1 Then
            GoTo Footer
        End If
     
    TabulateActivityForm.Show
        
Footer:

End Sub
