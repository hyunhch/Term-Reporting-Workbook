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
    If CheckReport(ReportSheet) > 2 Then
        Call MakeReportTable
        
        GoTo Footer
    End If
    
    Call ReportClearTotals
    
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
    
    'Pull in information from the cover sheet
    Call ReportCoverInfo(ReportSheet)
    
Footer:
    Call ResetProtection
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
