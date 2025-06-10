VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TabulateActivityForm 
   Caption         =   "Tabulate Activity"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670.001
   OleObjectBlob   =   "TabulateActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TabulateActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TabulateActivityCancelButton_Click()
'Hide the form

    TabulateActivityForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub TabulateActivityConfirmAllButton_Click()
'Tabulate everything displayed, regardless of selection

    Dim RecordsSheet As Worksheet
    Dim LabelCell As Range
    Dim i As Long
    Dim LabelString As String

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RecordsSheet = Worksheets("Records Page")
    Set LabelCell = RecordsSheet.Range("A1")

    'First tabualate the totals, then everything
    Call TabulateReportTotals
    Call TabulateAll
    
    TabulateActivityForm.Hide
   
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub TabulateActivityConfirmButton_Click()
'Recreate an activity sheet with the activity information and attendance
    
    Dim RecordsSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim c As Range
    Dim PracticeString As String
    Dim i As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Make sure an activity has been selected
    i = TabulateActivityListBox.ListIndex
        If i = -1 Then
            GoTo Footer
        End If

    Set RecordsSheet = Worksheets("Records Page")
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)

    'Loop through selected items
    For i = 0 To Me.TabulateActivityListBox.ListCount - 1
        If Me.TabulateActivityListBox.Selected(i) Then
            PracticeString = Me.TabulateActivityListBox.List(i, 0)
            Set c = RecordsLabelRange.Find(PracticeString, , xlValues, xlWhole)
            
            Call TabulateActivity(c)
        End If
    Next i
    
    'Tabulate the totals
    Call TabulateReportTotals
    
    TabulateActivityForm.Hide
   
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub TabulateActivityFilterTextBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & TabulateActivityFilterTextBox.Text & "*")
    Call TabulateActivityListBoxPopulate
    
    With TabulateActivityListBox
        For i = .ListCount - 1 To 0 Step -1
            If (Not (LCase(.List(i, 0)) Like testString)) _
            And (Not (LCase(.List(i, 1)) Like testString)) Then
                .RemoveItem i
            End If
        Next i
    End With
    
End Sub

Private Sub UserForm_Activate()
'Clear fields and define dimensions

    Me.TabulateActivityFilterTextBox.Value = ""
    
    TabulateActivityForm.Height = 300
    TabulateActivityForm.Width = 445
    
    TabulateActivityListBox.ColumnCount = 2
    TabulateActivityListBox.ColumnWidths = "220, 220"

    Call TabulateActivityListBoxPopulate
    
Footer:

End Sub

Private Sub UserForm_Deactivate()
'Bring up the Report Page and enable events

    Dim ReportSheet As Worksheet
    
    Set ReportSheet = Worksheets("Report Page")
    
    ReportSheet.Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub TabulateActivityListBoxPopulate()
'Populates the listbox with activities that have been completed

    Dim RecordsSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim RefLabelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
 
    Set RecordsSheet = Worksheets("Records Page")
    
    TabulateActivityListBox.Clear
    
    'Make columns in the list box
    With TabulateActivityListBox
        .ColumnCount = 2
        .ColumnWidths = "220, 220"
    
        'Define the lists of activities
        Set RefLabelRange = Range("ActivitiesList")
        Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
            If RecordsLabelRange(1, 1) = "V BREAK" Then
                GoTo Footer
            End If
        
        'Loop through to see which activies have saved attendance. Doing this here and not on ReportSheet because untabulated activities can be saved
        i = 0
        For Each c In RecordsLabelRange
            Set d = FindRecordsAttendance(RecordsSheet, , c)
            
            If Not d Is Nothing Then
                If IsChecked(d, "All") = True Then
                    .AddItem c.Value
                    .List(i, 1) = RefLabelRange.Find(c, , xlValues, xlWhole).Offset(0, -1)
                
                    i = i + 1
                End If
            End If
        Next c
    End With

Footer:

End Sub
