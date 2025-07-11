VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadActivityForm 
   Caption         =   "Load Activity"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670.001
   OleObjectBlob   =   "LoadActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadActivityCancelButton_Click()
'Hide the form

    LoadActivityForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityConfirmButton_Click()
'Recreate an activity sheet with the activity information and attendance
    
    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim c As Range
    Dim i As Long
    Dim PracticeString As String
    Dim CategoryString As String
    Dim NotesString As String
    Dim InfoArray() As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Make sure an activity has been selected
    i = LoadActivityListBox.ListIndex
    
    If i = -1 Then
        GoTo Footer
    End If

    Set RecordsSheet = Worksheets("Records Page")
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)

    'Loop through selections
    For i = 0 To Me.LoadActivityListBox.ListCount - 1
        If Not Me.LoadActivityListBox.Selected(i) Then
            GoTo NextRow
        End If
        
        PracticeString = LoadActivityListBox.List(i, 0)
        
        'See if there's an open Activity sheet
        Set ActivitySheet = FindActivitySheet(PracticeString)
            If Not ActivitySheet Is Nothing Then
                GoTo NextRow
            End If
        
        'Grab activity information. Notes come from the RecordsSheet
        Set c = RecordsLabelRange.Find(PracticeString, , xlValues, xlWhole).Offset(1, 0)
        
        CategoryString = LoadActivityListBox.List(i, 1)
        NotesString = c.Value
        
        'Create the array to pass
        ReDim InfoArray(1 To 3, 1 To 3)
        
        InfoArray(1, 1) = "Practice"
        InfoArray(2, 1) = "Category"
        InfoArray(3, 1) = "Notes"
        
        InfoArray(1, 2) = "A1"
        InfoArray(2, 2) = "A2"
        InfoArray(3, 2) = "A3"
    
        InfoArray(1, 3) = PracticeString
        InfoArray(2, 3) = CategoryString
        InfoArray(3, 3) = NotesString
        
        'Create a new sheet
        Call ActivityNewSheet(InfoArray, "Load")
        
NextRow:
    Next i
           
    LoadActivityForm.Hide
   
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityDeleteButton_Click()
'Delete the selected activities, removing it from the attendance and label sheets
    
    Dim RecordsSheet As Worksheet
    Dim TempLabelRange As Range
    Dim DelConfirm As Long
    Dim i As Long
    Dim j As Long
    Dim SelectedLabel As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set RecordsSheet = Worksheets("Records Page")

    'Make sure an activity is selected
    If LoadActivityListBox.ListIndex = -1 Then
        GoTo Footer
    End If
    
    'Count if one or multiple are selected
    j = 0
    For i = 0 To Me.LoadActivityListBox.ListCount - 1
        If Me.LoadActivityListBox.Selected(i) Then
            j = j + 1
        End If
    Next i
    
    'Give a warning
    If j = 1 Then
        DelConfirm = MsgBox("Are you sure you want to delete this activity? " & vbCr & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    Else
        DelConfirm = MsgBox("Are you sure you want to delete these activities? " & vbCr & _
        "This cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2)
    End If
    
    If DelConfirm <> vbYes Then
        GoTo Footer
    End If
    
    'Loop throughthe listbox and delete all selected item
    j = Me.LoadActivityListBox.ListCount - 1
    For i = j To 0 Step -1
        If Me.LoadActivityListBox.Selected(i) Then
            SelectedLabel = Me.LoadActivityListBox.List(i, 0)
            Set TempLabelRange = FindRecordsLabel(RecordsSheet).Find(SelectedLabel, , xlValues, xlWhole)
                        
            If Not TempLabelRange Is Nothing Then
                Call ActivityDelete(TempLabelRange)
            End If
        End If
    Next i

ListboxRemove:
    'Refresh the listbox items
    Call UserForm_Activate
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub LoadActivityFilterTextBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & LoadActivityFilterTextBox.Text & "*")
    Call LoadActivityListBoxPopulate
    
    With LoadActivityListBox
        For i = .ListCount - 1 To 0 Step -1
            If (Not (LCase(.List(i, 0)) Like testString)) _
            And (Not (LCase(.List(i, 1)) Like testString)) Then
                .RemoveItem i
            End If
        Next i
    End With
    
End Sub

Private Sub UserForm_Activate()
'Populate the list box with all saved activities

    Me.LoadActivityFilterTextBox.Value = ""
    
    LoadActivityForm.Height = 300
    LoadActivityForm.Width = 445
    
    LoadActivityListBox.ColumnCount = 2
    LoadActivityListBox.ColumnWidths = "220, 220"

    Call LoadActivityListBoxPopulate

Footer:

End Sub

Private Sub UserForm_Deactivate()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub LoadActivityListBoxPopulate()
'Populates the listbox with activities that haven't been completed

    Dim RecordsSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim RefLabelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
 
    Set RecordsSheet = Worksheets("Records Page")
    
    LoadActivityListBox.Clear
    
    'Make columns in the list box
    With LoadActivityListBox
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
