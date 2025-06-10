VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddStudentsForm 
   Caption         =   "Add Students"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670.001
   OleObjectBlob   =   "AddStudentsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddStudentsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddStudentsCancelButton_Click()
'Hide the form

    AddStudentsForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub AddStudentsConfirmButton_Click()
'Recreate an activity sheet with the activity information and attendance
    
    Dim RecordsSheet As Worksheet
    Dim ActivitySheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim AddStudentsRange As Range
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
    i = AddStudentsListBox.ListIndex
    
    If i = -1 Then
        GoTo Footer
    End If
    
    Set RecordsSheet = Worksheets("Records Page")
    Set RecordsLabelRange = FindRecordsLabel(RecordsSheet)
    
    'Grab activity information
    For i = 0 To Me.AddStudentsListBox.ListCount - 1
        If Not Me.AddStudentsListBox.Selected(i) Then
            GoTo NextRow
        End If
        
        PracticeString = Me.AddStudentsListBox.List(i, 0)
        
        'See if there's an open Activity sheet
        Set ActivitySheet = FindActivitySheet(PracticeString)
            If Not ActivitySheet Is Nothing Then
                GoTo SkipNewSheet
            End If
        
        'Grab activity information. Notes come from the RecordsSheet
        Set c = RecordsLabelRange.Find(PracticeString, , xlValues, xlWhole).Offset(1, 0)
        
        CategoryString = Me.AddStudentsListBox.List(i, 1)
        NotesString = c.Value
        
        GoTo MakeArray 'shouldn't be necessary since it's a single-select listbox
NextRow:
    Next i
    
MakeArray:
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
        Set ActivitySheet = ActivityNewSheet(InfoArray, "Load")

SkipNewSheet:
    'Copy over unique students
    Set AddStudentsRange = ActivityAddStudents(ActivitySheet)
        If AddStudentsRange Is Nothing Then
            MsgBox ("All checked students were already listed on the activity.")
            
            GoTo Footer
        End If
    
    'Uncheck all added students
    AddStudentsRange.ClearContents
    
    'Make a table
    Call CreateTable(ActivitySheet)
    
    i = AddStudentsRange.Cells.Count
    MsgBox (i & " students added.")
   
Footer:
    AddStudentsForm.Hide
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub AddStudentsFilterBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & AddStudentsFilterBox.Text & "*")
    Call AddStudentsListBoxPopulate
    
    With AddStudentsListBox
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

    Me.AddStudentsListBox.Clear
    Me.AddStudentsFilterBox.Value = ""
    
    AddStudentsForm.Height = 300
    AddStudentsForm.Width = 445
    
    AddStudentsListBox.ColumnCount = 2
    AddStudentsListBox.ColumnWidths = "220, 220"

    Call AddStudentsListBoxPopulate
    
Footer:

End Sub

Sub AddStudentsListBoxPopulate()
'Populates the listbox with activities that haven't been completed

    Dim RecordsSheet As Worksheet
    Dim RecordsLabelRange As Range
    Dim RefLabelRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
 
    Set RecordsSheet = Worksheets("Records Page")
    
    AddStudentsListBox.Clear
    
    'Make columns in the list box
    With AddStudentsListBox
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

Private Sub UserForm_Deactivate()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
