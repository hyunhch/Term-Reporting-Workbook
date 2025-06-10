VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewActivityForm 
   Caption         =   "New Activity Form"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   OleObjectBlob   =   "NewActivityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewActivityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub NewActivityCancelButton_Click()

    NewActivityForm.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub NewActivityConfirmButton_Click()
'Create a new sheet with the information given
'Checking for students and checked students comes previously

    Dim i As Long
    Dim PracticeString As String
    Dim CategoryString As String
    Dim NotesString As String
    Dim InfoArary() As Variant
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'First check that a practice has been selected. Notes are optional
    i = Me.NewActivitySelectListBox.ListIndex
    
    If i = -1 Then
        GoTo Footer
    End If
    
    PracticeString = NewActivitySelectListBox.List(i, 0)
    CategoryString = NewActivitySelectListBox.List(i, 1)
    NotesString = Me.NewActivityNotesBox.Value
    
    If Len(Trim(PracticeString)) = 0 Then
        MsgBox ("Please select an activity")
        GoTo Footer
    End If
        
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
    Call ActivityNewSheet(InfoArray)

    NewActivityForm.Hide
    
Footer:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub NewActivityFilterBox_Change()
'Dynamic filter for the activity list

    Dim i As Long
    Dim testString As String
    
    testString = LCase("*" & NewActivityFilterBox.Text & "*")
    Call NewActivityListBoxPopulate
    
    With NewActivitySelectListBox
        For i = .ListCount - 1 To 0 Step -1
            If (Not (LCase(.List(i, 0)) Like testString)) _
            And (Not (LCase(.List(i, 1)) Like testString)) Then
                .RemoveItem i
            End If
        Next i
    End With
    
End Sub

Private Sub UserForm_Activate()
'Clear anything in the date and description boxes when activated

    NewActivityFilterBox.Value = ""
    NewActivityNotesBox.Value = ""

    NewActivityForm.Height = 365
    NewActivityForm.Width = 407
    
    Call NewActivityListBoxPopulate

End Sub

Private Sub UserForm_Deactivate()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub UserForm_Initialize()

    NewActivityFilterBox.Value = ""
    NewActivityNotesBox.Value = ""

    NewActivityForm.Height = 365
    NewActivityForm.Width = 407
    
End Sub

Sub NewActivityListBoxPopulate()
'Populates the listbox with activities that haven't been completed

    Dim RecordsSheet As Worksheet
    Dim RecordsActivityRange As Range
    Dim ActivityRefRange As Range
    Dim c As Range
    Dim d As Range
    Dim i As Long
    
    Set RecordsSheet = Worksheets("Records Page")
    Set RecordsActivityRange = FindRecordsLabel(RecordsSheet)
    Set ActivityRefRange = Range("ActivitiesList")
    
    NewActivitySelectListBox.Clear
    
    With NewActivitySelectListBox
        .ColumnCount = 2
        .ColumnWidths = "180, 180"
    
        i = 0
        For Each c In RecordsActivityRange
            If IsChecked(FindRecordsAttendance(RecordsSheet, , c), "All") = False Then
                Set d = ActivityRefRange.Find(c.Value, , xlValues, xlWhole).Offset(0, -1) 'Category is one column behind
            
                .AddItem c.Value
                .List(i, 1) = d.Value
                
                i = i + 1
            End If
        Next c
    
    End With

End Sub

