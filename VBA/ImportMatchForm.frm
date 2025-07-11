VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportMatchForm 
   Caption         =   "Match Imported Activities"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11100
   OleObjectBlob   =   "ImportMatchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportMatchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ImportMatchCancelButton_Click()

    Me.Hide

End Sub

Private Sub ImportMatchConfirmButton_Click()
'Puts the cell reference from the matched labels into both arrays
'Calls for another iteration of populating the form

    Dim i As Long
    Dim j As Long
    Dim ColString As String
    Dim RowString As String
    Dim TempString As String
    Dim NewAddress As String
    Dim OldString As String
    Dim NewString As String
    Dim AddressArray As Variant
    
    'First make sure something is selected
    If ImportMatchNewListbox.ListIndex = -1 Then
        GoTo Footer
    End If
    
    'Grab the original label and the selected one
    OldString = ImportMatchOldTextbox.Value
    For j = 0 To ImportMatchNewListbox.ListCount - 1
        If ImportMatchNewListbox.Selected(j) = True Then
            NewString = ImportMatchNewListbox.List(j)
        End If
    Next j

    'Loop through the two indices and add in the cell references stored there
    For i = 1 To UBound(OldMatchArray, 2)
        If OldMatchArray(1, i) = OldString Then
            OldMatchArray(1, i) = NewString
        
            Exit For
        End If
    Next i
       
    'Different procedure if adding a new column
    If NewString = "Add new column" Then
        OldMatchArray(1, i) = OldString
    
        GoTo AddColumn
    End If
    
    For j = 1 To UBound(NewMatchArray, 2)
        If NewMatchArray(1, j) = NewString Then
            Exit For
        End If
    Next j

    GoTo MatchArrays
    
AddColumn:
    'Grab the address stored in the last element
    j = UBound(NewMatchArray, 2)
    TempString = NewMatchArray(2, j)
    
    'For now, skip if there are no "$" in the address. There always should be
    If Not InStr(TempString, "$") > 0 Then
        MsgBox ("There was a problem adding this column. It will need to be done manually")
        
        GoTo NextMatch
    End If
    
    'Go one column to the right and add into the array
    AddressArray = Split(TempString, "$")
    TempString = AddressArray(1)
    RowString = AddressArray(2)
    ColString = Chr(Asc(TempString) + 1)
    
    NewAddress = "$" & ColString & "$" & RowString
    
    j = j + 1
    ReDim Preserve NewMatchArray(1 To 3, 1 To j)
        NewMatchArray(1, j) = OldString
        NewMatchArray(2, j) = NewAddress

MatchArrays:
    OldMatchArray(3, i) = NewMatchArray(2, j)
    NewMatchArray(3, j) = OldMatchArray(2, i)
    
NextMatch:
    'Iterate for the next selection
    Call ImportMatchPopulate

Footer:

End Sub

Private Sub ImportMatchSkipButton_Click()
'Replaces the "0" in the old array with "1" so it isn't pulled in
'Calls for another iteration of populating the form

    Dim i As Long
    Dim j As Long
    Dim SearchString As String
    
    SearchString = ImportMatchOldTextbox.Value
    For i = 1 To UBound(OldMatchArray, 2)
        If OldMatchArray(1, i) = SearchString Then
            OldMatchArray(3, i) = 1
            
            GoTo Footer:
        End If
    Next i

Footer:
    Call ImportMatchPopulate

End Sub

Private Sub UserForm_Activate()
'Clear everything and populate with activities that didn't match when importing
    
    ImportMatchForm.Height = 298
    ImportMatchForm.Width = 567
    
    Call ImportMatchPopulate

End Sub

Private Sub UserForm_Deactivate()

    IsRoster = False
    Me.Hide

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub ImportMatchPopulate()
'Puts in all of the unmatched labels from the new workbook into a listbox
'Puts a single unmatched label from the old workbook into a textbox
'Closes if it runs out of items

    Dim i As Long
    Dim j As Long
    Dim IsRoster As Boolean

    'Check if this is the Roster. There will be "Add new column" as the first item
    If Me.ImportMatchNewListbox.ListCount > 0 Then
        If Me.ImportMatchNewListbox.List(0, 0) = "Add new column" Then
            IsRoster = True
        End If
    End If

    ImportMatchNewListbox.Clear
    ImportMatchOldTextbox.Value = ""
    
    If IsEmpty(NewMatchArray) Or IsEmpty(OldMatchArray) Then
        Me.Hide
        GoTo Footer
    End If
    
    'Insert "Add new column"
    If IsRoster = True Then
        Me.ImportMatchNewListbox.AddItem ("Add new column")
    End If
    
    'Put in all of the unmatched new labels
    For j = 1 To UBound(NewMatchArray, 2)
        If NewMatchArray(3, j) = 0 Then
            Me.ImportMatchNewListbox.AddItem NewMatchArray(1, j)
        End If
    Next j
    
    If Me.ImportMatchNewListbox.ListCount = 0 Then
        Me.Hide
        GoTo Footer
    End If

    'Loop through and show each of the old unmatched labels one at a time
    For i = 1 To UBound(OldMatchArray, 2)
        If OldMatchArray(3, i) = 0 Then
            Me.ImportMatchOldTextbox = OldMatchArray(1, i)
        End If
    Next i

    If ImportMatchOldTextbox.Value = "" Then
        Me.Hide
        GoTo Footer
    End If

Footer:

End Sub
