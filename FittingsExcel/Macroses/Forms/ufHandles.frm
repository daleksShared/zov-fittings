VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufHandles 
   Caption         =   "Выберите ручку"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2895
   OleObjectBlob   =   "ufHandles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufHandles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbSelect_Click()
    If cbHandles.ListIndex >= 0 Or cbNoHandle.Value Then Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Init_rsHandle
    
    cbHandles.Clear
    
    If rsHandle.RecordCount > 0 Then
        Dim HandleArray()
        ReDim HandleArray(rsHandle.RecordCount - 1)
        
        Dim i As Long
        rsHandle.MoveFirst
        For i = 0 To rsHandle.RecordCount - 1
            HandleArray(i) = rsHandle!Handle
            rsHandle.MoveNext
        Next i
        
        cbHandles.List = HandleArray
    End If
    
    'cbHandles.MatchRequired = True
End Sub
