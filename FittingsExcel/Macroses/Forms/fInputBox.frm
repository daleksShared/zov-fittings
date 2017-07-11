VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fInputBox 
   Caption         =   "UserForm2"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "fInputBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public result As Boolean

Private Sub cbSelect_Click()
    result = True
    Me.Hide
End Sub

Private Sub cbSkip_Click()
    result = False
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    result = False
End Sub
