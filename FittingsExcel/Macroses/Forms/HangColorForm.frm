VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HangColorForm 
   Caption         =   "Цвет завешек"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8235
   OleObjectBlob   =   "HangColorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HangColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public result As Boolean

Private Sub cbOK_Click()
    If cbHangColor.Text <> "" Then
        result = True
        Hide
    Else
        MsgBox "Необходимо выбрать цвет", vbExclamation, "Цвет завешек"
    End If
End Sub

Private Sub UserForm_Initialize()
    result = False
    
    Dim Hang()
    GetHangColors Hang
    cbHangColor.List = Hang
End Sub

