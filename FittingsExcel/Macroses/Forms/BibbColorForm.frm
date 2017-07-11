VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BibbColorForm 
   Caption         =   "Цвет заглушек"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8250
   OleObjectBlob   =   "BibbColorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BibbColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public result As Boolean


Private Sub cbOK_Click()
    If cbBibbColor.Text <> "" Then
        result = True
        Hide
    Else
        If MsgBox("Без заглушек", vbQuestion + vbQuestion + vbYesNo + vbDefaultButton2, "Цвет заглушек") = vbYes Then
            result = True
            Hide
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    result = False
    
    ' цвета заглушек
    Dim Bibb()
    GetBibbColors Bibb
    cbBibbColor.List = Bibb
    '...
End Sub
