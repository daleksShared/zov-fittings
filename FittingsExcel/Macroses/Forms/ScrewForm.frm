VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScrewForm 
   Caption         =   "Шуруп/винт для ручек заказа"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   OleObjectBlob   =   "ScrewForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScrewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public result As Boolean

Private Sub bOK_Click()
    If lbScrewLen.ListIndex >= 0 Then
        result = True
        Hide
    Else
        MsgBox "Необходимо выбрать оё из значений!", vbExclamation, "Длина шурупа"
    End If
End Sub



Private Sub UserForm_Initialize()
    result = False
    lbScrewLen.AddItem "22"
    lbScrewLen.AddItem "25"
    lbScrewLen.AddItem "28"
    lbScrewLen.AddItem "35"
    lbScrewLen.AddItem "40"
End Sub
