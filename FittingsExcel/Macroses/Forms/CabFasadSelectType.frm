VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CabFasadSelectType 
   Caption         =   "Тип эдемента фасада"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "CabFasadSelectType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CabFasadSelectType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public result As Boolean

Public Function GetFasadtype() As CabFasadType
     If Me.OptionButton1.Value = True Then
            GetFasadtype = Door
        ElseIf Me.OptionButton2.Value = True Then
            GetFasadtype = Nisha
        ElseIf Me.OptionButton3.Value = True Then
            GetFasadtype = Shuflyada
        End If
End Function

Public Sub SetFasadType(fasadtype As CabFasadType)

Select Case fasadtype
Case Door
    Me.OptionButton1.Value = True
Case Nisha
    Me.OptionButton2.Value = True
Case Shuflyada
    Me.OptionButton3.Value = True
End Select

End Sub

Private Sub BtnOk_Click()
        result = True
        Hide
End Sub

Private Sub UserForm_Activate()
    result = False
End Sub

Private Sub UserForm_Initialize()
    result = False
End Sub

