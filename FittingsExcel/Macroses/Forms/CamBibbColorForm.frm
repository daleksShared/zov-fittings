VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CamBibbColorForm 
   Caption         =   "�������� �����������"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   OleObjectBlob   =   "CamBibbColorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CamBibbColorForm"
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
        If MsgBox("��� ��������", vbQuestion + vbQuestion + vbYesNo + vbDefaultButton2, "���� ��������") = vbYes Then
            result = True
            Hide
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    result = False
    
    ' ����� ��������
    Dim CamBibb()
    GetCamBibbColors CamBibb
    cbBibbColor.List = CamBibb
    '...
End Sub

