VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewShipForm 
   Caption         =   "Новая отгрузка"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "NewShipForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewShipForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Public SNumber As String
Public SDate As Date
Public SNote As Variant
Public result As Boolean

Private Sub CancelButton_Click()
    result = False
    Hide
End Sub

Private Sub OkButton_Click()
    Dim L As Long
    L = Len(Trim(tbSNumber.Text))
    If L < 1 Or L > 20 Then Exit Sub
    SNumber = tbSNumber.Text
    SDate = dtpSDate.Value
    SNote = tbSNote.Text
    result = True
    Hide
End Sub

Private Sub UserForm_Initialize()
    result = False
    dtpSDate.Value = Date
    'tbSNumber.Text = Trim(Cells(1, 5))
    tbSNote.Text = Trim(Cells(1, 6))
End Sub
