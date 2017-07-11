VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WHOrderParamsForm 
   Caption         =   "Параметры заказа"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   OleObjectBlob   =   "WHOrderParamsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WHOrderParamsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private HandleArray()
Private LegArray()

Public result As Boolean

Private Sub btnOK_Click()
    result = False
    
    Dim bSpecified As Boolean
    bSpecified = True
    
    If cbHandle.Text = "" Then bSpecified = False
    If bSpecified Then If cbLeg.Text = "" Then bSpecified = False
    If bSpecified And Len(Trim(tbSetQty.Text)) > 0 Then
        If Not IsNumeric(tbSetQty.Text) Then
            bSpecified = False
        Else
            If CDec(tbSetQty.Text) <= 0 Or Round(tbSetQty.Text) <> CDec(tbSetQty.Text) Then bSpecified = False
        End If
    End If
    
    If bSpecified Then
        result = True
        Me.Hide
    Else
        result = False
        MsgBox "Определены не все параметры.", vbExclamation, "Параметры оптового заказа"
    End If
End Sub

Private Sub cbCancel_Click()
    result = False 'True '!
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    result = False

    Init_rsHandle
    Init_rsLeg
    
    ReDim HandleArray(0)
    ReDim LegArray(0)
    
    HandleArray(0) = "клиента"
    LegArray(0) = "клиента"
    
    Dim i As Integer
    If rsHandle.RecordCount > 0 Then
        ReDim Preserve HandleArray(rsHandle.RecordCount)
        rsHandle.MoveFirst
        For i = 0 To rsHandle.RecordCount - 1
            HandleArray(i + 1) = rsHandle!Handle
            rsHandle.MoveNext
        Next i
    End If

    If rsLeg.RecordCount > 0 Then
        ReDim Preserve LegArray(rsLeg.RecordCount)
        rsLeg.MoveFirst
        For i = 0 To rsLeg.RecordCount - 1
            LegArray(i + 1) = rsLeg!Leg
            rsLeg.MoveNext
        Next i
    End If
    
    cbHandle.List = HandleArray
    cbLeg.List = LegArray
End Sub
