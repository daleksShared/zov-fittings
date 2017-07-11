VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewTaskForm 
   Caption         =   "Новое задание"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   OleObjectBlob   =   "NewTaskForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewTaskForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Private Sub UserForm_Initialize()
    dtpTaskDate.Value = Date
End Sub


Private Sub btnAdd_Click()
    tbTaskNumber.Text = Trim(tbTaskNumber.Text)
    
    If tbTaskNumber.Text = "" Then
        MsgBox "Введите номер задания", vbInformation, "Добавление задания"
        Exit Sub
    ElseIf Len(tbTaskNumber.Text) > 6 Then
        MsgBox "Номер задания должен содержать не более 6 символов", vbInformation, "Ограничение на номер задания"
        Exit Sub
    End If
    
    Dim commAddNewTask As ADODB.Command
    Set commAddNewTask = New ADODB.Command
    commAddNewTask.ActiveConnection = GetConnection
    commAddNewTask.CommandType = adCmdStoredProc
    commAddNewTask.CommandText = "AddTask"
    
    commAddNewTask(1) = tbTaskNumber.Text
    commAddNewTask(2) = dtpTaskDate.Value
    If Trim(tbTaskNote.Text) = "" Then
        commAddNewTask(3) = Null
    Else
        commAddNewTask(3) = Trim(tbTaskNote.Text)
    End If
    commAddNewTask.Execute
    
    If commAddNewTask(4) = -1 Then MsgBox "Задание с таким номером уже существует." & vbCrLf & _
            "Введите другой номер задания.", vbExclamation, "Добавление задания"
    If commAddNewTask(4) = 0 Then MsgBox "Ошибка добавления задания", vbInformation, "Добавление задания"
    If commAddNewTask(4) > 0 Then Me.Hide

End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub
