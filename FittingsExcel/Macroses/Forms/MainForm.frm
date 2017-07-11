VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "ФУРНИТУРА"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text





Private Sub btnShip2Nest_Click()
    If MsgBox("Сделать все шкафы задания нестом?", vbQuestion + vbYesNo, "Обработка задания") = vbYes Then
        
        Dim commTask As ADODB.Command
        Set commTask = New ADODB.Command
        commTask.ActiveConnection = GetConnection
        commTask.CommandType = adCmdStoredProc
        commTask.CommandText = "SetShipCases2Nest"
        Dim i As Long
        For i = 0 To lbShips.ListCount - 1
            If lbShips.Selected(i) = True Then
                commTask("@ShipID") = ShipID(i)
                commTask.Execute
                  
                  Select Case commTask("@RETURN_VALUE")
                    Case 0
                        Application.Cursor = xlDefault
                        lbShips.Selected(i) = False
                    Case 3
                        Application.Cursor = xlDefault
                        MsgBox "Ошибка !!!", vbCritical, "Задание"
                        Exit Sub
                  End Select
                  
            End If
        Next i
    End If
End Sub

Private Sub btnShip2Restore_Click()
 If MsgBox("Вернуть состояние шкафов?", vbQuestion + vbYesNo, "Обработка задания") = vbYes Then
        
        Dim commTask As ADODB.Command
        Set commTask = New ADODB.Command
        commTask.ActiveConnection = GetConnection
        commTask.CommandType = adCmdStoredProc
        commTask.CommandText = "SetShipCases2Restore"
        Dim i As Long
        For i = 0 To lbShips.ListCount - 1
            If lbShips.Selected(i) = True Then
                commTask("@ShipID") = ShipID(i)
                commTask.Execute
                  
                  Select Case commTask("@RETURN_VALUE")
                    Case 0
                        Application.Cursor = xlDefault
                        lbShips.Selected(i) = False
                    Case 3
                        Application.Cursor = xlDefault
                        MsgBox "Ошибка !!!", vbCritical, "Задание"
                        Exit Sub
                  End Select
                  
            End If
        Next i
    End If
End Sub

Private Sub btnShip2Task_Click()
On Error GoTo err_btnShip2Task_Click
   ' If MsgBox("А ты перевел шкафы в нест??? Точно....?", vbYesNo, "Формирование задания") = vbYes Then

    If MsgBox("Экспортировать фурнитуру из отгрузок в задание?", vbYesNo, "Формирование задания") = vbYes Then
'        Dim WithPackets As Boolean
'
'        If MsgBox("Сформировать задание с пакетами?", vbYesNo + vbDefaultButton2 + vbQuestion, "Задание на фурнитуру") = vbYes Then
'            WithPackets = True
'        Else
'            WithPackets = False
'        End If
    
        Application.Cursor = xlWait
        btnShip2Task.Enabled = False
            
        Dim CreateTaskComm As ADODB.Command
        Set CreateTaskComm = New ADODB.Command
        CreateTaskComm.ActiveConnection = GetConnection
        CreateTaskComm.CommandType = adCmdStoredProc
        CreateTaskComm.CommandText = "CreateTask"
        CreateTaskComm.CommandTimeout = 120
        CreateTaskComm("@TaskID") = TaskID
        If cbSetQty.Value Then
            CreateTaskComm("@Opt") = 1
        Else
            CreateTaskComm("@Opt") = 0
        End If
'        CreateTaskComm("@WithPackets") = WithPackets
        
        Dim i As Long
        For i = 0 To lbShips.ListCount - 1
            If lbShips.Selected(i) = True Then
                CreateTaskComm("@ShipID") = ShipID(i)
                CreateTaskComm.Execute
                  
                  Select Case CreateTaskComm("@RETURN_VALUE")
                    Case 0
                        btnShip2Task.Enabled = False
                       ' MsgBox "Экспорт из отгрузки " & lbShips.List(i, 0) & " успешно завершен", vbInformation, "Формирование задания"
                    
                        lbShips.Selected(i) = False
                        btnShip2Task.Enabled = False
                    Case 1
                        Application.Cursor = xlDefault
                        MsgBox "Ошибка экспорта из отгрузки", vbCritical, "Экспорт фурнитуры"
                        Exit Sub
                    Case -1
                        Application.Cursor = xlDefault
                        MsgBox "Выбранное задание закрыто для редактирования", vbCritical, "Экспорт фурнитуры"
                        Exit Sub
                  End Select
                  
            End If
        Next i
    End If
 '   End If
    
    If cbSetQty.Value Then
        cbSetQty.Value = False
        LoadTasksList
    End If
    Application.Cursor = xlDefault
    Exit Sub
err_btnShip2Task_Click:
    Application.Cursor = xlDefault
    MsgBox Error, vbCritical
End Sub



Private Sub cb_getWithPackets_Click()
    If MsgBox("Создать задание с пакетами???", vbQuestion + vbYesNo, "Завершение задания") = vbYes Then
        
        Dim commDelTask As ADODB.Command
        Set commDelTask = New ADODB.Command
        commDelTask.ActiveConnection = GetConnection
        commDelTask.CommandType = adCmdStoredProc
        commDelTask.CommandText = "TaskConvertFurniture2Packets"
        commDelTask.CommandTimeout = 600
        commDelTask(1) = TaskID
        If commDelTask(1) > 0 Then commDelTask.Execute
        If commDelTask("@RETURN_VALUE") <> 0 Then
            MsgBox "ОШИБКА!!!", vbCritical, ""
        Else
            MsgBox "ОК!!!", vbInformation, ""
            LoadTasksList
        End If
    End If

End Sub

Private Sub lbTasks_Change()
    If lbTasks.ListIndex <> -1 And lbShips.ListIndex <> -1 Then
        btnShip2Task.Enabled = True
        btnShip2Nest.Enabled = True
        btnShip2Restore.Enabled = True
    Else
        btnShip2Task.Enabled = False
        btnShip2Nest.Enabled = False
        btnShip2Restore.Enabled = False
    End If
    
    If lbTasks.ListIndex <> -1 Then
        btnDeleteTask.Enabled = True
        btnClearTask.Enabled = True
        btnTask2Pack.Enabled = True
    Else
        btnDeleteTask.Enabled = False
        btnClearTask.Enabled = False
        btnTask2Pack.Enabled = False
    End If
End Sub



Private Sub UserForm_Initialize()
    LoadTasksList
    LoadShipsList
End Sub

Private Sub LoadTasksList()
    Dim commGetTasks As ADODB.Command
    Set commGetTasks = New ADODB.Command
    commGetTasks.ActiveConnection = GetConnection
    commGetTasks.CommandType = adCmdText
    commGetTasks.CommandText = "SELECT TOP(15) [Number], [Date],  ISNULL(Note,''), Draft, TaskID FROM Task ORDER BY TaskID DESC"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.Open commGetTasks, , adOpenDynamic, adLockReadOnly
    
    lbTasks.Clear
    If rs.RecordCount > 0 Then
        Dim TaskArray()
        ReDim TaskArray(rs.RecordCount - 1, 5)

        Dim i As Long
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            TaskArray(i, 0) = "№" & rs(0)
            TaskArray(i, 1) = Format(rs(1), "dd.mm.yy")
            TaskArray(i, 2) = rs(2)
            If rs(3) Then
                TaskArray(i, 3) = "X"
            End If
            TaskArray(i, 4) = rs(4)
            rs.MoveNext
        Next i
        
        lbTasks.List = TaskArray
    End If
    rs.Close
End Sub

Private Sub LoadShipsList()

    Dim commGetShips As ADODB.Command
    Set commGetShips = New ADODB.Command
    commGetShips.ActiveConnection = GetConnection
    commGetShips.CommandType = adCmdText
    commGetShips.CommandText = "SELECT Number, [Date], Note, ShipID FROM Ship WHERE Closed IS NULL ORDER BY ShipID DESC"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.Open commGetShips, , adOpenDynamic, adLockReadOnly
    
    lbShips.Clear
    If rs.RecordCount > 0 Then
        Dim ShipsArray()
        ReDim ShipsArray(rs.RecordCount - 1, 4)

        Dim i As Long
        If rs.RecordCount > 0 Then rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            ShipsArray(i, 0) = "№" & rs(0)
            ShipsArray(i, 1) = Format(rs(1), "dd.mm.yy")
            ShipsArray(i, 2) = rs(2)
            ShipsArray(i, 3) = rs(3)
            rs.MoveNext
        Next i
        
        lbShips.List = ShipsArray
    End If
    rs.Close
End Sub

            
Private Sub btnAddShip_Click()
    On Error GoTo err_btnAddShip_Click
    
    Dim FormShip As NewShipForm
    Set FormShip = New NewShipForm
    FormShip.Show
    If FormShip.result = False Then Exit Sub

    
    Dim AddComm As ADODB.Command
    Set AddComm = New ADODB.Command
    AddComm.ActiveConnection = GetConnection

    AddComm.CommandType = adCmdStoredProc
    AddComm.CommandText = "AddShip"
    AddComm(1) = FormShip.SNumber
    AddComm(2) = FormShip.SDate
    AddComm(3) = FormShip.SNote
    AddComm.Execute
    Select Case AddComm(0)
        Case 2
            MsgBox "Отгрузка №" & FormShip.SNumber & " уже существует.", vbInformation, "Новая отгрузка"
        Case Else
            LoadShipsList
    End Select
        
    Set FormShip = Nothing
    
    Exit Sub
err_btnAddShip_Click:
    MsgBox Error, vbCritical, "Добавление отгрузки"
End Sub



Private Sub btnClearShip_Click()

    If MsgBox("Вы действительно хотите удалить все элементы отгрузки?", vbQuestion + vbYesNo, "Очистка отгрузки") = vbYes Then
        On Error GoTo err_btnClearShip_Click
        
        Dim commClearShip As ADODB.Command
        Set commClearShip = New ADODB.Command
        commClearShip.ActiveConnection = GetConnection
        commClearShip.CommandType = adCmdStoredProc
        commClearShip.CommandText = "ClearShip"
        
        Dim i As Long
        For i = 0 To lbShips.ListCount - 1
            If lbShips.Selected(i) = True Then
                
                commClearShip(1) = ShipID(i)
                If commClearShip(1) > 0 Then
                    
                    commClearShip.Execute
                    lbShips.Selected(i) = False
                End If
            End If
        Next i
        
        LoadShipsList
    End If
    
    
    Exit Sub
err_btnClearShip_Click:
    MsgBox Error, vbCritical, "Удаление заказов из отгрузки"
End Sub

Private Sub btnDeleteShip_Click()
    
    If MsgBox("Вы действительно хотите удалить отгрузку?", vbQuestion + vbYesNo, "Удаление отгрузки") = vbYes Then
        On Error GoTo err_btnDeleteShip_Click
    
        Dim commDelShip As ADODB.Command
        Set commDelShip = New ADODB.Command
        commDelShip.ActiveConnection = GetConnection
        commDelShip.CommandType = adCmdStoredProc
        commDelShip.CommandText = "DeleteShip"
        
        Dim i As Long
        For i = 0 To lbShips.ListCount - 1
            If lbShips.Selected(i) = True Then
                
                commDelShip(1) = ShipID(i)
                If commDelShip(1) > 0 Then
                    
                    commDelShip.Execute
                    lbShips.Selected(i) = False
                End If
            End If
        Next i
                
        LoadShipsList
    End If
    
    Exit Sub
err_btnDeleteShip_Click:
    MsgBox Error, vbCritical, "Удаление отгрузки"
End Sub


Private Sub btnClearTask_Click()

    If MsgBox("Вы действительно хотите удалить все элементы задания?", vbQuestion + vbYesNo, "Очистка задания") = vbYes Then
        
        Dim commDelTask As ADODB.Command
        Set commDelTask = New ADODB.Command
        commDelTask.ActiveConnection = GetConnection
        commDelTask.CommandType = adCmdStoredProc
        commDelTask.CommandText = "ClearTask"
        commDelTask(1) = TaskID
        If commDelTask(1) > 0 Then commDelTask.Execute
        If commDelTask("@RETURN_VALUE") <> 0 Then
            MsgBox "Невозможно удалить элементы задания, т.к. задание уже выполняется", vbCritical, "Удаление элементов задания"
        Else
            LoadTasksList
        End If
    End If
End Sub


Private Sub btnDeleteTask_Click()

    If MsgBox("Вы действительно хотите удалить задание?", vbQuestion + vbYesNo, "Удаление задания") = vbYes Then
        
        Dim commDelTask As ADODB.Command
        Set commDelTask = New ADODB.Command
        commDelTask.ActiveConnection = GetConnection
        commDelTask.CommandType = adCmdStoredProc
        commDelTask.CommandText = "DeleteTask"
        commDelTask(1) = TaskID
        If commDelTask(1) > 0 Then
            commDelTask.Execute
            If commDelTask("@RETURN_VALUE") <> 0 Then
                MsgBox "Невозможно удалить задание, т.к. задание уже выполняется", vbCritical, "Удаление задания"
            Else
                LoadTasksList
            End If
        End If
    End If
End Sub









Public Property Get TaskID() As Long
    If lbTasks.ListIndex = -1 Then
        TaskID = 0
    Else
        TaskID = lbTasks.List(lbTasks.ListIndex, 4)
    End If
End Property

Public Property Get ShipID(Optional ByVal r As Long = -1) As Long
    If r >= 0 Then
        ShipID = lbShips.List(r, 3)
    Else
        If lbShips.ListIndex = -1 Then
            ShipID = 0
        Else
            ShipID = lbShips.List(lbShips.ListIndex, 3)
        End If
    End If
End Property

Public Property Get ShipNumber(Optional ByVal r As Long = -1) As String
    If r >= 0 Then
        ShipNumber = lbShips.List(r, 0)
    Else
        If lbShips.ListIndex = -1 Then
            ShipNumber = 0
        Else
            ShipNumber = lbShips.List(lbShips.ListIndex, 0)
        End If
    End If
    ShipNumber = Replace(ShipNumber, "№", "")
End Property

Public Property Get ShipDate(Optional ByVal r As Long = -1) As String
    If r >= 0 Then
        ShipDate = lbShips.List(r, 1)
    Else
        If lbShips.ListIndex = -1 Then
            ShipDate = ""
        Else
            ShipDate = lbShips.List(lbShips.ListIndex, 1)
        End If
    End If
End Property

Private Sub lbships_Change()
    If lbTasks.ListIndex <> -1 And lbShips.ListIndex <> -1 Then
        btnShip2Task.Enabled = True
        btnShip2Nest.Enabled = True
        btnShip2Restore.Enabled = True
    Else
        btnShip2Task.Enabled = False
        btnShip2Nest.Enabled = False
        btnShip2Restore.Enabled = False
    End If
    
    If lbShips.ListIndex <> -1 Then
        btnDeleteShip.Enabled = True
        btnClearShip.Enabled = True
    Else
        btnDeleteShip.Enabled = False
        btnClearShip.Enabled = False
    End If
End Sub


Private Sub btnAddTask_Click()
    Dim AddNewTaskForm As NewTaskForm
    Set AddNewTaskForm = New NewTaskForm
    AddNewTaskForm.Show
    LoadTasksList
End Sub


Private Sub btnOK_Click()
    Hide
End Sub


Private Sub btnTask2Pack_Click()
    If MsgBox("Завершить подготовку задания?", vbQuestion + vbYesNo, "Завершение задания") = vbYes Then
        
        Dim commDelTask As ADODB.Command
        Set commDelTask = New ADODB.Command
        commDelTask.ActiveConnection = GetConnection
        commDelTask.CommandType = adCmdStoredProc
        commDelTask.CommandText = "PrepareTask"
        commDelTask(1) = TaskID
        If commDelTask(1) > 0 Then commDelTask.Execute
        If commDelTask("@RETURN_VALUE") <> 0 Then
            'MsgBox "", vbCritical, ""
        Else
            LoadTasksList
        End If
    End If
End Sub


Private Sub btnExportDoors_Click()
        Dim commExportDoors As ADODB.Command
        Set commExportDoors = New ADODB.Command
        commExportDoors.ActiveConnection = GetConnection
        commExportDoors.CommandType = adCmdStoredProc
        commExportDoors.CommandText = "ExportDoors"
        
        Dim i As Long
        For i = 0 To lbShips.ListCount - 1
            If lbShips.Selected(i) = True Then
                commExportDoors(1) = ShipID(i)
                If commExportDoors(1) > 0 Then
                    commExportDoors.Execute
                    If commExportDoors("@RETURN_VALUE") = 0 Then
                        lbShips.Selected(i) = False
                    Else
                        MsgBox "ошибка", vbCritical, "Экспорт дверей"
                    End If
                End If
            End If
        Next
        
End Sub
