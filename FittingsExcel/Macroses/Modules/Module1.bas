Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

Private Customer As String, ShipNumber As String

Public Sub Добавить_фасады_по_элементам()
On Error GoTo err_Добавить_фасады_по_элементам
    
    ' веберем отгрузку
    Dim ShipID As Long
    MainForm.Show
    ShipID = MainForm.ShipID
    If ShipNumber <> MainForm.ShipNumber Then
        ShipNumber = MainForm.ShipNumber
        Customer = ""
    End If
    
    If ShipID = 0 Then Exit Sub
    
'    ' выберем клиента и заказ
'    Dim SelectOrder As SelectOrderForm
'    Set SelectOrder = New SelectOrderForm
'    SelectOrder.ShowForm ShipID
    
    Dim OrderNumber As String
'    Customer = SelectOrder.cbCustomer.Text
'    OrderNumber = SelectOrder.cbOrder.Text
    
    Do
        Customer = InputBox("Клиент (<=25 симв)", "Добавление фасадов по элементам", Customer)
    Loop While Customer = "" Or Len(Customer) > 25
    
    Dim face As String
    While face = "" Or Len(face) > 50
        face = InputBox("Фасад (<=50 симв)", "Добавление фасадов по элементам")
    Wend
'************
   
    Dim AddComm As ADODB.Command
    Set AddComm = New ADODB.Command
    AddComm.ActiveConnection = GetConnection
    AddComm.CommandType = adCmdText
    
    Dim casename As String
    Dim r As Range, qty As String, res As Long
    For Each r In Selection.Rows
        If r.Hidden = False Then
            casename = r.Cells(, 1)
            qty = r.Cells(, Selection.Columns.Count)
            If Not r.Hidden And Trim(casename) <> "" And Not IsEmpty(qty) And IsNumeric(qty) Then
                
                OrderNumber = casename
                While Len(OrderNumber) > 32 Or OrderNumber = ""
                    OrderNumber = InputBox("Заказ (<=32 симв)", "Добавление фасадов по элементам", OrderNumber)
                Wend
            
                AddComm.CommandText = "INSERT Doors(ShipNumber,Customer,OrderNumber,Face,SetQty,PackName) VALUES(?,?,?,?,?,?)"
                AddComm(0) = Left(ShipNumber, 20)
                AddComm(1) = Customer
                AddComm(2) = OrderNumber
                AddComm(3) = face
                AddComm(4) = CInt(qty)
                AddComm(5) = "А"
                AddComm.Execute res
                If res = 0 Then
                    MsgBox "Ошибка добавления фасадов", vbCritical, "Критическая ошибка времени выполнения"
                    r.Cells(, 1).Interior.ColorIndex = 3
                Else
                    r.Cells(, 1).Interior.ColorIndex = 37
                    r.Cells(, Selection.Columns.Count).Interior.ColorIndex = 37
                End If
               
            End If
        End If
    Next r
    
'************
    
    Exit Sub
    
err_Добавить_фасады_по_элементам:
    MsgBox Error, vbCritical, "Добавление фасадов по элементам"
End Sub


Public Sub Добавить_фасады_опт()
On Error GoTo err_Добавить_фасады_опт
    
    ' веберем отгрузку
    Dim ShipID As Long
    MainForm.Show
    ShipID = MainForm.ShipID
    If ShipNumber <> MainForm.ShipNumber Then
        ShipNumber = MainForm.ShipNumber
        Customer = ""
    End If
    
    If ShipID = 0 Then Exit Sub
    
'    ' выберем клиента и заказ
'    Dim SelectOrder As SelectOrderForm
'    Set SelectOrder = New SelectOrderForm
'    SelectOrder.ShowForm ShipID
    
    Dim OrderNumber As String
'    Customer = SelectOrder.cbCustomer.Text
'    OrderNumber = SelectOrder.cbOrder.Text
    
    Do
        Customer = InputBox("Клиент (<=25 симв)", "Добавление фасадов опт", Customer)
    Loop While Customer = "" Or Len(Customer) > 25
    
    Dim face As String
    While face = "" Or Len(face) > 50
        face = InputBox("Фасад (<=50 симв)", "Добавление фасадов опт")
    Wend
'************
   
    Dim AddComm As ADODB.Command
    Set AddComm = New ADODB.Command
    AddComm.ActiveConnection = GetConnection
    AddComm.CommandType = adCmdText
    
    Dim WidthComm As ADODB.Command
    Set WidthComm = New ADODB.Command
    WidthComm.ActiveConnection = GetConnection
    WidthComm.CommandType = adCmdStoredProc
        
    
    Dim casename As String
    Dim r As Range, qty As String, res As Long
    For Each r In Selection.Rows
        If r.Hidden = False Then
            casename = r.Cells(, 1)
            qty = r.Cells(, Selection.Columns.Count)
            If Not r.Hidden And Trim(casename) <> "" And Not IsEmpty(qty) And IsNumeric(qty) Then
                
                OrderNumber = casename & " " & qty & "шт"
                'If Len(OrderNumber) <= 13 Then OrderNumber = OrderNumber & "шт"
                While Len(OrderNumber) > 32 Or OrderNumber = ""
                    OrderNumber = InputBox("Заказ (<=32 симв)", "Добавление фасадов опт", OrderNumber)
                Wend
                
                'определим кол-во упаковок
                Dim SetQty As Integer, L As Long
                WidthComm.CommandText = "CaseWidth"
                WidthComm(1) = Left(casename, 70)
                WidthComm(2) = Left(casename, 20)
                WidthComm.Execute res
                
                SetQty = 0
                
                If res > 0 Then
                    L = WidthComm(0) * CInt(qty)
                    SetQty = L \ 3500
                    If L Mod 3500 > 0 Then SetQty = SetQty + 1
                End If
                
                While SetQty = 0 Or SetQty > 20 Or Not IsNumeric(SetQty)
                    SetQty = InputBox("Введите кол-во упаковок" & vbCrLf & vbCrLf & casename, "Добавление фасадов опт", SetQty)
                Wend
                
                
                Dim i As Integer
                For i = 1 To SetQty
                    AddComm.CommandText = "INSERT Doors(ShipNumber,Customer,OrderNumber,Face,SetQty,PackName) VALUES(?,?,?,?,?,?)"
                    AddComm(0) = Left(ShipNumber, 20)
                    AddComm(1) = Customer
                    AddComm(2) = OrderNumber
                    AddComm(3) = face
                    AddComm(4) = 1
                    AddComm(5) = Chr(191 + i)
                    AddComm.Execute res
                    If res = 0 Then
                        MsgBox "Ошибка добавления фасадов", vbCritical, "Критическая ошибка времени выполнения"
                        r.Cells(, 1).Interior.ColorIndex = 3
                    Else
                        r.Cells(, 1).Interior.ColorIndex = 37
                        r.Cells(, Selection.Columns.Count).Interior.ColorIndex = 37
                    End If
                Next i
               
            End If
        End If
    Next r
    
'************
    
    Exit Sub
    
err_Добавить_фасады_опт:
    MsgBox Error, vbCritical, "Добавление фасадов опт"
End Sub

Sub ОчисткаФитингс()
'
' ОчисткаФитингс Макрос
'
    Columns("I:AY").Select
    Selection.ClearContents
    Dim sh As Shape
    Dim wsHasShapes As Boolean
    wsHasShapes = False
    For Each sh In ActiveSheet.Shapes
        If InStr(1, sh.name, "comment", vbTextCompare) = 0 Then
            wsHasShapes = True
            Exit For
        End If
    Next sh
    If wsHasShapes Then
        ActiveSheet.Shapes.SelectAll
        Selection.Delete
    End If
    
    Range("A1").Select
    
    Columns("I:J").Select
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("I").Select
    Selection.ColumnWidth = 8
     With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
        
    Columns("J").Select
    Selection.ColumnWidth = 35
    
    Columns("K:O").Select
    Selection.ColumnWidth = 4
    Columns("P:P").Select
    Selection.ColumnWidth = 17
    Columns("Q:R").Select
    Selection.ColumnWidth = 3.5
    Columns("S:S").Select
    Selection.ColumnWidth = 13
    Columns("T:V").Select
    Selection.ColumnWidth = 8.5

    Columns("A:A").Select
    With Selection
        .VerticalAlignment = xlTop
    End With
    
    Range("A1").Select
    
End Sub
