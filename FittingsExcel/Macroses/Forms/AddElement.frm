VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddElement 
   Caption         =   "Добавить элементы каркасов"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "AddElement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private result As Boolean

Private rsElements As ADODB.Recordset
 
Private ElementArray()
    
    
    
Private Sub cbAdd_Click()
    result = False
    
    Dim bSpecified As Boolean
    bSpecified = True
    
    If cbElementName.Enabled And cbElementName.Text = "" Then bSpecified = False
    If bSpecified Then
        If Not IsNumeric(tbQty.Text) Then
            bSpecified = False
        Else
            If CDec(tbQty.Text) <= 0 Or Round(CDec(tbQty.Text)) <> CDec(tbQty.Text) Then bSpecified = False
        End If
    End If
    
    
    If bSpecified Then
        result = True
        Me.Hide
    Else
        result = False
        MsgBox "Элемент не добавлен." & vbCrLf & "Не все значения определены", vbExclamation, "Добавление элементов"
    End If
End Sub

Private Sub cbCancel_Click()
    result = False
    Me.Hide
End Sub


Private Sub UserForm_Activate()
    result = False
End Sub

Private Sub UserForm_Initialize()
    Dim comm As ADODB.Command
    Set comm = New ADODB.Command
    comm.ActiveConnection = GetConnection
    comm.CommandType = adCmdText
    comm.CommandText = "SELECT * FROM Element ORDER BY Name"

    Set rsElements = New ADODB.Recordset
    rsElements.CursorLocation = adUseClient
    rsElements.LockType = adLockBatchOptimistic
    rsElements.Open comm, , adOpenDynamic, adLockBatchOptimistic
    
    Dim i As Integer
    ReDim ElementArray(0)

    cbElementName.Clear
    If rsElements.RecordCount > 0 Then
        ReDim ElementArray(rsElements.RecordCount - 1)
        For i = 0 To rsElements.RecordCount - 1
            ElementArray(i) = rsElements!name
            rsElements.MoveNext
        Next i
        
        cbElementName.List = ElementArray
    End If

End Sub


Public Function AddElementToOrder(ByVal OrderId As Long, _
                                    ByVal name As String, _
                                    ByVal qty, _
                                    Optional ByVal caseID) As Boolean
                                    
                                    
    AddElementToOrder = False
    
    Dim bSpecified As Boolean
    bSpecified = False
    
    tbQty.Enabled = True
    
    If IsNumeric(qty) Then
        If qty = 0 Then
            AddElementToOrder = True
            Exit Function
        End If
        tbQty.Text = qty
    Else
        tbQty.Text = ""
        bSpecified = False
    End If
                                    
    cbAddNext.Value = False
    cbElementName.Enabled = True
    
    Dim i As Integer
    
    cbElementName.Text = ""
    
    If name = "" Then
        Me.Show
        If Not result Then Exit Function Else name = cbElementName.Text
    End If
    
    For i = 0 To cbElementName.ListCount - 1
        If cbElementName.List(i) = name Then
            cbElementName.Text = cbElementName.List(i)
            cbElementName.Enabled = False
            bSpecified = True
            Exit For
        End If
    Next
    
    If Not bSpecified Then
        MsgBox "ОШИБКА!!! НЕИЗВЕСТНЫЙ ЭЛЕМЕНТ", vbCritical
        Exit Function
    End If
    
    Do
        
        If cbElementName.Enabled And cbElementName.Text = "" Then bSpecified = False
        If bSpecified Then
            If Not IsNumeric(tbQty.Text) Then
                bSpecified = False
            Else
                If CDec(tbQty.Text) <= 0 Or Round(CDec(tbQty.Text)) <> CDec(tbQty.Text) Then bSpecified = False
            End If
        End If
        
        'если все элементы формы определены, то форму показывать не будем
        If bSpecified Then
            result = True
        Else
            If cbAddNext.Value Then
                cbAddNext.Value = False
                Me.Show 1
            Else
                Me.Show 1
            End If
        End If
                                    
        If result Then
            If IsMissing(OrderCaseID) = False And IsMissing(caseID) Then
                If OrderCaseID > 0 Then
                    caseID = getCaseIdbyOCID(OrderCaseID)
                End If
            End If
            If IsMissing(caseID) Then
                result = AddElement2Order(OrderId)
            Else
                result = AddElement2Order(OrderId, caseID)
            End If
        Else
            result = True 'была нажата отмена
        End If
            
         If cbAddNext.Value Then
            
            tbQty.Text = ""
            
            cbElementName.Enabled = True
            cbElementName.Text = ""
         Else
            Me.Hide
         End If
        
    Loop While cbAddNext.Value
End Function

Private Function AddElement2Order(ByVal OrderId As Long, _
                                    Optional ByVal caseID) As Boolean
                                    
    On Error GoTo err_AddElement2Order
    
    AddElement2Order = False
    Application.Cursor = xlWait
    
    Init_rsOrderElements
    
    Dim ElementID As Integer
    If rsElements.RecordCount > 0 Then rsElements.MoveFirst
    rsElements.Find "Name='" & cbElementName.Text & "'"
    If Not rsElements.EOF Then
        ElementID = rsElements!ElementID
    Else
        AddElement2Order = False
        MsgBox "Неизвестный элемент'" & cbElementName.Text & "'", vbCritical
        Exit Function
    End If
    
    
    'If rsOrderElements.RecordCount > 0 Then rsOrderElements.MoveFirst
    'rsOrderElements.Find "OrderID=" & OrderID
    'rsOrderElements.Find "ElementID=" & ElementID
    'If Not IsMissing(CaseID) Then rsOrderElements.Find "CaseID=" & CaseID
    'If rsOrderElements.EOF Then
        rsOrderElements.AddNew
        
        rsOrderElements!OrderId = OrderId
        rsOrderElements!ElementID = ElementID
        rsOrderElements!ocid = OrderCaseID
        
        rsOrderElements!qty = CDec(tbQty.Text)
        
        If Not IsMissing(caseID) Then
            rsOrderElements!caseID = caseID
            rsOrderElements!Standart = Not ActiveCell.Font.Bold
        Else
            rsOrderElements!Standart = False
        End If
    'Else
    ' rsOrderElements!Qty = rsOrderElements!Qty + CDec(tbQty.Text)
    'End If
    
    
    If Cells(ActiveCell.row, 10).Value <> "" Then
        Dim t As String
        t = ActiveCell.Offset(, 23).Value
        
        ActiveCell.Offset(, 23).Value = t & "; " & "Element=" & cbElementName.Text & ", QTY=" & tbQty.Text
    Else
        ActiveCell.Offset(, 23).Value = "Element=" & cbElementName.Text & ", QTY=" & tbQty.Text
    End If
        
    AddElement2Order = True
    Application.Cursor = xlDefault
    Exit Function
err_AddElement2Order:
    MsgBox "Ошибка добавления эелементов каркаса", vbCritical
    MsgBox Error, vbCritical
    AddElement2Order = False
    Application.Cursor = xlDefault
End Function


Public Sub AddElement()
    On Error GoTo err_ДобавитьЭлемент
    Set kitchenPropertyCurrent = Nothing
    Set casepropertyCurrent = Nothing
  
    ' веберем отгрузку
    'Dim TasksForm As MainForm
    Dim ShipID As Long
    'Set TasksForm = New MainForm
    MainForm.Show
    ShipID = MainForm.ShipID
    
    'Set TasksForm = Nothing
    If ShipID = 0 Then Exit Sub
    
    ' выберем клиента и заказ
    Dim SelectOrder As SelectOrderForm
    Set SelectOrder = New SelectOrderForm
    SelectOrder.ShowForm ShipID
    
    OrderCaseID = 0
    Dim SelectCase As SelectCaseForm
    Set SelectCase = New SelectCaseForm
    If SelectOrder.OrderId > 0 Then
        SelectCase.ShowForm SelectOrder.OrderId
    End If
    Set SelectCase = Nothing
    
    AddElementToOrder SelectOrder.OrderId, "", 1
    
    Set SelectOrder = Nothing
    
    If Not rsOrderElements Is Nothing Then
        rsOrderElements.UpdateBatch
        
        MsgBox "Успешно добавлено", vbInformation, "Добавление элементов"
    End If
    Exit Sub
    
err_ДобавитьЭлемент:
    MsgBox Error, vbCritical, "Добавление элементов"
End Sub
