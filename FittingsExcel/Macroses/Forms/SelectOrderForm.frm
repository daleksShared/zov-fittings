VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectOrderForm 
   Caption         =   "Выберите заказ"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   OleObjectBlob   =   "SelectOrderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectOrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public ShipID As Long, OrderId As Long
Private CustomerArray(), OrderArray()
Private rsOrders As ADODB.Recordset




Private Sub cbCustomer_Change()
    If cbCustomer.TopIndex >= 0 Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandText = "SELECT DISTINCT OrderID, Number FROM [Order] WHERE ShipID = ? AND Customer = ? ORDER BY Number"
    
        comm.Parameters(0) = ShipID
        comm.Parameters(1) = cbCustomer.Text
    
        Set rsOrders = New ADODB.Recordset
        rsOrders.CursorLocation = adUseClient
        rsOrders.Open comm, , adOpenDynamic, adLockBatchOptimistic
        
        ReDim OrderArray(0)
        
        Dim i As Integer
        If rsOrders.RecordCount > 0 Then
            ReDim OrderArray(rsOrders.RecordCount - 1)
            
            For i = 0 To rsOrders.RecordCount - 1
                OrderArray(i) = rsOrders!Number
                rsOrders.MoveNext
            Next i
            
            cbOrder.List = OrderArray
        End If
    End If
End Sub


Public Sub ShowForm(ByVal ShipID_ As Long)
    ShipID = ShipID_

    Dim comm As ADODB.Command
    Set comm = New ADODB.Command
    comm.ActiveConnection = GetConnection
    comm.CommandType = adCmdText
    comm.CommandText = "SELECT DISTINCT Customer FROM [Order] WHERE ShipID = ? ORDER BY Customer"

    comm.Parameters(0) = ShipID

    Dim rsCustomers As ADODB.Recordset
    Set rsCustomers = New ADODB.Recordset
    rsCustomers.CursorLocation = adUseClient
    rsCustomers.Open comm, , adOpenDynamic, adLockBatchOptimistic
    
    ReDim CustomerArray(0)
    
    Dim i As Integer
    If rsCustomers.RecordCount > 0 Then
        ReDim CustomerArray(rsCustomers.RecordCount - 1)
        
        For i = 0 To rsCustomers.RecordCount - 1
            CustomerArray(i) = rsCustomers!Customer
            rsCustomers.MoveNext
        Next i
        
        cbCustomer.List = CustomerArray
    End If
    
    Me.Show
End Sub

Private Sub cbOK_Click()
    OrderId = 0
    
    If Not rsOrders Is Nothing Then
        If rsOrders.RecordCount > 0 Then
            rsOrders.MoveFirst
            rsOrders.Find "Number='" & cbOrder.Text & "'"
            If Not rsOrders.EOF Then
                OrderId = rsOrders!OrderId
            End If
        End If
    End If
    
    If OrderId = 0 Then
        OrderId = AddOrder(ShipID, 0, cbCustomer.Text, cbOrder.Text)  ''' уточнить строку!!!
    End If
    Hide
End Sub


Private Sub UserForm_Initialize()
    OrderId = 0
End Sub
