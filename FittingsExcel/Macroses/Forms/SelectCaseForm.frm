VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectCaseForm 
   Caption         =   "Шкафы заказа"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "SelectCaseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectCaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OrderId As Long

Public Sub ShowForm(ByVal orderid_ As Long)
    OrderId = orderid_
OrderCaseID = 0
     Dim comm As ADODB.Command
    Set comm = New ADODB.Command
    comm.ActiveConnection = GetConnection
    comm.CommandType = adCmdText
    comm.CommandText = "SELECT OCID,CaseName FROM [Fittings].[dbo].[OrderCases] where OrderID=? order by OCID asc"
    comm.Parameters(0) = OrderId
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.Open comm, , adOpenDynamic, adLockReadOnly
    
    ListBox1.Clear
    If rs.RecordCount > 0 Then
        Dim CasesArray()
        ReDim CasesArray(rs.RecordCount - 1, 1)
        rs.MoveFirst
        Dim i As Long
         
        
        For i = 0 To rs.RecordCount - 1
            CasesArray(i, 0) = rs(0)
            CasesArray(i, 1) = rs(1)
            rs.MoveNext
        Next i
        
        ListBox1.List = CasesArray
    rs.Close
    End If
    
    If ListBox1.ListCount > 0 Then Me.Show
End Sub

Private Sub cbOK_Click()
For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then
       OrderCaseID = CLng(ListBox1.List(i, 0))
    End If
Next i
Me.Hide
'OrderCaseID=
End Sub

Private Sub cbOrderOnly_Click()
OrderCaseID = 0
Me.Hide
End Sub
