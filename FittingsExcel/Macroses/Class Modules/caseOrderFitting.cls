VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "caseOrderFitting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_orderid As Integer
Private p_fName As String
Private p_fOption As String
Private p_fLength As String
Private p_fQty As Integer
Private p_fRow As Integer
Private p_ismissingfOption As Boolean
Private p_ismissingfLength As Boolean
Private p_ismissingfQty As Boolean
Private p_qtyWithCaseQty As Boolean

Public Property Get ismissingfQty() As Boolean
    ismissingfQty = p_ismissingfQty
End Property


Public Property Get fRow() As Integer
    fRow = p_fRow
End Property
Public Property Let fRow(Value As Integer)
    p_fRow = Value
End Property

Public Property Get OrderId() As Integer
    OrderId = p_orderid
End Property
Public Property Let OrderId(Value As Integer)
    p_orderid = Value
End Property


Public Property Get fQty() As Integer
    fQty = p_fQty
End Property
Public Property Let fQty(Value As Integer)
    p_fQty = Value
    p_ismissingfQty = False
End Property


Public Property Get fName() As String
    fName = p_fName
End Property
Public Property Let fName(Value As String)
    p_fName = Value
    p_ismissingfOption = True
    p_fOption = Empty
    p_fLength = Empty
    p_ismissingfLength = True
    p_ismissingfQty = True
    p_qtyWithCaseQty = False
End Property



Public Property Get fOption() As String
    fOption = p_fOption
End Property
Public Property Let fOption(Value As String)
    p_fOption = Value
    p_ismissingfOption = False
End Property

Public Property Get fLength() As String
    fLength = p_fLength
End Property
Public Property Let fLength(Value As String)
    p_fLength = Value
    p_ismissingfLength = False
End Property

