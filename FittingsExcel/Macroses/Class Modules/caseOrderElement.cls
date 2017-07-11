VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "caseOrderElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_Name As String
Private p_Length As String
Private p_Qty As Integer
Private p_ismissingQty As Boolean
Public Property Get qty() As Integer
    qty = p_Qty
End Property
Public Property Let qty(Value As Integer)
    p_Qty = Value
    p_ismissingQty = False
End Property

Public Property Get name() As String
    name = p_Name
End Property
Public Property Let name(Value As String)
    p_Name = Value
    p_ismissingfQty = True
End Property

Public Property Get ismissingfQty() As Boolean
    ismissingQty = p_ismissingQty
End Property
