VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "caseParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private z_paramName As String
Private z_paramValue As String


Public Property Get paramName() As String
    paramName = z_paramName
End Property
Public Property Let paramName(Value As String)
    z_paramName = Value
End Property

Public Property Get paramValue() As String
    paramValue = z_paramValue
End Property
Public Property Let paramValue(Value As String)
    z_paramValue = Value
End Property
