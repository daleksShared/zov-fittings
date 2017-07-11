VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clSplitString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private zstr As String
Private zstart As Integer
Private zlength As Integer

Public Property Get start() As Integer
    start = zstart
End Property
Public Property Let start(Value As Integer)
    zstart = Value
End Property

Public Property Get str() As String
    str = zstr
End Property
Public Property Let str(Value As String)
    zstr = Value
    zlength = Len(zstr)
End Property

Public Property Get length() As Integer
    length = zlength
End Property
Public Property Let length(Value As Integer)
    zlength = Value
End Property
