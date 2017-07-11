VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "caseZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public casefasades As Collection
Public z_doorQty As Integer
Public z_windowsQty As Integer
Public z_shuflQty As Integer
Public z_id As Integer
Public z_isNisha As Boolean
Public z_rawstring As String

Public Sub init()
Set casefasades = New Collection
z_rawstring = ""
z_doorQty = 0
z_windowsQty = 0
z_shuflQty = 0
z_id = 0
z_isNisha = False
End Sub
