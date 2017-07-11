VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kitchenProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private z_fasadColor As String
Private z_dspColor As String
Private z_dspColorId As Integer
Private z_addGlassHolders As Boolean
Private z_OrderN As String
Private z_PProfil As Boolean
Private z_dspWidth As Integer
Private dd As Integer
Private z_CabMessage As String
Private z_CamBibbColor As String
Private z_changeCaseZaves As Integer
Private z_is18dsp As Boolean
Private z_OrderId As Long


Public Property Get OrderId() As Long
    OrderId = z_OrderId
End Property
Public Property Let OrderId(Value As Long)
    z_OrderId = Value
End Property


Public Property Get dspColorId() As Integer
    dspColorId = z_dspColorId
End Property
Public Property Let dspColorId(Value As Integer)
    z_dspColorId = Value
End Property
Public Property Get changeCaseZaves() As Integer
    If IsEmpty(z_changeCaseZaves) Or IsNull(z_changeCaseZaves) Then
        changeCaseZaves = 0
    Else
        changeCaseZaves = z_changeCaseZaves
    End If
End Property
Public Property Let changeCaseZaves(Value As Integer)
    z_changeCaseZaves = Value
End Property

Public Property Get CabMessage() As String
    CabMessage = z_CabMessage
End Property
Public Property Let CabMessage(Value As String)
    If Value = "" Then
        z_CabMessage = ""
    Else
        z_CabMessage = Value & "; "
    End If
End Property

Public Property Get CamBibbColor() As Variant
    CamBibbColor = z_CamBibbColor
End Property
Public Property Let CamBibbColor(Value As Variant)
    If IsNull(Value) Then
        z_CamBibbColor = ""
    ElseIf IsEmpty(Value) Then
        z_CamBibbColor = ""
    ElseIf Value = "" Then
        z_CamBibbColor = ""
    Else
        z_CamBibbColor = CStr(Value)
    End If
End Property

Public Property Let CabMessageAppend(Value As String)
If IsNull(z_CabMessage) Then z_CabMessage = ""
If IsEmpty(z_CabMessage) Then z_CabMessage = ""
    z_CabMessage = z_CabMessage & Value & "; "
End Property

Public Property Get dspWidth() As Integer
    dspWidth = z_dspWidth
End Property
Public Property Let dspWidth(Value As Integer)
    z_dspWidth = Value
     If z_dspWidth > 16 Then
            z_is18dsp = True
        Else
            z_is18dsp = False
    End If
End Property

Public Property Get fasadColor() As String
    fasadColor = z_fasadColor
End Property
Public Property Let fasadColor(Value As String)
    z_fasadColor = Value
End Property

Public Property Get dspColor() As String
    dspColor = z_dspColor
End Property
Public Property Let dspColor(Value As String)
    z_dspColor = Value
    
    dd = InStr(1, Value, "лондон", vbTextCompare)
    If dd = 0 Then dd = InStr(1, Value, "париж", vbTextCompare)
    If dd = 0 Then dd = InStr(1, Value, "глазго", vbTextCompare)
    If dd = 0 Then dd = InStr(1, Value, "система", vbTextCompare)
    If InStr(1, Value, "сис", vbTextCompare) > 0 _
        And (InStr(1, Value, "ral", vbTextCompare) > 0 Or InStr(1, Value, "рал", vbTextCompare) > 0) _
        Then
        dd = 1
        Else
        dd = 0
    End If
    If dd = 0 Then dd = InStr(1, Value, "женев", vbTextCompare)
    If dd = 0 Then dd = InStr(1, Value, "масс", vbTextCompare)
    If dd = 0 Then dd = InStr(1, Value, "прим", vbTextCompare)
    If dd > 0 Then
        z_addGlassHolders = False
    Else
        z_addGlassHolders = True
    End If
    If InStr(1, z_dspColor, "16", vbTextCompare) > 0 Then
        z_dspWidth = 16
    ElseIf InStr(1, z_dspColor, "легно", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "лаванда", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "платина", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "черешня", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "золото", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "марино", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "магия", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "листвен", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "авиньон", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "шпон", vbTextCompare) > 0 Or _
         InStr(1, z_dspColor, "18", vbTextCompare) > 0 Then
         z_dspWidth = 18
    Else
        z_dspWidth = 16
    End If
    
    
    If z_dspWidth > 16 Then
            z_is18dsp = True
        Else
            z_is18dsp = False
    End If
End Property

Public Property Get is18dsp() As Boolean
    is18dsp = z_is18dsp
End Property

Public Property Get addGlassHolders() As Boolean
    addGlassHolders = z_addGlassHolders
End Property
Public Property Let addGlassHolders(Value As Boolean)
    z_addGlassHolders = Value
End Property


Public Property Get OrderN() As String
    OrderN = z_OrderN
End Property
Public Property Let OrderN(Value As String)
    z_OrderN = Value
End Property


Public Property Get PProfil() As Boolean
    PProfil = z_PProfil
End Property
Public Property Let PProfil(Value As Boolean)
    z_PProfil = Value
    If z_PProfil Then z_addGlassHolders = False
End Property




