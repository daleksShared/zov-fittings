VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "caseProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private casename As String
Private Handle As String
Private HandleExtra As String
Private CaseColor As String
Private napravl As String
Private baseString As String
Private caseLetters As String
Private FasadesString As String
Private ShelfQty As Integer
Private Width As Integer
Private NQty As Integer
Private caseglub As Integer
Private dspbottom As Integer
Private caseHeight As Integer
Private changeZaves As Integer
Private DoorCount As Integer
Private windowcount As Integer
Private cabType As Integer '0-стандарт 1-ZK 2-Modul 3-Optima
Private cabLevel As Integer ' 1-низ 2-верх
Private sizeMin As Integer
Private sizeMax As Integer
Private FasadesCount As Integer
Private ShuflCount As Integer
Private cabDepth As Integer
Private cabHeigth As Integer
Private cabWidth As Integer
Private changeCaseKonfirmant As Integer
Private caseID As Long
Private NoFace As Boolean
Private Drawermount
Private Doormount
Private z_st_dsp As Boolean
Private haveNisha As Boolean
Private NishaQty As Integer
Private haveVVerh As Boolean
Private newname As String
Private newsystem As Boolean
Private CustomerDrawermount As Boolean
Private fullcn As String
Private addGlassHolders As Boolean
Private dvpNahlest As Boolean
Private delete_doormount As Boolean

Public Property Get p_dvpNahlest() As Boolean
    p_dvpNahlest = dvpNahlest
End Property
Public Property Let p_dvpNahlest(Value As Boolean)
    dvpNahlest = Value
End Property

Public Property Get p_delete_doormount() As Boolean
    p_delete_doormount = delete_doormount
End Property
Public Property Let p_delete_doormount(Value As Boolean)
    delete_doormount = Value
End Property

Public Property Get p_casename() As String
    p_casename = casename
End Property
Public Property Let p_casename(Value As String)
    casename = Value
End Property

Public Property Get p_Handle() As String
    p_Handle = Handle
End Property
Public Property Let p_Handle(Value As String)
    Handle = Value
End Property

Public Property Get p_HandleExtra() As String
    p_HandleExtra = HandleExtra
End Property
Public Property Let p_HandleExtra(Value As String)
    HandleExtra = Value
End Property

Public Property Get p_CaseColor() As String
    p_CaseColor = CaseColor
End Property
Public Property Let p_CaseColor(Value As String)
    CaseColor = Value
End Property

Public Property Get p_napravl() As String
    p_napravl = napravl
End Property
Public Property Let p_napravl(Value As String)
    napravl = Value
End Property

Public Property Get p_baseString() As String
    p_baseString = baseString
End Property
Public Property Let p_baseString(Value As String)
    baseString = Value
End Property

Public Property Get p_caseLetters() As String
    p_caseLetters = caseLetters
End Property
Public Property Let p_caseLetters(Value As String)
    caseLetters = Value
End Property

Public Property Get p_FasadesString() As String
    p_FasadesString = FasadesString
End Property
Public Property Let p_FasadesString(Value As String)
    FasadesString = Value
End Property

Public Property Get p_ShelfQty() As Integer
    p_ShelfQty = ShelfQty
End Property
Public Property Let p_ShelfQty(Value As Integer)
    ShelfQty = Value
End Property

Public Property Get p_Width() As Integer
    p_Width = Width
End Property
Public Property Let p_Width(Value As Integer)
    Width = Value
End Property

Public Property Get p_NQty() As Integer
    p_NQty = NQty
End Property
Public Property Let p_NQty(Value As Integer)
    NQty = Value
End Property

Public Property Get p_caseglub() As Integer
    p_caseglub = caseglub
End Property
Public Property Let p_caseglub(Value As Integer)
    caseglub = Value
End Property


Public Property Get p_dspbottom() As Integer
    p_dspbottom = dspbottom
End Property
Public Property Let p_dspbottom(Value As Integer)
    dspbottom = Value
End Property


Public Property Get p_caseHeight() As Integer
    p_caseHeight = caseHeight
End Property
Public Property Let p_caseHeight(Value As Integer)
    caseHeight = Value
End Property

Public Property Get p_changeZaves() As Integer
    If changeZaves = 1 And Left(fullcn, 3) = "ПЛХ" Then
        p_changeZaves = 0
    Else
        p_changeZaves = changeZaves
    End If
End Property
Public Property Let p_changeZaves(Value As Integer)
    changeZaves = Value
End Property

Public Property Get p_DoorCount() As Integer
    p_DoorCount = DoorCount
End Property
Public Property Let p_DoorCount(Value As Integer)
    DoorCount = Value
End Property
Public Property Get p_windowcount() As Integer
    p_windowcount = windowcount
End Property
Public Property Let p_windowcount(Value As Integer)
    windowcount = Value
End Property

Public Property Get p_cabTypeName() As String
    If cabType = 0 Then
        p_cabTypeName = "ZOV"
    ElseIf cabType = 1 Then
        p_cabTypeName = "ZK"
    ElseIf cabType = 2 Then
        p_cabTypeName = "MODUL"
    ElseIf cabType = 3 Then
        p_cabTypeName = "OPTIMA"
    Else
        p_cabTypeName = ""
    End If
End Property


'0-стандарт 1-ZK 2-zovmodul
Public Property Get p_cabType() As Integer
    p_cabType = cabType
End Property
Public Property Let p_cabType(Value As Integer)
    cabType = Value
End Property

' 1-низ 2-верх
Public Property Get p_cabLevel() As Integer
    p_cabLevel = cabLevel
End Property
Public Property Let p_cabLevel(Value As Integer)
    cabLevel = Value
End Property

Public Property Get p_sizeMin() As Integer
    p_sizeMin = sizeMin
End Property
Public Property Let p_sizeMin(Value As Integer)
    sizeMin = Value
End Property

Public Property Get p_sizeMax() As Integer
    p_sizeMax = sizeMax
End Property
Public Property Let p_sizeMax(Value As Integer)
    sizeMax = Value
End Property
Public Property Get p_FasadesCount() As Integer
    p_FasadesCount = FasadesCount
End Property
Public Property Let p_FasadesCount(Value As Integer)
    FasadesCount = Value
End Property
Public Property Get p_ShuflCount() As Integer
    p_ShuflCount = ShuflCount
End Property
Public Property Let p_ShuflCount(Value As Integer)
    ShuflCount = Value
End Property
Public Property Get p_cabDepth() As Integer
    p_cabDepth = cabDepth
End Property
Public Property Let p_cabDepth(Value As Integer)
    cabDepth = Value
End Property
Public Property Get p_cabHeigth() As Integer
    p_cabHeigth = cabHeigth
End Property
Public Property Let p_cabHeigth(Value As Integer)
    cabHeigth = Value
End Property
Public Property Get p_cabWidth() As Integer
    p_cabWidth = cabWidth
End Property
Public Property Let p_cabWidth(Value As Integer)
    cabWidth = Value
End Property
Public Property Get p_changeCaseKonfirmant() As Integer
    p_changeCaseKonfirmant = changeCaseKonfirmant
End Property
Public Property Let p_changeCaseKonfirmant(Value As Integer)
    changeCaseKonfirmant = Value
End Property
Public Property Get P_caseID() As Integer
    P_caseID = caseID
End Property
Public Property Let P_caseID(Value As Integer)
    caseID = Value
End Property

Public Property Get p_NoFace() As Boolean
    p_NoFace = NoFace
End Property
Public Property Let p_NoFace(Value As Boolean)
    NoFace = Value
End Property

Public Property Get p_Drawermount() As String
    p_Drawermount = Drawermount
End Property
Public Property Let p_Drawermount(Value As String)
    Drawermount = Value
End Property

Public Property Get p_Doormount() As String
    p_Doormount = Doormount
End Property
Public Property Let p_Doormount(Value As String)
    Doormount = Value
End Property


Public Property Get p_z_st_dsp() As Boolean
    p_z_st_dsp = z_st_dsp
End Property
Public Property Let p_z_st_dsp(Value As Boolean)
    z_st_dsp = Value
End Property
Public Property Get p_haveNisha() As Boolean
    p_haveNisha = haveNisha
End Property
Public Property Let p_haveNisha(Value As Boolean)
    haveNisha = Value
End Property
Public Property Get p_NishaQty() As Integer
    p_NishaQty = NishaQty
End Property
Public Property Let p_NishaQty(Value As Integer)
    NishaQty = Value
End Property
Public Property Get p_haveVVerh() As Boolean
    p_haveVVerh = haveVVerh
End Property
Public Property Let p_haveVVerh(Value As Boolean)
    haveVVerh = Value
End Property
Public Property Get p_newname() As String
    p_newname = newname
End Property
Public Property Let p_newname(Value As String)
    newname = Value
End Property
Public Property Get p_newsystem() As Boolean
    p_newsystem = newsystem
End Property
Public Property Let p_newsystem(Value As Boolean)
    newsystem = Value
End Property
Public Property Get p_CustomerDrawermount() As Boolean
    p_CustomerDrawermount = CustomerDrawermount
End Property
Public Property Let p_CustomerDrawermount(Value As Boolean)
    CustomerDrawermount = Value
End Property
Public Property Get p_fullcn() As String
    p_fullcn = fullcn
End Property
Public Property Let p_fullcn(Value As String)
    fullcn = Value
    casename = Value
    Dim tempStr As String
    tempStr = Value
    
    dspbottom = 0
     Dim dd As Integer
     If InStr(1, tempStr, "ДНО", vbTextCompare) > 0 And InStr(1, tempStr, "ДСП", vbTextCompare) > 0 Then
        dd = InStr(1, tempStr, "ДСП", vbTextCompare) - InStr(1, tempStr, "ДНО", vbTextCompare)
        If dd > 2 And dd < 20 Then
            dspbottom = 1
        End If
    End If
    
    cabType = 0
    changeCaseKonfirmant = 0
    
    If InStr(1, tempStr, "О", vbTextCompare) = 1 Then
        cabType = 3
        tempStr = "Ш" & Mid(tempStr, 2)
        fullcn = tempStr
        casename = tempStr
    ElseIf mRegexp.regexp_check(patCaseIsZovModul, tempStr) Then
        cabType = 2 '0-стандарт 1-ZK 2-zovmodul
        tempStr = Mid(Trim(tempStr), 1, Len(Trim(tempStr)) - 2)
    Else
        dd = InStr(1, tempStr, " ZK", vbTextCompare)
        If dd > 2 Then
            changeCaseKonfirmant = 1
            cabType = 1 '0-стандарт 1-ZK
        Else
            cabType = 0 '0-стандарт 1-ZK
        End If
    End If
    If InStr(1, tempStr, "фасад стекло", vbTextCompare) > 0 Then
            delete_doormount = True
            Else
            delete_doormount = False
    End If
    
    z_st_dsp = False
     dd = InStr(1, tempStr, "з-ст дсп", vbTextCompare)
    If dd = 0 Then dd = InStr(1, tempStr, "з ст дсп", vbTextCompare)
    If dd = 0 Then dd = InStr(1, tempStr, "зад. ст дсп", vbTextCompare)
    If dd = 0 Then dd = InStr(1, tempStr, "з ст шпон", vbTextCompare)
    If dd = 0 Then dd = InStr(1, tempStr, "з-ст шпон", vbTextCompare)
    If dd = 0 Then dd = InStr(1, tempStr, "зад. ст шпон", vbTextCompare)
    If dd > 0 Then
        z_st_dsp = True
    End If
    
    dvpNahlest = False
    If InStr(1, fullcn, "двп") > 0 Then
        If InStr(InStr(1, fullcn, "двп"), fullcn, "нахл") > 0 Then
            If InStr(InStr(1, fullcn, "двп"), fullcn, "нахл") - InStr(1, fullcn, "двп") < 8 Then
                dvpNahlest = True
            End If
        End If
    End If

    
    If mRegexp.regexp_check(patZKGetDepth, tempStr) Then
       casename = regexp_ReturnStringBySumOfMatches(patZKGetWithoutDepth, tempStr)
    End If
    
    If changeCaseKonfirmant = 0 Then
        dd = InStr(1, tempStr, "стяж", vbTextCompare)
        If dd > 0 Then
            changeCaseKonfirmant = 1
        End If
    End If
    
    cabDepth = 0
     If cabType = 1 And regexp_check(patZKGetDepth, tempStr) Then
        cabDepth = CInt(regexp_ReturnSearch(patZKGetDepth, tempStr))
        If cabDepth < 100 Then cabDepth = CInt(regexp_ReturnSearch(patZKGetDepth, tempStr) & 0)
    ElseIf mRegexp.regexp_check(patGetDepth, tempStr) Then
        cabDepth = CInt(mRegexp.regexp_ReturnSearch(patGetDepthValue, tempStr))
        If cabDepth < 100 Then cabDepth = CInt(mRegexp.regexp_ReturnSearch(patGetDepthValue, tempStr) & "0")
    End If
    
    'If cabDepth = 0 Then
        dd = InStr(1, tempStr, "Н", vbTextCompare)
        If dd = 2 Then
            cabLevel = 2
            If cabType = 1 And cabDepth = 0 Then
                cabDepth = 350
            ElseIf cabDepth = 0 Then
                cabDepth = 300
            End If
        End If
        dd = InStr(1, tempStr, "Л", vbTextCompare)
        If dd = 2 Then
            cabLevel = 1
            If cabDepth = 0 Then
                cabDepth = 570
            End If
        End If
   ' End If
    
    cabHeigth = 0
    If regexp_check(patGetBase, tempStr) Then
        baseString = "(" & mRegexp.regexp_ReturnSearch(patGetBaseValue, tempStr) & ")"
        cabHeigth = CInt(mRegexp.regexp_ReturnSearch(patGetBaseValue, tempStr))
    Else
        cabHeigth = 720
    End If
    
    cabWidth = 0
    If mRegexp.regexp_check(patGetWidth, tempStr) Then
        cabWidth = CInt(mRegexp.regexp_ReturnSearch(patGetWidthValue, tempStr) & "0")
    End If

    If cabType = 1 Then casename = Replace(casename, " ZK", "")
    If cabType = 2 Then casename = tempStr
    caseLetters = mRegexp.regexp_ReturnSearch(patGetFirstLetters, fullcn)
    If InStr(1, caseLetters, "В", vbTextCompare) > 1 Then addGlassHolders = True
    
End Property
Public Property Get p_addGlassHolders() As Boolean
    p_addGlassHolders = addGlassHolders
End Property
Public Property Let p_addGlassHolders(Value As Boolean)
    addGlassHolders = Value
End Property


Public Sub init()
Doormount = Empty
Drawermount = Empty
caseHeight = Empty
caseglub = Empty
CaseColor = Empty
NQty = Empty
ShelfQty = Empty
Width = Empty
HandleExtra = Empty
Handle = Empty
caseID = Empty
NoFace = Empty
cabDepth = 0
cabWidth = 0
cabHeigth = 0
changeZaves = 0
casename = ""
DoorCount = 0
FasadesCount = 0
ShuflCount = 0
FasadesString = ""
haveNisha = False
haveVVerh = False
napravl = ""
baseString = ""
windowcount = 0
sizeMin = 0
sizeMax = 0
newname = ""
newsystem = False
dspbottom = 0
changeCaseKonfirmant = 0
cabDepth = 0
cabHeigth = 0
CustomerDrawermount = False
fullcn = ""
NishaQty = 0
cabType = 0
z_st_dsp = False
cabLevel = Empty
cabType = Empty
addGlassHolders = False
dvpNahlest = False
End Sub


