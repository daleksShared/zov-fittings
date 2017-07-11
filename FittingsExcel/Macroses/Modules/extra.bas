Attribute VB_Name = "extra"
Option Explicit
Option Compare Text

Private conn As ADODB.Connection
Public ConnBarcode As ADODB.Connection
'Private Const ConnStr As String = "DRIVER=SQL Native Client;SERVER=server\zowmain;DATABASE=Fittings;Integrated_Security=True"
'Private Const ConnStr As String = "DRIVER=SQL server;SERVER=server\barcode;DATABASE=Fittings;"
'Public Const BCCS As String = "DRIVER={SQL Server};SERVER=server\barcode;DATABASE=Barcode;"
Private Const ConnStr As String = "DRIVER=SQL server;SERVER=ZSDB\MAIN;DATABASE=Fittings;"
'Private Const ConnStr As String = "DRIVER={SQL Server Native Client 10.0};Trusted_Connection=Yes;SERVER=ZSDB\MAIN;DATABASE=Fittings;Connection Timeout=30;"
Public Const BCCS As String = "DRIVER={SQL Server};SERVER=ZSDB\MAIN;DATABASE=Barcode;"
'Public Const BCCS As String = "DRIVER={SQL Server Native Client 10.0};Trusted_Connection=Yes;SERVER=ZSDB\MAIN;DATABASE=Barcode;Connection Timeout=30;"
'Public Const CCS  As String = "DRIVER={SQL Server Native Client 10.0};Trusted_Connection=Yes;SERVER=ZSDB\main;DATABASE=Cases;Connection Timeout=30;"

Public rsColor As ADODB.Recordset

Public FormShip As NewShipForm
Public FormFitting As AddFitting
Public FormElement As AddElement
Public FormColor As ColorForm
Public FormSearchReplace As frmSearchReplace

Public Const cHandle As String = "ручк"
Public Const cLeg As String = "ножк"
Public Const cPlank As String = "планк"
Public Const cGalog As String = "галог"
Public Const cSink As String = "мойка"
Public Const cStol As String = "стол " '!
Public Const cStul As String = "стул"
Public Const cStool As String = "табурет"
Public Const cNogi As String = "ноги"
Public Const cSit As String = "сидушка"


Public rsOrderFittings As ADODB.Recordset
Public rsOrderElements As ADODB.Recordset
Public rsOrderReplacements As ADODB.Recordset
Public rsCases As ADODB.Recordset
Public OrderCaseID As Long
'Public rsOrderCases As ADODB.Recordset
Public rsOrderCasesParams As ADODB.Recordset
Public rsHandle As ADODB.Recordset
Public rsLeg As ADODB.Recordset

'Public Type OCP
'     param_name As String
'     param_value As String
'End Type

Public paramsId As Integer


'переменные

Public FittingArray(), HandleArray(), LegArray()
Public OtbColors(), Doormount(), Plank(), Galog(), Rell(), Лифт(), Заглушки(), ЗаглЭксц(), Завешки(), Карго(), Полка()
Public OtbGorbColors(), vytyazhka_perfim()
Public zavesHL(), zavesHS(), zavesSensys(), zavesClipTop(), ploschadkaSensys()
Public tbLength()
Public tbkovrLength()
Public tbkovrOpt()
' гарнитура
'Private Stul(), SitK(), Крышка(), Спинка(), Палки()
'Private  Stol(),  Sink(), Sit(), SitColors(), SitKolib(), BackKolib()
Public StulNogi(), SW_bel(), SW(), LW(), PA(), Вставка(), Цоколь(), СоединительЦоколя(), Направляющие(), Отбойники(), Стекло() ', НогиСтол()
Public Sushk(), ГорбатаяМелочь(), Полкодержатель(), TOPLine(), МелочьОтб4м()
Public Stul_color_no(), Stul_color_1(), Stul_color_2()
Public MoikaColors()
'
Public params As Collection
Public param As caseParam

Public CaseFittingsCollection As Collection
Public caseFittingCurrent As caseOrderFitting

Public CaseElementsCollection As Collection
Public caseElementCurrent As caseOrderElement

Public casepropertyCurrent As caseProperty
Public kitchenPropertyCurrent As kitchenProperty

Public Sub addItem2param(param_name As String, Optional ByVal param_value As String = "")
Set param = New caseParam
param.paramName = param_name
param.paramValue = param_value
params.Add param
End Sub



Public Sub additem2caseFittings(ByVal OrderId As Long, _
                                    ByVal name As String, _
                                    ByVal qty, _
                                    Optional ByRef Opt = Empty, _
                                    Optional ByRef length = Empty, _
                                    Optional ByVal caseID, _
                                    Optional ByVal Standart As Boolean = False, _
                                    Optional ByVal RowN As Integer = 0)
Set caseFittingCurrent = New caseOrderFitting
caseFittingCurrent.fName = name
If IsNull(qty) = False And IsEmpty(qty) = False Then caseFittingCurrent.fQty = qty
If IsNull(Opt) = False Then caseFittingCurrent.fOption = Opt Else caseFittingCurrent.fOption = Empty
caseFittingCurrent.fLength = length
CaseFittingsCollection.Add caseFittingCurrent

End Sub
Public Function additem2caseElements(ByVal OrderId As Long, _
                                    ByVal name As String, _
                                    ByVal qty, _
                                    Optional ByVal caseID) As Boolean
Set caseElementCurrent = New caseOrderElement
caseElementCurrent.name = name
caseElementCurrent.qty = qty
CaseElementsCollection.Add caseElementCurrent
End Function

'Public rs As ADODB.Recordset
Public Property Get GetConnection() As ADODB.Connection

    If conn Is Nothing Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = ConnStr
    End If
    
    If conn.State <> 1 Then
        conn.Open
    End If

    Set GetConnection = conn
End Property



Sub InitBarcodeConnection(ByRef conn As ADODB.Connection)
    If conn Is Nothing Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = BCCS 'fLogin.BarcodeConnStr
        conn.Open
    ElseIf conn.State = 0 Then
        conn.Open
    End If
End Sub


Public Sub DelSymbol(ByRef str As String, ByVal pos As Integer)
    str = Left(str, pos - 1) & Mid(str, pos + 1)
End Sub

Public Sub DelTextLeft(ByRef str As String, Optional SymbolDeletedQty As Integer)
    SymbolDeletedQty = 0
    While Not IsNumeric(Left(str, 1)) And Len(str)
        str = Mid(str, 2)
        SymbolDeletedQty = SymbolDeletedQty + 1
    Wend
End Sub

Public Sub DelTextRight(ByRef str As String, Optional SymbolDeletedQty As Integer)
    SymbolDeletedQty = 0
    While Not IsNumeric(Right(str, 1)) And Len(str)
        str = Left(str, Len(str) - 1)
        SymbolDeletedQty = SymbolDeletedQty + 1
    Wend
End Sub

Public Sub CheckHandle(ByRef Handle)
    Init_rsHandle
    If IsNull(Handle) Then Handle = Empty
    If IsMissing(Handle) Then Handle = Empty
    
    Do
        If InStr(1, Handle, "фарфор", vbTextCompare) = 0 And _
             InStr(1, Handle, "модена", vbTextCompare) = 0 And _
             InStr(1, Handle, "барх", vbTextCompare) = 0 Then
            Dim k As Integer
            k = InStr(1, Handle, " на", vbTextCompare)
            If k > 1 Then Handle = Trim(Left(Handle, k))
            
            Handle = Replace(Handle, "/", "")
            Handle = Replace(Handle, " ", "")
            Handle = Replace(Handle, ".", "")
            Handle = Replace(Handle, "-", "")
        End If
        
        If rsHandle.RecordCount > 0 Then rsHandle.MoveFirst
        rsHandle.Find "Handle='" & Handle & "'"
        If rsHandle.EOF Then
            'MsgBox "Неизвестный тип ручек - " & Handle, vbCritical
            'Handle = InputBox("Введите тип ручек", "Ручки заказа по умолчанию", Handle)
            
            Dim fHandle As ufHandles
            Set fHandle = New ufHandles
            fHandle.cbHandles.Text = Handle
            fHandle.Show
            If fHandle.cbNoHandle.Value Then
                Handle = Null
                Exit Do
            Else
                Handle = fHandle.cbHandles.Value
            End If
            
'            If Handle = "клиента" Or Handle = "-" Then
'                Handle = Null
'                Exit Do
'            End If
            'Init_rsHandle False
        Else
            Exit Do
        End If
    Loop While 1
End Sub

Public Sub CheckLeg(ByRef Leg)
    Init_rsLeg
    
    Do
        If InStr(1, Leg, "пластик хром", vbTextCompare) = 0 And _
            InStr(1, Leg, "780 хром", vbTextCompare) = 0 And _
            InStr(1, Leg, "780 бронза", vbTextCompare) = 0 And _
            InStr(1, Leg, "волп", vbTextCompare) = 0 And _
            InStr(1, Leg, "черн", vbTextCompare) = 0 Then _
            Leg = Replace(Leg, " ", "")
        
        If rsLeg.RecordCount > 0 Then rsLeg.MoveFirst
        rsLeg.Find "Leg='" & Leg & "'"
        If rsLeg.EOF Then
            MsgBox "Неизвестный тип ножек - " & Leg, vbCritical
            Leg = InputBox("Введите тип ножек", "Ножки заказа по умолчанию", Leg)
            
            If Leg = "-" Then Leg = "клиента"
            If Leg = "клиента" Then
                Leg = Null
                Exit Sub
            End If
            
            Init_rsLeg False
        Else
            Exit Do
        End If
    Loop While 1
End Sub


Public Sub UpdateOrder(ByVal OrderId As Long, _
                         Optional ByVal HandleScrew, _
                         Optional ByVal HangColor, _
                         Optional ByVal BibbColor, _
                         Optional SetQty, _
                         Optional ByVal face, _
                         Optional ByVal CaseColor, _
                         Optional ByVal ColorId, _
                         Optional ByVal CamBibbColor _
                         )
                         
On Error GoTo err_UpdateOrder

    If (IsEmpty(CamBibbColor) Or IsMissing(CamBibbColor)) And _
        (IsEmpty(HandleScrew) Or IsMissing(HandleScrew)) And _
        (IsEmpty(HangColor) Or IsMissing(HangColor)) And _
        (IsEmpty(SetQty) Or IsMissing(SetQty)) And _
        (IsEmpty(BibbColor) Or IsMissing(BibbColor)) And _
        (IsEmpty(ColorId) Or IsMissing(ColorId)) And _
        (IsEmpty(CaseColor) Or IsMissing(CaseColor)) And _
        (IsEmpty(face) Or IsMissing(face)) Then Exit Sub
    
    
    Dim UpdateComm As ADODB.Command
    Set UpdateComm = New ADODB.Command
    UpdateComm.ActiveConnection = GetConnection
    UpdateComm.CommandType = adCmdStoredProc
    UpdateComm.CommandText = "UpdateOrder"
    
    UpdateComm.Parameters("@OrderID") = OrderId
    If Not IsMissing(HangColor) Then UpdateComm.Parameters("@HangColor") = HangColor
    If Not IsMissing(HandleScrew) Then UpdateComm.Parameters("@HandleScrew") = HandleScrew
    If Not IsMissing(BibbColor) Then UpdateComm.Parameters("@BibbColor") = BibbColor
    If Not IsMissing(CamBibbColor) Then
        UpdateComm.Parameters("@CamBibbColor") = CamBibbColor
    End If
    If Not IsMissing(SetQty) Then UpdateComm.Parameters("@SetQty") = CInt(SetQty)
    If Not IsMissing(face) Then
        If InStr(1, face, "клиент", vbTextCompare) > 0 Then face = ""
        UpdateComm.Parameters("@Face") = Left(face, 50)
        
    End If
    If Not IsMissing(CaseColor) Then UpdateComm.Parameters("@CColor") = Left(CaseColor, 20)
    If Not IsMissing(ColorId) Then
    UpdateComm.Parameters("@ColorId") = ColorId
      End If
    UpdateComm.Execute
    
    Exit Sub
err_UpdateOrder:
    MsgBox Error, vbCritical, "Обновление заказа (UpdateOrder)"
End Sub

Public Sub Init_rsHandle(Optional ByVal init As Boolean = True)
    If init And Not rsHandle Is Nothing Then
        If rsHandle.State = 0 Then init = False
    Else
        init = False
    End If
    
    If Not init Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandText = "SELECT * FROM Handle where Handle<>'1001'"
        
        Set rsHandle = New ADODB.Recordset
        rsHandle.CursorLocation = adUseClient
        rsHandle.LockType = adLockBatchOptimistic
        rsHandle.Open comm, , adOpenDynamic, adLockBatchOptimistic
    End If
End Sub

Public Sub Init_rsLeg(Optional ByVal init As Boolean = True)
    If init And Not rsLeg Is Nothing Then
        If rsLeg.State = 0 Then init = False
    Else
        init = False
    End If
    
    If Not init Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandText = "SELECT * FROM Leg"
        
        Set rsLeg = New ADODB.Recordset
        rsLeg.CursorLocation = adUseClient
        rsLeg.LockType = adLockBatchOptimistic
        rsLeg.Open comm, , adOpenDynamic, adLockBatchOptimistic
    End If
End Sub

'Public Sub Init_rs(Optional ByVal init As Boolean = True)
'    If init And Not rs Is Nothing Then
'        If rs.State = 0 Then init = False
'    Else
'        init = False
'    End If
'
'    If Not init Then
'        Dim comm As ADODB.Command
'        Set comm = New ADODB.Command
'        comm.ActiveConnection = GetConnection
'        comm.CommandType = adCmdText
'        comm.CommandText = "SELECT * FROM [Fittings].[dbo].[fGetCaseFromKB] () ORDER BY Qty DESC"
'
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        rs.Open comm, , adOpenDynamic, adLockReadOnly
'    End If
'End Sub

Public Sub Init_rsOrderElements(Optional ByVal init As Boolean = True)
    If init And Not rsOrderElements Is Nothing Then
        If rsOrderElements.State = 0 Then init = False
    Else
        init = False
    End If
    
    If Not init Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandText = "SELECT  * FROM [ЭлементыЗаказов]"
        
        Set rsOrderElements = New ADODB.Recordset
        rsOrderElements.CursorLocation = adUseClient
        rsOrderElements.LockType = adLockBatchOptimistic
        rsOrderElements.Open comm, , adOpenDynamic, adLockBatchOptimistic
    End If
End Sub

'Public Sub Init_rsOrderCases(Optional ByVal init As Boolean = True)
'    If init And Not rsOrderCases Is Nothing Then
'        If rsOrderCases.State = 0 Then init = False
'    Else
'        init = False
'    End If
'
'    If Not init Then
'        Dim comm As ADODB.Command
'        Set comm = New ADODB.Command
'        comm.ActiveConnection = GetConnection
'        comm.CommandType = adCmdText
'        comm.CommandTimeout = 90
'        comm.CommandText = "SELECT * FROM [ШкафыЗаказов]"
'
'        Set rsOrderCases = New ADODB.Recordset
'        rsOrderCases.CursorLocation = adUseClient
'        rsOrderCases.LockType = adLockBatchOptimistic
'        rsOrderCases.Open comm, , adOpenDynamic, adLockBatchOptimistic
'    End If
'End Sub
Public Sub Init_rsOrderCasesParams(Optional ByVal init As Boolean = True)
    If init And Not rsOrderCasesParams Is Nothing Then
        If rsOrderCasesParams.State = 0 Then init = False
    Else
        init = False
    End If
    
    If Not init Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandTimeout = 90
        comm.CommandText = "SELECT * FROM [OrderCasesParams]"
        
        Set rsOrderCasesParams = New ADODB.Recordset
        rsOrderCasesParams.CursorLocation = adUseClient
        rsOrderCasesParams.LockType = adLockBatchOptimistic
        rsOrderCasesParams.Open comm, , adOpenDynamic, adLockBatchOptimistic
    End If
End Sub

Public Sub Init_rsOrderFittings(Optional ByVal init As Boolean = True)
    If init And Not rsOrderFittings Is Nothing Then
        If rsOrderFittings.State = 0 Then init = False
    Else
        init = False
    End If
    
    If Not init Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandTimeout = 180
        comm.CommandText = "SELECT * FROM [ФурнитураЗаказов]"
        
        Set rsOrderFittings = New ADODB.Recordset
        rsOrderFittings.CursorLocation = adUseClient
        rsOrderFittings.LockType = adLockBatchOptimistic
        rsOrderFittings.Open comm, , adOpenDynamic, adLockBatchOptimistic
    End If
End Sub

Public Sub Init_rsCases(Optional ByVal init As Boolean = True)
    If init And Not rsCases Is Nothing Then
        If rsCases.State = 0 Then init = False
    Else
        init = False
    End If
    
    If Not init Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandText = "SELECT * FROM [Case]"
        
        Set rsCases = New ADODB.Recordset
        rsCases.CursorLocation = adUseClient
        rsCases.LockType = adLockBatchOptimistic
        rsCases.Open comm, , adOpenDynamic, adLockBatchOptimistic
    End If
End Sub
Public Sub Init_rsOrderReplaces(Optional ByVal init As Boolean = True)
'    If init And Not rsOrderReplacements Is Nothing Then
'        If rsOrderReplacements.State = 0 Then init = False
'    Else
'        init = False
'    End If
'
'    If Not init Then
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdText
        comm.CommandText = "SELECT [FindString] ,[ReplaceString],[AskOnFind],[isRegExp],[isfullStringSearch] " & _
                            "FROM [OrderReplacements] " & _
                            "where [Enabled] = 1 " & _
                            "order by ReplaceString asc,FindString asc"
        
        Set rsOrderReplacements = New ADODB.Recordset
        rsOrderReplacements.CursorLocation = adUseClient
        rsOrderReplacements.LockType = adLockBatchOptimistic
        rsOrderReplacements.Open comm, , adOpenDynamic, adLockBatchOptimistic
'    End If
End Sub


Public Sub ParseCase(ByRef casename As String, _
                        ByRef caseID As Integer, _
                        ByRef DoorQty, _
                        ByRef WindowQty, _
                        ByRef Drawermount, _
                        ByRef Doormount, _
                        ByRef BF As Boolean, _
                        ByRef Handle, _
                        ByRef HandleExtra, _
                        ByRef ShelfQty, _
                        ByRef W, _
                        ByRef NQty, _
                        ByVal CaseColor As String, _
                        ByRef caseglub As Integer, _
                        ByRef caseHeight As Integer _
                        )
    Dim caseFur As caseFurniture
    caseID = 0
    Dim tempString As String
    Dim caseElementsIterator As Integer
    Dim name As String, D
    
    
    
    Dim H, SVar, BVar, LVar, FrezBase As Boolean
    Dim ts As String, Nisha As Boolean, KF As Long, fQty, SHQty
    Dim k As Integer, L As Integer
    
    W = Empty
    FrezBase = False
If Not (casepropertyCurrent Is Nothing) Then
    If casepropertyCurrent.p_newsystem Then
        
        
        GoTo newsystem
        
    End If
End If
    If InStr(1, casename, "кв", vbTextCompare) > 0 And _
        InStr(1, casename, "/", vbTextCompare) = 0 Then
        
        casename = InputBox("Исправьте наименование", "Шкаф КВАДРО?", casename)
        
    End If
                        
On Error GoTo err_ParseCase
    
'    If regexp_check(patNewName, casename) Then
'        casename = parse_case(casename)
'    ElseIf regexp_check(patSHLK_check1, casename) Then
'    casename = regexp_replace(patSHLK_check1, casename, "/")
'
'    End If
    If regexp_check(patSHLK_check1, casename) Then casename = regexp_replace(patSHLK_check1, casename, "/")
    
    If casename = "ШНУ60/60" Then
        casename = "ШНУ60"
    ElseIf casename = "ШНУ60/60(600)" Then
        casename = "ШНУ60(600)"
    ElseIf casename = "ШНУ60/60(915)" Then
        casename = "ШНУ60(915)"
  '  ElseIf CaseName = "ШН60(915)Т/2" Then
   '     CaseName = "ШН60Т/2(915)"
    ElseIf casename = "ШЛУ90/90" Then
        casename = "ШЛУ90"
    ElseIf casename = "ШН60/2ВП" Then
        casename = "ШН60ВП"
    ElseIf casename = "ШНЗУЕ60" Then
        casename = "ШНЗУ60"
    ElseIf casename = "ШЛУ90/90Н" Then
        casename = "ШЛУ90Н"
    ElseIf casename = "ШСУ90/90" Then
        casename = "ШСУ90"
    ElseIf casename = "ШСУ90/90Н" Then
        casename = "ШСУ90Н"
    ElseIf Replace(Replace(Replace(casename, " ", ""), ".", ""), "/90", "") = "ШЛУ90/90Нновбаза" Then
        casename = "ШЛУН90/нб"
    ElseIf Replace(Replace(Replace(casename, " ", ""), ".", ""), "/90", "") = "ШЛУ90/90нбаза" Then
        casename = "ШЛУ90/нб"
    ElseIf Replace(Replace(Replace(casename, " ", ""), ".", ""), "/90", "") = "ШЛУ90/90Ннбаза" Then
        casename = "ШЛУН90/нб"
    ElseIf Replace(casename, " ", "") = "ШЛС30скос" Then
        casename = "ШЛ30скос"
    ElseIf Replace(casename, " ", "") = "ШЛУ" Then
        casename = "ШЛУ90"
    ElseIf Replace(casename, " ", "") = "ШСУ" Then
        casename = "ШСУ90"
    ElseIf Replace(casename, " ", "") = "ШНУ" Then
        casename = "ШНУ60"
    ElseIf Replace(casename, " ", "") = "ШНУ(600)" Then
        casename = "ШНУ60(600)"
    ElseIf Replace(casename, " ", "") = "ШНУ(915)" Then
        casename = "ШНУ60(915)"
    ElseIf Replace(casename, " ", "") = "ШЛУН" Then
        casename = "ШЛУ90Н"
    ElseIf Replace(casename, " ", "") = "ШСУН" Then
        casename = "ШСУ90Н"
    End If
    
    If Left(casename, 7) = "ШНЗУЕ65/65" Then
        casename = Replace(casename, "ШНЗУЕ65/65", "ШНЗУ65/65", 1, 1, vbTextCompare)
    End If
    
    If Left(casename, 3) = "ШНК" Then
        casename = Replace(casename, "ШНК", "ШНТ", 1, 1, vbTextCompare)
    End If
    If Left(casename, 3) = "ШНЮ" Then
        casename = Replace(casename, "ШНЮ", "ШН", 1, 1, vbTextCompare)
    End If
    If InStr(1, casename, "ШНУС60/60", vbTextCompare) = 1 Then
        casename = Replace(casename, "ШНУС60/60", "ШНУС60", 1, 1, vbTextCompare)
    End If
    If InStr(1, casename, "ШНУ60/60", vbTextCompare) = 1 Then
        casename = Replace(casename, "ШНУ60/60", "ШНУ60", 1, 1, vbTextCompare)
    End If
    If Left(casename, 4) = "ШНУС" Then
        casename = Replace(casename, "ШНУС", "ШНУ", 1, 1, vbTextCompare)

        Set caseFur = New caseFurniture
        caseFur.init
        caseFur.fName = "конфирмант"
        caseFur.qty = 6
        caseFurnCollection.Add caseFur
    
        Set caseFur = New caseFurniture
        caseFur.init
        caseFur.fName = "шкант"
        caseFur.qty = 6
        caseFurnCollection.Add caseFur
    End If
    
    
    '***************************************************************************************
    
    Dim dd As Integer
    
    
    While Not IsNumeric(Left(casename, 1)) And Len(casename) > 0
        name = name & Left(casename, 1)
        casename = Mid(casename, 2)
    Wend
    
    While IsNumeric(Left(casename, 1)) And Len(casename) > 0
        W = W & Left(casename, 1)
        casename = Mid(casename, 2)
    Wend

    k = InStr(1, casename, "б/ф", vbTextCompare)
    If k Then
        BF = True
        If k > 1 Then
            casename = Left(casename, k - 1) & Mid(casename, k + 3)
        Else
            casename = Mid(casename, k + 3)
        End If
    Else
        BF = False
    End If
    
    casename = Trim(Replace(casename, "тандем", "тб", , , vbTextCompare))
    
    
    k = InStr(1, casename, "/", vbTextCompare)
    If k And k < 8 Then
        Do
            Select Case Mid(casename, k, 1)
                
                Case "/", "1", "2", "3", "4", "-"
                        SVar = SVar & Mid(casename, k, 1)
                
                        If k > 1 Then
                            casename = Left(casename, k - 1) & Mid(casename, k + 1)
                        Else
                            casename = Mid(casename, k + 1)
                        End If
                
                Case "м", "т", "к", "а"
                    If Len(casename) > k Then
                    
                        Select Case Mid(casename, k, 4)
                            Case "мб-б", "тб-б", "мб-1", "кв-1", "тб-1", "2тбм"
                                SVar = SVar & Mid(casename, k, 4)
                        
                                If k > 1 Then
                                    casename = Left(casename, k - 1) & Mid(casename, k + 4)
                                Else
                                    casename = Mid(casename, k + 4)
                                End If
                                
                                Exit Do
                        End Select
                        
                        Select Case Mid(casename, k, 3)
                            Case "мбТ", "тбТ", "квТ", "тбм", "тбб", "мбб", "мбм", "арг"
                                SVar = SVar & Mid(casename, k, 3)
                                
                                If k > 1 Then
                                    casename = Left(casename, k - 1) & Mid(casename, k + 3)
                                Else
                                    casename = Mid(casename, k + 3)
                                End If
                                
                                Exit Do
                            
                        End Select
                        
                        Select Case Mid(casename, k, 2)
                            Case "мб", "тб", "кв"
                                SVar = SVar & Mid(casename, k, 2)
                            
                                If k > 1 Then
                                    casename = Left(casename, k - 1) & Mid(casename, k + 2)
                                Else
                                    casename = Mid(casename, k + 2)
                                End If
                            
                                Exit Do
                            Case Else
                                Exit Do
                        End Select
                    Else
                        
                        LVar = LVar & Mid(casename, k, 1)
                        If k > 1 Then
                            casename = Left(casename, k - 1) & Mid(casename, k + 2)
                        Else
                            casename = Mid(casename, k + 2)
                        End If
                        Exit Do
                    End If
                
                Case Else
                   Exit Do
            End Select
            
        Loop While Len(casename) >= k
    Else
        If Len(casename) >= 2 Then
            Select Case Left(casename, 3)
                Case "мбб", "мбм", "тбб", "тбм", "арг"
                    SVar = "/" & Left(casename, 3)
                    If Len(casename) > 3 Then casename = Mid(casename, 4) Else casename = ""
                Case Else
                    Select Case Left(casename, 2)
                    Case "мб", "тб"
                        SVar = "/" & Left(casename, 2)
                        If Len(casename) > 2 Then casename = Mid(casename, 3) Else casename = ""
                    End Select
            End Select
        End If
    End If
    If SVar = "/" Then SVar = Empty
        
        
    casename = Trim(casename)
    k = InStr(1, casename, "глуб.")
    
    If k Then
        casename = Left(casename, k - 1) & LTrim(Mid(casename, k + 5))
        casename = LTrim(casename)
        
    Else
        k = InStr(1, casename, "глуб")
        If k Then
        casename = Left(casename, k - 1) & LTrim(Mid(casename, k + 4))
        casename = LTrim(casename)
        End If
    End If
    
    If k Then
    L = 1
        While IsNumeric(Mid(casename, k, L)) And k + L - 1 < Len(casename)
            L = L + 1
        Wend
        If Not IsNumeric(Mid(casename, k, L)) Then L = L - 1
        If L > 0 Then D = CDec(Mid(casename, k, L))
            
            If D < 150 Then D = D * 10
        
        casename = Left(casename, k - 1) & Mid(casename, k + L)
        casename = LTrim(casename)
        
        L = InStr(1, casename, "см")
        If L Then casename = Left(casename, L - 1) & Mid(casename, L + 2)
    End If
    If IsEmpty(D) Then
        If SVar = "\нб" Then
            SVar = Empty
            D = 570 '530
        ElseIf InStr(1, casename, "база") Then
            D = 570 '530
        End If
        
        If InStr(1, casepropertyCurrent.p_fullcn, "Н", vbTextCompare) = 2 Then
            D = 300
            Else
            D = 570
        End If
        
    End If
   ' casepropertyCurrent.p_cabDepth = CInt(D)
    casename = Trim(casename)
    
    
   
    ' получаем кол-во фасадов и витрин
    Dim isWindow As Boolean
    If Asc(Right(name, 1)) = 194 Then
        isWindow = True
        name = Left(name, Len(name) - 1)
    Else
        isWindow = False
    End If
    
    ' DrawerMount
    Drawermount = GetDrawerMount()
    
    
    
    
    
    casename = Trim(casename)
    
'    k = InStr(1, CaseName, "(")
'    If k Then
'
'        l = InStr(1, CaseName, ")")
'        If l = 0 Then 'если нет закрывающей скобки
'            For l = k + 1 To Len(CaseName)
'                If Not IsNumeric(Mid(CaseName, l, 1)) Then Exit For
'            Next l
'        End If
'        BVar = Mid(CaseName, k, l - k + 1)
'
'        CaseName = Left(CaseName, k - 1) & Mid(CaseName, l + 1)
    If Len(casename) > 0 Then
        
        k = InStr(1, casename, "(")
        If k Then
            L = InStr(1, casename, ")")
            If L = 0 Then 'если нет закрывающей скобки
                For L = k + 1 To Len(casename)
                    If Not IsNumeric(Mid(casename, L, 1)) Then Exit For
                Next L
            End If
            BVar = Mid(casename, k, L - k + 1)
            
            casename = Left(casename, k - 1) & Mid(casename, L + 1)
            
            BVar = Replace(BVar, "(", "")
            BVar = Replace(BVar, ")", "")
        Else
            BVar = casename
            'CaseName = "" '???????????????
        End If
        
        If InStr(BVar, ",") > 0 Then
        
            BVar = Replace(BVar, " ", "")
            
            
            
            'вычисляем высоту пенала
            Dim tBvar As Variant
            tBvar = BVar
            H = 0
            KF = 0
            fQty = 0 '!!!!
            SHQty = 0
            NQty = 0
            Nisha = False
            While Len(tBvar) > 0
                k = InStr(tBvar, ",")
                If k > 0 Then
                    ts = Trim(Left(tBvar, k - 1))
                    tBvar = Mid(tBvar, k + 1)
                Else
                    ts = tBvar
                    tBvar = Null
                End If
                
                ts = Replace(ts, "вверх", "", , , vbTextCompare)
                ts = Replace(ts, "верх", "", , , vbTextCompare)
                ts = Replace(ts, "вниз", "", , , vbTextCompare)
                    
                'Проверим, не указаны ли шуфл/фасады/ниши через х
                L = InStr(ts, "х") 'рус х
                If L = 0 Then L = InStr(ts, "x") ' англ x
                If L = 0 Then L = InStr(ts, "/") ' просто /
                If L = 0 Then L = InStr(ts, "*") ' просто *
                If L > 0 Then L = CInt(Val(LTrim(Mid(ts, L + 1)))) Else L = 1 'кол-во фасадов
                While L > 4
                    L = InputBox("Введите кол-во фасадов (" & ts & ") по высоте", "Кол-во фасадов", L)
                Wend
                
                If InStr(1, ts, "н", vbTextCompare) = 0 Then
                    KF = KF + L
                    
                    If InStr(1, ts, "ш", vbTextCompare) > 0 Or L > 1 Then
                        SHQty = SHQty + L
                    Else
                        fQty = fQty + L
                    End If
                    
                    'If Nisha And KF > 0 Then KF = KF - 1
                    Nisha = False
                Else
                    'если ниша сверху или после ниши, добавляем к высоте высоту полика 16 (для ПС*, ПЛ*, ПН*)
                    If H = 0 Or Nisha Then H = H + 16 Else KF = KF + 1
                    H = H + 16 * (L - 1) 'если  ниши указаны ч/з х
                    Nisha = True
                    NQty = NQty + 1
                End If
                
                If Not Nisha And InStr(1, ts, "вит", vbTextCompare) > 0 Then
                    WindowQty = WindowQty + L
                End If
                
                'добавляем к высоте шкафа высоту ниши/фасада
                While Not IsNumeric(Left(ts, 1)) And Trim(ts) <> ""
                    ts = Mid(ts, 2)
                Wend
                H = H + CInt(Val(ts)) * L
            Wend
        
            'если ниша снизу, добавляем к высоте высоту полика 16 (для ?С*, ?Н*, но не ?Л*,т.к. у ?Л* снизу крышка)
            If Nisha And Mid(name, 2, 1) <> "Л" Then H = H + 16
            'добавляем зазор
            If Not Nisha Then
                H = H + (KF + 1) * 3
            Else
                H = H + KF * 3
            End If
            
            Select Case Mid(name, 2, 1)
                Case "С"
                    H = H + 97
                Case "Л"
                    H = H - 16
                Case "Н"
                    If H = 719 Then
                        H = 720
                    ElseIf H = 1271 Then
                        H = 1280
                    End If
            End Select
        Else
            H = BVar
        End If
        
    End If
    
    If InStr(1, H, "фр", vbTextCompare) Then
        H = Replace(H, "фр", "", 1, 1, vbTextCompare)
        If IsNumeric(H) Then FrezBase = True
    End If
    
    If IsNumeric(H) And Len(H) < 5 Then
        BVar = Empty
    Else
        H = Empty
    End If
    
    casename = Trim(casename)
    
    If k Then
        If Len(casename) >= 2 And IsEmpty(SVar) Then
            Select Case Left(casename, 3)
                Case "мбб", "мбм", "тбб", "тбм", "арг"
                    SVar = "/" & Left(casename, 3)
                    If Len(casename) > 3 Then casename = Mid(casename, 4) Else casename = ""
                Case Else
                    Select Case Left(casename, 2)
                    Case "мб", "тб"
                        SVar = "/" & Left(casename, 2)
                        If Len(casename) > 2 Then casename = Mid(casename, 3) Else casename = ""
                    End Select
            End Select
        End If
    End If
    If SVar = "/" Then SVar = Empty
    
    If Len(casename) > 0 Then
        If Asc(Left(casename, 1)) = 203 Then casename = Trim(Mid(casename, 2))
        
        If Len(casename) > 0 And (InStr(1, casename, "тех", vbTextCompare) <> 1) Then
            Select Case Asc(Left(casename, 1))
                Case 200 '"И"
                    LVar = "И"
                    casename = Trim(Mid(casename, 2))
                Case 210 '"Т"
                    LVar = "Т"
                    casename = Trim(Mid(casename, 2))
                Case 199 '"З"
                    LVar = "З"
                    casename = Trim(Mid(casename, 2))
                Case 192 '"А"
                    If InStr(1, casename, "арка", vbTextCompare) = 0 Then
                        LVar = "А"
                        casename = Trim(Mid(casename, 2))
                    End If
            End Select
        End If
    End If
   
    If LVar = "Т" And Not IsEmpty(SVar) Then
        SVar = SVar & LVar
        LVar = Empty
    End If
    
     If Not (casepropertyCurrent Is Nothing) Then
        If casepropertyCurrent.p_DoorCount > 0 Then
                DoorQty = casepropertyCurrent.p_DoorCount
        End If
        If casepropertyCurrent.p_windowcount > 0 Then
            WindowQty = casepropertyCurrent.p_windowcount
        End If
    End If
    
    
    
    If name = "ШНТ" Then
        SVar = SVar & "Т"
        name = "ШН"
    End If
    
    Select Case Left(name, 1)
      
        Case "Ш"
    
            If Not IsEmpty(H) Then
                If H > 820 Then
                    ShelfQty = 2
                ElseIf H < 500 Then
                    ShelfQty = 0
                End If
            End If
        
            Select Case Mid(name, 2, 1)
                Case "Н"
                    Select Case name
                    
                    
                    Case "ШН10 под БПД4 (б.1451) гл.570мм", _
                     "ШН10 под БПД4 (б.820) гл.483мм", _
                     "ШН10 под БПД4 (б.1442) гл.492мм", _
                     "ШН10 под БПД4 (б.718) гл.492мм", "ШНПД10"



                        Case "ШН", "ШНС"
                            If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                            If casepropertyCurrent.p_DoorCount > 0 Then DoorQty = casepropertyCurrent.p_DoorCount
                            Select Case SVar
                                
                                Case ""
                                    If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "скос") = 0 Then name = name & "915"

                                If Check_ВП(casename) Then
                                    name = name & " вп"
                                    Doormount = "+20"
                                End If
                            
'                                Case Empty
'                                    If IsEmpty(LVar) Then
'                                        If W <= 50 Then DoorQty = 1 Else DoorQty = 2
'                                    Else
'                                        Do
'                                            DoorQty = InputBox("введите кол-во фасадов", "кол-во фасадов")
'                                        Loop Until IsNumeric(DoorQty)
'                                    End If
                                Case "/1"
                                    If InStr(casename, "верх") > 0 And Check_ВП(casename) Then
                                        name = name & "/1 вп"
                                        Doormount = "110"
                                        DoorQty = InputBox("Уточните кол-во дверей", "Кол-во дверей шкафа", DoorQty)
                                    Else
                                    
                                        If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "скос") = 0 Then name = name & "915"
                                        DoorQty = 1
                                    End If
                                Case "/2"
                                
                                    If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "скос") = 0 Then name = name & "915"
                                    
                                    DoorQty = 2
                                    
                                    If InStr(1, casename, "HF") > 0 Then Doormount = Null
                                    If InStr(1, casename, "HK") > 0 Then Doormount = Null
                                    
                                Case "/1Т"
                                    DoorQty = 1
                                    name = name & " Т"
                                Case "/2Т"
                                    DoorQty = 2
                                    
                                    If InStr(1, Trim(casename), "HF") > 0 Then
                                        Doormount = Null
                                        name = name & " Т"
                                        'If DoorQty = 2 Then DoorQty = 1
                                        DoorQty = 0
                                    ElseIf InStr(1, Trim(casename), "FB-1") > 0 Then
                                        Doormount = Null
                                        name = name & " Т"
                                    ElseIf InStr(1, Trim(casename), "AV") > 0 Then
                                        Doormount = Null
                                        name = name & " Т"
                                    Else
                                        name = name & " 2Т"
                                    End If
                                
                                    
                                Case Else
                                
                                    If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "скос") = 0 Then name = name & "915"
                                
                                    If InStr(1, Trim(casename), "HK") > 0 Then ' + HK-S
                                        Doormount = Null
                                    End If
                            End Select
                            
                            ' для зоси с агатой
                            Select Case LVar
                                Case Empty
                                Case "А"
                                    name = "ШНВ А"
                                    DoorQty = 2
                                    WindowQty = 3
                                    LVar = Empty
                                Case "З"
                                    name = "ШНВ З"
                                    DoorQty = 0
                                    'DoorMount = "к-т к стеклу"
                                    LVar = Empty
                                Case "Т"
                                    'ShelfQty = 1  если 2- то 2
                                    name = name & " Т"
                                    'DoorQty = 1
                                Case Else
                                    name = ""
                                    MsgBox "Неизвестный шкаф", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                            If FrezBase Then name = name & " фр"
                            
                            If InStr(casename, "скос") Then
                                name = name & " скос"
                                If ShelfQty >= 2 Then name = name & "915"
                                
                             
                                Dim Wint As Integer
                                Wint = CInt(W) * 10
                                Dim Dint As Integer
                                Dint = 300
                                If Not IsEmpty(D) Then Dint = CInt(D)
                                If Dint < 70 Then Dint = Dint * 10
                                Doormount = Null
                                If Dint > 0 Then
                                    If Wint = 400 And Dint < 450 Then
                                        Doormount = "+20"
                                    ElseIf Dint = Wint Then
                                        Doormount = "-45"
                                    ElseIf (Dint - Wint) > 96 Then
                                        Doormount = "FGV45"
                                    End If
                                End If
'
'                                If W = 40 Then
'                                    If IsEmpty(D) Or D < 45 Then
'                                        Doormount = "+20"
'                                    Else
'                                        Doormount = "FGV45"
'                                    End If
'                                ElseIf W = 20 Then
'                                    Doormount = "FGV45"
'                                Else
'                                    Doormount = "-45"
'                                End If
                                DoorQty = 1
                            End If
                            
                        
                        Case "ШНГ"
                       
                            If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                            
                             If Not IsEmpty(H) Then
                                If H >= 500 Then
                                    ShelfQty = 2 ' здесь 2 значит +1, т.е. 1 полка
                                End If
                            End If
                            
                            Select Case SVar
                                Case Empty
                                    If Check_ВП(casename) Then
                                        name = name & " вп"
                                        Doormount = "+20"
                                    End If
                                Case "/1"
                                    DoorQty = 1
                                                                   
                                Case "/2"
                                    If Check_ВП(casename) Then
                                        name = name & " вп"
                                        Doormount = "+20"
                                    End If
                                    DoorQty = 2
                                Case Else
                                    name = ""
                                    MsgBox "Неизвестный шкаф", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select

                        Case "ШНП"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 0
                        Case "ШНУ"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                            Doormount = "FGV45"
                        Case "ШНЗ"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                        Case "ШНЗУ"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                        Case "ШНУГ"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                            Doormount = "Гармошка"
                        Case "ШНУР"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 2
                            Doormount = "175"
                        Case "ШНБТ"
                            If ShelfQty >= 2 Then
                                name = name & "915"
                                ShelfQty = Empty
                            End If
                        
                        Case Else
                        
                            name = ""
                            MsgBox "Неизвестный шкаф", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select ' Ш-Н-Name
                    
                Case "Л", "С"
                    Select Case name
                        Case "ШЛ", "ШС", "ШЛМ", "ШСМ", "ШЛЮ"
                            If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                            
                            
                            
                            
                            If InStr(casename, "скос") Then
                                'Dim Wint As Integer
                                Wint = CInt(W) * 10
                                'Dim Dint As Integer
                                Dint = 570
                                If Not IsEmpty(D) Then Dint = CInt(D)
                                If Dint < 70 Then Dint = Dint * 10
                                Doormount = Null
                                If Dint > 0 Then
                                    If Wint = 400 And Dint < 450 Then
                                        Doormount = "+20"
                                    ElseIf Dint = Wint Then
                                        Doormount = "-45"
                                    ElseIf (Dint - Wint) > 96 Then
                                        Doormount = "FGV45"
                                    End If
                                End If
                                DoorQty = 1
                            End If
                            
                            Select Case SVar
                                Case Empty
                                    If Check_ВП(casename) Then
                                        name = name & " вп"
                                        Doormount = "+20"
                                    End If
                                Case "/1Т"
                                    DoorQty = 1
                                    name = name & SVar
                                Case "/1"
                                    DoorQty = 1
                                Case "/2"
                                    DoorQty = 2
                                Case "/К"
                                    'ShelfQty = 0
                                    name = name & SVar
                                    DoorQty = 1
                                Case "/тбм", "/2тбм"
                                    name = name & SVar
                                    DoorQty = 0
                                    Drawermount = ""
                                Case Else
                                    name = ""
                                    MsgBox "Неизвестный шкаф", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                            Select Case LVar
                                Case Empty
                                Case "И"
                                    name = name & " И"
                                    LVar = Empty
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                Case "Т"
                                    'DoorQty = 1
                                    name = name & " Т"
                                    LVar = Empty
                                Case Else
                                    name = ""
                                    MsgBox "Неизвестный шкаф", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                            If InStr(casename, "скос") Then
                                name = name & " скос"
                                
                                Wint = CInt(W) * 10
                                'Dim Dint As Integer
                                Dint = 570
                                If Not IsEmpty(D) Then Dint = CInt(D)
                                If Dint < 70 Then Dint = Dint * 10
                                 Doormount = Null
                                If Dint > 0 Then
                                    If Wint = 400 And Dint < 450 Then
                                        Doormount = "+20"
                                    ElseIf Dint = Wint Then
                                        Doormount = "-45"
                                    ElseIf (Dint - Wint) > 96 Then
                                        Doormount = "FGV45"
                                    End If
                                End If
                                DoorQty = 1
                            End If
                            
                        Case "ШЛП"
                            DoorQty = 0
                            If IsEmpty(D) Then D = 570 '530
                        Case "ШСП"
                            DoorQty = 0
                            If IsEmpty(D) Then D = 530 '480
                        Case "ШЛГП", "ШСГП"
                            If SVar = "/тбб" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/тбм" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/мбб" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/мбм" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/тб" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/арг" Then
                                name = name & SVar
                                Drawermount = "500/78 ШЛГП"
                            ElseIf SVar = "/мб" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf InStr(casepropertyCurrent.p_fullcn, "кв") > 5 Then
                                name = name & "/кв"
                                If IsEmpty(D) Then D = 570 '530
                                Drawermount = CStr(GetDrawerMountKv())
                            ElseIf InStr(casepropertyCurrent.p_fullcn, "шар") > 5 Then
                                name = name & "/напр"
                                If IsEmpty(D) Then D = 570 '530
                                Drawermount = "шарик " & CStr(GetDrawerMount())
                            ElseIf InStr(casepropertyCurrent.p_fullcn, "рол") > 5 Then
                                name = name & "/напр"
                                If IsEmpty(D) Then D = 570 '530
                                Drawermount = "ролик " & CStr(GetDrawerMount())
                            ElseIf (InStr(casepropertyCurrent.p_fullcn, "оргаб") > 5 _
                                Or InStr(casepropertyCurrent.p_fullcn, "org") > 5 _
                                Or InStr(casepropertyCurrent.p_fullcn, "оргоб") > 5) Then
                                name = name & "/орг"
                                Drawermount = ""
                            ElseIf IsEmpty(D) Then
                                    D = 570 '530
                                    Drawermount = GetDrawerMount()
                            End If
                            ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                            DoorQty = 1
                        Case "ШЛУ", "ШСУ"
                            DoorQty = 2
                            Doormount = "175"
                            'Doormount = "FGV180"
                           name = name & SVar
                        Case "ШЛБ", "ШСБ"
                            DoorQty = 0
                        Case "ШЛУГ", "ШСУГ"
                            DoorQty = 1
                            Doormount = "гармошка"
                        Case "ШЛУН", "ШСУН"
                            DoorQty = 1
                            Doormount = "FGV45"
                            'Name = Name & SVar
                        Case "ШЛЗУ", "ШСЗУ"
                            DoorQty = 1
                            name = name & SVar
                        Case "ШЛЗ", "ШСЗ"
                            DoorQty = 2
                        Case "ШЛК", "ШСК"
                        
                            If IsEmpty(D) Then
                                If is18(CaseColor) Then
                                    
                                    D = 570 '530
                                    
                                End If
                            End If
                        
                        
                            Select Case SVar
                                Case "/1", "/1дв"
                                    name = name & "/1"
                                    DoorQty = 1
                                    
                                    Doormount = "110"
                                    
'                                    If InStr(1, CaseName, "cпр", vbTextCompare) > 0 And InStr(1, CaseName, "лев", vbTextCompare) > 0 And InStr(1, CaseName, "отк", vbTextCompare) > 0 Or _
'                                        InStr(1, CaseName, "cлев", vbTextCompare) > 0 And InStr(1, CaseName, "прав", vbTextCompare) > 0 And InStr(1, CaseName, "отк", vbTextCompare) > 0 Then
'                                        DoorMount = "110"
'                                    Else
'                                        DoorMount = "FGV180"
'                                    End If
                                Case "/2"
                                    name = name & "/1"
                                    DoorQty = 2
                                    Doormount = "110"
                                    
                                    
                                Case "/3-1", "/4", "/2-1"
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = GetDrawerMount()
                                    
                                Case "/3-1мб", "/4мб", "/2-1мб", "/2мб", "/2мб-1"
                                        
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = ""
                                    
                                Case "/3-1кв", "/4кв", "/2-1кв", "/2кв", "/2кв-1"
                                        
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = GetDrawerMountKv()
                                    
                                Case "/2-2мб", "/2-2тб"
                                    name = name & SVar
                                    DoorQty = 2
                                    Drawermount = ""
                                    
                                
                                Case "/2-2кв"
                                    name = name & SVar
                                    DoorQty = 2
                                    Drawermount = GetDrawerMountKv()
                                    
                                
                                Case Else
                                    name = ""
                                    MsgBox "Неизвестный шкаф", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                        Case "ШЛШМ"
                            Select Case SVar
                                Case "/тбм", "/2тбм"
                                    name = name & SVar
                                    DoorQty = 0
                                    Drawermount = ""
                                Case Else
                                    name = ""
                                    MsgBox "Неизвестный шкаф", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                        Case "ШЛШ", "ШСШ"
                            
                           
                            
                            If IsEmpty(D) Then
                                If is18(CaseColor) Then
                                    
                                    D = 570 '530
                                    
                                End If
                            End If
                       
                            Select Case SVar
                                Case Empty
                                                                        
                                    If IsEmpty(D) Or D = 530 Then
                                        Drawermount = 50
                                    Else
                                        Drawermount = GetDrawerMount()
                                    End If
                                    
                                    If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                                    
                                    If Check_ВП(casename) Then
                                        name = name & " вп"
                                        'DoorQty = 0
                                        Doormount = "+20"
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "квадро " & GetDrawerMountKv()
                                    End If
                                
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                                                   
                                Case "/тбм", "/2тбм"
                                    name = name & SVar
                                    DoorQty = 0
                                    Drawermount = ""
                                
                                Case "/1мб", "/мб", "/1тб", "/тб", "/1мб-б", "/мб-б", "/1тб-б", "/тб-б"
                                   
                                    
                                    If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = ""
                                    
                                Case "/1кв", "/кв"
                                    
                                    If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMountKv()
                                    
                                
                                    
                                Case "/1мбТ", "/мбТ", "/1тбТ", "/тбТ"
                                        
                                    DoorQty = 0
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    Drawermount = ""
                                
                                Case "/1квТ", "/квТ"                                         ' с нишей
                                        
                                    DoorQty = 0
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMountKv()
                                
                                Case "/1Т"
                                    DoorQty = 0
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMount()
                                    
                                Case "/2"
                                    Drawermount = GetDrawerMount()
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    If Check_ВП(casename) Then
                                        name = name & " вп"
                                        DoorQty = 0
                                        'DoorMount = "+20" НЕТ ДВЕРЕЙ!
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "квадро " & GetDrawerMountKv()
                                    End If
                                    
                                                                       
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                
                                Case "/2-1", "3"
                                        
                                    Drawermount = GetDrawerMount()
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    If Check_ВП(casename) Then
                                        name = name & " вп"
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "квадро " & GetDrawerMountKv()
                                    End If
                                
                                
                                Case "/2-1кв", "/1-2кв", "/3кв"
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    Drawermount = GetDrawerMountKv()
                                    
                                Case "/2-1мб", "/2-1тб", "/1-2мб", "/1-2тб", _
                                         "/3мб", "/3тб"
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    Drawermount = ""
                                    
                                Case "/2мб-1", "/2тб-1"
                                    
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = ""
                                
                                Case "/2кв-1"
                                    
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = GetDrawerMountKv()
                                
                                Case "/2мб", "/2тб"
                                    name = name & SVar
                                    DoorQty = 0
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = ""
                                    
                                Case "/2кв"
                                    name = name & SVar
                                    DoorQty = 0
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMountKv()
                                    
                                Case "/2-2"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = GetDrawerMount()
                                
                                Case "/2-2тб"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = ""
                                    
                                Case "/2-2мб"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = ""
                                    
                                Case "/2-2кв"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = GetDrawerMountKv()
                                
                                Case "/3-1", "/4"
                                    
                                    Drawermount = GetDrawerMount()
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    If Check_ВП(casename) Then
                                        name = name & " вп"
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "квадро " & GetDrawerMountKv()
                                    End If
                                
                                Case "/3-1мб", "/4мб", "/3-1тб", "/4тб"
                                        
                                    name = name & SVar
                                    DoorQty = 0
                                
                                    Drawermount = ""
                                
                                Case "/3-1кв", "/4кв"
                                        
                                    name = name & SVar
                                    DoorQty = 0
                                
                                    Drawermount = GetDrawerMountKv()
                                
                                Case Else
                                    name = ""
                                    MsgBox "Неизвестный шкаф", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            If Not (casepropertyCurrent Is Nothing) Then
                                If casepropertyCurrent.p_DoorCount > 0 Then
                                        DoorQty = casepropertyCurrent.p_DoorCount
                                End If
                                If casepropertyCurrent.p_windowcount > 0 Then
                                        WindowQty = casepropertyCurrent.p_windowcount
                                End If
                            End If
                            
                            
                            
                            
                            Case "ШЛБТ"
                            
                    
                        Case Else
                            name = ""
                            MsgBox "Неизвестный шкаф", vbCritical
                            ActiveCell.Interior.Color = vbRed
                        End Select ' Ш-Л/С-Name
                Case Else
            End Select 'Ш-
            
            If Not IsEmpty(fQty) Then
                If DoorQty <> fQty Then DoorQty = InputBox("Уточните кол-во дверей", "Кол-во дверей шкафа", fQty)
            End If
    
        Case "П"
            If name = "Портал П15" Then
            Else
            
            ShelfQty = Empty
            
            Do
                DoorQty = InputBox("Введите кол-во дверей", "Двери пенала", fQty)
            Loop Until IsNumeric(DoorQty)
            fQty = Empty

            Select Case name
                Case "ПН"
                    
                    Select Case SVar
                        Case Empty
                            
                        Case Else
                            MsgBox "Неизвестный шкаф", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                
                Case "ПНШ"
                    
                    
                    Select Case SVar
                        Case Empty
                            
                            If IsEmpty(D) Then D = 300
                            Drawermount = GetDrawerMount()
                            
                        Case "/2"
                        
                            If IsEmpty(D) Then D = 300
                            Drawermount = GetDrawerMount()
                            
                        Case Else
                            MsgBox "Неизвестный шкаф", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                    
                Case "ПС", "ПЛ", "ПЛХ"
                
                    Select Case SVar
                        Case Empty
                        
                        Case Else
                            MsgBox "Неизвестный шкаф", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                
                
                Case "ПСШ", "ПЛШ"
                
                    If IsEmpty(D) Then D = 570 ' 530
                
                    Select Case SVar
                        Case Empty
                            
                            Drawermount = GetDrawerMount()
                            
                        Case "/2мб", "/2тб", "/3-1мб", "/4мб", "/2-1мб", "/мб", "/тб", "/1-2мб", "/мб-м", "/тб-м", "/1мб-м", "/1тб-м", _
                                 "/2-1тб", "/3мб", "/3тб"
                                
                            name = name & SVar

                            Drawermount = ""

                        Case "/2кв", "/3-1кв", "/4кв", "/2-1кв", "/кв", "/1-2кв", "/3кв"
                                
                            name = name & SVar
                            
                            Drawermount = GetDrawerMountKv()

                        Case "/2-1", "/1-2", "/2", "/3-1", "/4", "/3"
                            name = name & SVar
                            
                            Drawermount = GetDrawerMount()
                        
                        Case Else
                            name = ""
                            MsgBox "Неизвестный шкаф", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                    
                                                  
                
                Case "ПЛД", "ПСД"
                    Select Case SVar
                        Case Empty, "/2мб"
                            name = name & SVar
                            

                            Drawermount = ""
                        
                        Case Else
                            name = ""
                            MsgBox "Неизвестный шкаф", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                    
                                                  
                    
                
                Case Else
                    MsgBox "Неизвестный шкаф", vbCritical
                    ActiveCell.Interior.Color = vbRed
            End Select
            End If
        Case Else
            MsgBox "Это не шкаф", vbCritical
    End Select
    If isWindow And IsEmpty(WindowQty) Then
        WindowQty = DoorQty
    End If

newsystem:

    If Not (casepropertyCurrent Is Nothing) Then
       
            If casepropertyCurrent.p_NishaQty > 0 And casepropertyCurrent.p_DoorCount = 0 Then
                DoorQty = 0
                WindowQty = casepropertyCurrent.p_windowcount
                Doormount = casepropertyCurrent.p_Doormount
            End If
            If casepropertyCurrent.p_newsystem Then
                caseID = 0
                caseID = CInt(GetCaseId(casepropertyCurrent.p_newname))
                If caseID = 0 And CaseElements.Count > 0 Then
                    tempString = ""
                    For caseElementsIterator = 1 To CaseElements.Count
                        tempString = tempString & CaseElements(caseElementsIterator).name & "," & CStr(CaseElements(caseElementsIterator).qty) & ";"
                       
                    Next caseElementsIterator
                     caseID = createCaseId(casepropertyCurrent.p_newname, tempString)
                End If
                
                name = casepropertyCurrent.p_newname
                D = casepropertyCurrent.p_cabDepth
                W = casepropertyCurrent.p_cabWidth
                H = casepropertyCurrent.p_cabHeigth
                DoorQty = casepropertyCurrent.p_DoorCount
                WindowQty = casepropertyCurrent.p_windowcount
                Doormount = casepropertyCurrent.p_Doormount
            End If
    End If
    If name <> "" And caseID = 0 Then
        Init_rsCases
        If rsCases.RecordCount > 0 Then rsCases.MoveFirst
        rsCases.Find "Name='" & name & "'"
        If rsCases.EOF Then
            ActiveCell.Interior.Color = vbRed
            caseID = 0
        Else
            caseID = rsCases!caseID
        End If
    End If
    
    ActiveCell.Offset(, 15).Value = name
    ActiveCell.Offset(, 16).Value = DoorQty
    ActiveCell.Offset(, 17).Value = WindowQty
    ActiveCell.Offset(, 18).Value = Drawermount
    ActiveCell.Offset(, 19).Value = Doormount
    ActiveCell.Offset(, 20).Value = HandleExtra
    ActiveCell.Offset(, 21).Value = BF
    ActiveCell.Offset(, 22).Value = ShelfQty
    If casepropertyCurrent.p_dspbottom > 0 Then ActiveCell.Offset(, 24).Value = "!Дно шуфляд ДСП!"
    
    If IsEmpty(D) = False Then
    caseglub = D
    ActiveCell.Offset(, 31).Value = D
    Else
    caseglub = 570
    ActiveCell.Offset(, 31).Value = 570
    End If
    'ActiveCell.Offset(, 23).Value = ЗАНЯТА!!
    
    HandleExtra = Empty
    casename = name
    caseHeight = H
    Exit Sub
err_ParseCase:
    MsgBox Error, vbCritical
End Sub

Private Function Check_ВП(ByVal casename As String) As Boolean
    Dim vp As Integer, dvp As Integer
    
    vp = InStr(1, casename, "ВП", vbBinaryCompare)
    dvp = InStr(1, casename, "ДВП", vbBinaryCompare)
    
    If vp > 0 And (vp - 1 <> dvp Or dvp = 0) Then
        Check_ВП = True
    Else
        Check_ВП = False
    End If

End Function


' Leg - только для гнутой системы!!!!!! иначе обрабатывать аналогично DefHandle!!!!
Public Function ParseShelving(ByVal name As String, _
                            ByRef caseID As Integer, _
                            ByRef Handle, _
                            ByRef Leg, _
                            ByRef Drawermount, _
                            ByRef Doormount, _
                            ByRef windowcount, _
                            ByRef bWithFittKit As Boolean, _
                            Optional ByRef CaseColor, _
                            Optional ByRef face) As Boolean
                            
On Error GoTo err_ParseShelving

    windowcount = 1 ' по умолчанию для всех стенок есть, кроме систем (см. ниже)
    
    

    Init_rsCases
    If rsCases.RecordCount > 0 Then rsCases.MoveFirst
    rsCases.Find "Name='" & name & "'"
    If rsCases.EOF Then
        ActiveCell.Interior.Color = vbRed
        caseID = 0
        
        ParseShelving = False
    Else
        caseID = rsCases!caseID
        
        name = rsCases!name
        Drawermount = rsCases!DrawerMountDefault
        Doormount = rsCases!DoorMountDefault
        bWithFittKit = rsCases!bWithFittKit
        
        If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = rsCases!HandleDefault
        If IsEmpty(Leg) Then Leg = rsCases!LegDefault
        
        ParseShelving = True
    End If
    
    ActiveCell.Offset(, 15).Value = name
    ActiveCell.Offset(, 18).Value = Drawermount
    ActiveCell.Offset(, 19).Value = Doormount
    ActiveCell.Offset(, 20).Value = bWithFittKit
        
    Exit Function
    
    
    Select Case Left(name, 1)
        
        Case "А" 'АЛЕКСАНДРА
        
            bWithFittKit = True
        
            If IsEmpty(Handle) Then Handle = "А025"
            Doormount = "110"
            
            Select Case name
                Case "А2"
                    Drawermount = "45"
                Case "А1", "А3"
                    Doormount = "равносторонний"
                Case "АЛШВ1", "АЛТВ1"
                    Drawermount = ""
                    Doormount = ""
                    Handle = ""
                    bWithFittKit = False
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "Неизвестная секция АЛЕКСАНДРА: " & name, vbCritical
            End Select
        
        
        Case "Д" 'ДАНИИЛ
        

            bWithFittKit = True

            If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
            Doormount = "110"
            
            Select Case name
                Case "Д1", "Д3"
                    Drawermount = "35"
                Case "Д2", "Д4", "Д5"
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "Неизвестная секция ДАНИИЛ: " & name, vbCritical
            End Select
            
            
        Case "М" 'МИЛЕНА
            
            Select Case name
                Case "МПВ", "МПГ", "МПБВ", "МНКВ"
                Drawermount = 40
                
             
                Case "МЗ", "ММЦ", "МКВ", "Медиацентр"
                
                
                
                Case Else
                
                    If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "8303"
                
                    Doormount = "110"
                
                    Select Case name
                    Case "М1", "М6", "М5", "М1верх", "М6верх", "М5верх", "М1низ", "М6низ", "М5низ"
                        Drawermount = "40"
                    Case "М2", "М3", "М4", "М2верх", "М3верх", "М4верх", "М2низ", "М3низ", "М4низ"
                    Case Else
                        ParseShelving = False
                        MsgBox "Неизвестная секция МИЛЕНА: " & name, vbCritical
                    End Select
            End Select
             
            
        Case "Н" 'НАТАЛИ
            
            bWithFittKit = True
        
            If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "8303"
            Doormount = "110"
            
            Select Case name
                Case "Н1", "Н4"
                Case "Н2"
                    Drawermount = 45
                Case "Н3"
                    Drawermount = 35
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "Неизвестная секция НАТАЛИ: " & name, vbCritical
            End Select
        Case "Ф" 'ФАРАОН
            
            Doormount = "110"
            
            If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
                        
            Select Case name
                Case "Ф1", "Ф2", "Ф4", "ФК1", "ФК2", "ФК4"
                Case "Ф3", "ФК3", "Ф3верх", "Ф3низ"
                    Drawermount = 40
                Case Else
                    ParseShelving = False
                    MsgBox "Неизвестная секция ФАРАОН: " & name, vbCritical
            End Select
        Case "К" 'КОРВАЛЬД + КРОВАТЬ
            
            Doormount = "110"
            
            If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
            
            Select Case name
                Case "КРОВАТЬ"
                Case "К1", "К2верх", "К3", "К3верх", "К3низ", "К4низ", "К5", "К5верх", "К5низ"
                Case "К1 (ПВХ)"
                    Leg = "465"
                Case "К2", "К2низ"
                    Drawermount = 40
                Case "К4", "К4верх"
                    Drawermount = 25
                
                Case "КСУ", "КП"
                Handle = Empty
                Leg = Empty
                Drawermount = Empty
                Case "КТСУ", "КТСР"
                Handle = "неликв"
                Drawermount = 40
                Case "КСР"
                Handle = "неликв"
                Drawermount = 40
                Leg = "2706"
                Case Else
                    ParseShelving = False
                    MsgBox "Неизвестная секция КОРВАЛЬД: " & name, vbCritical
            End Select
        Case "О" ' ОЛИМП/ОСКАР
            
            Doormount = "110"
            
            Select Case name
                Case "ОЛП1", "ОЛП3", "ОЛП2", "ОЛП4"
                
                    bWithFittKit = True
                
                    Drawermount = 45
                    If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
               
                Case "ОСК1", "ОСК2", "ОСК3", "ОСК4"
                    
                    If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
                
                Case Else
                    ParseShelving = False
                    MsgBox "Неизвестная секция ОЛИМП/ОСКАР: " & name, vbCritical
            End Select

        Case "В" 'ВИКТОРИЯ + Волна1000 + гнутая светлана + гнутая система
            
            Doormount = "110"
            
            
            Select Case name
                Case "ВЛШВ1", "ВЛШК1", "ВЛШТВ2", "ВЛТСТ2", "ВЛТСТ1", "ВЛТП1", "ВЛШТГ4", "ВЛШТГ5", "ВЛШТГ7", "ВЛШТГ6комби", "ВЛШТГ6полки", "ВЛШТГ6штанга", "ВЛШПР1", "ВЛШПР2"
                Doormount = ""
                Case "ВЛШБ1", "ВЛШБ2", "ВЛШБ3", "ВЛШБК1", "ВЛШБК2", "ВЛШБК3", "ВЛШБК4", "ВЛШКМ1", "ВЛШСП", "ВЛШСПМ", "ВЛШТ4"
                Doormount = ""
                Case "ВНШВ", "ВНТВ1"
                    Drawermount = "40"
                    Doormount = ""
                Case "ВЛТВ1", "ВЛШВ2", "ВЛШК2", "ВЛШТВ1", "ВЛШТГ2", "ВЛШГ1", "ВЛШГ2"
                    Drawermount = "35"
                    Doormount = "равносторонний"
                    Handle = "1Шур. Фарфор"
                    Leg = "черная 100"
                Case "ВЛШВ3"
                    Leg = "черная 100"
                    Handle = "1Шур. Фарфор"
                    Doormount = "равносторонний"
                Case "ВЛШТГ1"
                    Drawermount = "35"
                    Doormount = ""
                    Handle = "1Шур. Фарфор"
                    Leg = "черная 100"
                Case "ВНШТГ1"
                    Drawermount = "40"
                Case "ВЛТВ2"
                    Doormount = ""
                    Leg = "черная 100"
                Case "ВЛШУ1"
                    Drawermount = ""
                    Doormount = "равносторонний"
                    Handle = "1Шур. Фарфор"
                    Leg = "черная 100"
                Case "ВПСК3", "ВГСУ2", "ВПСН3", "ВГСШ11", "ВГСШ13", "ВГСШ12", _
                    "ВПСК2", "ВГСК6", "ВГСК7", "ВПСШ8", "ВПСШ10", "ВПСШ6", "ВПСШ3", "ВГСШ7", _
                    "ВГСН4", "ВТВ", "ВТП", "ВКД2", "ВСТ1", _
                    "ВПСД2", "ВПСД5В", "ВПСК4", "ВПСТВ11", "СНТВ10", "ВКД3"
                
                    Leg = "465" ' только для этого случая, иначе обрабатывать аналогично DefHandle!!!
                    
                    Handle = "Г06"
                    
                
                    Select Case name
                        Case "ВГСУ2", "ВПСК3", _
                            "ВПСД2", "ВПСК4"
                            Drawermount = "шарик 40"
                            Doormount = "+20"
                        
                        Case "ВПСТВ11"
                            Doormount = "+20"
                            Drawermount = "шарик 45"
                        
                        Case "ВПСШ10", "ВПСШ6"
                            Doormount = "+20"
                            Drawermount = "шарик 50"
                        
                        Case "ВГСШ12", "ВГСШ13", "ВГСК6", "ВГСК7"
                            Doormount = "-30"
                            Drawermount = "шарик 40"
                        
                        Case "ВПСК2", "ВТВ"
                            Doormount = "+20"
                            Drawermount = "шарик 50"
                        
                        Case "ВПСШ8", "ВПСШ3"
                            Doormount = "+20"
                            Drawermount = "шарик 40"
                            
                        Case "ВГСШ7", "ВГСШ11"
                            Doormount = "-30"
                            Drawermount = "шарик 50"
                                                
                        Case "ВКД2"
                            Doormount = "+20"
                            Drawermount = "шарик 50"
                            
                        Case "ВТП"
                            Doormount = "-30"
                            Drawermount = "шарик 40"
                    
                        Case "ВСТ1"
                            Doormount = "+20"
                            Drawermount = "шарик 35"
                    
                        Case "ВПСН3", "ВГСН4", _
                            "ВПСД5В", "ВКД3"
                    
                        Case Else
                            ParseShelving = False
                            MsgBox "Неизвестная секция ВОЛНА: " & name, vbCritical
                    End Select
                    
                
                
                Case "ВГС1", "ВПС2", "ВГС3", "ВПС4", "ВПСШ1", "ВПСШ2"
                    If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
                                        
                    Select Case name
                        
                        Case "ВГС1", "ВГС3"
                            Doormount = "-30"
                        
                        Case "ВПС2"
                            Drawermount = "шарик 45"
                    
                        Case "ВПС4"
                            'DoorMount = "+20"
                            
                        Case "ВПСШ2", "ВКД2", "ВПСШ1"
                       
                            Doormount = "+20"
                            Drawermount = "шарик 50"
                    
                        Case Else
                            ParseShelving = False
                            MsgBox "Неизвестная секция ГНУТАЯ СВЕТЛАНА: " & name, vbCritical
                    End Select
                    
                    
            Case Else
                    
'                    If IsEmpty(Handle) Then
'                        If InStr(1, Face, "ольха", vbTextCompare) > 0 Or InStr(1, Face, "груша", vbTextCompare) > 0 Then
'                            Handle = "1026"
'                        ElseIf InStr(1, Face, "орех", vbTextCompare) > 0 Or InStr(1, Face, "рустик", vbTextCompare) > 0 Then
'                            Handle = "1035"
'                        ElseIf InStr(1, Face, "махонь", vbTextCompare) > 0 Then
'                            Handle = "1007"
'                        End If
'                    End If
                    If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
                    
                    Select Case name
                        Case "В1", "В2верх", "В3", "В3верх", "В3низ", "В4", "В4низ", "В4верх", "В5", "В5низ", "В5верх", "В7А", "В7В"
                        Case "В8", "В8верх", "В8низ"
                            Drawermount = "40"
                        Case "В2", "В2низ", "В6", "В6верх", "В6низ", "В7", "В7Б"
                            Drawermount = "40"
                        Case Else
                            ParseShelving = False
                            MsgBox "Неизвестная секция ВИКТОРИЯ: " & name, vbCritical
                    End Select
            End Select

        Case "У" 'Угол
            If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
            
            Select Case name
                Case "Угол"
                    Doormount = "FGV45"
                Case Else
                    ParseShelving = False
                    MsgBox "Неизвестная секция Угол: " & name, vbCritical
            End Select
        Case "Ж"
            bWithFittKit = False
            
            Select Case name
                Case "ЖС1"
                Case Else
                    MsgBox "Неизвестная секция СИСТЕМА XXI век: " & name, vbCritical
            End Select
                    
        Case "С" 'СИСТЕМА XXI век + СВЕТЛАНА!!!
            
            bWithFittKit = True
            
            Doormount = "110"
            If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
            
'            Select Case Name
'                Case "С1", "С2", "С3", "С4"
'
'                    bWithFittKit = True
'
''                    If IsEmpty(Handle) Then Handle = "2903"
'                    If IsEmpty(Handle) Then If InStr(1, Face, "модена", vbTextCompare) = 0 Then Handle = "0603"
'
'                Case Else
'
''                    If IsEmpty(Handle) Then
''                        If InStr(1, Face, "ольха", vbTextCompare) > 0 Or _
''                            InStr(1, Face, "груша", vbTextCompare) > 0 Or _
''                            InStr(1, Face, "клен", vbTextCompare) > 0 Then
''                            Handle = "3826"
''                        ElseIf InStr(1, Face, "орех", vbTextCompare) > 0 Then
''                            Handle = "3835"
''                        End If
''                    End If
'                    If IsEmpty(Handle) Then If InStr(1, Face, "модена", vbTextCompare) = 0 Then Handle = "0603"
'
'            End Select
            
            Select Case name
                Case "С1", "С3", "С4", "С5"

'                    If IsEmpty(Handle) Then Handle = "2903"
                    'If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
                Case "С2"

'                    If IsEmpty(Handle) Then Handle = "2903"
                    'If IsEmpty(Handle) Then If InStr(1, face, "модена", vbTextCompare) = 0 Then Handle = "0603"
                    Drawermount = 45
                    
                Case "СШ5"
                    windowcount = Empty
                    
                Case "СП5", "СН1", "СН2", _
                        "СТВ", "СП4фр", "СП4ср", "СП2", "СП1", "СШ2", "СШ1", "СВ", "СБ", "СНТВ"
                    windowcount = Empty
                
                Case "СК4", "СК5", "СС2", "СП3", "СПр", "СК3", "СК2", "СШ4", "СШ3"
                    Drawermount = 40
                    windowcount = Empty
                    
                Case "СС3"
                    Drawermount = 45
                    windowcount = Empty
                
                Case "СС"
                    Drawermount = 45
                    windowcount = Empty
                
                Case "СК1"
                    Drawermount = 50
                    windowcount = Empty
                
                Case "СУ"
                    Doormount = "FGV45"
                    windowcount = Empty
                    
                Case "СУ (ПВХ)"
                    bWithFittKit = False
                
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 40"
                    windowcount = Empty
                    Doormount = "FGV45"
                                    
                Case "СКУ1 (ПВХ)"
                    bWithFittKit = False
                
                    Leg = "465"
                    Handle = "Г06"
                    windowcount = Empty
                    Doormount = "FGV45"
                                    
                Case "СП2 (ПВХ)", "СШМ1 (ПВХ)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 40"
                    windowcount = 1
                    Doormount = "FGV под аморт."
                                        
                Case "СНТВ5 (ПВХ)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    windowcount = Empty
                    Doormount = "FGV под аморт."

                Case "СТВ6 (ПВХ)", "СК6 (ПВХ)"
                                        
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 45"
                    windowcount = Empty
                    Doormount = "FGV под аморт."
                                    
                Case "СШ1 (ПВХ)", "СШ2 (ПВХ)", "СШ3 (ПВХ)", "СШ4 (ПВХ)", "СП1 (ПВХ)", "СП3 (ПВХ)", "СК2 (ПВХ)", _
                     "СПр (ПВХ)", "СТВ (ПВХ)", "СБ (ПВХ)", "СВ (ПВХ)", "СК5 (ПВХ)", "СП5 (ПВХ)", "СШ7(ПВХ)", "СНТВ8 (ПВХ)"
                
                    bWithFittKit = False
                    
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 40"
                    windowcount = Empty
                    Doormount = "FGV под аморт."
                    
                Case "СП4фр (ПВХ)", "СК3 ПВХ)", "СС2 ПВХ)", "СН1 (ПВХ)", _
                    "СК4 (ПВХ)", "СП1 (ПВХ)", "СН2 (ПВХ)", "СШ5 (ПВХ)", "СНТВ (ПВХ)", _
                    "СТВ2 (ПВХ)", "СНТВ2 (ПВХ)", "СНТВ4 (ПВХ)", "СНТВ10", "СТВ3 (ПВХ)"
                
                    bWithFittKit = False
                    
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 40"
                    windowcount = Empty
                    
                Case "СК1 (ПВХ)"
                    bWithFittKit = False
                    
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 50"
                    windowcount = Empty
                    
                Case "СШ8 (ПВХ)", "СА2 (ПВХ)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 50"
                    windowcount = 1
                
                Case "СНТВ6 (ПВХ)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 40"
                    windowcount = 1
                
                Case "СС (ПВХ)", "СС3 (ПВХ)", "Зеркало", _
                     "СТВ4 (ПВХ)", "СТВ5 (ПВХ)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 45"
                    windowcount = Empty
                    
                Case "СШ6 (ПВХ)", "СА1 (ПВХ)"
                    
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    windowcount = Empty
                    
                Case "СТВ7 (ПВХ)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 35"
                    windowcount = Empty
                    
                Case "СТВ8 (ПВХ)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "Г06"
                    Drawermount = "шарик 50"
                    Doormount = "+20"
                    windowcount = Empty
                    
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "Неизвестная секция СВЕТЛАНА/СИСТЕМА XXI век: " & name, vbCritical
            End Select
            
        Case "Т"
        
            bWithFittKit = False
            Doormount = "110"
            Leg = "465"
            Handle = "Г06"
            
            Select Case name
                Case "Т1"
                    Drawermount = "шарик 45"
                    
                Case "Т2", "Т3", "Т5", "Т6", "Т7", "Т9", "Т10"
                
                Case "Т4"
                    Drawermount = "шарик 40"
                    
                Case "ТУ"
                    Doormount = "FGV45"
                    
                Case Else
                    ParseShelving = False
                    MsgBox "Неизвестная секция ТАТЬЯНА: " & name, vbCritical
                    
            End Select
        
        Case "П"
        
            bWithFittKit = False
            Handle = "1006(160)"
            
            
            Select Case name
                Case "ПН10 под БПД4 (б.1442) гл.372мм", _
                     "ПН20 под БПД5 (б.2200)карго гл.620мм", _
                     "ПН20 под БПД5 (б.2200)карго гл.570мм", _
                     "ПН20 под БПД5 (б.2200) гл.570мм"
                    
            
            
            
                Case "ПС45(2200)стяж 14держ глуб.41", _
                    "ПС45(2200)стяж 2экспоз глуб.41"
                    Drawermount = "шарик 35"
                Case "ПСП1", "ПСП2", "ПСП3"
            
                Case "ПСТ7", "ПСШ2", "ПСТ1", "ПСТ3", "ПСТ6", "ПСТ5", "ПСТ4"
                
                    Leg = "2706"
                    Drawermount = "шарик 50" ' с доводч."
                    
                Case "ПСТ2"
                    
                    Leg = "2706"
                    Doormount = "софт с обр. пружиной"
                    Drawermount = "шарик 50" ' с доводч."
                                
                
                Case "ПСН1", "ПСН5", "ПСН2", "ПСН4", "ПСН3", "ПСН7", "ПСН6"
                    
                    Doormount = "софт"
                
                Case "ПСН8"
                    
                    Doormount = "софт для стекла"
                
                Case "ПСЖ1"
                    
                    Leg = "2706"
                
                Case "ПСК3"
                    
                    Leg = "2706"
                    Drawermount = "шарик 40" ' с доводч."
                    
                Case "ПСШ4"
                    
                    Leg = "2706"
                    Doormount = "софт"
                
                Case "ПСШ3", "ПСК5"
                    
                    Leg = "2706"
                    Drawermount = "шарик 40" ' с доводч."
                    Doormount = "софт"
                    
                Case "ПСК4"
                    
                    Leg = "2706"
                    Drawermount = "шарик 40" ' с доводч."
                    
                Case "ПСШ1", "ПСК2", "ПСК1"
                    
                    Leg = "2706"
                    Doormount = "софт с обр. пружиной"
                    
                    
                '********************************
                '********************************
                    
                'ПСК
                Case "ПСК(578)/4(203-4)", "ПСК(877)/4(176-4)", "ПСК(877)/4(223-4)", "ПСК(978)/4(176-4)", "ПСК(978)/4(223-4)", _
                        "ПСК(578)/4(176-3,283)", "ПСК(1277)/4(296-4)полки стекло"
                
                    Drawermount = "шарик 40"
                    Leg = "2706"
                    
                Case "ПСК(1277)/5(713,176-4)", "ПСК(1277)/5(901,223-4)", _
                        "ПСК(1876)/6(713,176-4,713)", "ПСК(1876)/6(901,223-4,901)", _
                        "ПСК(578)/2(640,176)", "ПСК(1277)/5(484-2,223,484-2)"
                
                    Drawermount = "шарик 40"
                    Leg = "2706"
                    Doormount = "софт"
                    
                Case "ПСК(578)/1(818)", _
                        "ПСК(1277)/3(596-2,1196)", "ПСК(1277)/4(596-4)", "ПСК(1277)/6(396-6)"
                
                    Leg = "2706"
                    Doormount = "софт"
                    
                ' ПСТ
                Case "ПСТ(1277)/2(223-2)", "ПСТ(1277)/2(396-2)", "ПСТ(678)/1(223)", "ПСТ(678)/1(396)", _
                    "ПСТ(678)/2(296-2)", "ПСТ(1876)/3(223-3)", "ПСТ(1876)/3(396,223,396)", _
                    "ПСТ(1876)/3(396-3)", "ПСТ(2476)/3(396,223,396)", "ПСТ(1876)/4(396,197-2,396)", _
                    "ПСТ(2476)/4(396,197-2,396)"
                
                    Leg = "2706"
                    Drawermount = "шарик 50"
                    
                Case "ПСТ(1876)/2(223-2)полка", "ПСТ(1876)/2(396-2)полка"
                
                    Leg = "2706"
                    Drawermount = "шарик 50"
                    
                Case "ПСТ(1876)/2(596-2)полка", "ПСТ(678)/1(596)", "ПСТ(1277)/2(596-2)", _
                    "ПСТ(1876)/3(596-3)"
                
                    Leg = "2706"
                    Doormount = "софт"
                    
                Case "ПСТ(1876)/3(596,296,596)", "ПСТ(2476)/3(596,296,596)", "ПСТ(1876)/4(596,296-2,596)", _
                    "ПСТ(2476)/4(596,296-2,596)"
                    
                    Drawermount = "шарик 50"
                    Leg = "2706"
                    Doormount = "софт"
                    
                
                ' ПСН
                Case "ПСН(478)/1(596)", "ПСН(478)/1(896)", "ПСН(478)/2(596-2)", "ПСН(678)/2(596-2)", _
                    "ПСН(478)/1(1196)", "ПСН(678)/1(396)", "ПСН(678)/1(396)бар", "ПСН(678)/1(596)", _
                    "ПСН(978)/1(396)", "ПСН(978)/1(396)бар", "ПСН(1277)/1(396)", "ПСН(1876)/1(396)полки", _
                    "ПСН(1876)/1(596)полки", "ПСН(2176)/1(396)полки", "ПСН(678)/2(396-2)", "ПСН(877)/2(1196-2)", _
                    "ПСН(1277)/2(396-2)", "ПСН(1277)/2(596-2)", "ПСН(1876)/2(396-2)полка", "ПСН(1876)/2(596-2)полка", _
                    "ПСН(1876)/3(396-3)", "ПСН(1876)/3(596-3)", "ПСН(1277)/4(396-4)", "ПСН(1277)/4(596-4)", _
                    "ПСН(376)/1(996)витрина"
                    
                    Leg = "2706"
                    Doormount = "софт"
                    
                'ПСНТ
                Case "ПСНТ(678-1200)/полки", "ПСНТ(678-1400)/полки", "ПСНТ(678-1573)/полки"
                
                    Leg = "2706"
                    
                    
                'ПСЛТ
                Case "ПСЛТ", "ПСЛТлев", "ПСЛТправ", "ПСЛК1прав", "ПСЛК2прав", "ПСЛК1лев", "ПСЛК2лев"
                
                    Leg = "2706"
                    Drawermount = "квадро 50"
                    
                    'ПСЛШ
               
                Case "ПСЛШ1прав", "ПСЛШ1лев", "ПСЛШ2прав", "ПСЛШ2лев"
                
                    Leg = "2706"
                    
     
                    
                'ПСШ
                Case "ПСШ(678)/1(1796)", "ПСШ(678)/3(596-3)", "ПСШ(1475)/4(1596-4)", "ПСШ(877)/6(396-2,1000-2,396-2)", _
                        "ПСШ(678)/2(596-2)полка", "ПСШ(877)/2(1796-2)", "ПСШ(1475)/4(1870-4)", "ПСШ(1475)/4(2074-4)", _
                        "ПСШ(1076)/6(496-2,1074-2,496-2)"
                
                    Leg = "2706"
                    Doormount = "софт"
                    
                Case "ПСШ(877)/5(897-2,1346,223-2)", "ПСШ(678)/3(748,296,748)", "ПСШ(678)/4(596-2,296-2)"
                
                    Drawermount = "шарик 40"
                    Leg = "2706"
                    Doormount = "софт"
                    
                Case "ПСШ(678)/2(296-2)полки"
                    
                    Leg = "2706"
                    Drawermount = "шарик 40"
                    
                Case "ПСШ(877)/4(1400-2,196-2)"
                
                    Leg = "2706"
                    Drawermount = "шарик 50"
                    Doormount = "софт"
                    
                Case "ПСШ(678)/полки"
                    
                    Leg = "2706"
                    
                Case "ПН20 под БПД5 (б.2200) гл.570мм", "ПН20 под БПД5 (б.2200)карго гл.570мм", "ПН20 под БПД5 (б.2200)карго гл.620мм", _
                "ПННД20", "ПН10 под БПД4 (б.1442) гл.372мм", "ПННД20 под БПД3"
                
                Case "ПЛ10 под БПД4 (б.2040) гл.570мм", "ПЛ20 под БПД5 (б.1184) гл.375мм", "ПЛНД20 под БПД5", "ПЛНД20карго под БПД5", _
                    "ПЛНД20", "ПЛНД20карго", "ПЛ20 под БПД5 (б.2200) гл.570мм"

                    Leg = "чёрная 100"
                    
            
                    
                Case Else
                        ParseShelving = False
                        MsgBox "Неизвестная секция ПУСТОТКА " & name, vbCritical
            End Select
        Case "Ш"
            bWithFittKit = False
            
            Drawermount = "квадро"
            Handle = "1006(160)"
            Select Case name
                Case "ШТВ-1", "ШТВ-2"
                Drawermount = "35"
                Doormount = "равносторонний"
                Handle = "Г06"
                
            
                Case "ШВ-1"
                Drawermount = "квадро 35"
                Doormount = "равносторонний"
                Handle = "Г06"
                Case "ШВ-2"
                Drawermount = "квадро 35"
                Doormount = "равносторонний"
                Handle = "Г06"
                
                Case "ШН10 под БПД4 (б.820) гл.483мм", "ШН10 под БПД4 (б.1442) гл.492мм", "ШН10 под БПД4 (б.1451) гл.570мм", _
                    "ШН10 под БПД4 (б.718) гл.492мм", "ШН10 под БПД4 (б.820) гл.483мм", "ШНПД10 под БПД4", "ШНПД10"


            End Select
                
        
            
        Case "З"
            Select Case name
                Case "Зеркало"
                    windowcount = 1
                    
                Case Else
                    ParseShelving = False
                    MsgBox "СИСТЕМА XXI век: " & name, vbCritical
            End Select
        
        Case Else
            ParseShelving = False
            MsgBox "Неизвестная секция", vbCritical
    End Select
    
    ActiveCell.Offset(, 15).Value = name
    ActiveCell.Offset(, 18).Value = Drawermount
    ActiveCell.Offset(, 19).Value = Doormount
    ActiveCell.Offset(, 20).Value = bWithFittKit
    
    'Name = InputBox("Неизвестная секция стенки", "Стенка", Name)
    Exit Function
err_ParseShelving:
    ParseShelving = False
    MsgBox Error, vbCritical, "Разбор секции"
End Function

Public Function FindFittings(ByVal OrderId As Long, _
                            ByVal row As Integer, _
                            ByVal Item As String, _
                            ByRef Fitting As String, _
                            Optional ByRef ShiftK As Integer = 1, _
                            Optional ByRef length, _
                            Optional ByRef FittOpt, _
                            Optional ByVal face, _
                            Optional ByRef HandleScrew, _
                            Optional ByRef changeCaseZaves, _
                            Optional ByVal SearchStringsCollection As Collection _
                            ) As Variant
                            

    If ( _
            ( _
                (InStr(ShiftK, Fitting, "на весь заказ петли", vbTextCompare) > 0 Or _
                InStr(ShiftK, Fitting, "на заказ завес", vbTextCompare) > 0 Or _
                InStr(ShiftK, Fitting, "на заказ петли", vbTextCompare) > 0) And _
                (InStr(ShiftK, Fitting, "Sens", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "Сенс", vbTextCompare) > 0 Or _
                InStr(ShiftK, Fitting, "Blumot", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "FGV", vbTextCompare) > 0) Or _
                (InStr(ShiftK, Fitting, "Петли ", vbTextCompare) > 0 And InStr(ShiftK, Fitting, " на весь заказ", vbTextCompare) > 0) Or _
                (InStr(ShiftK, Fitting, "Завес", vbTextCompare) > 0 And InStr(ShiftK, Fitting, "на весь заказ", vbTextCompare) > 0) _
             ) _
             And InStr(ShiftK, Fitting, "шт", vbTextCompare) = 0 _
            ) = False _
            Then
    Dim k As Integer, bExtra As Boolean, PlankColor
    
    If Item = cPlank Then PlankColor = FittOpt
    k = InStr(ShiftK, Fitting, Item, vbTextCompare)
    
    If k Then
        'посмотрим, дополнительно или на заказ
            
                If InStr(ShiftK, Fitting, "на зак", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "на  зак", vbTextCompare) > 0 Then bExtra = False Else bExtra = True
    
        
        'Dim FittOpt
        Dim en As Integer, p As Integer, ts As String, bPlus As Boolean
        
        Do
            ' начало строки есть, ищем ее конец
            ' признаком конца строки считаем ближайшую к началу (k) точку или плюс, иначе ее конец, если таковых нет
            en = InStr(k + Len(Item), Fitting, "+")
            p = InStr(k + Len(Item), Fitting, ".")
            If p = 0 Then p = Len(Fitting)
            If en = 0 Or (p > 0 And en > p) Then
                en = p
                bPlus = False
            Else
                bPlus = True
            End If
            
            
            
            ' попробуем выделить Option
            Dim x As Integer
            If InStr(k, Fitting, Item, vbTextCompare) Then
                ActiveCell.Characters(k, Len(Item)).Font.Color = vbBlue
            
                p = InStr(k + Len(Item), Fitting, " ")
                x = InStr(k + Len(Item), Fitting, "-") ' !!!здесь могут быть проблемы, если в наименовании типа использауется "-"
                If p = 0 Or (x > 0 And p > x) Then p = x
                
                k = p  ' увеличиваем на искомое слово
            End If
           
            If en > k Then
                ts = Mid(Fitting, k + 1, en - k) ' "кусочек" строки, который сейчас будем рассматривать
            Else
                ts = ""
            End If
            
            ' теперь работаем с ts
            
            ' ищем кол-во
            Dim qty, QtyPattern As String
            QtyPattern = "шт"
            p = InStr(1, ts, QtyPattern, vbTextCompare)
            If p = 0 Then
                QtyPattern = "к-т"
                p = InStr(1, ts, QtyPattern, vbTextCompare)
            End If
            If p = 0 Then
                QtyPattern = "комп"
                p = InStr(1, ts, QtyPattern, vbTextCompare)
            End If
            If p > 0 Then
                Dim tqty As String
                If p > 4 Then
                    x = 4
                ElseIf p < 4 Then
                    x = p - 1
                Else
                    x = 3
                End If
                
                tqty = Mid(ts, p - x, x)
                DelTextLeft tqty
                DelTextRight tqty, x
                If tqty <> "" And IsNumeric(tqty) And InStr(1, tqty, " ", vbTextCompare) = 0 Then
                    qty = CInt(tqty)
                    
                    ' выделим то, что нашли
                    ActiveCell.Characters(k + p, Len(QtyPattern)).Font.Bold = True
                    ActiveCell.Characters(k + p - x - Len(tqty), Len(tqty)).Font.Color = vbRed
                    ' вырежем ненужную информацию
                    ts = Left(ts, p - x - 1 - Len(tqty)) '& Mid(ts, p + Len(QtyPattern)) '!!!!!!!!!!!!!!!!
                Else 'If Not IsNumeric(tQty) Then
                    Do
                        Do
                            MsgBox "ОШИБКА! Кол-во=" & tqty, vbCritical
                            tqty = InputBox(ts, "Введите кол-во", tqty)
                        Loop While Not IsNumeric(tqty)
                        qty = CDec(tqty)
                    Loop Until qty >= 0
                End If
            End If
            
            ' ищем указание в см (не в метрах!), если есть
            If IsMissing(length) Then length = Empty
            p = InStr(1, ts, "см", vbTextCompare)
            If p > 0 Then
                Dim tLen As String
                If p < 5 Then x = p - 1 Else x = 4
                tLen = Mid(ts, p - x, x)
                DelTextLeft tLen
                DelTextRight tLen, x
                If tLen <> "" And IsNumeric(tLen) Then
                    length = CInt(tLen)
                    ' выделим то, что нашли
                    ActiveCell.Characters(k + p, 2).Font.Bold = True
                    ActiveCell.Characters(k + p - x - Len(tLen), Len(tLen)).Font.Color = vbRed
                    ' вырежем ненужную информацию
                    ts = Left(ts, p - x - 1 - Len(tLen)) '& Mid(ts, p + 2) '!!!!!!!!!!!!!!!!
                ElseIf Not IsNumeric(tLen) Then
                    If InStr(1, tLen, "х", vbTextCompare) = 0 Then
                        MsgBox "ОШИБКА! Длина=" & tLen, vbCritical
                        ActiveCell.Interior.Color = vbRed
                    End If
                End If
            End If
            
            If Right(ts, 1) = " " Then ts = RTrim(Left(ts, Len(ts) - 1))
            If Right(ts, 1) = "." Then ts = RTrim(Left(ts, Len(ts) - 1))
            If Right(ts, 1) = "+" Then ts = RTrim(Left(ts, Len(ts) - 1))
            If Right(ts, 1) = "-" Then ts = RTrim(Left(ts, Len(ts) - 1))
            
            
            ts = Trim(ts)
            
            If IsEmpty(FittOpt) Then
                If Len(ts) > 0 Then
                    FittOpt = ts
                    ActiveCell.Characters(k + 1, Len(ts)).Font.ColorIndex = 10
                End If
            ElseIf Len(ts) > 2 Then
                'MsgBox "Проверьте!!!новый тип"
                FittOpt = ts
                ActiveCell.Characters(k + 1, Len(ts)).Font.ColorIndex = 10
            End If
            
            Dim FittingName As String
            FittingName = Item
            
            
            If Item = cPlank Then ' для планок
                Dim pp As Integer
                
                pp = InStr(1, FittOpt, "уг", vbTextCompare)
                If pp > 0 Then
                    FittingName = "планка в угол"
                    
                    pp = InStr(pp, FittOpt, " ", vbTextCompare)
                    If pp > 0 Then FittOpt = Trim(Left(FittOpt, pp))
                    If pp = 0 Or Len(FittOpt) < 3 Then FittOpt = PlankColor
                    GoTo exit_if
                    'Length = TTThickness
                End If
                
                pp = InStr(1, FittOpt, "к стол", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "стол", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "м-у", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "м/у", vbTextCompare)
                If pp Then
                    FittingName = "планка м/у столешницами"
                    
                    pp = InStr(1, Left(FittOpt, pp), " ", vbTextCompare)
                    If pp > 0 Then FittOpt = Trim(Left(FittOpt, pp))
                    If pp = 0 Or Len(FittOpt) < 3 Then FittOpt = PlankColor
                    GoTo exit_if
                    'Length = TTThickness
                End If
                
                
                
                pp = InStr(1, FittOpt, "к газ", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "газ", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "к плит", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "к г/п", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "г/п", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "плит", vbTextCompare)
                If pp Then
                    FittingName = "планка к газ. плите"
                    
                    pp = InStr(pp, FittOpt, " ", vbTextCompare)
                    If pp > 0 Then FittOpt = Trim(Left(FittOpt, pp))
                    If pp = 0 Or Len(FittOpt) < 3 Then FittOpt = PlankColor
                    GoTo exit_if
                    'Length = TTThickness
                End If
                                
                If InStr(1, Fitting, "монт", vbTextCompare) Then
                    pp = InStr(1, Fitting, "планк", vbTextCompare)
                    If (pp > 0) Then
                        FittingName = "планка монтажная"
                        
                        If Not IsEmpty(length) And IsEmpty(qty) Then
                            qty = CInt(length) \ 200
                            length = Empty
                        End If
                    End If
                End If
                pp = InStr(1, Fitting, "клино", vbTextCompare)
                    If (pp < k) And (pp > 0) Then
                    length = Null
                    
                        pp = InStr(1, Fitting, "sens", vbTextCompare)
                        If pp = 0 Then pp = InStr(1, Fitting, "sensis", vbTextCompare)
                        If pp = 0 Then pp = InStr(1, Fitting, "сэнсис", vbTextCompare)
                        If pp = 0 Then pp = InStr(1, Fitting, "сенсис", vbTextCompare)
                        If pp > k Then
                        Dim plankoptpos As Integer
                            FittingName = "клиновая планка Sensys"
                            plankoptpos = InStr(pp, Fitting, "5", vbTextCompare)
                            If plankoptpos > pp Then FittOpt = "5гр"
                            plankoptpos = InStr(pp, Fitting, "10", vbTextCompare)
                            If plankoptpos > pp Then FittOpt = "10гр"
                    ElseIf pp = 0 Then
                            FittingName = "клиновая планка 5гр"
                            FittOpt = Null
                    End If
                End If
                
            ElseIf Item = "труба" Then
                If InStr(1, FittOpt, "крепл", vbTextCompare) > 0 And InStr(1, FittOpt, "стол", vbTextCompare) = 0 Then
                    FittingName = "крепление к трубе бар."
                ElseIf InStr(1, FittOpt, "стол", vbTextCompare) > 0 Then
                    FittingName = "крепление к стол тр. бар."
                End If
                
            ElseIf Item = "карг" Then
                FittingName = "карго"
            ElseIf Item = "завес Sens" Or _
                    Item = "петли Sens" Or _
                    Item = "петля Sens" Or _
                    Item = "завес Сенс" Or _
                    Item = "петли Сенс" Or _
                    Item = "петля Сенс" Or _
                    Item = "завес Сэнс" Or _
                    Item = "петли Сэнс" Or _
                    Item = "петля Сэнс" Or _
                    Item = "завесы Sens" Or _
                    Item = "завесы Сенс" Or _
                    Item = "завесы Сэнс" _
            Then
                FittingName = "Завес Sensys"
            ElseIf Item = "Уголки метал" Or Item = "Угол метал" Or Item = "Уголок метал" Then
                FittingName = "уголок мет."
            ElseIf Item = "мэйджик лайт" Then
                FittingName = "трансформатор"
             ElseIf Item = "стенд" Then
                FittingName = "ф-ра комплект"
                
            ElseIf Item = "мейджик лайт" Then
                FittingName = "трансформатор"
            ElseIf Item = "ухо" Then
                FittingName = "петля*"
            ElseIf Item = "уши" Then
                FittingName = "петля*"
            ElseIf Item = "штанга" Then
                FittingName = "штанга"
            ElseIf Item = "блок " Then
                FittingName = "трансформатор"
            ElseIf Item = "монт" Then
                FittingName = "планка монтажная"
            ElseIf Item = "вент" Then
                If InStr(1, ActiveCell.Value, "Реш", vbTextCompare) > 0 Then
                    FittingName = "Решетка вентиляционная"
                    If InStr(1, ActiveCell.Value, "Вал", vbTextCompare) > 0 Or _
                        InStr(1, ActiveCell.Value, "Volp", vbTextCompare) > 0 Or _
                        InStr(1, ActiveCell.Value, "Вол", vbTextCompare) > 0 Then
                        FittOpt = "ВОЛПАТО"
                    ElseIf InStr(1, ActiveCell.Value, "Решетка в", vbTextCompare) > 0 Then
                        FittingName = "Решетка вентиляционная"
                    End If
                End If
            ElseIf Item = "полка оборот" Then
                FittingName = "полка оборотная"
            ElseIf Item = "полка оборот" Then
                FittingName = "полка оборотная"
            ElseIf Item = "клипсы к цок" Or Item = "клипса к цок" Or Item = "клипсы цок" Or Item = "клипса цок" Or Item = "Клипсы на цок" Or Item = "Клипса на цок" _
                Or Item = "Клипсы для цок" Or Item = "Клипса для цок" _
            Then
                FittingName = "*крепл. цоколя универсал"
            ElseIf Item = "конф" Then
                If InStr(1, ActiveCell.Value, "Загл", vbTextCompare) > 0 Then
                    FittingName = "Заглушка"
                Else
                    FittingName = "Конфирмант"
                End If
            ElseIf Item = cStol Then
                FittingName = "стол китайский"
            ElseIf Item = cStool Then
                FittingName = cStul
            ElseIf Item = "нога" Then
                FittingName = cNogi
            ElseIf Item = "Экспоз" Then
                FittingName = "ВЫБРАТЬ Экспозитор"
            ElseIf Item = "Дугов" Then
                FittingName = "VS - Дуговой держатель"
            ElseIf Item = "петл" Then
                FittingName = "завес"
            ElseIf Item = "ведро" Then
                FittingName = "ведро"
                If InStr(1, ActiveCell.Value, "встро", vbTextCompare) > 0 And InStr(1, ActiveCell.Value, "5л", vbTextCompare) > 0 Then
                                FittOpt = "встроенное 5л"
                ElseIf InStr(1, ActiveCell.Value, "встро", vbTextCompare) > 0 And InStr(1, ActiveCell.Value, "11л", vbTextCompare) > 0 Then
                                FittOpt = "встроенное 11л"
                End If
            ElseIf Item = "площадк" Then
                FittingName = "площадка для завеса"
            ElseIf Item = "подпят" Or Item = "потпят" Then
                FittingName = "подпятник"
             ElseIf Item = "полир" Then
                FittingName = "полироль"
             ElseIf Item = "палир" Then
                FittingName = "полироль"
            ElseIf Item = "стык" Then
                FittingName = "соединитель цоколя"
            ElseIf Item = "соед" Then
                FittingName = "соединитель цоколя"
            ElseIf Item = "поддон" Then
                FittingName = "поддон"
      '          FittingName = "поддон алюминиевый"
            ElseIf Item = "диод" Then
                FittingName = "подсветка диодная"
            ElseIf Item = "аморт" Or Item = "амморт" Then
                FittingName = "амортизатор"
            ElseIf Item = "отбойник" Then
                FittingName = "отбойник"
                 If InStr(1, ActiveCell.Value, "ПВХ", vbTextCompare) > 0 Then
                    FittOpt = "ПВХ"
                End If
            ElseIf Item = "дюбел" Then
                FittingName = "_дюбел"
            ElseIf Item = "эксц" Then
                FittingName = "_эксц"
            ElseIf Item = "волшеб" Then
                FittingName = "карго"
            ElseIf Item = "push to open" Or Item = "push-to-open" Or Item = "пуш то опен" Or Item = "пуш ту опен" Then
                FittingName = "нажимной м-м Push-To-Open"
            ElseIf Item = "стеклодерж" Then
                FittingName = "стеклодержатель*"
            ElseIf Item = "полкодерж" Then
                FittingName = "полкодержатель"
                FittOpt = ""
               
'                If InStr(1, ActiveCell.Value, "Sekura", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "Секура", vbTextCompare) > 0 Then
'                     If InStr(1, ActiveCell.Value, "a 2-1", vbTextCompare) > 0 Then
'                    FittOpt = "Sekura 2-1"""
'                Else
'                    FittOpt = "Sekura 8 (для стекла)"
'                End If
            ElseIf Item = "пеликан" Then
                FittingName = "полкодержатель"
                FittOpt = ""
            ElseIf Item = "направл" Then
                If InStr(1, ActiveCell.Value, "квадро", vbTextCompare) > 0 Then
                    FittingName = "направляющие квадро"
                Else
                    FittingName = "направляющие"
                End If
            ElseIf Item = "метабокс" Then
                FittingName = "метабокс *"
            ElseIf Item = "Решетка вентиляционная" Then
                If InStr(1, ActiveCell.Value, "Вал", vbTextCompare) > 0 Or _
                InStr(1, ActiveCell.Value, "Volp", vbTextCompare) > 0 Or _
                InStr(1, ActiveCell.Value, "Вол", vbTextCompare) > 0 Then
                FittOpt = "ВОЛПАТО"
                
                ElseIf InStr(1, ActiveCell.Value, "Решетка в", vbTextCompare) > 0 Then
                    FittingName = "Решетка вентиляционная"
                End If
            ElseIf Item = "банк" Then
                If InStr(1, ActiveCell.Value, "банк", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "организ", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "тандем", vbTextCompare) > 0) Then
                    FittingName = "лоток д/тб"
                End If
            ElseIf Item = "тандембокс" Then
                FittingName = "тандембокс *"
            ElseIf Item = "доводчик к м" Or Item = "доводчик на м" Or Item = "доводчик для м" Or _
                    Item = "доводчики к м" Or Item = "доводчики на м" Or Item = "доводчики для м" Then
                FittingName = "доводчик на метабокс"
            ElseIf Item = "надставка" Then
                FittingName = "надставка д/тб"
            ElseIf Item = "система перегородок" Then
                FittingName = "разделитель д/тб с флажк."
            
            ElseIf Item = "стул" Then
                 If InStr(1, ActiveCell.Value, "стул", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "Ибица", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "Феликс", vbTextCompare) > 0) Then
                    FittingName = "стул И/Ф"
                ElseIf InStr(1, ActiveCell.Value, "стул", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "Женева", vbTextCompare) > 0) Then
                    FittingName = "стул Женева"
                ElseIf InStr(1, ActiveCell.Value, "стул", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "Браун", vbTextCompare) > 0) Then
                    FittingName = "стул Браун"
                ElseIf InStr(1, ActiveCell.Value, "стул", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "Юл", vbTextCompare) > 0) Then
                    FittingName = "стул Юлия"
                ElseIf InStr(1, ActiveCell.Value, "стул", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "Zebra", vbTextCompare) > 0) Then
                    FittingName = "стул Zebra"
                End If
            ElseIf Item = "меджик лайт" Or Item = "magic light" Then
                 If InStr(1, ActiveCell.Value, "меджик лайт", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "magic light", vbTextCompare) > 0 Then
                    FittingName = "трансформатор"
                    End If
            ElseIf Item = "электроблок" Then
                 If InStr(1, ActiveCell.Value, "электроблок", vbTextCompare) > 0 Then
                    FittingName = "трансформатор"
                End If
            ElseIf Item = "трансформатор" Then
                 If InStr(1, ActiveCell.Value, "трансформатор", vbTextCompare) > 0 Then
                    FittingName = "трансформатор"
                End If
            ElseIf Item = "лоток" Then
                 If IsEmpty(qty) Then qty = 1
                If InStr(1, ActiveCell.Value, "танд", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "тб", vbTextCompare) > 0 Then
                    FittingName = "лоток д/тб"
                ElseIf InStr(1, ActiveCell.Value, "ORGA", vbTextCompare) > 0 Then
                    FittingName = "лоток ORGALINE"
                 ElseIf InStr(1, ActiveCell.Value, "Arci", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "Архит", vbTextCompare) > 0 Then
                    FittingName = "лоток в Архитех"
                 Else
                 FittingName = "лоток"
                End If
            ElseIf Item = "коврик" Then
                FittingName = "коврик антискольжения"
            ElseIf Item = "сушк" Then
                If InStr(1, ActiveCell.Value, "одноур", vbTextCompare) > 0 Then
                    FittOpt = "одноуровневая хром"
                ElseIf InStr(1, ActiveCell.Value, "хром", vbTextCompare) > 0 Then
                    FittOpt = "хром"
                ElseIf InStr(1, ActiveCell.Value, "белая", vbTextCompare) > 0 Then
                    FittOpt = "белая"
                Else
                    FittOpt = ""
                End If
                If IsEmpty(qty) Then qty = 1
            
            End If
exit_if:
'            If InStr(1, Fitting, "стул ", vbTextCompare) > 0 Then
'            ActiveCell = Replace(ActiveCell.Text, ".", " ", InStr(1, ActiveCell.Text, "стул ", vbTextCompare) + 1)
'            End If
    
            If Item = cGalog Then
                If qty Mod 3 = 0 Then
                    FittingName = "галогенки 3"
                    qty = qty \ 3
                ElseIf qty Mod 5 = 0 Then
                    FittingName = "галогенки 5"
                    qty = qty \ 5
                End If
            End If
            
            If Item = "полос" And (InStr(1, ActiveCell.Value, "цокол", vbTextCompare) > 0 And InStr(1, ActiveCell.Value, "прозр", vbTextCompare) > 0) Then
                Dim find18 As Integer
                find18 = 0
                find18 = InStr(InStr(1, ActiveCell.Value, "цокол", vbTextCompare), ActiveCell.Value, "18", vbTextCompare)
                If find18 > 1 Then FittOpt = "18" Else FittOpt = "16"
                
                If InStr(1, ts, "3м", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 1
                ElseIf InStr(1, ts, "6м", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 2
                ElseIf InStr(1, ts, "9м", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 3
                ElseIf InStr(1, ts, "12м", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 4
                End If
            End If
            
                
            
           ' If Not IsEmpty(FittOpt) Then
'                If bExtra And (IsEmpty(Qty) And (IsEmpty(Length) _
'                    And (InStr(1, Item, cHandle) > 0 _
'                    Or InStr(1, Item, cLeg) > 0))) Then  ' если не указана на заказ или доп-но, то считаем, что указан а йрнитура для заказа по умолчанию (ножки, ручки)
                If Not IsMissing(FittOpt) And _
                    IsEmpty(qty) And _
                    (InStr(1, Item, cHandle, vbTextCompare) > 0 _
                     Or InStr(1, Item, cLeg, vbTextCompare) > 0) Then ' если не указана на заказ или доп-но, то считаем, что указан а йрнитура для заказа по умолчанию (ножки, ручки)
                    
                    If InStr(1, FittOpt, "клиент", vbTextCompare) Then
                        FindFittings = Null
                    Else
                        FindFittings = FittOpt
                    End If
                    
                Else
                    ' если это ручки, определимся с длиной шурупов
                    If InStr(1, Item, cHandle, vbTextCompare) > 0 Then
                             
                        If IsEmpty(HandleScrew) Then HandleScrew = GetHandleScrew(FittOpt, face)
                        If Not IsEmpty(HandleScrew) Then UpdateOrder OrderId, HandleScrew
                         
                        If bExtra Then ' если дополнительно
                        
                            CheckHandle FittOpt
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            
                        Else ' если на заказ
                        
                            FindFittings = Null
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            
                        End If
                                                     
                    Else
                        If bExtra Then ' если дополнительно
                            
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            If Item = cPlank Then PlankColor = FittOpt
                            
                        Else ' если на заказ
                            
                            FindFittings = Null
                            'FindFittings = FittOpt
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            
                        End If
                    End If
                End If
            
            
            
            k = en + 1  ' конец тек. строки, начало следующей
        Loop While k < Len(Fitting) And bPlus
        
        Fitting = Left(Fitting, ShiftK - 1) & Mid(Fitting, k)
        ShiftK = k
    End If
    Else
        
    
    
    
    '   InStr(ShiftK, Fitting, "Blumot", vbTextCompare) > 0)
        Dim t As String
        If (IsMissing(changeCaseZaves) = False) Then
         If (InStr(ShiftK, Fitting, "Sens", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "Сенсис", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "Sensyc", vbTextCompare) > 0) Then
                If changeCaseZaves <> 1 Then
                    changeCaseZaves = 1
                    kitchenPropertyCurrent.changeCaseZaves = 1
                   ' casepropertyCurrent.p_changeZaves = 1
                    ShiftK = Len(Fitting)
                    ActiveCell.Characters(k, Len(Item)).Font.Color = vbGreen
                    If Cells(ActiveCell.row, 10).Value <> "" Then
                        
                        t = Cells(ActiveCell.row, 10).Value
                        
                        Cells(ActiveCell.row, 10).Value = t & "!!!Смена завесов на СЕНСИС!!!"
                         Else
                        Cells(ActiveCell.row, 10).Value = "!!!Смена завесов на СЕНСИС!!!"
                    End If
                End If
            ElseIf (InStr(ShiftK, Fitting, "Blumot", vbTextCompare) > 0) Then
                If changeCaseZaves <> 2 Then
                kitchenPropertyCurrent.changeCaseZaves = 2
                    changeCaseZaves = 2
                    'casepropertyCurrent.p_changeZaves = 2
                    ShiftK = Len(Fitting)
                    ActiveCell.Characters(k, Len(Item)).Font.Color = vbGreen
                    If Cells(ActiveCell.row, 10).Value <> "" Then
                        t = Cells(ActiveCell.row, 10).Value
                        
                        Cells(ActiveCell.row, 10).Value = t & "!!!Смена завесов на БЛЮМОУШИН!!!"
                         Else
                        Cells(ActiveCell.row, 10).Value = "!!!Смена завесов на БЛЮМОУШИН!!!"
                    End If
                End If
            ElseIf (InStr(ShiftK, Fitting, "FGV", vbTextCompare) > 0) Then
                If changeCaseZaves <> 3 Then
                kitchenPropertyCurrent.changeCaseZaves = 3
                    changeCaseZaves = 3
                    'casepropertyCurrent.p_changeZaves = 3
                    ShiftK = Len(Fitting)
                    ActiveCell.Characters(k, Len(Item)).Font.Color = vbGreen
                    If Cells(ActiveCell.row, 10).Value <> "" Then
                        t = Cells(ActiveCell.row, 10).Value
                        
                        Cells(ActiveCell.row, 10).Value = t & "!!!Смена завесов на FGV!!!"
                         Else
                        Cells(ActiveCell.row, 10).Value = "!!!Смена завесов на FGV!!!"
                    End If
                End If
            
            End If
        End If
    End If
End Function


Public Function AddOrder(ByVal ShipID As Long, _
                         ByVal row As Integer, _
                         ByVal Customer As String, _
                         ByVal OrderN As String, _
                         Optional ByVal SetQty) As Long
On Error GoTo err_AddOrder

'***********************************************
'    Dim OrderUID As Long, k As Integer
'    k = InStr(2, OrderN, "-", vbTextCompare)
'    If k > 0 Then
'        Dim buf
'        buf = Split(OrderN, "-", 2, vbTextCompare)
'
'        If UBound(buf) = 1 Then
'            OrderUID = buf(0)
'            OrderN = buf(1)
'        Else
'            GoTo err_AddOrder
'        End If
'    Else
'        GoTo err_AddOrder
'    End If
'***********************************************

    
'    If IsEmpty(HandleScrew) Then HandleScrew = Null
'    If IsEmpty(HangColor) Then HangColor = Null

    
    Dim AddComm As ADODB.Command
    Set AddComm = New ADODB.Command
    AddComm.ActiveConnection = GetConnection
    AddComm.CommandType = adCmdStoredProc
    AddComm.CommandText = "AddOrder"
    
    '***********************************************
    'AddComm("@OrderID") = OrderUID
    '***********************************************
    AddComm.Parameters("@ShipID") = ShipID
    AddComm.Parameters("@Customer") = Left(Customer, 25)
    AddComm.Parameters("@Number") = OrderN
    AddComm.Parameters("@FirstRow") = row
'    AddComm.Parameters("@HangColor") = HangColor
'    AddComm.Parameters("@HandleScrew") = HandleScrew
'    If Not IsMissing(BibbColor) And Len(BibbColor) > 0 Then AddComm.Parameters("@BibbColor") = BibbColor
    If Not IsMissing(SetQty) Then If SetQty > 0 Then AddComm.Parameters("@SetQty") = SetQty
      
    AddComm.Execute
    AddOrder = AddComm.Parameters("@OrderID")
    
    Exit Function
err_AddOrder:
    MsgBox Error, vbCritical, "Добавление заказа"
    AddOrder = 0
End Function
'Public Sub AddCaseParams(ByVal OCID As Long, _
'                        ByVal param_name As String, _
'                        ByVal param_value As String)
'On Error GoTo err_AddCaseParams
'    Init_rsOrderCasesParams
'
'    rsOrderCasesParams.AddNew
'    rsOrderCasesParams!OCID = OCID
'    rsOrderCasesParams!param_name = param_name
'    rsOrderCasesParams!param_value = param_value
'Exit Sub
'err_AddCaseParams:
'    MsgBox Error, vbCritical, "Добавление свойства шкафа заказа"
'End Sub





'Public Function AddCase(ByVal OrderID As Long, _
'                    ByVal caseID As Long, _
'                    ByVal name As String, _
'                    ByVal CaseName As String, _
'                    ByVal Qty As Integer, _
'                    ByVal CaseHang, _
'                    ByVal Handle, _
'                    ByVal HandleExtra, _
'                    ByVal Leg, _
'                    ByVal DoorCount, _
'                    ByVal windowcount, _
'                    ByVal Drawermount, _
'                    ByVal Doormount, _
'                    ByVal Bibb As Integer, _
'                    ByVal glub As Integer, _
'                    Optional ByVal ShelfQty As Integer, _
'                    Optional ByVal bNeStandart As Boolean = True, _
'                    Optional ByVal Row, _
'                    Optional ByVal NoFace, _
'                    Optional ByVal caseChangeZaves As Integer = 0, _
'                    Optional ByVal caseChangeKonfirmant As Integer = 0, _
'                    Optional ByVal dspbottom As Integer = 0, _
'                    Optional ByVal caseHeight As Integer = 0 _
'                    ) As Long
'
'    On Error GoTo err_AddCase
'    Init_rsOrderCases
'    Dim OCID As Long
'    rsOrderCases.AddNew
'    rsOrderCases!OrderID = OrderID
'    rsOrderCases!caseID = caseID
'    rsOrderCases!CaseName = Left(CaseName, 70)
'    rsOrderCases!Qty = Qty
'    OCID = rsOrderCases!OCID
'
'    If Not IsMissing(Row) Then rsOrderCases!Row = Row
'    rsOrderCases!Standart = Not bNeStandart
'
'    If Not IsEmpty(CaseHang) Then rsOrderCases!CaseHang = CaseHang  ' по умолчания петля
'    rsOrderCases!Bibb = Bibb
'    If IsNull(Handle) Then
'        rsOrderCases!Handle = Null
'        rsOrderCases!HandleExtra = 0
'    Else
'        rsOrderCases!Handle = Handle
'        If Not IsEmpty(HandleExtra) Then rsOrderCases!HandleExtra = HandleExtra
'    End If
'    If Not IsEmpty(Leg) Then
'        rsOrderCases!CaseStand = Leg ' если Empty, значит по умолчанию, если Null - значит нет, есди значение, то тип ног
'    'Else
'       ' rsOrderCases!CaseStand = "черная"
'    End If
'    rsOrderCases!DoorCount = DoorCount
'
'    rsOrderCases!caseChangeZaves = caseChangeZaves
'    rsOrderCases!caseChangeKonfirmant = caseChangeKonfirmant
'    rsOrderCases!dspbottom = dspbottom
'
'
'    If Not IsEmpty(windowcount) Then rsOrderCases!windowcount = windowcount
'    If Not IsEmpty(Drawermount) Then rsOrderCases!Drawermount = Left(Drawermount, 20)
'    If Not IsEmpty(Doormount) Then rsOrderCases!Doormount = Doormount
'    If Not IsEmpty(NoFace) Then rsOrderCases!NoFace = NoFace
'    If Not IsEmpty(glub) Then
'        If glub > 0 Then
'            rsOrderCases!glub = glub
'            Else
'            rsOrderCases!glub = 570 '530
'        End If
'    End If
'    If bNeStandart Then
'        If ShelfQty >= 2 And Left(name, 1) <> "П" Then
'            Select Case name
'                Case "ШН 2Т"
'                   ' If Not IsEmpty(caseHeight) Then
'                                            'If caseHeight > 910 Then FormElement.AddElementToOrder OrderID, "полик", 2 * Qty, caseID Else
'                   FormElement.AddElementToOrder OrderID, "полик", Qty, caseID
'                   ' End If
'                Case "ШН915", "ШНП915", "ШНУ915", "ШН скос915", "ШНЗ915", "ШНЗУ915", "ШНУГ915", "ШНС915"
'                Case "ШЛП", "ШСП"   '"ШНП", "ШН скос"
'                    FormElement.AddElementToOrder OrderID, "полик", Qty, caseID
'                Case "ШНВ А"
'                    FormFitting.AddFittingToOrder OrderID, "крестик", Qty, , , caseID
'                'Case "ШНУГ" ' "ШНУ"
'                '    FormElement.AddElementToOrder OrderID, "полка угол", Qty, CaseID
'                Case Else
'                    FormElement.AddElementToOrder OrderID, "полка", Qty, caseID
'            End Select
'        End If
'    End If
'    rsOrderCases.Update
'    AddCase = rsOrderCases!OCID
'    Exit Function
'err_AddCase:
'    MsgBox Error, vbCritical, "Добавление шкафа в заказ"
'End Function
Public Sub AddCaseParamsbySp(ocid As Long, pname As String, pvalue As String)
On Error GoTo err_AddCaseParamsbySp
        Dim prm As ADODB.Parameter
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdStoredProc
        comm.CommandTimeout = 90
        comm.CommandText = "sp_addOrderCasesParams"
        comm.NamedParameters = True
        comm.Parameters.Append comm.CreateParameter("@OCID", adInteger, adParamInput, , ocid)
        comm.Parameters.Append comm.CreateParameter("@param_name", adVarChar, adParamInput, 50, Left(pname, 50))
        comm.Parameters.Append comm.CreateParameter("@param_value", adVarChar, adParamInput, 50, Left(pvalue, 50))
        comm.Execute

         Exit Sub
err_AddCaseParamsbySp:

    MsgBox Error, vbCritical, "Добавление шкафа в заказ"
End Sub
Public Function GetCaseId(name As String) As Integer
On Error GoTo err_GetCaseId
Dim prm As ADODB.Parameter
Dim comm As ADODB.Command
Set comm = New ADODB.Command

comm.ActiveConnection = GetConnection
comm.CommandType = adCmdStoredProc
comm.CommandTimeout = 90
comm.CommandText = "sp_getCaseId"
comm.NamedParameters = True
comm.Parameters.Append comm.CreateParameter("@name", adVarChar, adParamInput, 150, Left(name, 150))
comm.Parameters.Append comm.CreateParameter("@caseid", adInteger, adParamOutput)
comm.Execute
GetCaseId = comm.Parameters("@caseid")


Exit Function
err_GetCaseId:
GetCaseId = 0
MsgBox Error, vbCritical, "Поиск прототипа в базе"
End Function

Public Function createCaseId(name As String, args As String) As Integer
On Error GoTo err_createCaseId
Dim prm As ADODB.Parameter
Dim comm As ADODB.Command
Set comm = New ADODB.Command

comm.ActiveConnection = GetConnection
comm.CommandType = adCmdStoredProc
comm.CommandTimeout = 90
comm.CommandText = "sp_createCaseId"
comm.NamedParameters = True
comm.Parameters.Append comm.CreateParameter("@name", adVarChar, adParamInput, 150, Left(name, 150))
comm.Parameters.Append comm.CreateParameter("@legDefault", adVarChar, adParamInput, 20, "Чёрная 100")
comm.Parameters.Append comm.CreateParameter("@str", adVarChar, adParamInput, 512, args)

comm.Parameters.Append comm.CreateParameter("@caseid", adInteger, adParamOutput)
comm.Execute
createCaseId = comm.Parameters("@caseid")


Exit Function
err_createCaseId:
createCaseId = 0
MsgBox Error, vbCritical, "Создание прототипа в базе"
End Function


Public Function AddCaseBySp(ByVal OrderId As Long, _
                    ByVal caseID As Long, _
                    ByVal name As String, _
                    ByVal casename As String, _
                    ByVal qty As Integer, _
                    ByVal CaseHang, _
                    ByVal Handle, _
                    ByVal HandleExtra, _
                    ByVal Leg, _
                    ByVal DoorCount, _
                    ByVal windowcount, _
                    ByVal Drawermount, _
                    ByVal Doormount, _
                    ByVal Bibb As Integer, _
                    ByVal glub As Integer, _
                    Optional ByVal ShelfQty As Integer, _
                    Optional ByVal bNeStandart As Boolean = True, _
                    Optional ByVal row, _
                    Optional ByVal NoFace, _
                    Optional ByVal cabType As String = "ZOV" _
                    ) As Long
                    
On Error GoTo err_AddCaseSP
        AddCaseBySp = 0
        Dim prm As ADODB.Parameter
        Dim comm As ADODB.Command
        Set comm = New ADODB.Command
        
        comm.ActiveConnection = GetConnection
        comm.CommandType = adCmdStoredProc
        comm.CommandTimeout = 90
        comm.CommandText = "sp_AddCaseOrder"
        comm.NamedParameters = True
        comm.Parameters.Append comm.CreateParameter("@ocid", adInteger, adParamOutput)
        comm.Parameters.Append comm.CreateParameter("@orderid", adInteger, adParamInput, , OrderId)
        comm.Parameters.Append comm.CreateParameter("@caseID", adInteger, adParamInput, , caseID)
        comm.Parameters.Append comm.CreateParameter("@CaseName", adVarChar, adParamInput, 70, Left(casename, 70))
        comm.Parameters.Append comm.CreateParameter("@Qty", adTinyInt, adParamInput, , qty)
        If Not casepropertyCurrent Is Nothing Then
            If Not IsMissing(casepropertyCurrent.p_cabHeigth) Then
                comm.Parameters.Append comm.CreateParameter("@cabHeight", adInteger, adParamInput, , casepropertyCurrent.p_cabHeigth)
            End If
            If Not IsMissing(casepropertyCurrent.p_cabWidth) Then
                comm.Parameters.Append comm.CreateParameter("@cabWidth", adInteger, adParamInput, , casepropertyCurrent.p_cabWidth)
            End If
            If Not IsMissing(casepropertyCurrent.p_cabDepth) Then
                comm.Parameters.Append comm.CreateParameter("@cabDepth", adInteger, adParamInput, , casepropertyCurrent.p_cabDepth)
            End If
            If Not IsMissing(casepropertyCurrent.p_z_st_dsp) Then
                comm.Parameters.Append comm.CreateParameter("@z_st_dsp", adBoolean, adParamInput, , casepropertyCurrent.p_z_st_dsp)
            End If
            comm.Parameters.Append comm.CreateParameter("@caseChangeZaves", adInteger, adParamInput, , casepropertyCurrent.p_changeZaves)
            comm.Parameters.Append comm.CreateParameter("@caseChangeKonfirmant", adInteger, adParamInput, , casepropertyCurrent.p_changeCaseKonfirmant)
            comm.Parameters.Append comm.CreateParameter("@dspbottom", adInteger, adParamInput, , casepropertyCurrent.p_dspbottom)
        End If

        If Not IsMissing(row) Then
            comm.Parameters.Append comm.CreateParameter("@Row", adSmallInt, adParamInput, , row)
        End If

        comm.Parameters.Append comm.CreateParameter("@bNeStandart", adBoolean, adParamInput, , Not bNeStandart)
    
        If Not IsEmpty(CaseHang) Then
            comm.Parameters.Append comm.CreateParameter("@CaseHang", adVarChar, adParamInput, 15, CaseHang)
        End If

        comm.Parameters.Append comm.CreateParameter("@Bibb", adBoolean, adParamInput, , Bibb)
        
        If IsNull(Handle) Then
            comm.Parameters.Append comm.CreateParameter("@Handle", adVarChar, adParamInput, 20, Handle)
            comm.Parameters.Append comm.CreateParameter("@HandleExtra", adTinyInt, adParamInput, , 0)
        Else
            comm.Parameters.Append comm.CreateParameter("@Handle", adVarChar, adParamInput, 20, Handle)
            If Not IsEmpty(HandleExtra) Then
                comm.Parameters.Append comm.CreateParameter("@HandleExtra", adTinyInt, adParamInput, , HandleExtra)
                'rsOrderCases!HandleExtra = HandleExtra
                Else
                comm.Parameters.Append comm.CreateParameter("@HandleExtra", adTinyInt, adParamInput, , Null)
            End If
        End If
    
        If Not IsEmpty(Leg) Then
            comm.Parameters.Append comm.CreateParameter("@Leg", adVarChar, adParamInput, 15, Leg)
    '        Else
    '        comm.Parameters.Append comm.CreateParameter("@Leg", adVarChar, adParamInput, 15, Null)
            
        End If
        
        comm.Parameters.Append comm.CreateParameter("@DoorCount", adTinyInt, adParamInput, , DoorCount)
      
    
        If Not IsEmpty(windowcount) Then
            comm.Parameters.Append comm.CreateParameter("@windowcount", adTinyInt, adParamInput, , windowcount) 'rsOrderCases!windowcount = windowcount
        End If
        
        If Not IsEmpty(Drawermount) Then
            comm.Parameters.Append comm.CreateParameter("@Drawermount", adVarChar, adParamInput, 20, Drawermount) 'rsOrderCases!Drawermount = Left(Drawermount, 20)
        End If
        
        If Not casepropertyCurrent Is Nothing Then
            If casepropertyCurrent.p_delete_doormount = False Then
               If Not IsEmpty(Doormount) Then
                   comm.Parameters.Append comm.CreateParameter("@Doormount", adVarChar, adParamInput, 50, Doormount) 'rsOrderCases!Doormount = Doormount
               End If
            End If
        Else
            If Not IsEmpty(Doormount) Then
                comm.Parameters.Append comm.CreateParameter("@Doormount", adVarChar, adParamInput, 50, Doormount) 'rsOrderCases!Doormount = Doormount
            End If
        End If
        
        comm.Parameters.Append comm.CreateParameter("@cabType", adVarChar, adParamInput, 50, cabType)
        
        If Not IsEmpty(NoFace) Then
        comm.Parameters.Append comm.CreateParameter("@NoFace", adSmallInt, adParamInput, , NoFace) 'rsOrderCases!NoFace = NoFace
        End If

        If Not casepropertyCurrent Is Nothing Then
            If casepropertyCurrent.p_cabDepth > 0 Then
                    comm.Parameters.Append comm.CreateParameter("@glub", adInteger, adParamInput, , casepropertyCurrent.p_cabDepth)
            ElseIf Not IsEmpty(glub) Then
                If glub > 0 Then
                    comm.Parameters.Append comm.CreateParameter("@glub", adInteger, adParamInput, , glub)
                    Else
                     comm.Parameters.Append comm.CreateParameter("@glub", adInteger, adParamInput, , 570)
                End If
            Else
                comm.Parameters.Append comm.CreateParameter("@glub", adInteger, adParamInput, , 570)
            End If
        End If
 '   If bNeStandart Then
   ' End If
    
    comm.Execute
    
    AddCaseBySp = comm.Parameters("@ocid")
    
    Exit Function
err_AddCaseSP:
    AddCaseBySp = 0
    MsgBox Error, vbCritical, "Добавление шкафа в заказ"
End Function

Public Function GetHandleScrew(ByVal Handle, _
                               ByVal face) As Variant
            
                          
    If Not IsNull(Handle) And Not IsMissing(Handle) Then
            If Not (Handle = "ВЕРОНА" Or InStr(1, Handle, "МОДЕНА", vbTextCompare) > 0) Then
                If Not IsNull(face) And Not IsEmpty(face) Then
                If Len(face) > 3 Then
                    If InStr(1, face, "массив", vbTextCompare) > 0 And _
                           (InStr(1, Handle, "A025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "A-025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "А025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "А-025", vbTextCompare) > 0) Then
                          
                          GetHandleScrew = "40"
                    ElseIf InStr(1, face, "массив", vbTextCompare) = 0 And _
                           (InStr(1, Handle, "A025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "A-025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "А025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "А-025", vbTextCompare) > 0) _
                           Then
                          
                          GetHandleScrew = "35"
                          
                   
                    ElseIf InStr(1, face, "прим", vbTextCompare) > 0 Or _
                        InStr(1, face, "RAL", vbTextCompare) > 0 Or _
                        InStr(1, face, "Скарлет", vbTextCompare) > 0 Or _
                        InStr(1, face, "Женева", vbTextCompare) > 0 Or _
                        InStr(1, Replace(face, " ", ""), "Люкс2", vbTextCompare) > 0 Then
                        
                        GetHandleScrew = "25"
                
                
        
                    
                    ElseIf InStr(1, face, "массив", vbTextCompare) > 0 Or _
                            InStr(1, face, "Виола", vbTextCompare) > 0 Or _
                            InStr(1, face, "Марсе", vbTextCompare) > 0 Or _
                            InStr(1, face, "Вена", vbTextCompare) > 0 Or _
                            InStr(1, face, "акрил", vbTextCompare) > 0 Then
                        
                        GetHandleScrew = "28"
                        
                    ElseIf InStr(1, face, "система", vbTextCompare) > 0 Then
                    Else
                        GetHandleScrew = "22"
                    
                    End If
                End If 'If Len(Face) > 3 Then
            End If 'If Not IsNull(Face) And Not IsEmpty(Face) Then
             
             
            If IsEmpty(GetHandleScrew) Then
               Dim FormScrew As ScrewForm
               
               While GetHandleScrew = ""
                   Set FormScrew = New ScrewForm
                   FormScrew.Show 1
                   If FormScrew.result Then
                           GetHandleScrew = FormScrew.lbScrewLen.Value
                   End If
               Wend
               
               Set FormScrew = Nothing
            End If
        End If
     End If

End Function

Public Function GetHangColor(ByVal CaseColor) As String
'    If Not (IsEmpty(CaseColor) Or IsNull(CaseColor)) Then
'
'        If InStr(1, CaseColor, "бел", vbTextCompare) Then + пепел? + клен?
'            GetHangColor = "белая"
'        Else
'            GetHangColor = "желтая"
'        End If
'
'    End If

    If GetHangColor = "" Then
        Dim FormHangColor As HangColorForm
        Set FormHangColor = New HangColorForm
        
        
        While GetHangColor = ""
            Set FormHangColor = New HangColorForm
            FormHangColor.Caption = "Цвет завешек"
            If Not kitchenPropertyCurrent Is Nothing Then
                If kitchenPropertyCurrent.dspColor <> "" Then
                    FormHangColor.Caption = FormHangColor.Caption & " Бочки:" & kitchenPropertyCurrent.dspColor
                End If
            End If
            FormHangColor.Show 1
            If FormHangColor.result = True Then
                GetHangColor = FormHangColor.cbHangColor.Text
            End If
        Wend
        
        Set FormHangColor = Nothing
    End If
End Function

Public Function GetCamBibbColor(ByVal CaseColor) As Variant
   

    If GetCamBibbColor = "" Then
        Dim FormCamBibbColor As CamBibbColorForm
        
        While GetCamBibbColor = ""
            Set FormCamBibbColor = New CamBibbColorForm
            FormCamBibbColor.Caption = "Заглушки эксцентрика"
            If Not kitchenPropertyCurrent Is Nothing Then
                If kitchenPropertyCurrent.dspColor <> "" Then
                    FormCamBibbColor.Caption = FormCamBibbColor.Caption & " Бочки:" & kitchenPropertyCurrent.dspColor
                End If
            End If
            FormCamBibbColor.Show 1
            If FormCamBibbColor.result = True Then
                GetCamBibbColor = FormCamBibbColor.cbBibbColor.Text
            End If
            If GetCamBibbColor = "" Then
                GetCamBibbColor = Null
            End If
        Wend
        

        
        Set FormCamBibbColor = Nothing
    End If
End Function
Public Function GetBibbColor(ByVal CaseColor) As Variant
    If Not (IsEmpty(CaseColor) Or IsNull(CaseColor)) Then
    
'InStr(1, CaseColor, "бел", vbTextCompare) > 0 Or
        If InStr(1, CaseColor, "акация", vbTextCompare) > 0 Or _
            InStr(1, CaseColor, "см", vbTextCompare) > 0 Or _
            InStr(1, CaseColor, "тм", vbTextCompare) > 0 Then
            GetBibbColor = "белая"
'        ElseIf InStr(1, CaseColor, "груша", vbTextCompare) > 0 Then
'            GetBibbColor = "груша"
'        ElseIf InStr(1, CaseColor, "ольха", vbTextCompare) > 0 Then
'            GetBibbColor = "ольха"
        ElseIf InStr(1, CaseColor, "крем", vbTextCompare) > 0 Then
            GetBibbColor = "клен"
        ElseIf InStr(1, CaseColor, "платина", vbTextCompare) > 0 Then 'If InStr(1, CaseColor, "пепел", vbTextCompare) > 0 Or
            GetBibbColor = "пепел"
        ElseIf InStr(1, CaseColor, "каштан", vbTextCompare) > 0 Then
            GetBibbColor = "махонь"
        ElseIf InStr(1, CaseColor, "рустик", vbTextCompare) > 0 Then
            GetBibbColor = "орех"
        ElseIf InStr(1, CaseColor, "венге", vbTextCompare) > 0 Then
            GetBibbColor = "черный"
        ElseIf InStr(1, CaseColor, "негро", vbTextCompare) > 0 Then
            GetBibbColor = "черный"
        Else
            Dim BibbColors(), i As Integer
            GetBibbColors BibbColors
            For i = 0 To UBound(BibbColors) - 1
                If InStr(1, CaseColor, BibbColors(i), vbTextCompare) > 0 Then
                    GetBibbColor = BibbColors(i)
                End If
            Next
        End If

    End If

    If GetBibbColor = "" Then
        Dim FormBibbColor As BibbColorForm
        
        While GetBibbColor = ""
            Set FormBibbColor = New BibbColorForm
            FormBibbColor.Caption = "Цвет заглушек"
            If Not kitchenPropertyCurrent Is Nothing Then
                If kitchenPropertyCurrent.dspColor <> "" Then
                    FormBibbColor.Caption = FormBibbColor.Caption & " Бочки:" & kitchenPropertyCurrent.dspColor
                End If
            End If
            FormBibbColor.Show 1
            If FormBibbColor.result = True Then
                GetBibbColor = FormBibbColor.cbBibbColor.Text
            End If
            If GetBibbColor = "" Then
                GetBibbColor = Null
            End If
        Wend
        

        
        Set FormBibbColor = Nothing
    End If
End Function


Public Function GetPlankColor(ByVal TTColor) As String
    If Not (IsEmpty(TTColor) Or IsNull(TTColor)) Then

        If InStr(1, TTColor, "карара", vbTextCompare) > 0 Then
            GetPlankColor = "белый"
        ElseIf InStr(1, TTColor, "син", vbTextCompare) > 0 Then
            GetPlankColor = "синий"
        ElseIf InStr(1, TTColor, "зел", vbTextCompare) > 0 Then
            GetPlankColor = "зеленый"
        ElseIf InStr(1, TTColor, "дуб", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "ольх", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "бук", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "каф", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "проб", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "корень", vbTextCompare) > 0 Or _
                (InStr(1, TTColor, "беж", vbTextCompare) > 0 And InStr(1, TTColor, "гр", vbTextCompare) > 0) Then
            GetPlankColor = "бук"
        ElseIf InStr(1, TTColor, "махонь", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "рустик", vbTextCompare) > 0 Or _
                (InStr(1, TTColor, "красн", vbTextCompare) > 0 And InStr(1, TTColor, "гл", vbTextCompare) > 0) Or _
                InStr(1, TTColor, "гранит", vbTextCompare) > 0 And InStr(1, TTColor, "турин", vbTextCompare) = 0 Then
            GetPlankColor = "гранит"
        Else
            GetPlankColor = "хром"
        End If

    End If
End Function

Public Sub GetHangColors(ByRef HangColors())
    ReDim HangColors(5)
    HangColors(0) = "белая"
    HangColors(1) = "бук"
    HangColors(2) = "КАМАР806"
    HangColors(3) = "КАМАР807"
    HangColors(4) = "КАМАР808"
    
End Sub


Public Sub GetBibbColors(ByRef BibbColors())
    ReDim BibbColors(12)
    BibbColors(0) = "белая"
    BibbColors(1) = "бук"
    BibbColors(2) = "вишня"
    BibbColors(3) = "груша"
    BibbColors(4) = "дуб"
    BibbColors(5) = "клен"
    BibbColors(6) = "махонь"
    BibbColors(7) = "ольха"
    BibbColors(8) = "орех"
    BibbColors(9) = "пепел"
    BibbColors(10) = "черная"
    BibbColors(11) = "серый"
    BibbColors(12) = "бежевый"
    
    'SortArray BibbColors
End Sub
Public Sub GetCamBibbColors(ByRef CamBibbColors())
    ReDim CamBibbColors(10)
    Dim i As Integer
    i = 0
    CamBibbColors(i) = "Бежевая"
    i = i + 1
    CamBibbColors(i) = "Белая"
    i = i + 1
    CamBibbColors(i) = "Коричневая"
    i = i + 1
    CamBibbColors(i) = "Рыжая"
    i = i + 1
    CamBibbColors(i) = "Светло-беж"
    i = i + 1
    CamBibbColors(i) = "Светло-сер"
    i = i + 1
    CamBibbColors(i) = "Серая"
    i = i + 1
    CamBibbColors(i) = "Темно-беж"
    i = i + 1
    CamBibbColors(i) = "Темно-кор"
    i = i + 1
    CamBibbColors(i) = "Черная"
    i = i + 1
    CamBibbColors(i) = "Темно-сер"
    
    
End Sub

Public Function GetLegShelving(ByVal CaseColor As String) As String
    Select Case CaseColor
        Case "ольха", "груша"
            GetLegShelving = "ольха"
        Case "орех", "махонь", "рустик"
            GetLegShelving = "орех"
        Case "клен"
            GetLegShelving = "клен"
        Case Else
            GetLegShelving = CaseColor
    End Select
    CheckLeg GetLegShelving
End Function

Public Sub GetOtbColors(ByRef OtbColors())
    ReDim OtbColors(17)
    OtbColors(0) = "ХРОМ"
    OtbColors(1) = "БЕЖ ГР"
    OtbColors(2) = "БУК"
    OtbColors(3) = "ГРАНИТ"
    OtbColors(4) = "ДУБ"
    OtbColors(5) = "ЖЕЛТ КАМ"
    OtbColors(6) = "ЗЕЛЕНАЯ"
    OtbColors(7) = "ЗОЛОТАЯ"
    OtbColors(8) = "КАРАРА"
    OtbColors(9) = "МАХОНЬ"
    OtbColors(10) = "ОЛЬХА"
    OtbColors(11) = "ПРОБКА"
    OtbColors(12) = "РУСТИК"
    OtbColors(13) = "СИЗ КАМ"
    OtbColors(14) = "СИНЯЯ"
    OtbColors(15) = "ЧЕРН ГЛ"
    OtbColors(16) = "БЕЛ ГЛ"
    OtbColors(17) = "КРЕМ"
    'SortArray OtbColors
End Sub



Public Sub SortArray(ByRef MyArray())
    Dim lLoop As Long, lLoop2 As Long
    Dim str1 As String
    Dim str2 As String

    For lLoop = 0 To UBound(MyArray)

       For lLoop2 = lLoop To UBound(MyArray)

            If UCase(MyArray(lLoop2)) < UCase(MyArray(lLoop)) Then

                str1 = MyArray(lLoop)

                str2 = MyArray(lLoop2)

                MyArray(lLoop) = str2

                MyArray(lLoop2) = str1

            End If

        Next lLoop2
        
    Next lLoop
        
        
End Sub

Public Sub SortArrayByLengthDesc(ByRef MyArray())
    Dim lLoop As Long, lLoop2 As Long
    Dim str1 As String
    Dim str2 As String

    For lLoop = 0 To UBound(MyArray)

       For lLoop2 = lLoop To UBound(MyArray)

            If Len(MyArray(lLoop2)) > Len(MyArray(lLoop)) Then

                str1 = MyArray(lLoop)

                str2 = MyArray(lLoop2)

                MyArray(lLoop) = str2

                MyArray(lLoop2) = str1
            ElseIf (Len(MyArray(lLoop2)) = Len(MyArray(lLoop))) And (UCase(MyArray(lLoop2)) > UCase(MyArray(lLoop))) Then

                str1 = MyArray(lLoop)

                str2 = MyArray(lLoop2)

                MyArray(lLoop) = str2

                MyArray(lLoop2) = str1
            End If

        Next lLoop2
        
    Next lLoop
        
        
End Sub
Public Sub SortArrayDesc(ByRef MyArray())
    Dim lLoop As Long, lLoop2 As Long
    Dim str1 As String
    Dim str2 As String

    For lLoop = 0 To UBound(MyArray)

       For lLoop2 = lLoop To UBound(MyArray)

            If UCase(MyArray(lLoop2)) > UCase(MyArray(lLoop)) Then

                str1 = MyArray(lLoop)

                str2 = MyArray(lLoop2)

                MyArray(lLoop) = str2

                MyArray(lLoop2) = str1

            End If

        Next lLoop2
        
    Next lLoop
        
        
End Sub

Public Sub GetOtbGorbColors(ByRef OtbColors())
    ReDim OtbColors(44)
    Dim i As Integer
    i = 0
    OtbColors(i) = "ЖЕЛТ КАМ"
    i = i + 1
    OtbColors(i) = "СИЗ КАМ"
    i = i + 1
    OtbColors(i) = "НАКАРАДО"
    i = i + 1
    OtbColors(i) = "КАФЕЛЬ"
    i = i + 1
    OtbColors(i) = "КАМУШКИ"
    i = i + 1
    OtbColors(i) = "ЯСЕНЬ СВЕТЛЫЙ"
    i = i + 1
    OtbColors(i) = "ЯСЕНЬ ТЕМНЫЙ"
    i = i + 1
    OtbColors(i) = "КОРЕНЬ ГЛЯНЕЦ"
'    i = i + 1
'    OtbColors(i) = "ЛУННЫЙ КАМЕНЬ"
    i = i + 1
    OtbColors(i) = "СЕРАЯ КРОШКА"
'    i = i + 1
'    OtbColors(i) = "БЕЖ ГР ГЛЯНЕЦ"
    i = i + 1
    OtbColors(i) = "АЛЮМИНИЙ"
    i = i + 1
    OtbColors(i) = "ЧЕРН МРАМОР ГЛ"
    i = i + 1
    OtbColors(i) = "ОНИКС"
    i = i + 1
    OtbColors(i) = "ЯШМА"
    i = i + 1
    OtbColors(i) = "ИНЕЙ БЕЛЫЙ"
    i = i + 1
    OtbColors(i) = "ПЕСОЧНЫЙ ИНЕЙ"
    i = i + 1
    OtbColors(i) = "СЕРЫЙ ИНЕЙ"
    i = i + 1
    OtbColors(i) = "РЫЖИЙ ИНЕЙ"
    i = i + 1
    OtbColors(i) = "ТЕМНАЯ КРОШКА"
    i = i + 1
    OtbColors(i) = "АРКТИК"
    i = i + 1
    OtbColors(i) = "РАКУШЕЧНИК"
    i = i + 1
    OtbColors(i) = "ИЗВЕСТНЯК"
    i = i + 1
    OtbColors(i) = "КОРАЛЛ"
    i = i + 1
    OtbColors(i) = "КВАРЦ"
    i = i + 1
    OtbColors(i) = "ЧЕРНАЯ БРОНЗА"
    i = i + 1
    OtbColors(i) = "КРАСНЫЙ ИНЕЙ"
    i = i + 1
    OtbColors(i) = "СНОУ БЛЭК"
    i = i + 1
    OtbColors(i) = "СНОУ МИЛКИ"
    i = i + 1
    OtbColors(i) = "СНОУ УАЙТ"
    i = i + 1
    OtbColors(i) = "ГАЛИЦИА"
    i = i + 1
    OtbColors(i) = "ТУЯ"
    i = i + 1
    OtbColors(i) = "МРАМОР"
    i = i + 1
    OtbColors(i) = "БРЕШИА"
     i = i + 1
    OtbColors(i) = "БЕЖ ГР"
     i = i + 1
    OtbColors(i) = "ТУРИН ГРАНИТ"
     i = i + 1
    OtbColors(i) = "АРГИЛЛИТ БЕЛЫЙ"
     i = i + 1
    OtbColors(i) = "АРИЗОНА СЕРЫЙ"
     i = i + 1
    OtbColors(i) = "ВАЛЬМАСИНО"
     i = i + 1
    OtbColors(i) = "ГИАДА"
     i = i + 1
    OtbColors(i) = "ДАВОС ТРЮФЕЛЬ"
     i = i + 1
    OtbColors(i) = "КАНЗАС КОРИЧНЕВЫЙ"
     i = i + 1
    OtbColors(i) = "КАСЦИНА"
     i = i + 1
    OtbColors(i) = "КЕРАМИКА АНТРАЦИТ"
     i = i + 1
    OtbColors(i) = "ХРОМИКС АНТРАЦИТ"
     i = i + 1
    OtbColors(i) = "ХРОМИКС СЕРЕБРО"
    
    'SortArray OtbColors
End Sub


Public Function GetOtbColor(ByVal TTColor) As String
    If Not (IsEmpty(TTColor) Or IsNull(TTColor)) Then

        If InStr(1, TTColor, "карара", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "марок", vbTextCompare) > 0 Then
            GetOtbColor = "КАРАРА"
        ElseIf InStr(1, TTColor, "сиз", vbTextCompare) > 0 And _
                InStr(1, TTColor, "кам", vbTextCompare) > 0 Then
            GetOtbColor = "СИЗ КАМ"
        ElseIf InStr(1, TTColor, "жел", vbTextCompare) > 0 And _
                InStr(1, TTColor, "кам", vbTextCompare) > 0 Then
            GetOtbColor = "ЖЕЛТ КАМ"
        ElseIf InStr(1, TTColor, "беж", vbTextCompare) > 0 And _
                InStr(1, TTColor, "гр", vbTextCompare) > 0 Then
            GetOtbColor = "БЕЖ ГР"
        ElseIf InStr(1, TTColor, "рустик", vbTextCompare) > 0 Then
            GetOtbColor = "РУСТИК"
        ElseIf InStr(1, TTColor, "дуб", vbTextCompare) > 0 Then
            GetOtbColor = "ДУБ"
        ElseIf InStr(1, TTColor, "бук", vbTextCompare) > 0 Then
            GetOtbColor = "БУК"
        ElseIf InStr(1, TTColor, "дуб", vbTextCompare) > 0 Then
            GetOtbColor = "ДУБ"
        ElseIf InStr(1, TTColor, "ольха", vbTextCompare) > 0 Then
            GetOtbColor = "ОЛЬХА"
        ElseIf InStr(1, TTColor, "син", vbTextCompare) > 0 Then
            GetOtbColor = "СИНЯЯ"
        ElseIf InStr(1, TTColor, "зел", vbTextCompare) > 0 Then
            GetOtbColor = "ЗЕЛЕНАЯ"
        ElseIf InStr(1, TTColor, "пробк", vbTextCompare) > 0 Then
            GetOtbColor = "ПРОБКА"
        ElseIf InStr(1, TTColor, "крем", vbTextCompare) > 0 Then
            GetOtbColor = "КРЕМ"
        ElseIf InStr(1, TTColor, "махонь", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "гранит", vbTextCompare) > 0 Or _
                (InStr(1, TTColor, "крас", vbTextCompare) > 0 And InStr(1, TTColor, "гл", vbTextCompare) > 0) Then
            GetOtbColor = "ГРАНИТ"
        Else
            GetOtbColor = "ХРОМ"
        End If
    
    End If
End Function


Public Function getArchitehLength(Depth, karkas18 As Boolean) As Integer
     Dim is18karkas As Boolean
     Dim intDepth As Integer
    If IsMissing(Depth) Then intDepth = 0 Else intDepth = CInt(Depth)
     If intDepth < 130 Then intDepth = intDepth * 10
     
    If IsMissing(karkas18) Then is18karkas = False Else is18karkas = karkas18

    getArchitehLength = 0
    
    If is18karkas Then
        If intDepth >= 319 And intDepth < 518 Then
            getArchitehLength = 300
        ElseIf intDepth >= 518 Then
            getArchitehLength = 500
        End If
    Else
        If intDepth >= 303 And intDepth < 503 Then
            getArchitehLength = 300
        ElseIf intDepth >= 503 Then
            getArchitehLength = 500
        End If
    End If

End Function


Public Function GetDrawerMount() As Integer

Dim Depth As Integer
Depth = casepropertyCurrent.p_cabDepth

    GetDrawerMount = 0
    If (casepropertyCurrent.p_z_st_dsp Or casepropertyCurrent.p_dvpNahlest) Then
    
    If IsEmpty(Depth) Then
        GetDrawerMount = 50
    ElseIf Depth >= 504 Then
        GetDrawerMount = 50
    ElseIf Depth >= 454 Then
        GetDrawerMount = 45
    ElseIf Depth >= 404 Then
        GetDrawerMount = 40
    ElseIf Depth >= 354 Then
        GetDrawerMount = 35
    ElseIf Depth >= 304 Then
        GetDrawerMount = 30
    ElseIf Depth < 304 And Depth >= 254 Then 'If Depth >= 260 Then
        GetDrawerMount = 25
    End If
Else
     If IsEmpty(Depth) Then
        GetDrawerMount = 50
    ElseIf Depth >= 520 Then
        GetDrawerMount = 50
    ElseIf Depth >= 470 Then
        GetDrawerMount = 45
    ElseIf Depth >= 420 Then
        GetDrawerMount = 40
    ElseIf Depth >= 370 Then
        GetDrawerMount = 35
    ElseIf Depth >= 320 Then
        GetDrawerMount = 30
    ElseIf Depth < 320 And Depth >= 270 Then 'If Depth >= 260 Then
        GetDrawerMount = 25
    End If
    
   End If

End Function
Public Function GetDrawerMountTB() As String
Dim Depth As Integer
Depth = casepropertyCurrent.p_cabDepth
GetDrawerMountTB = ""
If Not (casepropertyCurrent.p_z_st_dsp Or casepropertyCurrent.p_dvpNahlest) Then
    If Depth >= 295 And Depth <= 330 Then
        GetDrawerMountTB = 260
    ElseIf Depth >= 331 And Depth <= 380 Then
        GetDrawerMountTB = 260
    ElseIf Depth >= 381 And Depth <= 450 Then
        GetDrawerMountTB = 350
    ElseIf Depth >= 451 And Depth <= 500 Then
        GetDrawerMountTB = 420
    ElseIf Depth >= 501 Then
        GetDrawerMountTB = 470
    Else
        GetDrawerMountTB = ""
    End If
Else
    If Depth >= 279 And Depth <= 314 Then
        GetDrawerMountTB = 260
    ElseIf Depth >= 316 And Depth <= 364 Then
        GetDrawerMountTB = 260
    ElseIf Depth >= 365 And Depth <= 434 Then
        GetDrawerMountTB = 350
    ElseIf Depth >= 435 And Depth <= 484 Then
        GetDrawerMountTB = 420
    ElseIf Depth >= 485 Then
        GetDrawerMountTB = 470
    Else
        GetDrawerMountTB = ""
    End If

End If
   
    

End Function
Public Function GetDrawerMountTB_vnutr_bol(Width, karkas18 As Boolean) As String
GetDrawerMountTB_vnutr_bol = ""
If karkas18 Then
    If Width = 40 Or Width = 400 Then
        GetDrawerMountTB_vnutr_bol = "400 (285/341)"
    ElseIf Width = 45 Or Width = 450 Then
        GetDrawerMountTB_vnutr_bol = "450 (435/391)"
    ElseIf Width = 50 Or Width = 500 Then
        GetDrawerMountTB_vnutr_bol = "500 (385/441)"
    ElseIf Width = 55 Or Width = 550 Then
        GetDrawerMountTB_vnutr_bol = "550 (435/491)"
    ElseIf Width = 60 Or Width = 600 Then
        GetDrawerMountTB_vnutr_bol = "600 (485/541)"
    ElseIf Width = 65 Or Width = 650 Then
        GetDrawerMountTB_vnutr_bol = "650 (535/591)"
    ElseIf Width = 70 Or Width = 700 Then
        GetDrawerMountTB_vnutr_bol = "700 (585/641)"
    ElseIf Width = 75 Or Width = 750 Then
        GetDrawerMountTB_vnutr_bol = "750 (635/691)"
    ElseIf Width = 80 Or Width = 800 Then
        GetDrawerMountTB_vnutr_bol = "800 (685/741)"
    ElseIf Width = 85 Or Width = 850 Then
        GetDrawerMountTB_vnutr_bol = "850 (735/791)"
    ElseIf Width = 90 Or Width = 900 Then
        GetDrawerMountTB_vnutr_bol = "900 (785/841)"
    ElseIf Width = 95 Or Width = 950 Then
        GetDrawerMountTB_vnutr_bol = "950 (835/891)"
    ElseIf Width = 100 Or Width = 1000 Then
        GetDrawerMountTB_vnutr_bol = "1000 (885/941)"
    ElseIf Width = 105 Or Width = 1050 Then
        GetDrawerMountTB_vnutr_bol = "1050 (935/991)"
    ElseIf Width = 110 Or Width = 1100 Then
        GetDrawerMountTB_vnutr_bol = "1100 (985/1041)"
    ElseIf Width = 115 Or Width = 1150 Then
        GetDrawerMountTB_vnutr_bol = "1150 (1035/1091)"
    ElseIf Width = 120 Or Width = 1200 Then
        GetDrawerMountTB_vnutr_bol = "1200 (1085/1141)"

    Else
        GetDrawerMountTB_vnutr_bol = ""
    End If
Else


    If Width = 40 Or Width = 400 Then
        GetDrawerMountTB_vnutr_bol = "400 (289/345)"
    ElseIf Width = 45 Or Width = 450 Then
        GetDrawerMountTB_vnutr_bol = "450 (339/395)"
    ElseIf Width = 50 Or Width = 500 Then
        GetDrawerMountTB_vnutr_bol = "500 (389/445)"
    ElseIf Width = 55 Or Width = 550 Then
        GetDrawerMountTB_vnutr_bol = "550 (439/495)"
    ElseIf Width = 60 Or Width = 600 Then
        GetDrawerMountTB_vnutr_bol = "600 (489/545)"
    ElseIf Width = 65 Or Width = 650 Then
        GetDrawerMountTB_vnutr_bol = "650 (539/595)"
    ElseIf Width = 70 Or Width = 700 Then
        GetDrawerMountTB_vnutr_bol = "700 (589/645)"
    ElseIf Width = 75 Or Width = 750 Then
        GetDrawerMountTB_vnutr_bol = "750 (639/695)"
    ElseIf Width = 80 Or Width = 800 Then
        GetDrawerMountTB_vnutr_bol = "800 (689/745)"
    ElseIf Width = 85 Or Width = 850 Then
        GetDrawerMountTB_vnutr_bol = "850 (739/795)"
    ElseIf Width = 90 Or Width = 900 Then
        GetDrawerMountTB_vnutr_bol = "900 (789/845)"
    ElseIf Width = 95 Or Width = 950 Then
        GetDrawerMountTB_vnutr_bol = "950 (839/895)"
    ElseIf Width = 100 Or Width = 1000 Then
        GetDrawerMountTB_vnutr_bol = "1000 (889/945)"
    ElseIf Width = 105 Or Width = 1050 Then
        GetDrawerMountTB_vnutr_bol = "1050 (939/995)"
    ElseIf Width = 110 Or Width = 1100 Then
        GetDrawerMountTB_vnutr_bol = "1100 (989/1045)"
    ElseIf Width = 115 Or Width = 1150 Then
        GetDrawerMountTB_vnutr_bol = "1150 (1039/1095)"
    ElseIf Width = 120 Or Width = 1200 Then
        GetDrawerMountTB_vnutr_bol = "1200 (1089/1145)"
    Else
        GetDrawerMountTB_vnutr_bol = ""
    End If


End If
   
    

End Function
Public Function GetDrawerMountTB_vnutr_mal(Width, karkas18 As Boolean) As String
GetDrawerMountTB_vnutr_mal = ""

If karkas18 Then
    If Width = 40 Or Width = 400 Then
        GetDrawerMountTB_vnutr_mal = "400 (285)"
    ElseIf Width = 45 Or Width = 450 Then
        GetDrawerMountTB_vnutr_mal = "450 (335)"
    ElseIf Width = 50 Or Width = 500 Then
        GetDrawerMountTB_vnutr_mal = "500 (385)"
    ElseIf Width = 55 Or Width = 550 Then
        GetDrawerMountTB_vnutr_mal = "550 (435)"
    ElseIf Width = 60 Or Width = 600 Then
        GetDrawerMountTB_vnutr_mal = "600 (485)"
    ElseIf Width = 65 Or Width = 650 Then
        GetDrawerMountTB_vnutr_mal = "650 (535)"
    ElseIf Width = 70 Or Width = 700 Then
        GetDrawerMountTB_vnutr_mal = "700 (585)"
    ElseIf Width = 75 Or Width = 750 Then
        GetDrawerMountTB_vnutr_mal = "750 (635)"
    ElseIf Width = 80 Or Width = 800 Then
        GetDrawerMountTB_vnutr_mal = "800 (685)"
    ElseIf Width = 85 Or Width = 850 Then
        GetDrawerMountTB_vnutr_mal = "850 (735)"
    ElseIf Width = 90 Or Width = 900 Then
        GetDrawerMountTB_vnutr_mal = "900 (785)"
    ElseIf Width = 95 Or Width = 950 Then
        GetDrawerMountTB_vnutr_mal = "950 (835)"
    ElseIf Width = 100 Or Width = 1000 Then
        GetDrawerMountTB_vnutr_mal = "1000 (885)"
    ElseIf Width = 105 Or Width = 1050 Then
        GetDrawerMountTB_vnutr_mal = "1050 (935)"
    ElseIf Width = 110 Or Width = 1100 Then
        GetDrawerMountTB_vnutr_mal = "1100 (985)"
    ElseIf Width = 115 Or Width = 1150 Then
        GetDrawerMountTB_vnutr_mal = "1150 (1035)"
    ElseIf Width = 120 Or Width = 1200 Then
        GetDrawerMountTB_vnutr_mal = "1200 (1085)"
    Else
        GetDrawerMountTB_vnutr_mal = ""
    End If
Else


    If Width = 40 Or Width = 400 Then
        GetDrawerMountTB_vnutr_mal = "400 (289)"
    ElseIf Width = 45 Or Width = 450 Then
        GetDrawerMountTB_vnutr_mal = "450 (339)"
    ElseIf Width = 50 Or Width = 500 Then
        GetDrawerMountTB_vnutr_mal = "500 (389)"
    ElseIf Width = 55 Or Width = 550 Then
        GetDrawerMountTB_vnutr_mal = "550 (439)"
    ElseIf Width = 60 Or Width = 600 Then
        GetDrawerMountTB_vnutr_mal = "600 (489)"
    ElseIf Width = 65 Or Width = 650 Then
        GetDrawerMountTB_vnutr_mal = "650 (539)"
    ElseIf Width = 70 Or Width = 700 Then
        GetDrawerMountTB_vnutr_mal = "700 (589)"
    ElseIf Width = 75 Or Width = 750 Then
        GetDrawerMountTB_vnutr_mal = "750 (639)"
    ElseIf Width = 80 Or Width = 800 Then
        GetDrawerMountTB_vnutr_mal = "800 (689)"
    ElseIf Width = 85 Or Width = 850 Then
        GetDrawerMountTB_vnutr_mal = "850 (739)"
    ElseIf Width = 90 Or Width = 900 Then
        GetDrawerMountTB_vnutr_mal = "900 (789)"
    ElseIf Width = 95 Or Width = 950 Then
        GetDrawerMountTB_vnutr_mal = "950 (839)"
    ElseIf Width = 100 Or Width = 1000 Then
        GetDrawerMountTB_vnutr_mal = "1000 (889)"
    ElseIf Width = 105 Or Width = 1050 Then
        GetDrawerMountTB_vnutr_mal = "1050 (939)"
    ElseIf Width = 110 Or Width = 1100 Then
        GetDrawerMountTB_vnutr_mal = "1100 (989)"
    ElseIf Width = 115 Or Width = 1150 Then
        GetDrawerMountTB_vnutr_mal = "1150 (1039)"
    ElseIf Width = 120 Or Width = 1200 Then
        GetDrawerMountTB_vnutr_mal = "1200 (1089)"
    Else
        GetDrawerMountTB_vnutr_mal = ""
    End If


End If
   
    

End Function
Public Function GetDrawerMountMb() As String
GetDrawerMountMb = ""
Dim Depth As Integer
Depth = casepropertyCurrent.p_cabDepth
If Not (casepropertyCurrent.p_z_st_dsp Or casepropertyCurrent.p_dvpNahlest) Then
    If Depth >= 471 And Depth <= 520 Then
     GetDrawerMountMb = 450
    ElseIf Depth >= 521 And Depth <= 1000 Then
     GetDrawerMountMb = 500
    End If
Else
    If Depth >= 455 And Depth <= 504 Then
     GetDrawerMountMb = 450
    ElseIf Depth >= 505 And Depth <= 1000 Then
     GetDrawerMountMb = 500
    End If
End If

End Function
Public Function GetDrawerMountKv() As Integer
Dim Depth As Integer
Depth = casepropertyCurrent.p_cabDepth
GetDrawerMountKv = 0
   If Not (casepropertyCurrent.p_z_st_dsp Or casepropertyCurrent.p_dvpNahlest) Then
        If IsEmpty(Depth) Then
            GetDrawerMountKv = 50
        ElseIf Depth >= 530 Then
            GetDrawerMountKv = 50
        ElseIf Depth >= 480 And Depth < 530 Then
            GetDrawerMountKv = 45
        ElseIf Depth >= 430 And Depth < 480 Then
            GetDrawerMountKv = 40
        ElseIf Depth >= 380 And Depth < 430 Then
            GetDrawerMountKv = 35
        ElseIf Depth >= 330 And Depth < 380 Then
        GetDrawerMountKv = 30
        ElseIf Depth >= 280 And Depth < 330 Then
        GetDrawerMountKv = 25
        End If
    Else
        If IsEmpty(Depth) Then
            GetDrawerMountKv = 50
        ElseIf Depth >= 514 Then
            GetDrawerMountKv = 50
        ElseIf Depth >= 464 And Depth < 514 Then
            GetDrawerMountKv = 45
        ElseIf Depth >= 414 And Depth < 464 Then
            GetDrawerMountKv = 40
        ElseIf Depth >= 364 And Depth < 414 Then
            GetDrawerMountKv = 35
        ElseIf Depth >= 314 And Depth < 364 Then
        GetDrawerMountKv = 30
        ElseIf Depth >= 264 And Depth < 314 Then
        GetDrawerMountKv = 25
        End If
  End If

'    ElseIf Depth > 405 Then
'        GetDrawerMountKv = 40
'    ElseIf Depth > 355 Then
'        GetDrawerMountKv = 35
'    ElseIf Depth > 305 Then
'        GetDrawerMountKv = 30
'    Else 'If Depth >= 260 Then
'        GetDrawerMount = 25
  
End Function

Public Function GetHandleExtra(ByVal Handle) As Integer
    Dim he
    
    If IsNull(Handle) Then
        he = 0
        
    Else
     
        If Not IsNull(Handle) And Not IsEmpty(Handle) Then
            If InStr(1, Handle, "модена", vbTextCompare) Or Handle = "верона" Then
                he = 0
            Else
        
                Init_rsHandle
                
                If rsHandle.RecordCount > 0 Then rsHandle.MoveFirst
                rsHandle.Find "Handle='" & Handle & "'"
                If Not rsHandle.EOF Then
                    '03-10-11 исп только для стенок
                    'If rsHandle!Drilling >= 160 Then he = 0 Else he = 1
                    If rsHandle!Drilling > 160 Then he = 0 Else he = 1
                End If
            
            End If
        End If
        
        If IsEmpty(he) Then
            If MsgBox("Две ручки на шуфляду?", vbDefaultButton3 Or vbQuestion Or vbYesNo, "Ручки на шуфляду") = vbYes Then
                he = 1
            Else
                he = 0
            End If
        End If
    End If
    
    GetHandleExtra = he
End Function


Public Function CheckHandleExtra(ByVal Handle) As Variant
    CheckHandleExtra = Empty

    If Not IsNull(Handle) Then
        If InStr(1, Replace(ActiveCell.Offset(, 2), " ", ""), "на1руч", vbTextCompare) Or _
            InStr(1, Replace(ActiveCell.Offset(, 3), " ", ""), "на1руч", vbTextCompare) > 0 Or _
            InStr(1, Replace(ActiveCell.Offset(, 4), " ", ""), "на1руч", vbTextCompare) > 0 Or InStr(1, Replace(ActiveCell.Offset(, 5), " ", ""), "на1руч", vbTextCompare) > 0 Then
            
            CheckHandleExtra = 0
        ElseIf InStr(1, Replace(ActiveCell.Offset(, 2), " ", ""), "на2руч", vbTextCompare) > 0 Or _
                InStr(1, Replace(ActiveCell.Offset(, 3), " ", ""), "на2руч", vbTextCompare) > 0 Or _
                InStr(1, Replace(ActiveCell.Offset(, 4), " ", ""), "на2руч", vbTextCompare) > 0 Or InStr(1, Replace(ActiveCell.Offset(, 5), " ", ""), "на2руч", vbTextCompare) > 0 Then
            
            CheckHandleExtra = 1
        End If
    End If
End Function


Sub Задания_и_отгрузки()
    MainForm.Show
End Sub
Public Function getCaseIdbyOCID(ByVal ocid As Long)
    Dim comm As ADODB.Command
    Set comm = New ADODB.Command
    comm.ActiveConnection = GetConnection
    comm.CommandType = adCmdText
    comm.CommandText = "SELECT TOP 1 CaseId FROM [Fittings].[dbo].[OrderCases] where OCID=? order by OCID asc"
    comm.Parameters(0) = ocid
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.Open comm, , adOpenDynamic, adLockReadOnly
    
    
    If rs.RecordCount > 0 Then
        
        rs.MoveFirst
         getCaseIdbyOCID = CLng(rs(0))
    Else
        getCaseIdbyOCID = 0
    End If
    
    rs.Close
End Function

Public Function GetColorID(ByRef ColorName, ByRef BibbColor, ByRef CamBibbColor) As Integer
    If IsNull(ColorName) Then ColorName = ""
    If Not rsColor Is Nothing Then
    
        rsColor.MoveFirst
        rsColor.Find "parserName='" & ColorName & "'", , adSearchForward, 1
        If rsColor.EOF Then
            rsColor.MoveFirst
            rsColor.Find "parserSecondName like '" & ColorName & "'", , adSearchForward, 1
        End If
        If rsColor.EOF Then
            rsColor.MoveFirst
            rsColor.Find "parserTripleName like '" & ColorName & "'", , adSearchForward, 1
        End If
        If rsColor.EOF Then
            rsColor.MoveFirst
            rsColor.Find "ColorName like '" & ColorName & "'", , adSearchForward, 1
        End If
        If rsColor.EOF Then
            GetColorID = 0
        End If
        
        If rsColor.EOF = False Then
            GetColorID = rsColor!ColorId
            ColorName = rsColor!ColorName
            BibbColor = rsColor!BibbColor
            CamBibbColor = rsColor!CamBibbColor
        Else
             'подправим
            ColorName = Replace(ColorName, "клён", "клен")
            ColorName = Replace(ColorName, "белое", "белый")
            ColorName = Replace(ColorName, "белые", "белый")
            ColorName = Replace(ColorName, "-16мм", "")
            ColorName = Replace(ColorName, "-18мм", " 18")
            ColorName = Replace(ColorName, "-16мм", "")
            ColorName = Replace(ColorName, " 18мм", " 18")
            ColorName = Replace(ColorName, "-18", " 18")
            ColorName = Replace(ColorName, "-16", "")
            ColorName = Replace(ColorName, " 16", "")
            
            If ColorName = "рустик" Then ColorName = "рустикаль"
        
            rsColor.MoveFirst
            rsColor.Find "parserName like '" & ColorName & "'", , adSearchForward, 1
            If rsColor.EOF Then
                rsColor.MoveFirst
                rsColor.Find "parserSecondName like '" & ColorName & "'", , adSearchForward, 1
            End If
            If rsColor.EOF Then
            rsColor.MoveFirst
            rsColor.Find "parserTripleName like '" & ColorName & "'", , adSearchForward, 1
            End If
            If rsColor.EOF Then
                rsColor.MoveFirst
                rsColor.Find "ColorName like '" & ColorName & "'", , adSearchForward, 1
            End If
            
            If rsColor.EOF = False Then
                GetColorID = rsColor!ColorId
                ColorName = rsColor!ColorName
                BibbColor = rsColor!BibbColor
                CamBibbColor = rsColor!CamBibbColor
            Else
                GetColorID = 0
            End If
        End If
        
       If BibbColor = "" Then BibbColor = Empty
       If CamBibbColor = "" Then CamBibbColor = Empty
'
'        rsColor.Find "ColorName='" & ColorName & "'", , adSearchForward, 1
'        If rsColor.EOF Then GetColorID = 0 Else GetColorID = rsColor!colorid
    Else
        GetColorID = 0
        BibbColor = Empty
        CamBibbColor = Empty
    End If
End Function


