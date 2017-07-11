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

Public Const cHandle As String = "����"
Public Const cLeg As String = "����"
Public Const cPlank As String = "�����"
Public Const cGalog As String = "�����"
Public Const cSink As String = "�����"
Public Const cStol As String = "���� " '!
Public Const cStul As String = "����"
Public Const cStool As String = "�������"
Public Const cNogi As String = "����"
Public Const cSit As String = "�������"


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


'����������

Public FittingArray(), HandleArray(), LegArray()
Public OtbColors(), Doormount(), Plank(), Galog(), Rell(), ����(), ��������(), ��������(), �������(), �����(), �����()
Public OtbGorbColors(), vytyazhka_perfim()
Public zavesHL(), zavesHS(), zavesSensys(), zavesClipTop(), ploschadkaSensys()
Public tbLength()
Public tbkovrLength()
Public tbkovrOpt()
' ���������
'Private Stul(), SitK(), ������(), ������(), �����()
'Private  Stol(),  Sink(), Sit(), SitColors(), SitKolib(), BackKolib()
Public StulNogi(), SW_bel(), SW(), LW(), PA(), �������(), ������(), �����������������(), ������������(), ���������(), ������() ', ��������()
Public Sushk(), ��������������(), ��������������(), TOPLine(), ���������4�()
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
        If InStr(1, Handle, "������", vbTextCompare) = 0 And _
             InStr(1, Handle, "������", vbTextCompare) = 0 And _
             InStr(1, Handle, "����", vbTextCompare) = 0 Then
            Dim k As Integer
            k = InStr(1, Handle, " ��", vbTextCompare)
            If k > 1 Then Handle = Trim(Left(Handle, k))
            
            Handle = Replace(Handle, "/", "")
            Handle = Replace(Handle, " ", "")
            Handle = Replace(Handle, ".", "")
            Handle = Replace(Handle, "-", "")
        End If
        
        If rsHandle.RecordCount > 0 Then rsHandle.MoveFirst
        rsHandle.Find "Handle='" & Handle & "'"
        If rsHandle.EOF Then
            'MsgBox "����������� ��� ����� - " & Handle, vbCritical
            'Handle = InputBox("������� ��� �����", "����� ������ �� ���������", Handle)
            
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
            
'            If Handle = "�������" Or Handle = "-" Then
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
        If InStr(1, Leg, "������� ����", vbTextCompare) = 0 And _
            InStr(1, Leg, "780 ����", vbTextCompare) = 0 And _
            InStr(1, Leg, "780 ������", vbTextCompare) = 0 And _
            InStr(1, Leg, "����", vbTextCompare) = 0 And _
            InStr(1, Leg, "����", vbTextCompare) = 0 Then _
            Leg = Replace(Leg, " ", "")
        
        If rsLeg.RecordCount > 0 Then rsLeg.MoveFirst
        rsLeg.Find "Leg='" & Leg & "'"
        If rsLeg.EOF Then
            MsgBox "����������� ��� ����� - " & Leg, vbCritical
            Leg = InputBox("������� ��� �����", "����� ������ �� ���������", Leg)
            
            If Leg = "-" Then Leg = "�������"
            If Leg = "�������" Then
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
        If InStr(1, face, "������", vbTextCompare) > 0 Then face = ""
        UpdateComm.Parameters("@Face") = Left(face, 50)
        
    End If
    If Not IsMissing(CaseColor) Then UpdateComm.Parameters("@CColor") = Left(CaseColor, 20)
    If Not IsMissing(ColorId) Then
    UpdateComm.Parameters("@ColorId") = ColorId
      End If
    UpdateComm.Execute
    
    Exit Sub
err_UpdateOrder:
    MsgBox Error, vbCritical, "���������� ������ (UpdateOrder)"
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
        comm.CommandText = "SELECT  * FROM [���������������]"
        
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
'        comm.CommandText = "SELECT * FROM [������������]"
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
        comm.CommandText = "SELECT * FROM [����������������]"
        
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
    If InStr(1, casename, "��", vbTextCompare) > 0 And _
        InStr(1, casename, "/", vbTextCompare) = 0 Then
        
        casename = InputBox("��������� ������������", "���� ������?", casename)
        
    End If
                        
On Error GoTo err_ParseCase
    
'    If regexp_check(patNewName, casename) Then
'        casename = parse_case(casename)
'    ElseIf regexp_check(patSHLK_check1, casename) Then
'    casename = regexp_replace(patSHLK_check1, casename, "/")
'
'    End If
    If regexp_check(patSHLK_check1, casename) Then casename = regexp_replace(patSHLK_check1, casename, "/")
    
    If casename = "���60/60" Then
        casename = "���60"
    ElseIf casename = "���60/60(600)" Then
        casename = "���60(600)"
    ElseIf casename = "���60/60(915)" Then
        casename = "���60(915)"
  '  ElseIf CaseName = "��60(915)�/2" Then
   '     CaseName = "��60�/2(915)"
    ElseIf casename = "���90/90" Then
        casename = "���90"
    ElseIf casename = "��60/2��" Then
        casename = "��60��"
    ElseIf casename = "�����60" Then
        casename = "����60"
    ElseIf casename = "���90/90�" Then
        casename = "���90�"
    ElseIf casename = "���90/90" Then
        casename = "���90"
    ElseIf casename = "���90/90�" Then
        casename = "���90�"
    ElseIf Replace(Replace(Replace(casename, " ", ""), ".", ""), "/90", "") = "���90/90��������" Then
        casename = "����90/��"
    ElseIf Replace(Replace(Replace(casename, " ", ""), ".", ""), "/90", "") = "���90/90�����" Then
        casename = "���90/��"
    ElseIf Replace(Replace(Replace(casename, " ", ""), ".", ""), "/90", "") = "���90/90������" Then
        casename = "����90/��"
    ElseIf Replace(casename, " ", "") = "���30����" Then
        casename = "��30����"
    ElseIf Replace(casename, " ", "") = "���" Then
        casename = "���90"
    ElseIf Replace(casename, " ", "") = "���" Then
        casename = "���90"
    ElseIf Replace(casename, " ", "") = "���" Then
        casename = "���60"
    ElseIf Replace(casename, " ", "") = "���(600)" Then
        casename = "���60(600)"
    ElseIf Replace(casename, " ", "") = "���(915)" Then
        casename = "���60(915)"
    ElseIf Replace(casename, " ", "") = "����" Then
        casename = "���90�"
    ElseIf Replace(casename, " ", "") = "����" Then
        casename = "���90�"
    End If
    
    If Left(casename, 7) = "�����65/65" Then
        casename = Replace(casename, "�����65/65", "����65/65", 1, 1, vbTextCompare)
    End If
    
    If Left(casename, 3) = "���" Then
        casename = Replace(casename, "���", "���", 1, 1, vbTextCompare)
    End If
    If Left(casename, 3) = "���" Then
        casename = Replace(casename, "���", "��", 1, 1, vbTextCompare)
    End If
    If InStr(1, casename, "����60/60", vbTextCompare) = 1 Then
        casename = Replace(casename, "����60/60", "����60", 1, 1, vbTextCompare)
    End If
    If InStr(1, casename, "���60/60", vbTextCompare) = 1 Then
        casename = Replace(casename, "���60/60", "���60", 1, 1, vbTextCompare)
    End If
    If Left(casename, 4) = "����" Then
        casename = Replace(casename, "����", "���", 1, 1, vbTextCompare)

        Set caseFur = New caseFurniture
        caseFur.init
        caseFur.fName = "����������"
        caseFur.qty = 6
        caseFurnCollection.Add caseFur
    
        Set caseFur = New caseFurniture
        caseFur.init
        caseFur.fName = "�����"
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

    k = InStr(1, casename, "�/�", vbTextCompare)
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
    
    casename = Trim(Replace(casename, "������", "��", , , vbTextCompare))
    
    
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
                
                Case "�", "�", "�", "�"
                    If Len(casename) > k Then
                    
                        Select Case Mid(casename, k, 4)
                            Case "��-�", "��-�", "��-1", "��-1", "��-1", "2���"
                                SVar = SVar & Mid(casename, k, 4)
                        
                                If k > 1 Then
                                    casename = Left(casename, k - 1) & Mid(casename, k + 4)
                                Else
                                    casename = Mid(casename, k + 4)
                                End If
                                
                                Exit Do
                        End Select
                        
                        Select Case Mid(casename, k, 3)
                            Case "���", "���", "���", "���", "���", "���", "���", "���"
                                SVar = SVar & Mid(casename, k, 3)
                                
                                If k > 1 Then
                                    casename = Left(casename, k - 1) & Mid(casename, k + 3)
                                Else
                                    casename = Mid(casename, k + 3)
                                End If
                                
                                Exit Do
                            
                        End Select
                        
                        Select Case Mid(casename, k, 2)
                            Case "��", "��", "��"
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
                Case "���", "���", "���", "���", "���"
                    SVar = "/" & Left(casename, 3)
                    If Len(casename) > 3 Then casename = Mid(casename, 4) Else casename = ""
                Case Else
                    Select Case Left(casename, 2)
                    Case "��", "��"
                        SVar = "/" & Left(casename, 2)
                        If Len(casename) > 2 Then casename = Mid(casename, 3) Else casename = ""
                    End Select
            End Select
        End If
    End If
    If SVar = "/" Then SVar = Empty
        
        
    casename = Trim(casename)
    k = InStr(1, casename, "����.")
    
    If k Then
        casename = Left(casename, k - 1) & LTrim(Mid(casename, k + 5))
        casename = LTrim(casename)
        
    Else
        k = InStr(1, casename, "����")
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
        
        L = InStr(1, casename, "��")
        If L Then casename = Left(casename, L - 1) & Mid(casename, L + 2)
    End If
    If IsEmpty(D) Then
        If SVar = "\��" Then
            SVar = Empty
            D = 570 '530
        ElseIf InStr(1, casename, "����") Then
            D = 570 '530
        End If
        
        If InStr(1, casepropertyCurrent.p_fullcn, "�", vbTextCompare) = 2 Then
            D = 300
            Else
            D = 570
        End If
        
    End If
   ' casepropertyCurrent.p_cabDepth = CInt(D)
    casename = Trim(casename)
    
    
   
    ' �������� ���-�� ������� � ������
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
'        If l = 0 Then '���� ��� ����������� ������
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
            If L = 0 Then '���� ��� ����������� ������
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
            
            
            
            '��������� ������ ������
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
                
                ts = Replace(ts, "�����", "", , , vbTextCompare)
                ts = Replace(ts, "����", "", , , vbTextCompare)
                ts = Replace(ts, "����", "", , , vbTextCompare)
                    
                '��������, �� ������� �� ����/������/���� ����� �
                L = InStr(ts, "�") '��� �
                If L = 0 Then L = InStr(ts, "x") ' ���� x
                If L = 0 Then L = InStr(ts, "/") ' ������ /
                If L = 0 Then L = InStr(ts, "*") ' ������ *
                If L > 0 Then L = CInt(Val(LTrim(Mid(ts, L + 1)))) Else L = 1 '���-�� �������
                While L > 4
                    L = InputBox("������� ���-�� ������� (" & ts & ") �� ������", "���-�� �������", L)
                Wend
                
                If InStr(1, ts, "�", vbTextCompare) = 0 Then
                    KF = KF + L
                    
                    If InStr(1, ts, "�", vbTextCompare) > 0 Or L > 1 Then
                        SHQty = SHQty + L
                    Else
                        fQty = fQty + L
                    End If
                    
                    'If Nisha And KF > 0 Then KF = KF - 1
                    Nisha = False
                Else
                    '���� ���� ������ ��� ����� ����, ��������� � ������ ������ ������ 16 (��� ��*, ��*, ��*)
                    If H = 0 Or Nisha Then H = H + 16 Else KF = KF + 1
                    H = H + 16 * (L - 1) '����  ���� ������� �/� �
                    Nisha = True
                    NQty = NQty + 1
                End If
                
                If Not Nisha And InStr(1, ts, "���", vbTextCompare) > 0 Then
                    WindowQty = WindowQty + L
                End If
                
                '��������� � ������ ����� ������ ����/������
                While Not IsNumeric(Left(ts, 1)) And Trim(ts) <> ""
                    ts = Mid(ts, 2)
                Wend
                H = H + CInt(Val(ts)) * L
            Wend
        
            '���� ���� �����, ��������� � ������ ������ ������ 16 (��� ?�*, ?�*, �� �� ?�*,�.�. � ?�* ����� ������)
            If Nisha And Mid(name, 2, 1) <> "�" Then H = H + 16
            '��������� �����
            If Not Nisha Then
                H = H + (KF + 1) * 3
            Else
                H = H + KF * 3
            End If
            
            Select Case Mid(name, 2, 1)
                Case "�"
                    H = H + 97
                Case "�"
                    H = H - 16
                Case "�"
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
    
    If InStr(1, H, "��", vbTextCompare) Then
        H = Replace(H, "��", "", 1, 1, vbTextCompare)
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
                Case "���", "���", "���", "���", "���"
                    SVar = "/" & Left(casename, 3)
                    If Len(casename) > 3 Then casename = Mid(casename, 4) Else casename = ""
                Case Else
                    Select Case Left(casename, 2)
                    Case "��", "��"
                        SVar = "/" & Left(casename, 2)
                        If Len(casename) > 2 Then casename = Mid(casename, 3) Else casename = ""
                    End Select
            End Select
        End If
    End If
    If SVar = "/" Then SVar = Empty
    
    If Len(casename) > 0 Then
        If Asc(Left(casename, 1)) = 203 Then casename = Trim(Mid(casename, 2))
        
        If Len(casename) > 0 And (InStr(1, casename, "���", vbTextCompare) <> 1) Then
            Select Case Asc(Left(casename, 1))
                Case 200 '"�"
                    LVar = "�"
                    casename = Trim(Mid(casename, 2))
                Case 210 '"�"
                    LVar = "�"
                    casename = Trim(Mid(casename, 2))
                Case 199 '"�"
                    LVar = "�"
                    casename = Trim(Mid(casename, 2))
                Case 192 '"�"
                    If InStr(1, casename, "����", vbTextCompare) = 0 Then
                        LVar = "�"
                        casename = Trim(Mid(casename, 2))
                    End If
            End Select
        End If
    End If
   
    If LVar = "�" And Not IsEmpty(SVar) Then
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
    
    
    
    If name = "���" Then
        SVar = SVar & "�"
        name = "��"
    End If
    
    Select Case Left(name, 1)
      
        Case "�"
    
            If Not IsEmpty(H) Then
                If H > 820 Then
                    ShelfQty = 2
                ElseIf H < 500 Then
                    ShelfQty = 0
                End If
            End If
        
            Select Case Mid(name, 2, 1)
                Case "�"
                    Select Case name
                    
                    
                    Case "��10 ��� ���4 (�.1451) ��.570��", _
                     "��10 ��� ���4 (�.820) ��.483��", _
                     "��10 ��� ���4 (�.1442) ��.492��", _
                     "��10 ��� ���4 (�.718) ��.492��", "����10"



                        Case "��", "���"
                            If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                            If casepropertyCurrent.p_DoorCount > 0 Then DoorQty = casepropertyCurrent.p_DoorCount
                            Select Case SVar
                                
                                Case ""
                                    If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "����") = 0 Then name = name & "915"

                                If Check_��(casename) Then
                                    name = name & " ��"
                                    Doormount = "+20"
                                End If
                            
'                                Case Empty
'                                    If IsEmpty(LVar) Then
'                                        If W <= 50 Then DoorQty = 1 Else DoorQty = 2
'                                    Else
'                                        Do
'                                            DoorQty = InputBox("������� ���-�� �������", "���-�� �������")
'                                        Loop Until IsNumeric(DoorQty)
'                                    End If
                                Case "/1"
                                    If InStr(casename, "����") > 0 And Check_��(casename) Then
                                        name = name & "/1 ��"
                                        Doormount = "110"
                                        DoorQty = InputBox("�������� ���-�� ������", "���-�� ������ �����", DoorQty)
                                    Else
                                    
                                        If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "����") = 0 Then name = name & "915"
                                        DoorQty = 1
                                    End If
                                Case "/2"
                                
                                    If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "����") = 0 Then name = name & "915"
                                    
                                    DoorQty = 2
                                    
                                    If InStr(1, casename, "HF") > 0 Then Doormount = Null
                                    If InStr(1, casename, "HK") > 0 Then Doormount = Null
                                    
                                Case "/1�"
                                    DoorQty = 1
                                    name = name & " �"
                                Case "/2�"
                                    DoorQty = 2
                                    
                                    If InStr(1, Trim(casename), "HF") > 0 Then
                                        Doormount = Null
                                        name = name & " �"
                                        'If DoorQty = 2 Then DoorQty = 1
                                        DoorQty = 0
                                    ElseIf InStr(1, Trim(casename), "FB-1") > 0 Then
                                        Doormount = Null
                                        name = name & " �"
                                    ElseIf InStr(1, Trim(casename), "AV") > 0 Then
                                        Doormount = Null
                                        name = name & " �"
                                    Else
                                        name = name & " 2�"
                                    End If
                                
                                    
                                Case Else
                                
                                    If ShelfQty >= 2 And Not FrezBase And LVar = "" And InStr(casename, "����") = 0 Then name = name & "915"
                                
                                    If InStr(1, Trim(casename), "HK") > 0 Then ' + HK-S
                                        Doormount = Null
                                    End If
                            End Select
                            
                            ' ��� ���� � ������
                            Select Case LVar
                                Case Empty
                                Case "�"
                                    name = "��� �"
                                    DoorQty = 2
                                    WindowQty = 3
                                    LVar = Empty
                                Case "�"
                                    name = "��� �"
                                    DoorQty = 0
                                    'DoorMount = "�-� � ������"
                                    LVar = Empty
                                Case "�"
                                    'ShelfQty = 1  ���� 2- �� 2
                                    name = name & " �"
                                    'DoorQty = 1
                                Case Else
                                    name = ""
                                    MsgBox "����������� ����", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                            If FrezBase Then name = name & " ��"
                            
                            If InStr(casename, "����") Then
                                name = name & " ����"
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
                            
                        
                        Case "���"
                       
                            If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                            
                             If Not IsEmpty(H) Then
                                If H >= 500 Then
                                    ShelfQty = 2 ' ����� 2 ������ +1, �.�. 1 �����
                                End If
                            End If
                            
                            Select Case SVar
                                Case Empty
                                    If Check_��(casename) Then
                                        name = name & " ��"
                                        Doormount = "+20"
                                    End If
                                Case "/1"
                                    DoorQty = 1
                                                                   
                                Case "/2"
                                    If Check_��(casename) Then
                                        name = name & " ��"
                                        Doormount = "+20"
                                    End If
                                    DoorQty = 2
                                Case Else
                                    name = ""
                                    MsgBox "����������� ����", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select

                        Case "���"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 0
                        Case "���"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                            Doormount = "FGV45"
                        Case "���"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                        Case "����"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                        Case "����"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 1
                            Doormount = "��������"
                        Case "����"
                            If ShelfQty >= 2 Then name = name & "915"
                            DoorQty = 2
                            Doormount = "175"
                        Case "����"
                            If ShelfQty >= 2 Then
                                name = name & "915"
                                ShelfQty = Empty
                            End If
                        
                        Case Else
                        
                            name = ""
                            MsgBox "����������� ����", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select ' �-�-Name
                    
                Case "�", "�"
                    Select Case name
                        Case "��", "��", "���", "���", "���"
                            If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                            
                            
                            
                            
                            If InStr(casename, "����") Then
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
                                    If Check_��(casename) Then
                                        name = name & " ��"
                                        Doormount = "+20"
                                    End If
                                Case "/1�"
                                    DoorQty = 1
                                    name = name & SVar
                                Case "/1"
                                    DoorQty = 1
                                Case "/2"
                                    DoorQty = 2
                                Case "/�"
                                    'ShelfQty = 0
                                    name = name & SVar
                                    DoorQty = 1
                                Case "/���", "/2���"
                                    name = name & SVar
                                    DoorQty = 0
                                    Drawermount = ""
                                Case Else
                                    name = ""
                                    MsgBox "����������� ����", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                            Select Case LVar
                                Case Empty
                                Case "�"
                                    name = name & " �"
                                    LVar = Empty
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                Case "�"
                                    'DoorQty = 1
                                    name = name & " �"
                                    LVar = Empty
                                Case Else
                                    name = ""
                                    MsgBox "����������� ����", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                            If InStr(casename, "����") Then
                                name = name & " ����"
                                
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
                            
                        Case "���"
                            DoorQty = 0
                            If IsEmpty(D) Then D = 570 '530
                        Case "���"
                            DoorQty = 0
                            If IsEmpty(D) Then D = 530 '480
                        Case "����", "����"
                            If SVar = "/���" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/���" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/���" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/���" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/��" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf SVar = "/���" Then
                                name = name & SVar
                                Drawermount = "500/78 ����"
                            ElseIf SVar = "/��" Then
                                name = name & SVar
                                Drawermount = ""
                            ElseIf InStr(casepropertyCurrent.p_fullcn, "��") > 5 Then
                                name = name & "/��"
                                If IsEmpty(D) Then D = 570 '530
                                Drawermount = CStr(GetDrawerMountKv())
                            ElseIf InStr(casepropertyCurrent.p_fullcn, "���") > 5 Then
                                name = name & "/����"
                                If IsEmpty(D) Then D = 570 '530
                                Drawermount = "����� " & CStr(GetDrawerMount())
                            ElseIf InStr(casepropertyCurrent.p_fullcn, "���") > 5 Then
                                name = name & "/����"
                                If IsEmpty(D) Then D = 570 '530
                                Drawermount = "����� " & CStr(GetDrawerMount())
                            ElseIf (InStr(casepropertyCurrent.p_fullcn, "�����") > 5 _
                                Or InStr(casepropertyCurrent.p_fullcn, "org") > 5 _
                                Or InStr(casepropertyCurrent.p_fullcn, "�����") > 5) Then
                                name = name & "/���"
                                Drawermount = ""
                            ElseIf IsEmpty(D) Then
                                    D = 570 '530
                                    Drawermount = GetDrawerMount()
                            End If
                            ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                            DoorQty = 1
                        Case "���", "���"
                            DoorQty = 2
                            Doormount = "175"
                            'Doormount = "FGV180"
                           name = name & SVar
                        Case "���", "���"
                            DoorQty = 0
                        Case "����", "����"
                            DoorQty = 1
                            Doormount = "��������"
                        Case "����", "����"
                            DoorQty = 1
                            Doormount = "FGV45"
                            'Name = Name & SVar
                        Case "����", "����"
                            DoorQty = 1
                            name = name & SVar
                        Case "���", "���"
                            DoorQty = 2
                        Case "���", "���"
                        
                            If IsEmpty(D) Then
                                If is18(CaseColor) Then
                                    
                                    D = 570 '530
                                    
                                End If
                            End If
                        
                        
                            Select Case SVar
                                Case "/1", "/1��"
                                    name = name & "/1"
                                    DoorQty = 1
                                    
                                    Doormount = "110"
                                    
'                                    If InStr(1, CaseName, "c��", vbTextCompare) > 0 And InStr(1, CaseName, "���", vbTextCompare) > 0 And InStr(1, CaseName, "���", vbTextCompare) > 0 Or _
'                                        InStr(1, CaseName, "c���", vbTextCompare) > 0 And InStr(1, CaseName, "����", vbTextCompare) > 0 And InStr(1, CaseName, "���", vbTextCompare) > 0 Then
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
                                    
                                Case "/3-1��", "/4��", "/2-1��", "/2��", "/2��-1"
                                        
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = ""
                                    
                                Case "/3-1��", "/4��", "/2-1��", "/2��", "/2��-1"
                                        
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = GetDrawerMountKv()
                                    
                                Case "/2-2��", "/2-2��"
                                    name = name & SVar
                                    DoorQty = 2
                                    Drawermount = ""
                                    
                                
                                Case "/2-2��"
                                    name = name & SVar
                                    DoorQty = 2
                                    Drawermount = GetDrawerMountKv()
                                    
                                
                                Case Else
                                    name = ""
                                    MsgBox "����������� ����", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                        Case "����"
                            Select Case SVar
                                Case "/���", "/2���"
                                    name = name & SVar
                                    DoorQty = 0
                                    Drawermount = ""
                                Case Else
                                    name = ""
                                    MsgBox "����������� ����", vbCritical
                                    ActiveCell.Interior.Color = vbRed
                            End Select
                            
                        Case "���", "���"
                            
                           
                            
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
                                    
                                    If Check_��(casename) Then
                                        name = name & " ��"
                                        'DoorQty = 0
                                        Doormount = "+20"
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "������ " & GetDrawerMountKv()
                                    End If
                                
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                                                   
                                Case "/���", "/2���"
                                    name = name & SVar
                                    DoorQty = 0
                                    Drawermount = ""
                                
                                Case "/1��", "/��", "/1��", "/��", "/1��-�", "/��-�", "/1��-�", "/��-�"
                                   
                                    
                                    If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = ""
                                    
                                Case "/1��", "/��"
                                    
                                    If W <= 50 Then DoorQty = 1 Else DoorQty = 2
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMountKv()
                                    
                                
                                    
                                Case "/1���", "/���", "/1���", "/���"
                                        
                                    DoorQty = 0
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    Drawermount = ""
                                
                                Case "/1���", "/���"                                         ' � �����
                                        
                                    DoorQty = 0
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMountKv()
                                
                                Case "/1�"
                                    DoorQty = 0
                                    name = name & SVar
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMount()
                                    
                                Case "/2"
                                    Drawermount = GetDrawerMount()
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    If Check_��(casename) Then
                                        name = name & " ��"
                                        DoorQty = 0
                                        'DoorMount = "+20" ��� ������!
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "������ " & GetDrawerMountKv()
                                    End If
                                    
                                                                       
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                
                                Case "/2-1", "3"
                                        
                                    Drawermount = GetDrawerMount()
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    If Check_��(casename) Then
                                        name = name & " ��"
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "������ " & GetDrawerMountKv()
                                    End If
                                
                                
                                Case "/2-1��", "/1-2��", "/3��"
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    Drawermount = GetDrawerMountKv()
                                    
                                Case "/2-1��", "/2-1��", "/1-2��", "/1-2��", _
                                         "/3��", "/3��"
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    Drawermount = ""
                                    
                                Case "/2��-1", "/2��-1"
                                    
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = ""
                                
                                Case "/2��-1"
                                    
                                    name = name & SVar
                                    DoorQty = 1
                                    Drawermount = GetDrawerMountKv()
                                
                                Case "/2��", "/2��"
                                    name = name & SVar
                                    DoorQty = 0
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = ""
                                    
                                Case "/2��"
                                    name = name & SVar
                                    DoorQty = 0
                                    ' 20/07/2009 If W >= 60 And Not IsNull(Handle) Then HandleExtra = GetHandleExtra(Handle)
                                    
                                    Drawermount = GetDrawerMountKv()
                                    
                                Case "/2-2"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = GetDrawerMount()
                                
                                Case "/2-2��"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = ""
                                    
                                Case "/2-2��"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = ""
                                    
                                Case "/2-2��"
                                    name = name & SVar
                                    DoorQty = 2
                                    
                                    Drawermount = GetDrawerMountKv()
                                
                                Case "/3-1", "/4"
                                    
                                    Drawermount = GetDrawerMount()
                                    
                                    name = name & SVar
                                    DoorQty = 0
                                    
                                    If Check_��(casename) Then
                                        name = name & " ��"
                                        If IsEmpty(D) Then D = 570 '530
                                        Drawermount = "������ " & GetDrawerMountKv()
                                    End If
                                
                                Case "/3-1��", "/4��", "/3-1��", "/4��"
                                        
                                    name = name & SVar
                                    DoorQty = 0
                                
                                    Drawermount = ""
                                
                                Case "/3-1��", "/4��"
                                        
                                    name = name & SVar
                                    DoorQty = 0
                                
                                    Drawermount = GetDrawerMountKv()
                                
                                Case Else
                                    name = ""
                                    MsgBox "����������� ����", vbCritical
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
                            
                            
                            
                            
                            Case "����"
                            
                    
                        Case Else
                            name = ""
                            MsgBox "����������� ����", vbCritical
                            ActiveCell.Interior.Color = vbRed
                        End Select ' �-�/�-Name
                Case Else
            End Select '�-
            
            If Not IsEmpty(fQty) Then
                If DoorQty <> fQty Then DoorQty = InputBox("�������� ���-�� ������", "���-�� ������ �����", fQty)
            End If
    
        Case "�"
            If name = "������ �15" Then
            Else
            
            ShelfQty = Empty
            
            Do
                DoorQty = InputBox("������� ���-�� ������", "����� ������", fQty)
            Loop Until IsNumeric(DoorQty)
            fQty = Empty

            Select Case name
                Case "��"
                    
                    Select Case SVar
                        Case Empty
                            
                        Case Else
                            MsgBox "����������� ����", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                
                Case "���"
                    
                    
                    Select Case SVar
                        Case Empty
                            
                            If IsEmpty(D) Then D = 300
                            Drawermount = GetDrawerMount()
                            
                        Case "/2"
                        
                            If IsEmpty(D) Then D = 300
                            Drawermount = GetDrawerMount()
                            
                        Case Else
                            MsgBox "����������� ����", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                    
                Case "��", "��", "���"
                
                    Select Case SVar
                        Case Empty
                        
                        Case Else
                            MsgBox "����������� ����", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                
                
                Case "���", "���"
                
                    If IsEmpty(D) Then D = 570 ' 530
                
                    Select Case SVar
                        Case Empty
                            
                            Drawermount = GetDrawerMount()
                            
                        Case "/2��", "/2��", "/3-1��", "/4��", "/2-1��", "/��", "/��", "/1-2��", "/��-�", "/��-�", "/1��-�", "/1��-�", _
                                 "/2-1��", "/3��", "/3��"
                                
                            name = name & SVar

                            Drawermount = ""

                        Case "/2��", "/3-1��", "/4��", "/2-1��", "/��", "/1-2��", "/3��"
                                
                            name = name & SVar
                            
                            Drawermount = GetDrawerMountKv()

                        Case "/2-1", "/1-2", "/2", "/3-1", "/4", "/3"
                            name = name & SVar
                            
                            Drawermount = GetDrawerMount()
                        
                        Case Else
                            name = ""
                            MsgBox "����������� ����", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                    
                                                  
                
                Case "���", "���"
                    Select Case SVar
                        Case Empty, "/2��"
                            name = name & SVar
                            

                            Drawermount = ""
                        
                        Case Else
                            name = ""
                            MsgBox "����������� ����", vbCritical
                            ActiveCell.Interior.Color = vbRed
                    End Select
                    
                                                  
                    
                
                Case Else
                    MsgBox "����������� ����", vbCritical
                    ActiveCell.Interior.Color = vbRed
            End Select
            End If
        Case Else
            MsgBox "��� �� ����", vbCritical
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
    If casepropertyCurrent.p_dspbottom > 0 Then ActiveCell.Offset(, 24).Value = "!��� ������ ���!"
    
    If IsEmpty(D) = False Then
    caseglub = D
    ActiveCell.Offset(, 31).Value = D
    Else
    caseglub = 570
    ActiveCell.Offset(, 31).Value = 570
    End If
    'ActiveCell.Offset(, 23).Value = ������!!
    
    HandleExtra = Empty
    casename = name
    caseHeight = H
    Exit Sub
err_ParseCase:
    MsgBox Error, vbCritical
End Sub

Private Function Check_��(ByVal casename As String) As Boolean
    Dim vp As Integer, dvp As Integer
    
    vp = InStr(1, casename, "��", vbBinaryCompare)
    dvp = InStr(1, casename, "���", vbBinaryCompare)
    
    If vp > 0 And (vp - 1 <> dvp Or dvp = 0) Then
        Check_�� = True
    Else
        Check_�� = False
    End If

End Function


' Leg - ������ ��� ������ �������!!!!!! ����� ������������ ���������� DefHandle!!!!
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

    windowcount = 1 ' �� ��������� ��� ���� ������ ����, ����� ������ (��. ����)
    
    

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
        
        If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = rsCases!HandleDefault
        If IsEmpty(Leg) Then Leg = rsCases!LegDefault
        
        ParseShelving = True
    End If
    
    ActiveCell.Offset(, 15).Value = name
    ActiveCell.Offset(, 18).Value = Drawermount
    ActiveCell.Offset(, 19).Value = Doormount
    ActiveCell.Offset(, 20).Value = bWithFittKit
        
    Exit Function
    
    
    Select Case Left(name, 1)
        
        Case "�" '����������
        
            bWithFittKit = True
        
            If IsEmpty(Handle) Then Handle = "�025"
            Doormount = "110"
            
            Select Case name
                Case "�2"
                    Drawermount = "45"
                Case "�1", "�3"
                    Doormount = "��������������"
                Case "����1", "����1"
                    Drawermount = ""
                    Doormount = ""
                    Handle = ""
                    bWithFittKit = False
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "����������� ������ ����������: " & name, vbCritical
            End Select
        
        
        Case "�" '������
        

            bWithFittKit = True

            If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
            Doormount = "110"
            
            Select Case name
                Case "�1", "�3"
                    Drawermount = "35"
                Case "�2", "�4", "�5"
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "����������� ������ ������: " & name, vbCritical
            End Select
            
            
        Case "�" '������
            
            Select Case name
                Case "���", "���", "����", "����"
                Drawermount = 40
                
             
                Case "��", "���", "���", "����������"
                
                
                
                Case Else
                
                    If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "8303"
                
                    Doormount = "110"
                
                    Select Case name
                    Case "�1", "�6", "�5", "�1����", "�6����", "�5����", "�1���", "�6���", "�5���"
                        Drawermount = "40"
                    Case "�2", "�3", "�4", "�2����", "�3����", "�4����", "�2���", "�3���", "�4���"
                    Case Else
                        ParseShelving = False
                        MsgBox "����������� ������ ������: " & name, vbCritical
                    End Select
            End Select
             
            
        Case "�" '������
            
            bWithFittKit = True
        
            If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "8303"
            Doormount = "110"
            
            Select Case name
                Case "�1", "�4"
                Case "�2"
                    Drawermount = 45
                Case "�3"
                    Drawermount = 35
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "����������� ������ ������: " & name, vbCritical
            End Select
        Case "�" '������
            
            Doormount = "110"
            
            If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
                        
            Select Case name
                Case "�1", "�2", "�4", "��1", "��2", "��4"
                Case "�3", "��3", "�3����", "�3���"
                    Drawermount = 40
                Case Else
                    ParseShelving = False
                    MsgBox "����������� ������ ������: " & name, vbCritical
            End Select
        Case "�" '�������� + �������
            
            Doormount = "110"
            
            If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
            
            Select Case name
                Case "�������"
                Case "�1", "�2����", "�3", "�3����", "�3���", "�4���", "�5", "�5����", "�5���"
                Case "�1 (���)"
                    Leg = "465"
                Case "�2", "�2���"
                    Drawermount = 40
                Case "�4", "�4����"
                    Drawermount = 25
                
                Case "���", "��"
                Handle = Empty
                Leg = Empty
                Drawermount = Empty
                Case "����", "����"
                Handle = "������"
                Drawermount = 40
                Case "���"
                Handle = "������"
                Drawermount = 40
                Leg = "2706"
                Case Else
                    ParseShelving = False
                    MsgBox "����������� ������ ��������: " & name, vbCritical
            End Select
        Case "�" ' �����/�����
            
            Doormount = "110"
            
            Select Case name
                Case "���1", "���3", "���2", "���4"
                
                    bWithFittKit = True
                
                    Drawermount = 45
                    If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
               
                Case "���1", "���2", "���3", "���4"
                    
                    If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
                
                Case Else
                    ParseShelving = False
                    MsgBox "����������� ������ �����/�����: " & name, vbCritical
            End Select

        Case "�" '�������� + �����1000 + ������ �������� + ������ �������
            
            Doormount = "110"
            
            
            Select Case name
                Case "����1", "����1", "�����2", "�����2", "�����1", "����1", "�����4", "�����5", "�����7", "�����6�����", "�����6�����", "�����6������", "�����1", "�����2"
                Doormount = ""
                Case "����1", "����2", "����3", "�����1", "�����2", "�����3", "�����4", "�����1", "�����", "������", "����4"
                Doormount = ""
                Case "����", "����1"
                    Drawermount = "40"
                    Doormount = ""
                Case "����1", "����2", "����2", "�����1", "�����2", "����1", "����2"
                    Drawermount = "35"
                    Doormount = "��������������"
                    Handle = "1���. ������"
                    Leg = "������ 100"
                Case "����3"
                    Leg = "������ 100"
                    Handle = "1���. ������"
                    Doormount = "��������������"
                Case "�����1"
                    Drawermount = "35"
                    Doormount = ""
                    Handle = "1���. ������"
                    Leg = "������ 100"
                Case "�����1"
                    Drawermount = "40"
                Case "����2"
                    Doormount = ""
                    Leg = "������ 100"
                Case "����1"
                    Drawermount = ""
                    Doormount = "��������������"
                    Handle = "1���. ������"
                    Leg = "������ 100"
                Case "����3", "����2", "����3", "����11", "����13", "����12", _
                    "����2", "����6", "����7", "����8", "����10", "����6", "����3", "����7", _
                    "����4", "���", "���", "���2", "���1", _
                    "����2", "����5�", "����4", "�����11", "����10", "���3"
                
                    Leg = "465" ' ������ ��� ����� ������, ����� ������������ ���������� DefHandle!!!
                    
                    Handle = "�06"
                    
                
                    Select Case name
                        Case "����2", "����3", _
                            "����2", "����4"
                            Drawermount = "����� 40"
                            Doormount = "+20"
                        
                        Case "�����11"
                            Doormount = "+20"
                            Drawermount = "����� 45"
                        
                        Case "����10", "����6"
                            Doormount = "+20"
                            Drawermount = "����� 50"
                        
                        Case "����12", "����13", "����6", "����7"
                            Doormount = "-30"
                            Drawermount = "����� 40"
                        
                        Case "����2", "���"
                            Doormount = "+20"
                            Drawermount = "����� 50"
                        
                        Case "����8", "����3"
                            Doormount = "+20"
                            Drawermount = "����� 40"
                            
                        Case "����7", "����11"
                            Doormount = "-30"
                            Drawermount = "����� 50"
                                                
                        Case "���2"
                            Doormount = "+20"
                            Drawermount = "����� 50"
                            
                        Case "���"
                            Doormount = "-30"
                            Drawermount = "����� 40"
                    
                        Case "���1"
                            Doormount = "+20"
                            Drawermount = "����� 35"
                    
                        Case "����3", "����4", _
                            "����5�", "���3"
                    
                        Case Else
                            ParseShelving = False
                            MsgBox "����������� ������ �����: " & name, vbCritical
                    End Select
                    
                
                
                Case "���1", "���2", "���3", "���4", "����1", "����2"
                    If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
                                        
                    Select Case name
                        
                        Case "���1", "���3"
                            Doormount = "-30"
                        
                        Case "���2"
                            Drawermount = "����� 45"
                    
                        Case "���4"
                            'DoorMount = "+20"
                            
                        Case "����2", "���2", "����1"
                       
                            Doormount = "+20"
                            Drawermount = "����� 50"
                    
                        Case Else
                            ParseShelving = False
                            MsgBox "����������� ������ ������ ��������: " & name, vbCritical
                    End Select
                    
                    
            Case Else
                    
'                    If IsEmpty(Handle) Then
'                        If InStr(1, Face, "�����", vbTextCompare) > 0 Or InStr(1, Face, "�����", vbTextCompare) > 0 Then
'                            Handle = "1026"
'                        ElseIf InStr(1, Face, "����", vbTextCompare) > 0 Or InStr(1, Face, "������", vbTextCompare) > 0 Then
'                            Handle = "1035"
'                        ElseIf InStr(1, Face, "������", vbTextCompare) > 0 Then
'                            Handle = "1007"
'                        End If
'                    End If
                    If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
                    
                    Select Case name
                        Case "�1", "�2����", "�3", "�3����", "�3���", "�4", "�4���", "�4����", "�5", "�5���", "�5����", "�7�", "�7�"
                        Case "�8", "�8����", "�8���"
                            Drawermount = "40"
                        Case "�2", "�2���", "�6", "�6����", "�6���", "�7", "�7�"
                            Drawermount = "40"
                        Case Else
                            ParseShelving = False
                            MsgBox "����������� ������ ��������: " & name, vbCritical
                    End Select
            End Select

        Case "�" '����
            If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
            
            Select Case name
                Case "����"
                    Doormount = "FGV45"
                Case Else
                    ParseShelving = False
                    MsgBox "����������� ������ ����: " & name, vbCritical
            End Select
        Case "�"
            bWithFittKit = False
            
            Select Case name
                Case "��1"
                Case Else
                    MsgBox "����������� ������ ������� XXI ���: " & name, vbCritical
            End Select
                    
        Case "�" '������� XXI ��� + ��������!!!
            
            bWithFittKit = True
            
            Doormount = "110"
            If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
            
'            Select Case Name
'                Case "�1", "�2", "�3", "�4"
'
'                    bWithFittKit = True
'
''                    If IsEmpty(Handle) Then Handle = "2903"
'                    If IsEmpty(Handle) Then If InStr(1, Face, "������", vbTextCompare) = 0 Then Handle = "0603"
'
'                Case Else
'
''                    If IsEmpty(Handle) Then
''                        If InStr(1, Face, "�����", vbTextCompare) > 0 Or _
''                            InStr(1, Face, "�����", vbTextCompare) > 0 Or _
''                            InStr(1, Face, "����", vbTextCompare) > 0 Then
''                            Handle = "3826"
''                        ElseIf InStr(1, Face, "����", vbTextCompare) > 0 Then
''                            Handle = "3835"
''                        End If
''                    End If
'                    If IsEmpty(Handle) Then If InStr(1, Face, "������", vbTextCompare) = 0 Then Handle = "0603"
'
'            End Select
            
            Select Case name
                Case "�1", "�3", "�4", "�5"

'                    If IsEmpty(Handle) Then Handle = "2903"
                    'If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
                Case "�2"

'                    If IsEmpty(Handle) Then Handle = "2903"
                    'If IsEmpty(Handle) Then If InStr(1, face, "������", vbTextCompare) = 0 Then Handle = "0603"
                    Drawermount = 45
                    
                Case "��5"
                    windowcount = Empty
                    
                Case "��5", "��1", "��2", _
                        "���", "��4��", "��4��", "��2", "��1", "��2", "��1", "��", "��", "����"
                    windowcount = Empty
                
                Case "��4", "��5", "��2", "��3", "���", "��3", "��2", "��4", "��3"
                    Drawermount = 40
                    windowcount = Empty
                    
                Case "��3"
                    Drawermount = 45
                    windowcount = Empty
                
                Case "��"
                    Drawermount = 45
                    windowcount = Empty
                
                Case "��1"
                    Drawermount = 50
                    windowcount = Empty
                
                Case "��"
                    Doormount = "FGV45"
                    windowcount = Empty
                    
                Case "�� (���)"
                    bWithFittKit = False
                
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 40"
                    windowcount = Empty
                    Doormount = "FGV45"
                                    
                Case "���1 (���)"
                    bWithFittKit = False
                
                    Leg = "465"
                    Handle = "�06"
                    windowcount = Empty
                    Doormount = "FGV45"
                                    
                Case "��2 (���)", "���1 (���)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 40"
                    windowcount = 1
                    Doormount = "FGV ��� �����."
                                        
                Case "����5 (���)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    windowcount = Empty
                    Doormount = "FGV ��� �����."

                Case "���6 (���)", "��6 (���)"
                                        
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 45"
                    windowcount = Empty
                    Doormount = "FGV ��� �����."
                                    
                Case "��1 (���)", "��2 (���)", "��3 (���)", "��4 (���)", "��1 (���)", "��3 (���)", "��2 (���)", _
                     "��� (���)", "��� (���)", "�� (���)", "�� (���)", "��5 (���)", "��5 (���)", "��7(���)", "����8 (���)"
                
                    bWithFittKit = False
                    
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 40"
                    windowcount = Empty
                    Doormount = "FGV ��� �����."
                    
                Case "��4�� (���)", "��3 ���)", "��2 ���)", "��1 (���)", _
                    "��4 (���)", "��1 (���)", "��2 (���)", "��5 (���)", "���� (���)", _
                    "���2 (���)", "����2 (���)", "����4 (���)", "����10", "���3 (���)"
                
                    bWithFittKit = False
                    
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 40"
                    windowcount = Empty
                    
                Case "��1 (���)"
                    bWithFittKit = False
                    
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 50"
                    windowcount = Empty
                    
                Case "��8 (���)", "��2 (���)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 50"
                    windowcount = 1
                
                Case "����6 (���)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 40"
                    windowcount = 1
                
                Case "�� (���)", "��3 (���)", "�������", _
                     "���4 (���)", "���5 (���)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 45"
                    windowcount = Empty
                    
                Case "��6 (���)", "��1 (���)"
                    
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    windowcount = Empty
                    
                Case "���7 (���)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 35"
                    windowcount = Empty
                    
                Case "���8 (���)"
                    
                    bWithFittKit = False
                    Leg = "465"
                    Handle = "�06"
                    Drawermount = "����� 50"
                    Doormount = "+20"
                    windowcount = Empty
                    
                Case Else
                    bWithFittKit = False
                    ParseShelving = False
                    MsgBox "����������� ������ ��������/������� XXI ���: " & name, vbCritical
            End Select
            
        Case "�"
        
            bWithFittKit = False
            Doormount = "110"
            Leg = "465"
            Handle = "�06"
            
            Select Case name
                Case "�1"
                    Drawermount = "����� 45"
                    
                Case "�2", "�3", "�5", "�6", "�7", "�9", "�10"
                
                Case "�4"
                    Drawermount = "����� 40"
                    
                Case "��"
                    Doormount = "FGV45"
                    
                Case Else
                    ParseShelving = False
                    MsgBox "����������� ������ �������: " & name, vbCritical
                    
            End Select
        
        Case "�"
        
            bWithFittKit = False
            Handle = "1006(160)"
            
            
            Select Case name
                Case "��10 ��� ���4 (�.1442) ��.372��", _
                     "��20 ��� ���5 (�.2200)����� ��.620��", _
                     "��20 ��� ���5 (�.2200)����� ��.570��", _
                     "��20 ��� ���5 (�.2200) ��.570��"
                    
            
            
            
                Case "��45(2200)���� 14���� ����.41", _
                    "��45(2200)���� 2������ ����.41"
                    Drawermount = "����� 35"
                Case "���1", "���2", "���3"
            
                Case "���7", "���2", "���1", "���3", "���6", "���5", "���4"
                
                    Leg = "2706"
                    Drawermount = "����� 50" ' � ������."
                    
                Case "���2"
                    
                    Leg = "2706"
                    Doormount = "���� � ���. ��������"
                    Drawermount = "����� 50" ' � ������."
                                
                
                Case "���1", "���5", "���2", "���4", "���3", "���7", "���6"
                    
                    Doormount = "����"
                
                Case "���8"
                    
                    Doormount = "���� ��� ������"
                
                Case "���1"
                    
                    Leg = "2706"
                
                Case "���3"
                    
                    Leg = "2706"
                    Drawermount = "����� 40" ' � ������."
                    
                Case "���4"
                    
                    Leg = "2706"
                    Doormount = "����"
                
                Case "���3", "���5"
                    
                    Leg = "2706"
                    Drawermount = "����� 40" ' � ������."
                    Doormount = "����"
                    
                Case "���4"
                    
                    Leg = "2706"
                    Drawermount = "����� 40" ' � ������."
                    
                Case "���1", "���2", "���1"
                    
                    Leg = "2706"
                    Doormount = "���� � ���. ��������"
                    
                    
                '********************************
                '********************************
                    
                '���
                Case "���(578)/4(203-4)", "���(877)/4(176-4)", "���(877)/4(223-4)", "���(978)/4(176-4)", "���(978)/4(223-4)", _
                        "���(578)/4(176-3,283)", "���(1277)/4(296-4)����� ������"
                
                    Drawermount = "����� 40"
                    Leg = "2706"
                    
                Case "���(1277)/5(713,176-4)", "���(1277)/5(901,223-4)", _
                        "���(1876)/6(713,176-4,713)", "���(1876)/6(901,223-4,901)", _
                        "���(578)/2(640,176)", "���(1277)/5(484-2,223,484-2)"
                
                    Drawermount = "����� 40"
                    Leg = "2706"
                    Doormount = "����"
                    
                Case "���(578)/1(818)", _
                        "���(1277)/3(596-2,1196)", "���(1277)/4(596-4)", "���(1277)/6(396-6)"
                
                    Leg = "2706"
                    Doormount = "����"
                    
                ' ���
                Case "���(1277)/2(223-2)", "���(1277)/2(396-2)", "���(678)/1(223)", "���(678)/1(396)", _
                    "���(678)/2(296-2)", "���(1876)/3(223-3)", "���(1876)/3(396,223,396)", _
                    "���(1876)/3(396-3)", "���(2476)/3(396,223,396)", "���(1876)/4(396,197-2,396)", _
                    "���(2476)/4(396,197-2,396)"
                
                    Leg = "2706"
                    Drawermount = "����� 50"
                    
                Case "���(1876)/2(223-2)�����", "���(1876)/2(396-2)�����"
                
                    Leg = "2706"
                    Drawermount = "����� 50"
                    
                Case "���(1876)/2(596-2)�����", "���(678)/1(596)", "���(1277)/2(596-2)", _
                    "���(1876)/3(596-3)"
                
                    Leg = "2706"
                    Doormount = "����"
                    
                Case "���(1876)/3(596,296,596)", "���(2476)/3(596,296,596)", "���(1876)/4(596,296-2,596)", _
                    "���(2476)/4(596,296-2,596)"
                    
                    Drawermount = "����� 50"
                    Leg = "2706"
                    Doormount = "����"
                    
                
                ' ���
                Case "���(478)/1(596)", "���(478)/1(896)", "���(478)/2(596-2)", "���(678)/2(596-2)", _
                    "���(478)/1(1196)", "���(678)/1(396)", "���(678)/1(396)���", "���(678)/1(596)", _
                    "���(978)/1(396)", "���(978)/1(396)���", "���(1277)/1(396)", "���(1876)/1(396)�����", _
                    "���(1876)/1(596)�����", "���(2176)/1(396)�����", "���(678)/2(396-2)", "���(877)/2(1196-2)", _
                    "���(1277)/2(396-2)", "���(1277)/2(596-2)", "���(1876)/2(396-2)�����", "���(1876)/2(596-2)�����", _
                    "���(1876)/3(396-3)", "���(1876)/3(596-3)", "���(1277)/4(396-4)", "���(1277)/4(596-4)", _
                    "���(376)/1(996)�������"
                    
                    Leg = "2706"
                    Doormount = "����"
                    
                '����
                Case "����(678-1200)/�����", "����(678-1400)/�����", "����(678-1573)/�����"
                
                    Leg = "2706"
                    
                    
                '����
                Case "����", "�������", "��������", "����1����", "����2����", "����1���", "����2���"
                
                    Leg = "2706"
                    Drawermount = "������ 50"
                    
                    '����
               
                Case "����1����", "����1���", "����2����", "����2���"
                
                    Leg = "2706"
                    
     
                    
                '���
                Case "���(678)/1(1796)", "���(678)/3(596-3)", "���(1475)/4(1596-4)", "���(877)/6(396-2,1000-2,396-2)", _
                        "���(678)/2(596-2)�����", "���(877)/2(1796-2)", "���(1475)/4(1870-4)", "���(1475)/4(2074-4)", _
                        "���(1076)/6(496-2,1074-2,496-2)"
                
                    Leg = "2706"
                    Doormount = "����"
                    
                Case "���(877)/5(897-2,1346,223-2)", "���(678)/3(748,296,748)", "���(678)/4(596-2,296-2)"
                
                    Drawermount = "����� 40"
                    Leg = "2706"
                    Doormount = "����"
                    
                Case "���(678)/2(296-2)�����"
                    
                    Leg = "2706"
                    Drawermount = "����� 40"
                    
                Case "���(877)/4(1400-2,196-2)"
                
                    Leg = "2706"
                    Drawermount = "����� 50"
                    Doormount = "����"
                    
                Case "���(678)/�����"
                    
                    Leg = "2706"
                    
                Case "��20 ��� ���5 (�.2200) ��.570��", "��20 ��� ���5 (�.2200)����� ��.570��", "��20 ��� ���5 (�.2200)����� ��.620��", _
                "����20", "��10 ��� ���4 (�.1442) ��.372��", "����20 ��� ���3"
                
                Case "��10 ��� ���4 (�.2040) ��.570��", "��20 ��� ���5 (�.1184) ��.375��", "����20 ��� ���5", "����20����� ��� ���5", _
                    "����20", "����20�����", "��20 ��� ���5 (�.2200) ��.570��"

                    Leg = "������ 100"
                    
            
                    
                Case Else
                        ParseShelving = False
                        MsgBox "����������� ������ �������� " & name, vbCritical
            End Select
        Case "�"
            bWithFittKit = False
            
            Drawermount = "������"
            Handle = "1006(160)"
            Select Case name
                Case "���-1", "���-2"
                Drawermount = "35"
                Doormount = "��������������"
                Handle = "�06"
                
            
                Case "��-1"
                Drawermount = "������ 35"
                Doormount = "��������������"
                Handle = "�06"
                Case "��-2"
                Drawermount = "������ 35"
                Doormount = "��������������"
                Handle = "�06"
                
                Case "��10 ��� ���4 (�.820) ��.483��", "��10 ��� ���4 (�.1442) ��.492��", "��10 ��� ���4 (�.1451) ��.570��", _
                    "��10 ��� ���4 (�.718) ��.492��", "��10 ��� ���4 (�.820) ��.483��", "����10 ��� ���4", "����10"


            End Select
                
        
            
        Case "�"
            Select Case name
                Case "�������"
                    windowcount = 1
                    
                Case Else
                    ParseShelving = False
                    MsgBox "������� XXI ���: " & name, vbCritical
            End Select
        
        Case Else
            ParseShelving = False
            MsgBox "����������� ������", vbCritical
    End Select
    
    ActiveCell.Offset(, 15).Value = name
    ActiveCell.Offset(, 18).Value = Drawermount
    ActiveCell.Offset(, 19).Value = Doormount
    ActiveCell.Offset(, 20).Value = bWithFittKit
    
    'Name = InputBox("����������� ������ ������", "������", Name)
    Exit Function
err_ParseShelving:
    ParseShelving = False
    MsgBox Error, vbCritical, "������ ������"
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
                (InStr(ShiftK, Fitting, "�� ���� ����� �����", vbTextCompare) > 0 Or _
                InStr(ShiftK, Fitting, "�� ����� �����", vbTextCompare) > 0 Or _
                InStr(ShiftK, Fitting, "�� ����� �����", vbTextCompare) > 0) And _
                (InStr(ShiftK, Fitting, "Sens", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "����", vbTextCompare) > 0 Or _
                InStr(ShiftK, Fitting, "Blumot", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "FGV", vbTextCompare) > 0) Or _
                (InStr(ShiftK, Fitting, "����� ", vbTextCompare) > 0 And InStr(ShiftK, Fitting, " �� ���� �����", vbTextCompare) > 0) Or _
                (InStr(ShiftK, Fitting, "�����", vbTextCompare) > 0 And InStr(ShiftK, Fitting, "�� ���� �����", vbTextCompare) > 0) _
             ) _
             And InStr(ShiftK, Fitting, "��", vbTextCompare) = 0 _
            ) = False _
            Then
    Dim k As Integer, bExtra As Boolean, PlankColor
    
    If Item = cPlank Then PlankColor = FittOpt
    k = InStr(ShiftK, Fitting, Item, vbTextCompare)
    
    If k Then
        '���������, ������������� ��� �� �����
            
                If InStr(ShiftK, Fitting, "�� ���", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "��  ���", vbTextCompare) > 0 Then bExtra = False Else bExtra = True
    
        
        'Dim FittOpt
        Dim en As Integer, p As Integer, ts As String, bPlus As Boolean
        
        Do
            ' ������ ������ ����, ���� �� �����
            ' ��������� ����� ������ ������� ��������� � ������ (k) ����� ��� ����, ����� �� �����, ���� ������� ���
            en = InStr(k + Len(Item), Fitting, "+")
            p = InStr(k + Len(Item), Fitting, ".")
            If p = 0 Then p = Len(Fitting)
            If en = 0 Or (p > 0 And en > p) Then
                en = p
                bPlus = False
            Else
                bPlus = True
            End If
            
            
            
            ' ��������� �������� Option
            Dim x As Integer
            If InStr(k, Fitting, Item, vbTextCompare) Then
                ActiveCell.Characters(k, Len(Item)).Font.Color = vbBlue
            
                p = InStr(k + Len(Item), Fitting, " ")
                x = InStr(k + Len(Item), Fitting, "-") ' !!!����� ����� ���� ��������, ���� � ������������ ���� ������������� "-"
                If p = 0 Or (x > 0 And p > x) Then p = x
                
                k = p  ' ����������� �� ������� �����
            End If
           
            If en > k Then
                ts = Mid(Fitting, k + 1, en - k) ' "�������" ������, ������� ������ ����� �������������
            Else
                ts = ""
            End If
            
            ' ������ �������� � ts
            
            ' ���� ���-��
            Dim qty, QtyPattern As String
            QtyPattern = "��"
            p = InStr(1, ts, QtyPattern, vbTextCompare)
            If p = 0 Then
                QtyPattern = "�-�"
                p = InStr(1, ts, QtyPattern, vbTextCompare)
            End If
            If p = 0 Then
                QtyPattern = "����"
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
                    
                    ' ������� ��, ��� �����
                    ActiveCell.Characters(k + p, Len(QtyPattern)).Font.Bold = True
                    ActiveCell.Characters(k + p - x - Len(tqty), Len(tqty)).Font.Color = vbRed
                    ' ������� �������� ����������
                    ts = Left(ts, p - x - 1 - Len(tqty)) '& Mid(ts, p + Len(QtyPattern)) '!!!!!!!!!!!!!!!!
                Else 'If Not IsNumeric(tQty) Then
                    Do
                        Do
                            MsgBox "������! ���-��=" & tqty, vbCritical
                            tqty = InputBox(ts, "������� ���-��", tqty)
                        Loop While Not IsNumeric(tqty)
                        qty = CDec(tqty)
                    Loop Until qty >= 0
                End If
            End If
            
            ' ���� �������� � �� (�� � ������!), ���� ����
            If IsMissing(length) Then length = Empty
            p = InStr(1, ts, "��", vbTextCompare)
            If p > 0 Then
                Dim tLen As String
                If p < 5 Then x = p - 1 Else x = 4
                tLen = Mid(ts, p - x, x)
                DelTextLeft tLen
                DelTextRight tLen, x
                If tLen <> "" And IsNumeric(tLen) Then
                    length = CInt(tLen)
                    ' ������� ��, ��� �����
                    ActiveCell.Characters(k + p, 2).Font.Bold = True
                    ActiveCell.Characters(k + p - x - Len(tLen), Len(tLen)).Font.Color = vbRed
                    ' ������� �������� ����������
                    ts = Left(ts, p - x - 1 - Len(tLen)) '& Mid(ts, p + 2) '!!!!!!!!!!!!!!!!
                ElseIf Not IsNumeric(tLen) Then
                    If InStr(1, tLen, "�", vbTextCompare) = 0 Then
                        MsgBox "������! �����=" & tLen, vbCritical
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
                'MsgBox "���������!!!����� ���"
                FittOpt = ts
                ActiveCell.Characters(k + 1, Len(ts)).Font.ColorIndex = 10
            End If
            
            Dim FittingName As String
            FittingName = Item
            
            
            If Item = cPlank Then ' ��� ������
                Dim pp As Integer
                
                pp = InStr(1, FittOpt, "��", vbTextCompare)
                If pp > 0 Then
                    FittingName = "������ � ����"
                    
                    pp = InStr(pp, FittOpt, " ", vbTextCompare)
                    If pp > 0 Then FittOpt = Trim(Left(FittOpt, pp))
                    If pp = 0 Or Len(FittOpt) < 3 Then FittOpt = PlankColor
                    GoTo exit_if
                    'Length = TTThickness
                End If
                
                pp = InStr(1, FittOpt, "� ����", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "����", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "�-�", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "�/�", vbTextCompare)
                If pp Then
                    FittingName = "������ �/� ������������"
                    
                    pp = InStr(1, Left(FittOpt, pp), " ", vbTextCompare)
                    If pp > 0 Then FittOpt = Trim(Left(FittOpt, pp))
                    If pp = 0 Or Len(FittOpt) < 3 Then FittOpt = PlankColor
                    GoTo exit_if
                    'Length = TTThickness
                End If
                
                
                
                pp = InStr(1, FittOpt, "� ���", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "���", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "� ����", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "� �/�", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "�/�", vbTextCompare)
                If pp = 0 Then pp = InStr(1, FittOpt, "����", vbTextCompare)
                If pp Then
                    FittingName = "������ � ���. �����"
                    
                    pp = InStr(pp, FittOpt, " ", vbTextCompare)
                    If pp > 0 Then FittOpt = Trim(Left(FittOpt, pp))
                    If pp = 0 Or Len(FittOpt) < 3 Then FittOpt = PlankColor
                    GoTo exit_if
                    'Length = TTThickness
                End If
                                
                If InStr(1, Fitting, "����", vbTextCompare) Then
                    pp = InStr(1, Fitting, "�����", vbTextCompare)
                    If (pp > 0) Then
                        FittingName = "������ ���������"
                        
                        If Not IsEmpty(length) And IsEmpty(qty) Then
                            qty = CInt(length) \ 200
                            length = Empty
                        End If
                    End If
                End If
                pp = InStr(1, Fitting, "�����", vbTextCompare)
                    If (pp < k) And (pp > 0) Then
                    length = Null
                    
                        pp = InStr(1, Fitting, "sens", vbTextCompare)
                        If pp = 0 Then pp = InStr(1, Fitting, "sensis", vbTextCompare)
                        If pp = 0 Then pp = InStr(1, Fitting, "������", vbTextCompare)
                        If pp = 0 Then pp = InStr(1, Fitting, "������", vbTextCompare)
                        If pp > k Then
                        Dim plankoptpos As Integer
                            FittingName = "�������� ������ Sensys"
                            plankoptpos = InStr(pp, Fitting, "5", vbTextCompare)
                            If plankoptpos > pp Then FittOpt = "5��"
                            plankoptpos = InStr(pp, Fitting, "10", vbTextCompare)
                            If plankoptpos > pp Then FittOpt = "10��"
                    ElseIf pp = 0 Then
                            FittingName = "�������� ������ 5��"
                            FittOpt = Null
                    End If
                End If
                
            ElseIf Item = "�����" Then
                If InStr(1, FittOpt, "�����", vbTextCompare) > 0 And InStr(1, FittOpt, "����", vbTextCompare) = 0 Then
                    FittingName = "��������� � ����� ���."
                ElseIf InStr(1, FittOpt, "����", vbTextCompare) > 0 Then
                    FittingName = "��������� � ���� ��. ���."
                End If
                
            ElseIf Item = "����" Then
                FittingName = "�����"
            ElseIf Item = "����� Sens" Or _
                    Item = "����� Sens" Or _
                    Item = "����� Sens" Or _
                    Item = "����� ����" Or _
                    Item = "����� ����" Or _
                    Item = "����� ����" Or _
                    Item = "����� ����" Or _
                    Item = "����� ����" Or _
                    Item = "����� ����" Or _
                    Item = "������ Sens" Or _
                    Item = "������ ����" Or _
                    Item = "������ ����" _
            Then
                FittingName = "����� Sensys"
            ElseIf Item = "������ �����" Or Item = "���� �����" Or Item = "������ �����" Then
                FittingName = "������ ���."
            ElseIf Item = "������� ����" Then
                FittingName = "�������������"
             ElseIf Item = "�����" Then
                FittingName = "�-�� ��������"
                
            ElseIf Item = "������� ����" Then
                FittingName = "�������������"
            ElseIf Item = "���" Then
                FittingName = "�����*"
            ElseIf Item = "���" Then
                FittingName = "�����*"
            ElseIf Item = "������" Then
                FittingName = "������"
            ElseIf Item = "���� " Then
                FittingName = "�������������"
            ElseIf Item = "����" Then
                FittingName = "������ ���������"
            ElseIf Item = "����" Then
                If InStr(1, ActiveCell.Value, "���", vbTextCompare) > 0 Then
                    FittingName = "������� ��������������"
                    If InStr(1, ActiveCell.Value, "���", vbTextCompare) > 0 Or _
                        InStr(1, ActiveCell.Value, "Volp", vbTextCompare) > 0 Or _
                        InStr(1, ActiveCell.Value, "���", vbTextCompare) > 0 Then
                        FittOpt = "�������"
                    ElseIf InStr(1, ActiveCell.Value, "������� �", vbTextCompare) > 0 Then
                        FittingName = "������� ��������������"
                    End If
                End If
            ElseIf Item = "����� ������" Then
                FittingName = "����� ���������"
            ElseIf Item = "����� ������" Then
                FittingName = "����� ���������"
            ElseIf Item = "������ � ���" Or Item = "������ � ���" Or Item = "������ ���" Or Item = "������ ���" Or Item = "������ �� ���" Or Item = "������ �� ���" _
                Or Item = "������ ��� ���" Or Item = "������ ��� ���" _
            Then
                FittingName = "*�����. ������ ���������"
            ElseIf Item = "����" Then
                If InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 Then
                    FittingName = "��������"
                Else
                    FittingName = "����������"
                End If
            ElseIf Item = cStol Then
                FittingName = "���� ���������"
            ElseIf Item = cStool Then
                FittingName = cStul
            ElseIf Item = "����" Then
                FittingName = cNogi
            ElseIf Item = "������" Then
                FittingName = "������� ����������"
            ElseIf Item = "�����" Then
                FittingName = "VS - ������� ���������"
            ElseIf Item = "����" Then
                FittingName = "�����"
            ElseIf Item = "�����" Then
                FittingName = "�����"
                If InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0 And InStr(1, ActiveCell.Value, "5�", vbTextCompare) > 0 Then
                                FittOpt = "���������� 5�"
                ElseIf InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0 And InStr(1, ActiveCell.Value, "11�", vbTextCompare) > 0 Then
                                FittOpt = "���������� 11�"
                End If
            ElseIf Item = "�������" Then
                FittingName = "�������� ��� ������"
            ElseIf Item = "������" Or Item = "������" Then
                FittingName = "���������"
             ElseIf Item = "�����" Then
                FittingName = "��������"
             ElseIf Item = "�����" Then
                FittingName = "��������"
            ElseIf Item = "����" Then
                FittingName = "����������� ������"
            ElseIf Item = "����" Then
                FittingName = "����������� ������"
            ElseIf Item = "������" Then
                FittingName = "������"
      '          FittingName = "������ �����������"
            ElseIf Item = "����" Then
                FittingName = "��������� �������"
            ElseIf Item = "�����" Or Item = "������" Then
                FittingName = "�����������"
            ElseIf Item = "��������" Then
                FittingName = "��������"
                 If InStr(1, ActiveCell.Value, "���", vbTextCompare) > 0 Then
                    FittOpt = "���"
                End If
            ElseIf Item = "�����" Then
                FittingName = "_�����"
            ElseIf Item = "����" Then
                FittingName = "_����"
            ElseIf Item = "������" Then
                FittingName = "�����"
            ElseIf Item = "push to open" Or Item = "push-to-open" Or Item = "��� �� ����" Or Item = "��� �� ����" Then
                FittingName = "�������� �-� Push-To-Open"
            ElseIf Item = "����������" Then
                FittingName = "���������������*"
            ElseIf Item = "���������" Then
                FittingName = "��������������"
                FittOpt = ""
               
'                If InStr(1, ActiveCell.Value, "Sekura", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "������", vbTextCompare) > 0 Then
'                     If InStr(1, ActiveCell.Value, "a 2-1", vbTextCompare) > 0 Then
'                    FittOpt = "Sekura 2-1"""
'                Else
'                    FittOpt = "Sekura 8 (��� ������)"
'                End If
            ElseIf Item = "�������" Then
                FittingName = "��������������"
                FittOpt = ""
            ElseIf Item = "�������" Then
                If InStr(1, ActiveCell.Value, "������", vbTextCompare) > 0 Then
                    FittingName = "������������ ������"
                Else
                    FittingName = "������������"
                End If
            ElseIf Item = "��������" Then
                FittingName = "�������� *"
            ElseIf Item = "������� ��������������" Then
                If InStr(1, ActiveCell.Value, "���", vbTextCompare) > 0 Or _
                InStr(1, ActiveCell.Value, "Volp", vbTextCompare) > 0 Or _
                InStr(1, ActiveCell.Value, "���", vbTextCompare) > 0 Then
                FittOpt = "�������"
                
                ElseIf InStr(1, ActiveCell.Value, "������� �", vbTextCompare) > 0 Then
                    FittingName = "������� ��������������"
                End If
            ElseIf Item = "����" Then
                If InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "�������", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "������", vbTextCompare) > 0) Then
                    FittingName = "����� �/��"
                End If
            ElseIf Item = "����������" Then
                FittingName = "���������� *"
            ElseIf Item = "�������� � �" Or Item = "�������� �� �" Or Item = "�������� ��� �" Or _
                    Item = "��������� � �" Or Item = "��������� �� �" Or Item = "��������� ��� �" Then
                FittingName = "�������� �� ��������"
            ElseIf Item = "���������" Then
                FittingName = "��������� �/��"
            ElseIf Item = "������� �����������" Then
                FittingName = "����������� �/�� � �����."
            
            ElseIf Item = "����" Then
                 If InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "������", vbTextCompare) > 0) Then
                    FittingName = "���� �/�"
                ElseIf InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "������", vbTextCompare) > 0) Then
                    FittingName = "���� ������"
                ElseIf InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0) Then
                    FittingName = "���� �����"
                ElseIf InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "��", vbTextCompare) > 0) Then
                    FittingName = "���� ����"
                ElseIf InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 And (InStr(1, ActiveCell.Value, "Zebra", vbTextCompare) > 0) Then
                    FittingName = "���� Zebra"
                End If
            ElseIf Item = "������ ����" Or Item = "magic light" Then
                 If InStr(1, ActiveCell.Value, "������ ����", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "magic light", vbTextCompare) > 0 Then
                    FittingName = "�������������"
                    End If
            ElseIf Item = "�����������" Then
                 If InStr(1, ActiveCell.Value, "�����������", vbTextCompare) > 0 Then
                    FittingName = "�������������"
                End If
            ElseIf Item = "�������������" Then
                 If InStr(1, ActiveCell.Value, "�������������", vbTextCompare) > 0 Then
                    FittingName = "�������������"
                End If
            ElseIf Item = "�����" Then
                 If IsEmpty(qty) Then qty = 1
                If InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "��", vbTextCompare) > 0 Then
                    FittingName = "����� �/��"
                ElseIf InStr(1, ActiveCell.Value, "ORGA", vbTextCompare) > 0 Then
                    FittingName = "����� ORGALINE"
                 ElseIf InStr(1, ActiveCell.Value, "Arci", vbTextCompare) > 0 Or InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0 Then
                    FittingName = "����� � �������"
                 Else
                 FittingName = "�����"
                End If
            ElseIf Item = "������" Then
                FittingName = "������ ��������������"
            ElseIf Item = "����" Then
                If InStr(1, ActiveCell.Value, "������", vbTextCompare) > 0 Then
                    FittOpt = "������������� ����"
                ElseIf InStr(1, ActiveCell.Value, "����", vbTextCompare) > 0 Then
                    FittOpt = "����"
                ElseIf InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0 Then
                    FittOpt = "�����"
                Else
                    FittOpt = ""
                End If
                If IsEmpty(qty) Then qty = 1
            
            End If
exit_if:
'            If InStr(1, Fitting, "���� ", vbTextCompare) > 0 Then
'            ActiveCell = Replace(ActiveCell.Text, ".", " ", InStr(1, ActiveCell.Text, "���� ", vbTextCompare) + 1)
'            End If
    
            If Item = cGalog Then
                If qty Mod 3 = 0 Then
                    FittingName = "��������� 3"
                    qty = qty \ 3
                ElseIf qty Mod 5 = 0 Then
                    FittingName = "��������� 5"
                    qty = qty \ 5
                End If
            End If
            
            If Item = "�����" And (InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0 And InStr(1, ActiveCell.Value, "�����", vbTextCompare) > 0) Then
                Dim find18 As Integer
                find18 = 0
                find18 = InStr(InStr(1, ActiveCell.Value, "�����", vbTextCompare), ActiveCell.Value, "18", vbTextCompare)
                If find18 > 1 Then FittOpt = "18" Else FittOpt = "16"
                
                If InStr(1, ts, "3�", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 1
                ElseIf InStr(1, ts, "6�", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 2
                ElseIf InStr(1, ts, "9�", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 3
                ElseIf InStr(1, ts, "12�", vbTextCompare) > 0 Then
                    If IsEmpty(qty) Then qty = 4
                End If
            End If
            
                
            
           ' If Not IsEmpty(FittOpt) Then
'                If bExtra And (IsEmpty(Qty) And (IsEmpty(Length) _
'                    And (InStr(1, Item, cHandle) > 0 _
'                    Or InStr(1, Item, cLeg) > 0))) Then  ' ���� �� ������� �� ����� ��� ���-��, �� �������, ��� ������ � �������� ��� ������ �� ��������� (�����, �����)
                If Not IsMissing(FittOpt) And _
                    IsEmpty(qty) And _
                    (InStr(1, Item, cHandle, vbTextCompare) > 0 _
                     Or InStr(1, Item, cLeg, vbTextCompare) > 0) Then ' ���� �� ������� �� ����� ��� ���-��, �� �������, ��� ������ � �������� ��� ������ �� ��������� (�����, �����)
                    
                    If InStr(1, FittOpt, "������", vbTextCompare) Then
                        FindFittings = Null
                    Else
                        FindFittings = FittOpt
                    End If
                    
                Else
                    ' ���� ��� �����, ����������� � ������ �������
                    If InStr(1, Item, cHandle, vbTextCompare) > 0 Then
                             
                        If IsEmpty(HandleScrew) Then HandleScrew = GetHandleScrew(FittOpt, face)
                        If Not IsEmpty(HandleScrew) Then UpdateOrder OrderId, HandleScrew
                         
                        If bExtra Then ' ���� �������������
                        
                            CheckHandle FittOpt
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            
                        Else ' ���� �� �����
                        
                            FindFittings = Null
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            
                        End If
                                                     
                    Else
                        If bExtra Then ' ���� �������������
                            
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            If Item = cPlank Then PlankColor = FittOpt
                            
                        Else ' ���� �� �����
                            
                            FindFittings = Null
                            'FindFittings = FittOpt
                            If Not FormFitting.AddFittingToOrder(OrderId, FittingName, qty, FittOpt, length, , , row) Then Exit Function
                            
                        End If
                    End If
                End If
            
            
            
            k = en + 1  ' ����� ���. ������, ������ ���������
        Loop While k < Len(Fitting) And bPlus
        
        Fitting = Left(Fitting, ShiftK - 1) & Mid(Fitting, k)
        ShiftK = k
    End If
    Else
        
    
    
    
    '   InStr(ShiftK, Fitting, "Blumot", vbTextCompare) > 0)
        Dim t As String
        If (IsMissing(changeCaseZaves) = False) Then
         If (InStr(ShiftK, Fitting, "Sens", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "������", vbTextCompare) > 0 Or InStr(ShiftK, Fitting, "Sensyc", vbTextCompare) > 0) Then
                If changeCaseZaves <> 1 Then
                    changeCaseZaves = 1
                    kitchenPropertyCurrent.changeCaseZaves = 1
                   ' casepropertyCurrent.p_changeZaves = 1
                    ShiftK = Len(Fitting)
                    ActiveCell.Characters(k, Len(Item)).Font.Color = vbGreen
                    If Cells(ActiveCell.row, 10).Value <> "" Then
                        
                        t = Cells(ActiveCell.row, 10).Value
                        
                        Cells(ActiveCell.row, 10).Value = t & "!!!����� ������� �� ������!!!"
                         Else
                        Cells(ActiveCell.row, 10).Value = "!!!����� ������� �� ������!!!"
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
                        
                        Cells(ActiveCell.row, 10).Value = t & "!!!����� ������� �� ���������!!!"
                         Else
                        Cells(ActiveCell.row, 10).Value = "!!!����� ������� �� ���������!!!"
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
                        
                        Cells(ActiveCell.row, 10).Value = t & "!!!����� ������� �� FGV!!!"
                         Else
                        Cells(ActiveCell.row, 10).Value = "!!!����� ������� �� FGV!!!"
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
    MsgBox Error, vbCritical, "���������� ������"
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
'    MsgBox Error, vbCritical, "���������� �������� ����� ������"
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
'    If Not IsEmpty(CaseHang) Then rsOrderCases!CaseHang = CaseHang  ' �� ��������� �����
'    rsOrderCases!Bibb = Bibb
'    If IsNull(Handle) Then
'        rsOrderCases!Handle = Null
'        rsOrderCases!HandleExtra = 0
'    Else
'        rsOrderCases!Handle = Handle
'        If Not IsEmpty(HandleExtra) Then rsOrderCases!HandleExtra = HandleExtra
'    End If
'    If Not IsEmpty(Leg) Then
'        rsOrderCases!CaseStand = Leg ' ���� Empty, ������ �� ���������, ���� Null - ������ ���, ���� ��������, �� ��� ���
'    'Else
'       ' rsOrderCases!CaseStand = "������"
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
'        If ShelfQty >= 2 And Left(name, 1) <> "�" Then
'            Select Case name
'                Case "�� 2�"
'                   ' If Not IsEmpty(caseHeight) Then
'                                            'If caseHeight > 910 Then FormElement.AddElementToOrder OrderID, "�����", 2 * Qty, caseID Else
'                   FormElement.AddElementToOrder OrderID, "�����", Qty, caseID
'                   ' End If
'                Case "��915", "���915", "���915", "�� ����915", "���915", "����915", "����915", "���915"
'                Case "���", "���"   '"���", "�� ����"
'                    FormElement.AddElementToOrder OrderID, "�����", Qty, caseID
'                Case "��� �"
'                    FormFitting.AddFittingToOrder OrderID, "�������", Qty, , , caseID
'                'Case "����" ' "���"
'                '    FormElement.AddElementToOrder OrderID, "����� ����", Qty, CaseID
'                Case Else
'                    FormElement.AddElementToOrder OrderID, "�����", Qty, caseID
'            End Select
'        End If
'    End If
'    rsOrderCases.Update
'    AddCase = rsOrderCases!OCID
'    Exit Function
'err_AddCase:
'    MsgBox Error, vbCritical, "���������� ����� � �����"
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

    MsgBox Error, vbCritical, "���������� ����� � �����"
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
MsgBox Error, vbCritical, "����� ��������� � ����"
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
comm.Parameters.Append comm.CreateParameter("@legDefault", adVarChar, adParamInput, 20, "׸���� 100")
comm.Parameters.Append comm.CreateParameter("@str", adVarChar, adParamInput, 512, args)

comm.Parameters.Append comm.CreateParameter("@caseid", adInteger, adParamOutput)
comm.Execute
createCaseId = comm.Parameters("@caseid")


Exit Function
err_createCaseId:
createCaseId = 0
MsgBox Error, vbCritical, "�������� ��������� � ����"
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
    MsgBox Error, vbCritical, "���������� ����� � �����"
End Function

Public Function GetHandleScrew(ByVal Handle, _
                               ByVal face) As Variant
            
                          
    If Not IsNull(Handle) And Not IsMissing(Handle) Then
            If Not (Handle = "������" Or InStr(1, Handle, "������", vbTextCompare) > 0) Then
                If Not IsNull(face) And Not IsEmpty(face) Then
                If Len(face) > 3 Then
                    If InStr(1, face, "������", vbTextCompare) > 0 And _
                           (InStr(1, Handle, "A025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "A-025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "�025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "�-025", vbTextCompare) > 0) Then
                          
                          GetHandleScrew = "40"
                    ElseIf InStr(1, face, "������", vbTextCompare) = 0 And _
                           (InStr(1, Handle, "A025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "A-025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "�025", vbTextCompare) > 0 Or _
                    InStr(1, Handle, "�-025", vbTextCompare) > 0) _
                           Then
                          
                          GetHandleScrew = "35"
                          
                   
                    ElseIf InStr(1, face, "����", vbTextCompare) > 0 Or _
                        InStr(1, face, "RAL", vbTextCompare) > 0 Or _
                        InStr(1, face, "�������", vbTextCompare) > 0 Or _
                        InStr(1, face, "������", vbTextCompare) > 0 Or _
                        InStr(1, Replace(face, " ", ""), "����2", vbTextCompare) > 0 Then
                        
                        GetHandleScrew = "25"
                
                
        
                    
                    ElseIf InStr(1, face, "������", vbTextCompare) > 0 Or _
                            InStr(1, face, "�����", vbTextCompare) > 0 Or _
                            InStr(1, face, "�����", vbTextCompare) > 0 Or _
                            InStr(1, face, "����", vbTextCompare) > 0 Or _
                            InStr(1, face, "�����", vbTextCompare) > 0 Then
                        
                        GetHandleScrew = "28"
                        
                    ElseIf InStr(1, face, "�������", vbTextCompare) > 0 Then
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
'        If InStr(1, CaseColor, "���", vbTextCompare) Then + �����? + ����?
'            GetHangColor = "�����"
'        Else
'            GetHangColor = "������"
'        End If
'
'    End If

    If GetHangColor = "" Then
        Dim FormHangColor As HangColorForm
        Set FormHangColor = New HangColorForm
        
        
        While GetHangColor = ""
            Set FormHangColor = New HangColorForm
            FormHangColor.Caption = "���� �������"
            If Not kitchenPropertyCurrent Is Nothing Then
                If kitchenPropertyCurrent.dspColor <> "" Then
                    FormHangColor.Caption = FormHangColor.Caption & " �����:" & kitchenPropertyCurrent.dspColor
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
            FormCamBibbColor.Caption = "�������� �����������"
            If Not kitchenPropertyCurrent Is Nothing Then
                If kitchenPropertyCurrent.dspColor <> "" Then
                    FormCamBibbColor.Caption = FormCamBibbColor.Caption & " �����:" & kitchenPropertyCurrent.dspColor
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
    
'InStr(1, CaseColor, "���", vbTextCompare) > 0 Or
        If InStr(1, CaseColor, "������", vbTextCompare) > 0 Or _
            InStr(1, CaseColor, "��", vbTextCompare) > 0 Or _
            InStr(1, CaseColor, "��", vbTextCompare) > 0 Then
            GetBibbColor = "�����"
'        ElseIf InStr(1, CaseColor, "�����", vbTextCompare) > 0 Then
'            GetBibbColor = "�����"
'        ElseIf InStr(1, CaseColor, "�����", vbTextCompare) > 0 Then
'            GetBibbColor = "�����"
        ElseIf InStr(1, CaseColor, "����", vbTextCompare) > 0 Then
            GetBibbColor = "����"
        ElseIf InStr(1, CaseColor, "�������", vbTextCompare) > 0 Then 'If InStr(1, CaseColor, "�����", vbTextCompare) > 0 Or
            GetBibbColor = "�����"
        ElseIf InStr(1, CaseColor, "������", vbTextCompare) > 0 Then
            GetBibbColor = "������"
        ElseIf InStr(1, CaseColor, "������", vbTextCompare) > 0 Then
            GetBibbColor = "����"
        ElseIf InStr(1, CaseColor, "�����", vbTextCompare) > 0 Then
            GetBibbColor = "������"
        ElseIf InStr(1, CaseColor, "�����", vbTextCompare) > 0 Then
            GetBibbColor = "������"
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
            FormBibbColor.Caption = "���� ��������"
            If Not kitchenPropertyCurrent Is Nothing Then
                If kitchenPropertyCurrent.dspColor <> "" Then
                    FormBibbColor.Caption = FormBibbColor.Caption & " �����:" & kitchenPropertyCurrent.dspColor
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

        If InStr(1, TTColor, "������", vbTextCompare) > 0 Then
            GetPlankColor = "�����"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetPlankColor = "�����"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetPlankColor = "�������"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "����", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "���", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "���", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "����", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "������", vbTextCompare) > 0 Or _
                (InStr(1, TTColor, "���", vbTextCompare) > 0 And InStr(1, TTColor, "��", vbTextCompare) > 0) Then
            GetPlankColor = "���"
        ElseIf InStr(1, TTColor, "������", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "������", vbTextCompare) > 0 Or _
                (InStr(1, TTColor, "�����", vbTextCompare) > 0 And InStr(1, TTColor, "��", vbTextCompare) > 0) Or _
                InStr(1, TTColor, "������", vbTextCompare) > 0 And InStr(1, TTColor, "�����", vbTextCompare) = 0 Then
            GetPlankColor = "������"
        Else
            GetPlankColor = "����"
        End If

    End If
End Function

Public Sub GetHangColors(ByRef HangColors())
    ReDim HangColors(5)
    HangColors(0) = "�����"
    HangColors(1) = "���"
    HangColors(2) = "�����806"
    HangColors(3) = "�����807"
    HangColors(4) = "�����808"
    
End Sub


Public Sub GetBibbColors(ByRef BibbColors())
    ReDim BibbColors(12)
    BibbColors(0) = "�����"
    BibbColors(1) = "���"
    BibbColors(2) = "�����"
    BibbColors(3) = "�����"
    BibbColors(4) = "���"
    BibbColors(5) = "����"
    BibbColors(6) = "������"
    BibbColors(7) = "�����"
    BibbColors(8) = "����"
    BibbColors(9) = "�����"
    BibbColors(10) = "������"
    BibbColors(11) = "�����"
    BibbColors(12) = "�������"
    
    'SortArray BibbColors
End Sub
Public Sub GetCamBibbColors(ByRef CamBibbColors())
    ReDim CamBibbColors(10)
    Dim i As Integer
    i = 0
    CamBibbColors(i) = "�������"
    i = i + 1
    CamBibbColors(i) = "�����"
    i = i + 1
    CamBibbColors(i) = "����������"
    i = i + 1
    CamBibbColors(i) = "�����"
    i = i + 1
    CamBibbColors(i) = "������-���"
    i = i + 1
    CamBibbColors(i) = "������-���"
    i = i + 1
    CamBibbColors(i) = "�����"
    i = i + 1
    CamBibbColors(i) = "�����-���"
    i = i + 1
    CamBibbColors(i) = "�����-���"
    i = i + 1
    CamBibbColors(i) = "������"
    i = i + 1
    CamBibbColors(i) = "�����-���"
    
    
End Sub

Public Function GetLegShelving(ByVal CaseColor As String) As String
    Select Case CaseColor
        Case "�����", "�����"
            GetLegShelving = "�����"
        Case "����", "������", "������"
            GetLegShelving = "����"
        Case "����"
            GetLegShelving = "����"
        Case Else
            GetLegShelving = CaseColor
    End Select
    CheckLeg GetLegShelving
End Function

Public Sub GetOtbColors(ByRef OtbColors())
    ReDim OtbColors(17)
    OtbColors(0) = "����"
    OtbColors(1) = "��� ��"
    OtbColors(2) = "���"
    OtbColors(3) = "������"
    OtbColors(4) = "���"
    OtbColors(5) = "���� ���"
    OtbColors(6) = "�������"
    OtbColors(7) = "�������"
    OtbColors(8) = "������"
    OtbColors(9) = "������"
    OtbColors(10) = "�����"
    OtbColors(11) = "������"
    OtbColors(12) = "������"
    OtbColors(13) = "��� ���"
    OtbColors(14) = "�����"
    OtbColors(15) = "���� ��"
    OtbColors(16) = "��� ��"
    OtbColors(17) = "����"
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
    OtbColors(i) = "���� ���"
    i = i + 1
    OtbColors(i) = "��� ���"
    i = i + 1
    OtbColors(i) = "��������"
    i = i + 1
    OtbColors(i) = "������"
    i = i + 1
    OtbColors(i) = "�������"
    i = i + 1
    OtbColors(i) = "����� �������"
    i = i + 1
    OtbColors(i) = "����� ������"
    i = i + 1
    OtbColors(i) = "������ ������"
'    i = i + 1
'    OtbColors(i) = "������ ������"
    i = i + 1
    OtbColors(i) = "����� ������"
'    i = i + 1
'    OtbColors(i) = "��� �� ������"
    i = i + 1
    OtbColors(i) = "��������"
    i = i + 1
    OtbColors(i) = "���� ������ ��"
    i = i + 1
    OtbColors(i) = "�����"
    i = i + 1
    OtbColors(i) = "����"
    i = i + 1
    OtbColors(i) = "���� �����"
    i = i + 1
    OtbColors(i) = "�������� ����"
    i = i + 1
    OtbColors(i) = "����� ����"
    i = i + 1
    OtbColors(i) = "����� ����"
    i = i + 1
    OtbColors(i) = "������ ������"
    i = i + 1
    OtbColors(i) = "������"
    i = i + 1
    OtbColors(i) = "����������"
    i = i + 1
    OtbColors(i) = "���������"
    i = i + 1
    OtbColors(i) = "������"
    i = i + 1
    OtbColors(i) = "�����"
    i = i + 1
    OtbColors(i) = "������ ������"
    i = i + 1
    OtbColors(i) = "������� ����"
    i = i + 1
    OtbColors(i) = "���� ����"
    i = i + 1
    OtbColors(i) = "���� �����"
    i = i + 1
    OtbColors(i) = "���� ����"
    i = i + 1
    OtbColors(i) = "�������"
    i = i + 1
    OtbColors(i) = "���"
    i = i + 1
    OtbColors(i) = "������"
    i = i + 1
    OtbColors(i) = "������"
     i = i + 1
    OtbColors(i) = "��� ��"
     i = i + 1
    OtbColors(i) = "����� ������"
     i = i + 1
    OtbColors(i) = "�������� �����"
     i = i + 1
    OtbColors(i) = "������� �����"
     i = i + 1
    OtbColors(i) = "����������"
     i = i + 1
    OtbColors(i) = "�����"
     i = i + 1
    OtbColors(i) = "����� �������"
     i = i + 1
    OtbColors(i) = "������ ����������"
     i = i + 1
    OtbColors(i) = "�������"
     i = i + 1
    OtbColors(i) = "�������� ��������"
     i = i + 1
    OtbColors(i) = "������� ��������"
     i = i + 1
    OtbColors(i) = "������� �������"
    
    'SortArray OtbColors
End Sub


Public Function GetOtbColor(ByVal TTColor) As String
    If Not (IsEmpty(TTColor) Or IsNull(TTColor)) Then

        If InStr(1, TTColor, "������", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "�����", vbTextCompare) > 0 Then
            GetOtbColor = "������"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 And _
                InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetOtbColor = "��� ���"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 And _
                InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetOtbColor = "���� ���"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 And _
                InStr(1, TTColor, "��", vbTextCompare) > 0 Then
            GetOtbColor = "��� ��"
        ElseIf InStr(1, TTColor, "������", vbTextCompare) > 0 Then
            GetOtbColor = "������"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetOtbColor = "���"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetOtbColor = "���"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetOtbColor = "���"
        ElseIf InStr(1, TTColor, "�����", vbTextCompare) > 0 Then
            GetOtbColor = "�����"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetOtbColor = "�����"
        ElseIf InStr(1, TTColor, "���", vbTextCompare) > 0 Then
            GetOtbColor = "�������"
        ElseIf InStr(1, TTColor, "�����", vbTextCompare) > 0 Then
            GetOtbColor = "������"
        ElseIf InStr(1, TTColor, "����", vbTextCompare) > 0 Then
            GetOtbColor = "����"
        ElseIf InStr(1, TTColor, "������", vbTextCompare) > 0 Or _
                InStr(1, TTColor, "������", vbTextCompare) > 0 Or _
                (InStr(1, TTColor, "����", vbTextCompare) > 0 And InStr(1, TTColor, "��", vbTextCompare) > 0) Then
            GetOtbColor = "������"
        Else
            GetOtbColor = "����"
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
            If InStr(1, Handle, "������", vbTextCompare) Or Handle = "������" Then
                he = 0
            Else
        
                Init_rsHandle
                
                If rsHandle.RecordCount > 0 Then rsHandle.MoveFirst
                rsHandle.Find "Handle='" & Handle & "'"
                If Not rsHandle.EOF Then
                    '03-10-11 ��� ������ ��� ������
                    'If rsHandle!Drilling >= 160 Then he = 0 Else he = 1
                    If rsHandle!Drilling > 160 Then he = 0 Else he = 1
                End If
            
            End If
        End If
        
        If IsEmpty(he) Then
            If MsgBox("��� ����� �� �������?", vbDefaultButton3 Or vbQuestion Or vbYesNo, "����� �� �������") = vbYes Then
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
        If InStr(1, Replace(ActiveCell.Offset(, 2), " ", ""), "��1���", vbTextCompare) Or _
            InStr(1, Replace(ActiveCell.Offset(, 3), " ", ""), "��1���", vbTextCompare) > 0 Or _
            InStr(1, Replace(ActiveCell.Offset(, 4), " ", ""), "��1���", vbTextCompare) > 0 Or InStr(1, Replace(ActiveCell.Offset(, 5), " ", ""), "��1���", vbTextCompare) > 0 Then
            
            CheckHandleExtra = 0
        ElseIf InStr(1, Replace(ActiveCell.Offset(, 2), " ", ""), "��2���", vbTextCompare) > 0 Or _
                InStr(1, Replace(ActiveCell.Offset(, 3), " ", ""), "��2���", vbTextCompare) > 0 Or _
                InStr(1, Replace(ActiveCell.Offset(, 4), " ", ""), "��2���", vbTextCompare) > 0 Or InStr(1, Replace(ActiveCell.Offset(, 5), " ", ""), "��2���", vbTextCompare) > 0 Then
            
            CheckHandleExtra = 1
        End If
    End If
End Function


Sub �������_�_��������()
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
             '���������
            ColorName = Replace(ColorName, "���", "����")
            ColorName = Replace(ColorName, "�����", "�����")
            ColorName = Replace(ColorName, "�����", "�����")
            ColorName = Replace(ColorName, "-16��", "")
            ColorName = Replace(ColorName, "-18��", " 18")
            ColorName = Replace(ColorName, "-16��", "")
            ColorName = Replace(ColorName, " 18��", " 18")
            ColorName = Replace(ColorName, "-18", " 18")
            ColorName = Replace(ColorName, "-16", "")
            ColorName = Replace(ColorName, " 16", "")
            
            If ColorName = "������" Then ColorName = "���������"
        
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


