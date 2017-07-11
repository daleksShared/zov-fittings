Attribute VB_Name = "parserDop"
Option Explicit
Option Compare Text

Public Sub getDrawerMountItem(ByVal localis18 As Boolean, _
                                    ByVal drawerString As String, _
                                    ByVal draweroption As String, _
                                    ByVal qty As Integer, _
                                    Optional ByVal fasadHeight As Integer = 0, _
                                    Optional ByVal fasadWidth As Integer = 0, _
                                    Optional ByVal fasadDepth As Integer = 0, _
                                    Optional ByVal elementOption As String = "" _
                                    )
Dim caseFur As caseFurniture
Dim caseElementItem As caseElement

Dim fTempLength As Integer
Dim naprVar As Variant
Dim naprList() As String
Dim naprQty As Integer
Dim naprOpt As String
'If Mid(drawerString, 1, 1) = "+" Then drawerString = "���" & drawerString
naprList = regexp_ReturnSearchArray(patCaseFasadesNapravlList, drawerString)
Dim naprItem As String
For Each naprVar In naprList
    naprItem = naprVar

If naprItem = "" Then naprItem = "���"
naprQty = qty
naprOpt = ""
If regexp_check(patGetNumberFirst, naprItem) Then
    naprQty = CInt(regexp_ReturnSearch(patGetNumberFirst, naprItem))
End If
If regexp_check(patGetNumberLast, naprItem) Then
    naprOpt = (regexp_ReturnSearch(patGetNumberLast, naprItem))
End If
If regexp_check(patGetStringTrimNumbers, naprItem) Then
    naprItem = (regexp_ReturnSearch(patGetStringTrimNumbers, naprItem))
End If

If naprOpt = "" And draweroption <> "" Then naprOpt = draweroption
If InStr(naprItem, "���") > 0 Or InStr(naprItem, "��-�����") > 0 Or InStr(naprItem, "��-�����") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","

    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 140 And fasadHeight < 210 Then
        caseFur.fName = "�������� �����"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = naprOpt & "0"
        Else
            caseFur.fLength = CStr(GetDrawerMountMb())
            If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    ElseIf fasadHeight >= 210 And fasadHeight < 714 Then
        caseFur.fName = "�������� �������"
        caseFur.fType = "drawermount"
         If naprOpt <> "" Then
        caseFur.fLength = naprOpt
        Else
        caseFur.fLength = CStr(GetDrawerMountMb())
        If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    End If
    caseFurnCollection.Add caseFur
    
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�������� �� ��������"
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ��� ����� 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "��") > 0 And InStr(naprItem, "�����") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","

    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.fName = "���������� ��� �����"
    caseFur.fType = "drawermount"
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ���"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
    
ElseIf naprItem = "��" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 135 And fasadHeight < 214 Then
        caseFur.fName = "���������� �����"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = naprOpt & "0"
        Else
            caseFur.fLength = CStr(GetDrawerMountTB())
            If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    ElseIf fasadHeight >= 215 And fasadHeight < 714 Then
        caseFur.fName = "���������� �������"
        caseFur.fType = "drawermount"
         If naprOpt <> "" Then
        caseFur.fLength = naprOpt & "0"
        Else
        caseFur.fLength = CStr(GetDrawerMountTB())
        If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    End If
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ��� ����� 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf naprItem = "��" And Len(naprItem) = 2 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 140 And fasadHeight < 210 Then
        caseFur.fName = "�������� �����"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = naprOpt & "0"
        Else
            caseFur.fLength = CStr(GetDrawerMountMb())
            If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    ElseIf fasadHeight >= 210 And fasadHeight < 714 Then
        caseFur.fName = "�������� �������"
        caseFur.fType = "drawermount"
         If naprOpt <> "" Then
        caseFur.fLength = naprOpt
        Else
        caseFur.fLength = CStr(GetDrawerMountMb())
        If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    End If
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ��� ����� 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "���") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If InStr(naprItem, "����") > 0 Then
        If localis18 Then
            caseFur.fName = "���������� �����. 18 ���"
            Else
            caseFur.fName = "���������� �����. 16 ���"
        End If
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
         caseFur.fLength = naprOpt & "0"
        Else
        caseFur.fLength = CStr(GetDrawerMountTB())
        
        End If
        
        caseFur.fOption = CStr(GetDrawerMountTB_vnutr_mal(fasadWidth, localis18))
        caseFurnCollection.Add caseFur
        
        Set caseElementItem = New caseElement
        caseElementItem.init
        caseElementItem.name = "~��� ��� ����� 1"
        caseElementItem.qty = caseFur.qty
        CaseElements.Add caseElementItem
    ElseIf InStr(naprItem, "����") > 0 Then
        If localis18 Then
            caseFur.fName = "���������� �����. 18 ���"
            Else
            caseFur.fName = "���������� �����. 16 ���"
        End If
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = naprOpt & "0"
        Else
            caseFur.fLength = CStr(GetDrawerMountTB())
        End If
        caseFur.fOption = CStr(GetDrawerMountTB_vnutr_bol(fasadWidth, localis18))
            
        caseFurnCollection.Add caseFur
        Set caseElementItem = New caseElement
        caseElementItem.init
        caseElementItem.name = "~��� ��� ����� 1"
        caseElementItem.qty = caseFur.qty
        CaseElements.Add caseElementItem
    End If
    
ElseIf InStr(naprItem, "����") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "����� � ������ ����"
    If fasadWidth = 60 Or fasadWidth = 600 Then
        caseFur.fOption = "60"
    ElseIf fasadWidth = 80 Or fasadWidth = 800 Then
        caseFur.fOption = "80"
    ElseIf fasadWidth = 90 Or fasadWidth = 900 Then
        caseFur.fOption = "90"
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ��� ����� 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
    
ElseIf InStr(naprItem, "��") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "������������ ������"
    caseFur.fType = "drawermount"
    If naprOpt <> "" Then
        caseFur.fOption = naprOpt
    Else
        caseFur.fOption = GetDrawerMountKv()
    End If

    If caseFur.fOption = 0 Then caseFur.fOption = ""
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ���"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "��") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "������������ ������"
    caseFur.fType = "drawermount"
    If naprOpt <> "" Then
        caseFur.fOption = naprOpt
    Else
        caseFur.fOption = GetDrawerMountKv()
    End If
    If caseFur.fOption = "0" Then caseFur.fOption = ""
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ��� ��"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "���") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "������������"
    caseFur.fType = "drawermount"
    If naprOpt <> "" Then
        caseFur.fOption = "����� " & naprOpt
    Else
        caseFur.fOption = "����� " & GetDrawerMount()
    End If
    If caseFur.fOption = "����� 0" Then caseFur.fOption = ""
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init

    caseElementItem.name = "~��� ���"
    If elementOption = "��������" Then caseElementItem.name = "������� ��������"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "���") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "������������"
    caseFur.fType = "drawermount"
    If naprOpt <> "" Then
        caseFur.fOption = "����� " & naprOpt
    Else
        caseFur.fOption = "����� " & GetDrawerMount()
    End If
    If caseFur.fOption = "����� 0" Then caseFur.fOption = ""
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ���"
    If elementOption = "��������" Then caseElementItem.name = "������� ��������"

    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(1, naprItem, "���", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� �������"
    If getArchitehLength(fasadDepth, localis18) = 500 Then
        caseFur.fLength = "500/78 ����"
        If Right(naprItem, 2) = "-�" Then
            caseFur.fOption = "��������"
        Else
            caseFur.fOption = "�����"
        End If
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "����") = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� ������� �����"
    If getArchitehLength(fasadDepth, localis18) = 500 Then
        caseFur.fLength = "500/186 ������"
    End If
    If Right(naprItem, 2) = "-�" Then
        caseFur.fOption = "��������"
    Else
        caseFur.fOption = "�����"
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "���1�", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� ������� �����"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/186 1����"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/186 1����"
    End If
    If Right(naprItem, 2) = "-�" Then
        caseFur.fOption = "��������"
    Else
        caseFur.fOption = "�����"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "���", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� ������� �����"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/94 ���"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/94 ���"
    End If
        If Right(naprItem, 2) = "-�" Then
        caseFur.fOption = "��������"
    Else
        caseFur.fOption = "�����"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "���", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� �������"
    caseFur.fLength = "500/186 ������"
    If Right(naprItem, 2) = "-�" Then
        caseFur.fOption = "��������"
    Else
        caseFur.fOption = "�����"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "��1�", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� �������"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/186 1����"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/186 1����"
    End If
    If Right(naprItem, 2) = "-�" Then
        caseFur.fOption = "��������"
    Else
        caseFur.fOption = "�����"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "��2�", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� �������"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/250 2����"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/250 2����"
    End If
    If Right(naprItem, 2) = "-�" Then
        caseFur.fOption = "��������"
    Else
        caseFur.fOption = "�����"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "��", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� �������"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/94 ���"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/94 ���"
    End If
    If Right(naprItem, 2) = "-�" Then
        caseFur.fOption = "��������"
    Else
        caseFur.fOption = "�����"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf naprItem = "���" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - ������� ��� �����"
    If casepropertyCurrent.p_cabWidth = 800 Then
        caseFur.fOption = "800"
    ElseIf casepropertyCurrent.p_cabWidth = 900 Then
        caseFur.fOption = "900"
    End If
'    If fas = 500 Then
'        caseFur.foption = "500/94 ���"
'    ElseIf fTempLength = 300 Then
'        caseFur.foption = "300/94 ���"
'    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "C" Or naprItem = "M" Or naprItem = "D" Or naprItem = "N") Or _
        naprItem = "��C" Or naprItem = "��M" Or naprItem = "��D" _
    Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� ��� ���"
    If localis18 Then
        If casepropertyCurrent.p_cabDepth >= 319 And casepropertyCurrent.p_cabDepth < 519 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "��", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 519 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "��", "")
        End If
    Else
        If casepropertyCurrent.p_cabDepth >= 303 And casepropertyCurrent.p_cabDepth < 503 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "��", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 503 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "��", "")
        End If
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf naprItem = "C-�����" Or naprItem = "M-�����" Or naprItem = "D-�����" Or naprItem = "N-�����" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� ��� ��� ��� �����"
    If localis18 Then
        If casepropertyCurrent.p_cabDepth >= 319 And casepropertyCurrent.p_cabDepth < 519 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "-�����", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 519 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "-�����", "")
        End If
    Else
        If casepropertyCurrent.p_cabDepth >= 303 And casepropertyCurrent.p_cabDepth < 503 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "-�����", "")
        ElseIf fasadWidth >= 503 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "-�����", "")
        End If
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "inC" Or naprItem = "inM" Or naprItem = "inD" Or naprItem = "inN") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "�� ��� ��� ����"
    If localis18 Then
        caseFur.fOption = "18"
        If casepropertyCurrent.p_cabDepth >= 319 And casepropertyCurrent.p_cabDepth < 519 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "in", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 519 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "in", "")
        End If
    Else
        caseFur.fOption = "16"
        If casepropertyCurrent.p_cabDepth >= 303 And casepropertyCurrent.p_cabDepth < 503 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "in", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 503 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "in", "")
        End If
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "���") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - ��������� �������"
        If casepropertyCurrent.p_cabWidth = 450 Then
            caseFur.fOption = "450 ��� ��.��. ��� �.��."
        ElseIf casepropertyCurrent.p_cabWidth = 600 Then
            caseFur.fOption = "600 ��� ��.��. ��� �.��."
        ElseIf casepropertyCurrent.p_cabWidth = 900 Then
            caseFur.fOption = "900 ��� ��.��. ��� �.��."
        End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - ����� ����� ��� ���"
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "�����") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - ��������� �������"
        If casepropertyCurrent.p_cabWidth = 450 Then
            caseFur.fOption = "450 ��� ��.��. ��� �.��."
        ElseIf casepropertyCurrent.p_cabWidth = 600 Then
            caseFur.fOption = "600 ��� ��.��. ��� �.��."
        ElseIf casepropertyCurrent.p_cabWidth = 900 Then
            caseFur.fOption = "900 ��� ��.��. ��� �.��."
        End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "���") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - ��������� �������"
        If casepropertyCurrent.p_cabWidth = 450 Then
            caseFur.fOption = "450 �� ���� �� � �������"
        ElseIf casepropertyCurrent.p_cabWidth = 600 Then
            caseFur.fOption = "600 �� ���� �� � �������"
        ElseIf casepropertyCurrent.p_cabWidth = 900 Then
            caseFur.fOption = "900 �� ���� �� � �������"
        End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf naprItem = "��" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 140 And fasadHeight < 210 Then
        caseFur.fName = "�������� �����"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = "������" & naprOpt & "0"
        Else
            caseFur.fLength = "������" & CStr(GetDrawerMountMb())
        End If
    ElseIf fasadHeight >= 210 And fasadHeight < 714 Then
        caseFur.fName = "�������� �������"
        caseFur.fType = "drawermount"
         If naprOpt <> "" Then
        caseFur.fLength = "������" & naprOpt
        Else
        caseFur.fLength = "������" & CStr(GetDrawerMountMb())
        If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    End If
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~��� ��� ����� 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "���") > 0 Or InStr(naprItem, "����") > 0 Or InStr(naprItem, "��") > 0 Then
'    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & "��" & ","
'    Set caseFur = New caseFurniture
'    caseFur.init
'    caseFur.Qty = naprQty
'    caseFur.fName = "��������� �� " & naprItem
'    caseFurnCollection.Add caseFur
Else
     Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = naprItem
    caseFur.fOption = naprOpt
    caseFurnCollection.Add caseFur
End If

Next naprVar

    
    


End Sub
                                    



