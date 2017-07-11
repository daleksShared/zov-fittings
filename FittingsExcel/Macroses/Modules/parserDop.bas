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
'If Mid(drawerString, 1, 1) = "+" Then drawerString = "шар" & drawerString
naprList = regexp_ReturnSearchArray(patCaseFasadesNapravlList, drawerString)
Dim naprItem As String
For Each naprVar In naprList
    naprItem = naprVar

If naprItem = "" Then naprItem = "шар"
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
If InStr(naprItem, "мбд") > 0 Or InStr(naprItem, "мб-довод") > 0 Or InStr(naprItem, "мб-довод") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","

    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 140 And fasadHeight < 210 Then
        caseFur.fName = "метабокс малый"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = naprOpt & "0"
        Else
            caseFur.fLength = CStr(GetDrawerMountMb())
            If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    ElseIf fasadHeight >= 210 And fasadHeight < 714 Then
        caseFur.fName = "метабокс большой"
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
    caseFur.fName = "доводчик на метабокс"
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~фур шуф ручка 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "тб") > 0 And InStr(naprItem, "мойка") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","

    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.fName = "тандембокс под мойку"
    caseFur.fType = "drawermount"
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~фур шуф"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
    
ElseIf naprItem = "тб" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 135 And fasadHeight < 214 Then
        caseFur.fName = "тандембокс малый"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = naprOpt & "0"
        Else
            caseFur.fLength = CStr(GetDrawerMountTB())
            If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    ElseIf fasadHeight >= 215 And fasadHeight < 714 Then
        caseFur.fName = "тандембокс большой"
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
    caseElementItem.name = "~фур шуф ручка 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf naprItem = "мб" And Len(naprItem) = 2 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 140 And fasadHeight < 210 Then
        caseFur.fName = "метабокс малый"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = naprOpt & "0"
        Else
            caseFur.fLength = CStr(GetDrawerMountMb())
            If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    ElseIf fasadHeight >= 210 And fasadHeight < 714 Then
        caseFur.fName = "метабокс большой"
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
    caseElementItem.name = "~фур шуф ручка 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "тбв") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If InStr(naprItem, "тбвм") > 0 Then
        If localis18 Then
            caseFur.fName = "тандембокс внутр. 18 мал"
            Else
            caseFur.fName = "тандембокс внутр. 16 мал"
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
        caseElementItem.name = "~фур шуф ручка 1"
        caseElementItem.qty = caseFur.qty
        CaseElements.Add caseElementItem
    ElseIf InStr(naprItem, "тбвб") > 0 Then
        If localis18 Then
            caseFur.fName = "тандембокс внутр. 18 бол"
            Else
            caseFur.fName = "тандембокс внутр. 16 бол"
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
        caseElementItem.name = "~фур шуф ручка 1"
        caseElementItem.qty = caseFur.qty
        CaseElements.Add caseElementItem
    End If
    
ElseIf InStr(naprItem, "сушк") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "Сушка в нижний шкаф"
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
    caseElementItem.name = "~фур шуф ручка 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
    
ElseIf InStr(naprItem, "кв") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "направляющие Квадро"
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
    caseElementItem.name = "~фур шуф"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "вп") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "направляющие Квадро"
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
    caseElementItem.name = "~фур шуф вп"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "шар") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "направляющие"
    caseFur.fType = "drawermount"
    If naprOpt <> "" Then
        caseFur.fOption = "шарик " & naprOpt
    Else
        caseFur.fOption = "шарик " & GetDrawerMount()
    End If
    If caseFur.fOption = "шарик 0" Then caseFur.fOption = ""
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init

    caseElementItem.name = "~фур шуф"
    If elementOption = "имитация" Then caseElementItem.name = "шуфляда имитация"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "рол") > 0 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "направляющие"
    caseFur.fType = "drawermount"
    If naprOpt <> "" Then
        caseFur.fOption = "ролик " & naprOpt
    Else
        caseFur.fOption = "ролик " & GetDrawerMount()
    End If
    If caseFur.fOption = "ролик 0" Then caseFur.fOption = ""
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~фур шуф"
    If elementOption = "имитация" Then caseElementItem.name = "шуфляда имитация"

    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(1, naprItem, "арг", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех"
    If getArchitehLength(fasadDepth, localis18) = 500 Then
        caseFur.fLength = "500/78 ШЛГП"
        If Right(naprItem, 2) = "-а" Then
            caseFur.fOption = "Антрацит"
        Else
            caseFur.fOption = "Белый"
        End If
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "арвс") = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех внутр"
    If getArchitehLength(fasadDepth, localis18) = 500 Then
        caseFur.fLength = "500/186 стекло"
    End If
    If Right(naprItem, 2) = "-а" Then
        caseFur.fOption = "Антрацит"
    Else
        caseFur.fOption = "Белый"
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "арв1р", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех внутр"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/186 1релл"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/186 1релл"
    End If
    If Right(naprItem, 2) = "-а" Then
        caseFur.fOption = "Антрацит"
    Else
        caseFur.fOption = "Белый"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "арв", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех внутр"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/94 мал"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/94 мал"
    End If
        If Right(naprItem, 2) = "-а" Then
        caseFur.fOption = "Антрацит"
    Else
        caseFur.fOption = "Белый"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "арс", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех"
    caseFur.fLength = "500/186 стекло"
    If Right(naprItem, 2) = "-а" Then
        caseFur.fOption = "Антрацит"
    Else
        caseFur.fOption = "Белый"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "ар1р", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/186 1релл"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/186 1релл"
    End If
    If Right(naprItem, 2) = "-а" Then
        caseFur.fOption = "Антрацит"
    Else
        caseFur.fOption = "Белый"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "ар2р", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/250 2релл"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/250 2релл"
    End If
    If Right(naprItem, 2) = "-а" Then
        caseFur.fOption = "Антрацит"
    Else
        caseFur.fOption = "Белый"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf InStr(1, naprItem, "ар", vbTextCompare) = 1 Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "ТБ Архитех"
    fTempLength = getArchitehLength(fasadDepth, localis18)
    If fTempLength = 500 Then
        caseFur.fLength = "500/94 мал"
    ElseIf fTempLength = 300 Then
        caseFur.fLength = "300/94 мал"
    End If
    If Right(naprItem, 2) = "-а" Then
        caseFur.fOption = "Антрацит"
    Else
        caseFur.fOption = "Белый"
    End If

    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf naprItem = "ВКМ" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - Корзина под мойку"
    If casepropertyCurrent.p_cabWidth = 800 Then
        caseFur.fOption = "800"
    ElseIf casepropertyCurrent.p_cabWidth = 900 Then
        caseFur.fOption = "900"
    End If
'    If fas = 500 Then
'        caseFur.foption = "500/94 мал"
'    ElseIf fTempLength = 300 Then
'        caseFur.foption = "300/94 мал"
'    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "C" Or naprItem = "M" Or naprItem = "D" Or naprItem = "N") Or _
        naprItem = "анC" Or naprItem = "анM" Or naprItem = "анD" _
    Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "Тб Ант бел"
    If localis18 Then
        If casepropertyCurrent.p_cabDepth >= 319 And casepropertyCurrent.p_cabDepth < 519 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "ан", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 519 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "ан", "")
        End If
    Else
        If casepropertyCurrent.p_cabDepth >= 303 And casepropertyCurrent.p_cabDepth < 503 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "ан", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 503 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "ан", "")
        End If
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf naprItem = "C-МОЙКА" Or naprItem = "M-МОЙКА" Or naprItem = "D-МОЙКА" Or naprItem = "N-МОЙКА" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "Тб Ант бел под мойку"
    If localis18 Then
        If casepropertyCurrent.p_cabDepth >= 319 And casepropertyCurrent.p_cabDepth < 519 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "-МОЙКА", "")
        ElseIf casepropertyCurrent.p_cabDepth >= 519 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "-МОЙКА", "")
        End If
    Else
        If casepropertyCurrent.p_cabDepth >= 303 And casepropertyCurrent.p_cabDepth < 503 Then
            caseFur.fLength = "300" & "/" & Replace(naprItem, "-МОЙКА", "")
        ElseIf fasadWidth >= 503 Then
            caseFur.fLength = "500" & "/" & Replace(naprItem, "-МОЙКА", "")
        End If
    End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "inC" Or naprItem = "inM" Or naprItem = "inD" Or naprItem = "inN") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "Тб Ант бел внут"
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
ElseIf (naprItem = "ВКФ") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - Выдвижная корзина"
        If casepropertyCurrent.p_cabWidth = 450 Then
            caseFur.fOption = "450 для фр.кр. без ф.кр."
        ElseIf casepropertyCurrent.p_cabWidth = 600 Then
            caseFur.fOption = "600 для фр.кр. без ф.кр."
        ElseIf casepropertyCurrent.p_cabWidth = 900 Then
            caseFur.fOption = "900 для фр.кр. без ф.кр."
        End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
    
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - Фронт крепл вдв крз"
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "ВКФвн") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - Выдвижная корзина"
        If casepropertyCurrent.p_cabWidth = 450 Then
            caseFur.fOption = "450 для фр.кр. без ф.кр."
        ElseIf casepropertyCurrent.p_cabWidth = 600 Then
            caseFur.fOption = "600 для фр.кр. без ф.кр."
        ElseIf casepropertyCurrent.p_cabWidth = 900 Then
            caseFur.fOption = "900 для фр.кр. без ф.кр."
        End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf (naprItem = "ВКР") Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    caseFur.fName = "VS - Выдвижная корзина"
        If casepropertyCurrent.p_cabWidth = 450 Then
            caseFur.fOption = "450 на расп дв с планкой"
        ElseIf casepropertyCurrent.p_cabWidth = 600 Then
            caseFur.fOption = "600 на расп дв с планкой"
        ElseIf casepropertyCurrent.p_cabWidth = 900 Then
            caseFur.fOption = "900 на расп дв с планкой"
        End If
    caseFur.qty = naprQty
    caseFurnCollection.Add caseFur
ElseIf naprItem = "мк" Then
    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & naprItem & ","
    Set caseFur = New caseFurniture
    caseFur.init
    caseFur.qty = naprQty
    If fasadHeight >= 140 And fasadHeight < 210 Then
        caseFur.fName = "метабокс малый"
        caseFur.fType = "drawermount"
        If naprOpt <> "" Then
            caseFur.fLength = "оптима" & naprOpt & "0"
        Else
            caseFur.fLength = "оптима" & CStr(GetDrawerMountMb())
        End If
    ElseIf fasadHeight >= 210 And fasadHeight < 714 Then
        caseFur.fName = "метабокс большой"
        caseFur.fType = "drawermount"
         If naprOpt <> "" Then
        caseFur.fLength = "оптима" & naprOpt
        Else
        caseFur.fLength = "оптима" & CStr(GetDrawerMountMb())
        If caseFur.fLength = "" Then caseFur.fLength = ""
        End If
    End If
    caseFurnCollection.Add caseFur
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "~фур шуф ручка 1"
    caseElementItem.qty = caseFur.qty
    CaseElements.Add caseElementItem
ElseIf InStr(naprItem, "лев") > 0 Or InStr(naprItem, "прав") > 0 Or InStr(naprItem, "дв") > 0 Then
'    casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & naprQty & "дв" & ","
'    Set caseFur = New caseFurniture
'    caseFur.init
'    caseFur.Qty = naprQty
'    caseFur.fName = "фурнитура на " & naprItem
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
                                    



