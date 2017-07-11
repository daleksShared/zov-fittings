Attribute VB_Name = "parser"
Option Explicit
Option Compare Text
Const NaprMainDefault As String = "шарик"
Public splitString As New Collection
Public splitStringItem As clSplitString
Public casefasades As New Collection
Public CaseElements As New Collection
Public caseFurnCollection As New Collection
Public casezones As New Collection

Private Sub trest()
Dim str As Variant
GetOtbGorbColors OtbGorbColors

'resStr = ""
'mRegexp.regexp_ReturnSearchCollection patSplitPattern, "Клипсы к цоколю пластик-24шт"

'For Each str In mRegexp.regexp_ReturnSearchArray(patSplitPattern, "!!! К цоколь пластик ХРОМ + 2 наружн угла + 1уг внутр + 3загл.")
'    resStr = RTrim(resStr & " " & LTrim(str))
'Next str
Dim resStr As String
resStr = parseUncloseBrackets("!!! Отбортовка ЧЕРНАЯ БРОНЗА (горб) 4м + 2загл(лев+прав) + 2уг внутр.")

While InStr(1, resStr, "(") > 0 And InStr(1, resStr, ")") > InStr(1, resStr, "(")
resStr = parseUncloseBrackets(resStr)
Wend
MsgBox (resStr)

'в первую очередь узнаю тип отбортовки.




'MsgBox (TestRegExp(patCaseFasades, "ШЛК80(713/1прав+176/1,176/1,176/1,176/1)"))
'ШЛК80(713/1прав+176/1,176/1,176/1,176/1)        ШЛК80/4
'MsgBox regexp_ReturnSearch("([0-9]+)", "ш328мб")
'MsgBox parseSKOBOCHKI("ШЛК80(713/1прав+176/1,176/1,176/1,176/1)")
'MsgBox parse_SHLSH("ШЛШ60(б.610)(ниша ,ш283мб) глуб35см")
'MsgBox parse_SHLSH("ШЛШ60(915)(193/1,355/1,355/1)тб")
'MsgBox Replace("ШН90(355/1,355/1)вверх", regexp_ReturnSearch(patGetSkobochki, "ШН90(355/1,355/1)вверх"), "/" & parseGetSumFromSKOBOCHKI(regexp_ReturnSearch(patGetSkobochki, "ШН90(355/1,355/1)вверх")))
End Sub
Function getColorForArrayFromString(ByVal rawstr As String, ByRef colorlist()) As String
Dim cCol As Collection
Set cCol = New Collection

SortArrayByLengthDesc colorlist
Dim i As Integer
Dim j As Integer
Dim tempStr As String
Dim tempStr2 As String

Dim tempStrArray() As String

For i = LBound(colorlist) To UBound(colorlist)
    tempStrArray() = Split(colorlist(i), " ")
    tempStr2 = ""
    For j = LBound(tempStrArray) To UBound(tempStrArray)
        Select Case UCase(Right(tempStrArray(j), 2))
            Case "ЯЯ", "АЯ", "ИЙ", "ЫЙ"
                tempStrArray(j) = Mid(tempStrArray(j), 1, Len(tempStrArray(j)) - 2)
        End Select
    If j < UBound(tempStrArray) Then tempStr2 = tempStr2 & tempStrArray(j) & " " Else tempStr2 = tempStr2 & tempStrArray(j)
    Next j
    
    cCol.Add tempStr2
Next i

Dim RawWord As String
Dim WordColor As String
Dim WordColors() As String
Dim strVar As Variant
Dim strVar2 As Variant

Dim WordColorsQty() As Integer
Dim WordColorsMatchQty() As Integer
Dim WordColorsMatch() As Single

ReDim WordColorsQty(cCol.Count)
ReDim WordColorsMatchQty(cCol.Count)
For i = 1 To cCol.Count
    WordColorsQty(i) = 0
    WordColorsMatchQty(i) = 0
Next i

ReDim WordColorsMatch(cCol.Count)
Dim WordColorsIterator As Integer

Dim isMatch As Boolean

Dim FullColor As String

For i = 1 To cCol.Count
    WordColors() = Split(cCol(i), " ", -1, vbTextCompare)
    WordColorsQty(i) = UBound(WordColors()) + 1
    For Each strVar In WordColors
        WordColor = CStr(strVar)
        isMatch = False
        For Each strVar2 In Split(rawstr, " ", -1, vbTextCompare)
            RawWord = CStr(strVar2)
            If Len(RawWord) > 2 And InStr(1, RawWord, WordColor, vbTextCompare) = 1 Then
                isMatch = True
                Exit For
            ElseIf Len(RawWord) > 2 And InStr(1, WordColor, RawWord, vbTextCompare) = 1 Then
                isMatch = True
                Exit For
            End If
        Next strVar2
        If isMatch Then
            WordColorsMatchQty(i) = WordColorsMatchQty(i) + 1
        End If
    Next strVar
    If WordColorsMatchQty(i) = WordColorsQty(i) Then
        getColorForArrayFromString = cCol(i)
        Exit For
    End If
Next i


End Function
Function parseUncloseBrackets(ByVal rawstring As String, Optional ByVal storeParentQty As Boolean = False)
Dim localPatGetQty As String
localPatGetQty = "^(\d+)\s*.*?"
Dim i As Integer
i = 1
Dim ii As Integer
Dim BracketStart As Integer
Dim BracketFinish As Integer


    rawstring = Replace(rawstring, "  ", " ")
    rawstring = Replace(rawstring, " (", "(")


'переменная x из строки .....x(a1+a2+a3)
Dim firstPart As String
Dim firstPartQty As String
Dim partsIterator As Integer

'найдём, открыв скобку
While i < Len(rawstring) And Mid(rawstring, i, 1) <> "("
    i = i + 1
Wend
BracketStart = i

While i < Len(rawstring) And Mid(rawstring, i, 1) <> ")"
    i = i + 1
Wend
BracketFinish = i

If ((BracketStart > 0) And (BracketStart < BracketFinish)) = False Then
    Exit Function
End If
i = BracketStart - 1

If InStr(1, Mid(rawstring, BracketStart + 1, BracketFinish - BracketStart - 1), "+") = 0 Then
parseUncloseBrackets = Trim(Mid(rawstring, 1, BracketStart - 1)) & " " & Trim(Mid(rawstring, BracketStart + 1, BracketFinish - BracketStart - 1)) & " " & (Mid(rawstring, BracketFinish + 1))
Exit Function
End If


While i > 0 And InStr(1, "0123456789-+", Mid(rawstring, i, 1)) = 0
i = i - 1
Wend



'нашли чтото перед скобками


Dim newStr As String
newStr = ""
newStr = Mid(rawstring, 1, i - 1)

If i > 0 Then
    Dim strVar As Variant
    Dim delQtyFromFirstPart As Boolean
    delQtyFromFirstPart = False
    Dim parts() As String
    Dim partsQty() As Integer
    ii = i
    If storeParentQty = False Then
        While IsNumeric(Mid(rawstring, ii, 1))
            ii = ii + 1
        Wend
    End If
    
    parts() = Split(Mid(rawstring, BracketStart + 1, BracketFinish - BracketStart - 1), "+")
    ReDim partsQty(UBound(parts))
    For partsIterator = 0 To UBound(parts)
        If mRegexp.regexp_ReturnSearch(localPatGetQty, parts(partsIterator)) <> "" Then
            partsQty(partsIterator) = CInt(mRegexp.regexp_ReturnSearch(localPatGetQty, parts(partsIterator)))
            parts(partsIterator) = Mid(parts(partsIterator), Len(mRegexp.regexp_ReturnSearch(localPatGetQty, parts(partsIterator))) + 1)
        Else
            partsQty(partsIterator) = 1
        End If
    Next partsIterator
     
     
    If storeParentQty = False Then
        For partsIterator = 0 To UBound(parts)
                newStr = newStr & CStr(partsQty(partsIterator)) & Mid(rawstring, ii, BracketStart - ii) & " " & Trim(CStr(parts(partsIterator))) & " + "
        Next partsIterator
    Else
        For partsIterator = 0 To UBound(parts)
            newStr = newStr & Mid(rawstring, i, BracketStart - i) & " " & Trim(CStr(parts(partsIterator))) & " + "
        Next partsIterator
    End If
    

End If
If BracketFinish < Len(rawstring) Then
    If InStr(1, (Mid(rawstring, BracketFinish + 1, 2)), "+", vbTextCompare) > 0 And InStr(1, (Mid(rawstring, BracketFinish + 1, 2)), "+", vbTextCompare) < 3 Then
        newStr = newStr & Trim(Mid(rawstring, BracketFinish + 1 + InStr(1, Trim(Mid(rawstring, BracketFinish + 1, 2)), "+", vbTextCompare) + 1))
       
    Else
     newStr = newStr & Mid(rawstring, BracketFinish + 1)
    End If
Else
    newStr = Mid(newStr, 1, Len(newStr) - 3)
End If
parseUncloseBrackets = newStr
End Function
Sub parseTsokol(ByVal OrderId As Long, ByVal row As Integer)
Dim localPatGetLength As String
Dim localPatGetQty As String
localPatGetQty = "(\d+)\s*(?:шт)"
localPatGetLength = "(\d+)\s?(?:(м|m)(?!(м|m)))"

Dim cPart As clSplitString
Dim name As String
Dim i As Integer
Dim k As Integer, k1 As Integer, k2 As Integer, k3 As Integer
Dim qty, Opt, rawstring As String, isTsokol As Boolean, isTsokolVolpato As Boolean
Dim strVar1 As String, strVar2 As String, strVar3 As String
Dim checkCokolInd As Integer
Dim CokolInd As Integer
If splitString.Count = 0 Then Exit Sub

Dim tsokolHeigth As Integer
tsokolHeigth = 100

For i = 1 To splitString.Count
    If (i Mod 2) = 0 Then
        ActiveCell.Characters(splitString(i).start, splitString(i).length).Font.Color = vbBlue
    Else
        ActiveCell.Characters(splitString(i).start, splitString(i).length).Font.Color = vbRed
    End If
Next i

'в первом куске будет или сам цоколь или указание на него и цвет
Set cPart = splitString(1)
isTsokol = True
isTsokolVolpato = False
Opt = Null
qty = 0
rawstring = cPart.str
If Left(rawstring, 1) = "№" Then
    While Left(rawstring, 1) <> " " And Left(rawstring, 1) <> "."
    rawstring = Mid(rawstring, 2)
    Wend
    If Left(rawstring, 1) = "." Then rawstring = Mid(rawstring, 2)
    If Left(rawstring, 1) = " " Then rawstring = Mid(rawstring, 2)
    
End If

While InStr(1, rawstring, "!", vbTextCompare) = 1
rawstring = Mid(rawstring, 2)
Wend
While InStr(1, rawstring, " ", vbTextCompare) = 1
rawstring = Mid(rawstring, 2)
Wend

rawstring = Replace(rawstring, "ё", "е", , , vbTextCompare)
'теперь узнаю: это цоколь? или это к цоколю?
k = 0
k = InStr(1, rawstring, "к цок", vbTextCompare)
If k = 0 Then k = InStr(1, rawstring, "на цок", vbTextCompare)
If k = 0 Then k = InStr(1, rawstring, "для цок", vbTextCompare)
If k > 0 Then
    'это что-то для цоколя
    isTsokol = False
End If

    ' в любом случае разберусь с цветом и типом цоколя
    'найду конец для "пластик"
    k1 = 0
    k1 = InStr(k + 4, rawstring, "плас")
    If k1 > 0 Then
        name = "цоколь пластик"
        While Mid(rawstring, k1, 1) <> " " And k1 < Len(rawstring)
        k1 = k1 + 1
        Wend
    End If
    
    'а может имеем дело с Волпато?
    If k1 = 0 Then
        k1 = InStr(k + 4, rawstring, "волп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "вольп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "вальп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "валп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "volp")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "valp")
        If k1 > 0 Then
            While Mid(rawstring, k1, 1) <> " " And k1 < Len(rawstring)
                k1 = k1 + 1
            Wend
        name = "цоколь Волпато"
        End If
    End If
    
    'обычно следом идёт цвет
    If k1 > 0 And isTsokolVolpato = False Then
        k2 = 0
        k2 = InStr(k1, rawstring, "ХРОМ", vbTextCompare)
        If k2 > 0 Then Opt = "ХРОМ"
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "БЕЛЫЙ", vbTextCompare)
            
            If k2 > 0 Then
                If InStr(k1, rawstring, "МАТ", vbTextCompare) Then
                    Opt = "БЕЛЫЙмат"
                Else
                    Opt = "БЕЛЫЙгл"
                End If

            End If
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "БУК", vbTextCompare)
            If k2 > 0 Then Opt = "БУК"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ВЕНГЕ", vbTextCompare)
            If k2 > 0 Then Opt = "ВЕНГЕ"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ГРУША", vbTextCompare)
            If k2 > 0 Then Opt = "ГРУША"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "КЛЕН", vbTextCompare)
            If k2 > 0 Then Opt = "КЛЕН"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ОЛЬХА", vbTextCompare)
            If k2 > 0 Then Opt = "ОЛЬХА"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ОРЕХ", vbTextCompare)
            If k2 > 0 Then Opt = "ОРЕХ"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "МАХОНЬ", vbTextCompare)
            If k2 > 0 Then Opt = "МАХОНЬ"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "КРЕМ", vbTextCompare)
            If k2 > 0 Then Opt = "КРЕМ"
        End If
         If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ЧЕРН", vbTextCompare)
            If k2 > 0 Then Opt = "ЧЁРНЫЙгл"
        End If
         If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "РУСТИК", vbTextCompare)
            If k2 > 0 Then Opt = "РУСТИК"
        End If
         If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "Сосна", vbTextCompare)
            If k2 > 0 Then Opt = "СОСНАсветл"
        End If
         If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "Ясень", vbTextCompare)
            If k2 > 0 Then Opt = "ЯСЕНЬ"
        End If
        If k2 > 0 Then
            If InStr(k1, rawstring, "150", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "15см", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "=15", vbTextCompare) > 0 _
            Then
                tsokolHeigth = 150
                Opt = Opt & "150"
            Else
                Opt = Opt & "100"
            End If
        End If
    ElseIf k1 > 0 And isTsokolVolpato = True Then
         If InStr(k1, rawstring, "150", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "15см", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "=15", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "0/15", vbTextCompare) > 0 _
            Then
                tsokolHeigth = 150
                name = name & "150 Ал"
            Else
                name = name & "100 Ал"
            End If
    
    End If


    'название и тип/цвет есть. определюсь с количеством если имею дело с цоколем
    k3 = Len(rawstring)
    While k3 > 0 And InStr(1, " )", Mid(rawstring, k3, 1)) = 0
        k3 = k3 - 1
    Wend
    
    
    'получил какоето обозначение колва
    'проверю наличие шт
    qty = 0
    If mRegexp.regexp_ReturnSearch(localPatGetQty, Mid(rawstring, k3)) <> "" Then
        qty = CInt(mRegexp.regexp_ReturnSearch(localPatGetQty, Mid(rawstring, k3)))
    End If
    
    
    'проверю наличие шт
    Dim mqty As Integer
    mqty = 0
    
    If mRegexp.regexp_ReturnSearch(localPatGetLength, Mid(rawstring, k1)) <> "" Then
        mqty = CInt(mRegexp.regexp_ReturnSearch(localPatGetLength, Mid(rawstring, k1)))
    End If
    
    If mqty > 0 Then
        If (mqty Mod 3) <> 0 Then mqty = 0 'трехметровый цоколь
    End If
    If mqty = 0 And qty = 0 Then
        qty = Empty
    ElseIf mqty > 0 And qty = 0 Then
        qty = mqty / 3
    End If
    If qty = 0 Then qty = Null
    If isTsokol Then
        FormFitting.AddFittingToOrder OrderId, name, qty, Opt, , , , row
    ElseIf InStr(1, rawstring, "заг", vbTextCompare) > 0 Then 'Если не цоколь то проверю что же это?
        If qty = 0 Then
            k1 = InStr(1, rawstring, "заг", vbTextCompare)
            qty = CInt(mRegexp.regexp_ReturnSearch(".*\s*(\d+)\s*$", Mid(rawstring, 1, k1 - 1)))
        End If
        FormFitting.AddFittingToOrder OrderId, "заглушка к цоколю", qty, Opt, , , , row
    ElseIf InStr(1, rawstring, "уг", vbTextCompare) > 0 Then
        If qty = 0 Then
            k1 = InStr(1, rawstring, "уг", vbTextCompare)
            qty = CInt(mRegexp.regexp_ReturnSearch(".*\s*(\d+)\s*$", Mid(rawstring, 1, k1 - 1)))
        End If
        If isTsokolVolpato Then
            If tsokolHeigth = 150 Then
                FormFitting.AddFittingToOrder OrderId, "угол90 к цок Волпато150", qty, Opt, , , , row
            Else
                FormFitting.AddFittingToOrder OrderId, "угол90 к цоколю Волпато", qty, Opt, , , , row
            End If
        ElseIf InStr(1, Left(rawstring, 8), "90", vbTextCompare) Then
            FormFitting.AddFittingToOrder OrderId, "угол90* к цоколю", qty, Opt, , , , row
        ElseIf InStr(1, Left(rawstring, 9), "135", vbTextCompare) Then
            FormFitting.AddFittingToOrder OrderId, "угол135* к цоколю", qty, Opt, , , , row
        Else
            FormFitting.AddFittingToOrder OrderId, "угол90* к цоколю", Empty, Opt, , , , row
        End If
    ElseIf InStr(1, Left(rawstring, 4), "соед", vbTextCompare) > 0 Or InStr(1, Left(rawstring, 4), "стык", vbTextCompare) > 0 Then
        If qty = 0 Then
            k1 = InStr(1, Left(rawstring, 4), "соед", vbTextCompare) Or InStr(1, Left(rawstring, 4), "стык", vbTextCompare)
            qty = CInt(mRegexp.regexp_ReturnSearch(".*\s*(\d+)\s*$", Mid(rawstring, 1, k1 - 1)))
        End If
        FormFitting.AddFittingToOrder OrderId, "соединитель цоколя", qty, Opt, , , , row
    End If
    
    If splitString.Count > 1 Then
    For i = 2 To splitString.Count
        Set cPart = splitString(i)
        rawstring = cPart.str
        qty = 0
        If mRegexp.regexp_check("^(\d+).*", rawstring) Then
            qty = CInt(mRegexp.regexp_ReturnSearch("^(\d+).*", rawstring))
        Else
            qty = Empty
        End If
        k1 = 1
        While k1 < Len(rawstring) And IsNumeric(Mid(rawstring, k1, 1))
            k1 = k1 + 1
        Wend
        rawstring = Mid(rawstring, k1)
        If InStr(1, rawstring, "заг", vbTextCompare) > 0 Then
            If isTsokolVolpato Then
            FormFitting.AddFittingToOrder OrderId, "угл+загл к отб Волпато", qty, Opt, , , , row
            Else
            FormFitting.AddFittingToOrder OrderId, "заглушка к цоколю", qty, Opt, , , , row
            End If
        ElseIf InStr(1, rawstring, "уг", vbTextCompare) > 0 Then
            If isTsokolVolpato Then
                If tsokolHeigth = 150 Then
                    FormFitting.AddFittingToOrder OrderId, "угол90 к цок Волпато150", qty, Opt, , , , row
                Else
                    FormFitting.AddFittingToOrder OrderId, "угол90 к цоколю Волпато", qty, Opt, , , , row
                End If
            ElseIf InStr(1, Left(rawstring, 8), "90", vbTextCompare) Then
                FormFitting.AddFittingToOrder OrderId, "угол90* к цоколю", qty, Opt, , , , row
            ElseIf InStr(1, Left(rawstring, 9), "135", vbTextCompare) Then
                FormFitting.AddFittingToOrder OrderId, "угол135* к цоколю", qty, Opt, , , , row
            Else
                FormFitting.AddFittingToOrder OrderId, "угол90* к цоколю", qty, Opt, , , , row
            End If
        ElseIf InStr(1, Left(rawstring, 4), "соед", vbTextCompare) > 0 Or InStr(1, Left(rawstring, 4), "стык", vbTextCompare) > 0 Then
            FormFitting.AddFittingToOrder OrderId, "соединитель цоколя", qty, Opt, , , , row
        End If
    Next i
    End If


        
        
End Sub
Sub parseOtbort(ByVal OrderId As Long, ByVal row As Integer)
Dim localPatGetLength As String
Dim localPatGetQty As String
localPatGetQty = "(\d+)\s*(?:шт)"
localPatGetLength = "(\d+)\s*(?:м)"

Dim cPart As clSplitString
Dim name As String
Dim i As Integer
Dim k As Integer, k1 As Integer, k2 As Integer, k3 As Integer
Dim qty As Integer, Opt, rawstring As String, isTsokol As Boolean, isTsokolVolpato As Boolean
Dim strVar1 As String, strVar2 As String, strVar3 As String
Dim checkCokolInd As Integer
Dim CokolInd As Integer
If splitString.Count = 0 Then Exit Sub

Dim tsokolHeigth As Integer
tsokolHeigth = 100

For i = 1 To splitString.Count
    If (i Mod 2) = 0 Then
        ActiveCell.Characters(splitString(i).start, splitString(i).length).Font.Color = vbBlue
    Else
        ActiveCell.Characters(splitString(i).start, splitString(i).length).Font.Color = vbRed
    End If
Next i

'в первом куске будет или сам цоколь или указание на него и цвет
Set cPart = splitString(1)
isTsokol = True
isTsokolVolpato = False
Opt = Null
qty = 0
rawstring = cPart.str
While InStr(1, rawstring, "!", vbTextCompare) = 1
rawstring = Mid(rawstring, 2)
Wend
While InStr(1, rawstring, " ", vbTextCompare) = 1
rawstring = Mid(rawstring, 2)
Wend
rawstring = Replace(rawstring, "ё", "е", , , vbTextCompare)
'теперь узнаю: это цоколь? или это к цоколю?
k = 0
k = InStr(1, rawstring, "к цок", vbTextCompare)
If k = 0 Then k = InStr(1, rawstring, "на цок", vbTextCompare)
If k = 0 Then k = InStr(1, rawstring, "для цок", vbTextCompare)
If k > 0 Then
    'это что-то для цоколя
    isTsokol = False
End If

    ' в любом случае разберусь с цветом и типом цоколя
    'найду конец для "пластик"
    k1 = 0
    k1 = InStr(k + 4, rawstring, "плас")
    If k1 > 0 Then
        name = "цоколь пластик"
        While Mid(rawstring, k1, 1) <> " " And k1 < Len(rawstring)
        k1 = k1 + 1
        Wend
    End If
    
    'а может имеем дело с Волпато?
    If k1 = 0 Then
        k1 = InStr(k + 4, rawstring, "волп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "вольп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "вальп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "валп")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "volp")
        If k1 = 0 Then k1 = InStr(k + 4, rawstring, "valp")
        If k1 > 0 Then
            While Mid(rawstring, k1, 1) <> " " And k1 < Len(rawstring)
                k1 = k1 + 1
            Wend
        name = "цоколь Волпато"
        End If
    
    End If
    
    'обычно следом идёт цвет
    If k1 > 0 And isTsokolVolpato = False Then
        k2 = 0
        k2 = InStr(k1, rawstring, "ХРОМ", vbTextCompare)
        If k2 > 0 Then Opt = "ХРОМ"
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "БЕЛЫЙ", vbTextCompare)
            If k2 > 0 Then Opt = "БЕЛЫЙгл"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "БУК", vbTextCompare)
            If k2 > 0 Then Opt = "БУК"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ВЕНГЕ", vbTextCompare)
            If k2 > 0 Then Opt = "ВЕНГЕ"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ГРУША", vbTextCompare)
            If k2 > 0 Then Opt = "ГРУША"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "КЛЕН", vbTextCompare)
            If k2 > 0 Then Opt = "КЛЕН"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ОЛЬХА", vbTextCompare)
            If k2 > 0 Then Opt = "ОЛЬХА"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ОРЕХ", vbTextCompare)
            If k2 > 0 Then Opt = "ОРЕХ"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "МАХОНЬ", vbTextCompare)
            If k2 > 0 Then Opt = "МАХОНЬ"
        End If
        If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "КРЕМ", vbTextCompare)
            If k2 > 0 Then Opt = "КРЕМ"
        End If
         If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "ЧЕРН", vbTextCompare)
            If k2 > 0 Then Opt = "ЧЁРНЫЙгл"
        End If
         If IsNull(Opt) Then
            k2 = InStr(k1, rawstring, "РУСТИК", vbTextCompare)
            If k2 > 0 Then Opt = "РУСТИК"
        End If
        If k2 > 0 Then
            If InStr(k1, rawstring, "150", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "15см", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "=15", vbTextCompare) > 0 _
            Then
                tsokolHeigth = 150
                Opt = Opt & "150"
            Else
                Opt = Opt & "100"
            End If
        End If
    ElseIf k1 > 0 And isTsokolVolpato = True Then
         If InStr(k1, rawstring, "150", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "15см", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "=15", vbTextCompare) > 0 _
                Or InStr(k1, rawstring, "0/15", vbTextCompare) > 0 _
            Then
                tsokolHeigth = 150
                name = name & "150 Ал"
            Else
                name = name & "100 Ал"
            End If
    
    End If


    'название и тип/цвет есть. определюсь с количеством если имею дело с цоколем
    k3 = Len(rawstring)
    While k3 > 0 And InStr(1, " )", Mid(rawstring, k3, 1)) = 0
        k3 = k3 - 1
    Wend
    
    
    'получил какоето обозначение колва
    'проверю наличие шт
    qty = 0
    If mRegexp.regexp_ReturnSearch(localPatGetQty, Mid(rawstring, k3)) <> "" Then
        qty = CInt(mRegexp.regexp_ReturnSearch(localPatGetQty, Mid(rawstring, k3)))
    End If
    
    
    'проверю наличие шт
    Dim mqty As Integer
    mqty = 0
    
    If mRegexp.regexp_ReturnSearch(localPatGetLength, Mid(rawstring, k3)) <> "" Then
        mqty = CInt(mRegexp.regexp_ReturnSearch(localPatGetLength, Mid(rawstring, k3)))
    End If
    
    If mqty > 0 Then
        If (mqty Mod 3) <> 0 Then mqty = 0 'трехметровый цоколь
    End If
    If mqty = 0 And qty = 0 Then
        qty = Empty
    ElseIf mqty > 0 And qty = 0 Then
        qty = mqty / 3
    End If
    
    If isTsokol Then
        FormFitting.AddFittingToOrder OrderId, name, qty, Opt, , , , row
    ElseIf InStr(1, rawstring, "заг", vbTextCompare) > 0 Then 'Если не цоколь то проверю что же это?
        If qty = 0 Then
            k1 = InStr(1, rawstring, "заг", vbTextCompare)
            qty = CInt(mRegexp.regexp_ReturnSearch(".*\s*(\d+)\s*$", Mid(rawstring, 1, k1 - 1)))
        End If
        FormFitting.AddFittingToOrder OrderId, "заглушка к цоколю", qty, Opt, , , , row
    ElseIf InStr(1, rawstring, "уг", vbTextCompare) > 0 Then
        If qty = 0 Then
            k1 = InStr(1, rawstring, "уг", vbTextCompare)
            qty = CInt(mRegexp.regexp_ReturnSearch(".*\s*(\d+)\s*$", Mid(rawstring, 1, k1 - 1)))
        End If
        If isTsokolVolpato Then
            If tsokolHeigth = 150 Then
                FormFitting.AddFittingToOrder OrderId, "угол90 к цок Волпато150", qty, Opt, , , , row
            Else
                FormFitting.AddFittingToOrder OrderId, "угол90 к цоколю Волпато", qty, Opt, , , , row
            End If
        ElseIf InStr(1, Left(rawstring, 8), "90", vbTextCompare) Then
            FormFitting.AddFittingToOrder OrderId, "угол90* к цоколю", qty, Opt, , , , row
        ElseIf InStr(1, Left(rawstring, 9), "135", vbTextCompare) Then
            FormFitting.AddFittingToOrder OrderId, "угол135* к цоколю", qty, Opt, , , , row
        Else
            FormFitting.AddFittingToOrder OrderId, "угол90* к цоколю", Empty, Opt, , , , row
        End If
    ElseIf InStr(1, Left(rawstring, 4), "соед", vbTextCompare) > 0 Or InStr(1, Left(rawstring, 4), "стык", vbTextCompare) > 0 Then
        If qty = 0 Then
            k1 = InStr(1, Left(rawstring, 4), "соед", vbTextCompare) Or InStr(1, Left(rawstring, 4), "стык", vbTextCompare)
            qty = CInt(mRegexp.regexp_ReturnSearch(".*\s*(\d+)\s*$", Mid(rawstring, 1, k1 - 1)))
        End If
        FormFitting.AddFittingToOrder OrderId, "соединитель цоколя", qty, Opt, , , , row
    End If
    
    If splitString.Count > 1 Then
    For i = 2 To splitString.Count
        Set cPart = splitString(i)
        rawstring = cPart.str
        qty = 0
        If mRegexp.regexp_check("^(\d+).*", rawstring) Then
            qty = CInt(mRegexp.regexp_ReturnSearch("^(\d+).*", rawstring))
        Else
            qty = Empty
        End If
        k1 = 1
        While k1 < Len(rawstring) And IsNumeric(Mid(rawstring, k1, 1))
            k1 = k1 + 1
        Wend
        rawstring = Mid(rawstring, k1)
        If InStr(1, rawstring, "загл", vbTextCompare) > 0 Then
            If isTsokolVolpato Then
            FormFitting.AddFittingToOrder OrderId, "угл+загл к отб Волпато", qty, Opt, , , , row
            Else
            FormFitting.AddFittingToOrder OrderId, "заглушка к цоколю", qty, Opt, , , , row
            End If
        ElseIf InStr(1, rawstring, "уг", vbTextCompare) > 0 Then
            If isTsokolVolpato Then
                If tsokolHeigth = 150 Then
                    FormFitting.AddFittingToOrder OrderId, "угол90 к цок Волпато150", qty, Opt, , , , row
                Else
                    FormFitting.AddFittingToOrder OrderId, "угол90 к цоколю Волпато", qty, Opt, , , , row
                End If
            ElseIf InStr(1, Left(rawstring, 8), "90", vbTextCompare) Then
                FormFitting.AddFittingToOrder OrderId, "угол90* к цоколю", qty, Opt, , , , row
            ElseIf InStr(1, Left(rawstring, 9), "135", vbTextCompare) Then
                FormFitting.AddFittingToOrder OrderId, "угол135* к цоколю", qty, Opt, , , , row
            Else
                FormFitting.AddFittingToOrder OrderId, "угол к цоколю", Empty, Opt, , , , row
            End If
        ElseIf InStr(1, Left(rawstring, 4), "соед", vbTextCompare) > 0 Or InStr(1, Left(rawstring, 4), "стык", vbTextCompare) > 0 Then
            FormFitting.AddFittingToOrder OrderId, "соединитель цоколя", qty, Opt, , , , row
        End If
    Next i
    End If


        
        
End Sub
Public Function parseSKOBOCHKI(ByVal inputname As String) As String
    Dim RetStr As String
    Dim myPattern As String
    Dim mystring As String
    
    myPattern = "^[А-ЯаЯ]+[0-9]+(?:[(][0-9]+[)])?[\(]{1}(.*[,].*)+[\)]{1}"
    RetStr = inputname
    mystring = inputname

    Dim objRegExp As regexp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim objSubmatches As SubMatches
    Set objRegExp = New regexp
    objRegExp.Pattern = myPattern
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    If (objRegExp.Test(mystring) = True) Then
        Set colMatches = objRegExp.Execute(mystring)
        If colMatches.Count > 0 Then
            Set objMatch = colMatches.Item(0)
            Set objSubmatches = objMatch.SubMatches
            If objSubmatches.Count = 1 Then
                RetStr = Replace(inputname, "(" & objSubmatches.Item(0) & ")", "/" & CStr(regexp_count(patZPT, objSubmatches.Item(0)) + 1))
            End If
        End If
    End If
    
    parseSKOBOCHKI = RetStr
End Function
Public Function parseGetSumFromSKOBOCHKI(ByVal inputname As String) As Integer
    Dim Ret As Integer
    Dim myPattern As String
    Dim mystring As String
    
    myPattern = patCountInSkobochkiAfterSlash
    Ret = 0
    mystring = inputname
    Dim i  As Integer
    Dim objRegExp As regexp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim objSubmatches As SubMatches
    Set objRegExp = New regexp
    objRegExp.Pattern = myPattern
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    If (objRegExp.Test(mystring) = True) Then
        Set colMatches = objRegExp.Execute(mystring)

        For Each objMatch In colMatches
        Set objSubmatches = objMatch.SubMatches
        For i = 0 To objSubmatches.Count - 1
            If IsNumeric(objSubmatches.Item(i)) Then
                Ret = Ret + CInt(objSubmatches.Item(i))
            End If
        Next i
        Next
    End If
    
    parseGetSumFromSKOBOCHKI = Ret
End Function

Public Function parseShtQtyfromString(ByVal rawdata As String)
Dim localStrValue As String
localStrValue = regexp_ReturnSearch(patGetShtQtyfromStringBegin, rawdata)
If IsNumeric(localStrValue) Then
    parseShtQtyfromString = CInt(localStrValue)
    If parseShtQtyfromString <= 0 Then parseShtQtyfromString = Empty
Else
    parseShtQtyfromString = Empty
End If
End Function


Public Function parse_case(ByVal inputname As String) As String
    Dim result As String
    Dim stringBeforeFasades As String
    stringBeforeFasades = ""
    Dim stringAfterFasades As String
    stringAfterFasades = ""
    
    result = inputname
    
    
    
    If Mid(inputname, 1, 4) = "ШЛГП" And mRegexp.regexp_check(patCaseFasades, casepropertyCurrent.p_casename) Then
        If mRegexp.regexp_check(patGetBase, casepropertyCurrent.p_casename) Then casepropertyCurrent.p_baseString = "(" & mRegexp.regexp_ReturnSearch(patGetBaseValue, casepropertyCurrent.p_casename) & ")"
        parseFasadesSHLGP (mRegexp.regexp_ReturnSearch(patCaseFasades, casepropertyCurrent.p_casename))
        result = mRegexp.regexp_replace(patCaseFasades, casepropertyCurrent.p_casename, casepropertyCurrent.p_FasadesString)
    ElseIf Mid(inputname, 1, 3) = "ШЛШ" And mRegexp.regexp_check(patCaseFasades, casepropertyCurrent.p_casename) Then
         
        parseFasadesSHLSH (mRegexp.regexp_ReturnSearch(patCaseFasades, casepropertyCurrent.p_casename))
        result = casepropertyCurrent.p_newname
'        result = mRegexp.regexp_replace(patCaseFasades, inputname, casepropertyCurrent.p_FasadesString)
    ElseIf regexp_check(patSHL_check2, inputname) Then
        result = Replace(result, regexp_ReturnSearch(patSHL_check2, inputname), "")
    ElseIf Mid(inputname, 1, 2) = "ШН" And mRegexp.regexp_check(patCaseFasades, inputname) Then 'Or Mid(inputname, 1, 2) = "ШН" Then
        parseFasadesSHN (mRegexp.regexp_ReturnSearch(patCaseFasades, inputname))
        result = mRegexp.regexp_replace(patCaseFasadesOnlyString, inputname, casepropertyCurrent.p_FasadesString)
        stringAfterFasades = mRegexp.regexp_ReturnSearch(patCaseStringAfterFasades, inputname)
        If (casepropertyCurrent.p_baseString <> "") Then
            result = Replace(result, casepropertyCurrent.p_baseString, "") & casepropertyCurrent.p_baseString
            If stringAfterFasades <> "" Then result = Replace(result, stringAfterFasades, "") & stringAfterFasades
        End If
    ElseIf Mid(inputname, 1, 2) = "ШЛ" Or regexp_check(patSHL_check1, inputname) Then
        parseFasadesSHN (mRegexp.regexp_ReturnSearch(patCaseFasades, inputname))
        result = parse_SHL(inputname)
    End If
    casepropertyCurrent.p_newname = result
    parse_case = result
End Function

Function parse_SHL(ByVal casename_input As String) As String
Dim casename As String

If mRegexp.regexp_check(patGetBase, casename_input) Then
    casename = mRegexp.regexp_replace(patGetBase, casename_input, "")
Else
    casename = casename_input
End If
'casename = case_removeBaze(casename_input)

Dim cnameNew As String
 cnameNew = ""

Dim CnameNewFinishPart As String
CnameNewFinishPart = ""

Dim caseWidth As String
Dim cpos As Integer
cpos = 1

While Not IsNumeric(Mid(casename, cpos, 1))
cnameNew = cnameNew & Mid(casename, cpos, 1)
cpos = cpos + 1

Wend
While IsNumeric(Mid(casename, cpos, 1))
cnameNew = cnameNew & Mid(casename, cpos, 1)
cpos = cpos + 1

Wend

Dim start_range As Integer
Dim end_range As Integer
Dim range_data As String
Dim range_data_pos As Integer


Dim case_fasades As New Collection
Dim cf As casefasade



If Mid(casename, cpos, 1) = "(" Then
    cpos = cpos + 1
    start_range = cpos
    end_range = cpos
    While Mid(casename, cpos, 1) <> ")"
        If Mid(casename, cpos, 1) = "," Or Mid(casename, cpos + 1, 1) = ")" Then
            Set cf = New casefasade
            end_range = cpos
            range_data = Replace(Mid(casename, start_range, end_range - start_range + 1), ",", "")
            
            'разберу фасад
            If InStr(1, Replace(range_data, " ", ""), "/", vbTextCompare) > 0 Then
                cf.name = Mid(range_data, 1, InStr(1, range_data, "/", vbTextCompare) - 1)
                    If IsNumeric(cf.name) Then
                        cf.size = CInt(cf.name)
                    Else
                        cf.size = 0
                    End If
                cf.dopinfo = Mid(range_data, InStr(1, range_data, "/", vbTextCompare) + 1, Len(range_data) - InStr(1, range_data, "/", vbTextCompare))
                range_data_pos = 1
                cf.qty_string = ""
                While IsNumeric(Mid(cf.dopinfo, range_data_pos, 1))
                cf.qty_string = cf.qty_string & Mid(cf.dopinfo, range_data_pos, 1)
                   range_data_pos = range_data_pos + 1
                Wend
                If Len(cf.qty_string) > 0 Then
                cf.qty = CInt(cf.qty_string)
                   If range_data_pos > Len(cf.dopinfo) Then
                      cf.dopinfo = ""
                      Else
                       cf.dopinfo = Mid(cf.dopinfo, range_data_pos, Len(cf.dopinfo) - range_data_pos + 1)
                   End If
                End If
         
                case_fasades.Add cf
                Else
                If regexp_check("ниш", range_data) Or regexp_check("нд", range_data) Then
                    CnameNewFinishPart = "Т"
                Else
                    range_data = regexp_replace("^(ш)", range_data, "")
                    
                    If regexp_check(patNumberWithSlash, range_data) = False And regexp_check(patNumber, range_data) = True Then
                        cf.name = regexp_ReturnSearch(patNumber, range_data)
                        If IsNumeric(cf.name) Then
                            cf.size = CInt(cf.name)
                        Else
                            cf.size = 0
                        End If
                        cf.dopinfo = regexp_replace(patNumber, range_data, "")
                        cf.qty_string = ""
                        cf.qty = 1
                        case_fasades.Add cf
                    End If
'
'                    cf.name = Replace(range_data, " ", "")
'                    cf.size = 0
'                    cf.dopinfo = ""
'                    cf.qty_string = ""
'                    case_fasades.Add cf
                End If
            End If
         
            start_range = cpos + 1
        End If
        cpos = cpos + 1
    Wend
End If
Dim facades_qty As Integer
facades_qty = 0
Dim dopiska As String
dopiska = "("
If case_fasades.Count > 0 Then
    Dim list_pos As Integer
    Dim max As Integer
    Dim min As Integer
    min = 9999
    max = 0
    Dim curfasad As casefasade
    For list_pos = 1 To case_fasades.Count
        If case_fasades(list_pos).size > 0 Then
            If case_fasades(list_pos).size > max Then max = case_fasades(list_pos).size
            If case_fasades(list_pos).size < min Then min = case_fasades(list_pos).size
        End If
        facades_qty = facades_qty + case_fasades(list_pos).qty
        If Len(case_fasades(list_pos).dopinfo) > 0 Then
            If (Mid(dopiska, Len(dopiska), 1) <> ",") And (CStr(Mid(dopiska, Len(dopiska), 1)) <> "(") Then
            dopiska = dopiska & ","
            End If
            dopiska = dopiska & case_fasades(list_pos).dopinfo
        End If
    Next list_pos
    If min = max Then
        cnameNew = cnameNew & "/" & CStr(case_fasades.Count)
    Else
        Dim maxcount As Integer
        maxcount = 0
        Dim mincount As Integer
        mincount = 0
        
        For list_pos = 1 To case_fasades.Count
                If case_fasades(list_pos).size >= max Then maxcount = maxcount + 1
                If case_fasades(list_pos).size < max Then mincount = mincount + 1
        Next list_pos
        
        If maxcount = 1 And mincount = 2 Then
            cnameNew = cnameNew & "/3"
        ElseIf maxcount = 2 And mincount = 1 Then
            cnameNew = cnameNew & "/2-1"
        ElseIf maxcount = 3 And mincount = 1 Then
            cnameNew = cnameNew & "/3-1"
        Else
            cnameNew = cnameNew & "/" & CStr(facades_qty)
        End If
    End If
    
End If
dopiska = dopiska & ")"
If dopiska <> "()" Then
    cnameNew = cnameNew & Mid(dopiska, 2, Len(dopiska) - 2)
    Else
cpos = cpos - 1
End If
cnameNew = cnameNew & CnameNewFinishPart
If (Len(casename) - cpos) > 0 Then
    cnameNew = cnameNew & Mid(casename, cpos + 1, Len(casename) - cpos)
End If
parse_SHL = cnameNew

End Function
Function parse_SHN(ByVal casename_input As String) As String
Dim casename As String

If mRegexp.regexp_check(patGetBase, casename_input) Then
    casename = mRegexp.regexp_replace(patGetBase, casename_input, "")
Else
    casename = casename_input
End If
'casename = case_removeBaze(casename_input)

Dim cnameNew As String
 cnameNew = ""

Dim CnameNewFinishPart As String
CnameNewFinishPart = ""

Dim caseWidth As String
Dim cpos As Integer
cpos = 1

While Not IsNumeric(Mid(casename, cpos, 1))
cnameNew = cnameNew & Mid(casename, cpos, 1)
cpos = cpos + 1

Wend
While IsNumeric(Mid(casename, cpos, 1))
cnameNew = cnameNew & Mid(casename, cpos, 1)
cpos = cpos + 1

Wend

Dim start_range As Integer
Dim end_range As Integer
Dim range_data As String
Dim range_data_pos As Integer


Dim case_fasades As New Collection
Dim cf As casefasade



If Mid(casename, cpos, 1) = "(" Then
    cpos = cpos + 1
    start_range = cpos
    end_range = cpos
    While Mid(casename, cpos, 1) <> ")"
        If Mid(casename, cpos, 1) = "," Or Mid(casename, cpos + 1, 1) = ")" Then
            Set cf = New casefasade
            end_range = cpos
            range_data = Replace(Mid(casename, start_range, end_range - start_range + 1), ",", "")
            
            'разберу фасад
            If InStr(1, Replace(range_data, " ", ""), "/", vbTextCompare) > 0 Then
                cf.name = Mid(range_data, 1, InStr(1, range_data, "/", vbTextCompare) - 1)
                    If IsNumeric(cf.name) Then
                        cf.size = CInt(cf.name)
                    Else
                        cf.size = 0
                    End If
                cf.dopinfo = Mid(range_data, InStr(1, range_data, "/", vbTextCompare) + 1, Len(range_data) - InStr(1, range_data, "/", vbTextCompare))
                range_data_pos = 1
                cf.qty_string = ""
                While IsNumeric(Mid(cf.dopinfo, range_data_pos, 1))
                cf.qty_string = cf.qty_string & Mid(cf.dopinfo, range_data_pos, 1)
                   range_data_pos = range_data_pos + 1
                Wend
                If Len(cf.qty_string) > 0 Then
                cf.qty = CInt(cf.qty_string)
                   If range_data_pos > Len(cf.dopinfo) Then
                      cf.dopinfo = ""
                      Else
                       cf.dopinfo = Mid(cf.dopinfo, range_data_pos, Len(cf.dopinfo) - range_data_pos + 1)
                   End If
                End If
         
                case_fasades.Add cf
                Else
                If regexp_check("ниш", range_data) Or regexp_check("нд", range_data) Then
                    CnameNewFinishPart = "Т"
                Else
                    range_data = regexp_replace("^(ш)", range_data, "")
                    
                    If regexp_check(patNumberWithSlash, range_data) = False And regexp_check(patNumber, range_data) = True Then
                        cf.name = regexp_ReturnSearch(patNumber, range_data)
                        If IsNumeric(cf.name) Then
                            cf.size = CInt(cf.name)
                        Else
                            cf.size = 0
                        End If
                        cf.dopinfo = regexp_replace(patNumber, range_data, "")
                        cf.qty_string = ""
                        cf.qty = 1
                        case_fasades.Add cf
                    End If
'
'                    cf.name = Replace(range_data, " ", "")
'                    cf.size = 0
'                    cf.dopinfo = ""
'                    cf.qty_string = ""
'                    case_fasades.Add cf
                End If
            End If
         
            start_range = cpos + 1
        End If
        cpos = cpos + 1
    Wend
End If
Dim facades_qty As Integer
facades_qty = 0
Dim dopiska As String
dopiska = "("
If case_fasades.Count > 0 Then
    Dim list_pos As Integer
    Dim max As Integer
    Dim min As Integer
    min = 9999
    max = 0
    Dim curfasad As casefasade
    For list_pos = 1 To case_fasades.Count
        If case_fasades(list_pos).size > 0 Then
            If case_fasades(list_pos).size > max Then max = case_fasades(list_pos).size
            If case_fasades(list_pos).size < min Then min = case_fasades(list_pos).size
        End If
        facades_qty = facades_qty + case_fasades(list_pos).qty
        If Len(case_fasades(list_pos).dopinfo) > 0 Then
            If (Mid(dopiska, Len(dopiska), 1) <> ",") And (CStr(Mid(dopiska, Len(dopiska), 1)) <> "(") Then
            dopiska = dopiska & ","
            End If
            dopiska = dopiska & case_fasades(list_pos).dopinfo
        End If
    Next list_pos
    If min = max Then
        cnameNew = cnameNew & "/" & CStr(case_fasades.Count)
    Else
        Dim maxcount As Integer
        maxcount = 0
        Dim mincount As Integer
        mincount = 0
        
        For list_pos = 1 To case_fasades.Count
                If case_fasades(list_pos).size >= max Then maxcount = maxcount + 1
                If case_fasades(list_pos).size < max Then mincount = mincount + 1
        Next list_pos
        
        If maxcount = 1 And mincount = 2 Then
            cnameNew = cnameNew & "/3"
        ElseIf maxcount = 2 And mincount = 1 Then
            cnameNew = cnameNew & "/2-1"
        ElseIf maxcount = 3 And mincount = 1 Then
            cnameNew = cnameNew & "/3-1"
        Else
            cnameNew = cnameNew & "/" & CStr(facades_qty)
        End If
    End If
    
End If
dopiska = dopiska & ")"
If dopiska <> "()" Then
    cnameNew = cnameNew & Mid(dopiska, 2, Len(dopiska) - 2)
End If
cnameNew = cnameNew & CnameNewFinishPart
If (Len(casename) - cpos) > 0 Then
    cnameNew = cnameNew & Mid(casename, cpos + 1, Len(casename) - cpos)
End If
parse_SHN = cnameNew

End Function
Sub parse_RLFull()

'Dim patCheckTbTBV As String
'patCheckTbTBV = "тб[+]\d+тбв"
Dim inputstring As String
Dim storeinputstring As String
Dim patlocalGetOption As String
patlocalGetOption = "\d+(?:[+]\d+)?"

Dim patCheckVP As String
patlocalGetOption = "\d+(?:[+]\d+)?"

 casepropertyCurrent.p_newname = "РЛ "
'On Error GoTo err_parseFasades
Dim max As Integer
Dim min As Integer
Dim cf_count As Integer
Dim cf_curItem_index As Integer

min = 9999
max = 0
'Dim AddT As Boolean
'AddT = False
While CaseElements.Count > 0
CaseElements.Remove (1)
Wend
Dim caseElementItem As New caseElement

casepropertyCurrent.p_newsystem = True

Dim caseFur As caseFurniture

While casefasades.Count > 0
casefasades.Remove (1)
Wend
If Len(casepropertyCurrent.p_fullcn) = 0 Then
    GoTo err_parseFasades
End If
Dim cnameNew As String
cnameNew = ""
Dim cc() As String
Dim c As Variant
Dim cc_element As String

Dim cc_count As Integer
Dim cf As casefasade

Dim maxcount As Integer
maxcount = 0
Dim mincount As Integer
mincount = 0

'стартовые элементы каркаса
Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "цоколь верхний"
caseElementItem.qty = 2
CaseElements.Add caseElementItem

Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "бочок ШЛ"
caseElementItem.qty = 2
CaseElements.Add caseElementItem

Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "крышка ШЛ"
caseElementItem.qty = 1
CaseElements.Add caseElementItem


Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "ДВП"
caseElementItem.qty = 1
CaseElements.Add caseElementItem

inputstring = mRegexp.regexp_ReturnSearch(patCaseFasadesOnlyString, casepropertyCurrent.p_fullcn)
inputstring = Mid(inputstring, 2, Len(inputstring) - 2)
cc = Split(inputstring, ",")
cc_count = 0

For Each c In cc
    cc_element = CStr(c)
    cc_count = cc_count + 1
    Set cf = New casefasade
    cf.init
    If mRegexp.regexp_check(patCaseFasadesIsNisha, cc_element) Then
        cf.isNisha = True
    ElseIf mRegexp.regexp_check(patCaseFasadesIsDver, cc_element) Then
        cf.isDveri = True
    ElseIf mRegexp.regexp_check(patCaseFasadesIsShufl, cc_element) Then
        cf.isShuflyada = True
        If InStr(1, cc_element, "имит") > 0 Then
        cf.dopinfo = "имитация"
        End If
    Else
       cf.isShuflyada = True
    End If
    If mRegexp.regexp_check(patCaseFasadesIsVitr, cc_element) Then
        cf.isVitr = True
    End If
    
    If mRegexp.regexp_check(patCaseFasadesQty, cc_element) Then
        cf.qty = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesQty, cc_element))
    Else
        cf.qty = 1
    End If
    If mRegexp.regexp_check(patCaseFasadesNapravl, cc_element) Then
        cf.isShuflyada = True
       ' cf.napravl = cc_element
        cf.napravl = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravl, cc_element)
        If (cf.napravl <> "") Then
            cf.fCustomerFur = mRegexp.regexp_check(patCaseFasadesNapravlCustomer, cf.napravl)
        End If
        
'        If mRegexp.regexp_check(patlocalGetOption, cf.napravl) Then
'            cf.foption = mRegexp.regexp_ReturnSearch(patNumber, cf.napravl)
'            cf.napravl = mRegexp.regexp_replace(patNumber, cf.napravl, "")
'        End If
        
'        If mRegexp.regexp_check(patCheckTbTBV, cf.napravl) Then
'            cf.napravl = mRegexp.regexp_replace(patNumber, cf.napravl, "")
'        End If
    End If
    If mRegexp.regexp_check(patCaseFasadesWidth, cc_element) Then
        cf.size = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesWidth, cc_element))
'        If cf.size >= 570 Then
'            cf.isShuflyada = False
'            cf.isDveri = True
'        End If
    End If
    
    casefasades.Add cf
Next c

Dim cf_item As casefasade

If casefasades.Count > 0 Then
   
    
    For Each cf_item In casefasades
        If cf_item.isShuflyada Then
            If cf_item.size >= 570 And cf_item.napravl = "" Then
                cf_item.isShuflyada = False
                cf_item.isDveri = True
            End If
        End If
        If cf_item.isDveri Then
            If IsEmpty(casepropertyCurrent.p_DoorCount) Then
                casepropertyCurrent.p_DoorCount = cf_item.qty
                Else
                casepropertyCurrent.p_DoorCount = casepropertyCurrent.p_DoorCount + cf_item.qty
            End If
        End If
        If cf_item.isShuflyada Then
            If cf_item.size > 0 Then
                If cf_item.size > max Then max = cf_item.size
                If cf_item.size < min Then min = cf_item.size
            End If
            casepropertyCurrent.p_ShuflCount = casepropertyCurrent.p_ShuflCount + cf_item.qty
        End If
        If cf_item.isVitr Then
            If IsEmpty(casepropertyCurrent.p_windowcount) Then
                casepropertyCurrent.p_windowcount = cf_item.qty
            ElseIf casepropertyCurrent.p_windowcount = 0 Then
                casepropertyCurrent.p_windowcount = cf_item.qty
            Else
               casepropertyCurrent.p_windowcount = casepropertyCurrent.p_windowcount + cf_item.qty
            End If
        End If
        
        If IsEmpty(casepropertyCurrent.p_FasadesCount) Then
            casepropertyCurrent.p_FasadesCount = cf_item.qty
            Else
            casepropertyCurrent.p_FasadesCount = casepropertyCurrent.p_FasadesCount + cf_item.qty
        End If
       
        
        
    Next cf_item
       ' If NaprMain = "" Then NaprMain = NaprMainDefault
         cf_curItem_index = 0
        For Each cf_item In casefasades
         cf_curItem_index = cf_curItem_index + 1
        If cf_item.isShuflyada Then
          
          
          If cf_item.dopinfo = "имитация" Then
            casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & "имит" & cf_item.qty & cf_item.napravl & ","
            Set caseElementItem = New caseElement
            caseElementItem.init
            caseElementItem.name = "шуфляда имитация"
            caseElementItem.qty = cf_item.qty
            CaseElements.Add caseElementItem
          Else
         '   If cf_item.napravl = "" Then cf_item.napravl = NaprMain
         '   If cf_item.foption = "" Then cf_item.foption = NaprMainLength
          Call parserDop.getDrawerMountItem(main.is18(casepropertyCurrent.p_CaseColor), _
                             cf_item.napravl, _
                             cf_item.fOption, _
                             cf_item.qty, _
                             cf_item.size, _
                             casepropertyCurrent.p_cabWidth, _
                             casepropertyCurrent.p_cabDepth, _
                             cf_item.dopinfo _
                             )
          End If
        End If
        If cf_item.isNisha Then
            casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & cf_item.qty & "ниш" & ","
            If cf_curItem_index < casefasades.Count Then
                Set caseElementItem = New caseElement
                caseElementItem.init
                caseElementItem.name = "полик"
                caseElementItem.qty = 1
                CaseElements.Add caseElementItem
            End If
        End If
        If cf_item.isDveri Then
            casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & cf_item.qty & "дв" & ","
            Set caseElementItem = New caseElement
            caseElementItem.init
            caseElementItem.name = "дверка"
            caseElementItem.qty = cf_item.qty
            CaseElements.Add caseElementItem
        End If
    Next cf_item

        
    
    
End If
If IsEmpty(casepropertyCurrent.p_DoorCount) Then casepropertyCurrent.p_DoorCount = 0
If casepropertyCurrent.p_DoorCount > 0 Then casepropertyCurrent.p_Doormount = "110"

   If Mid(casepropertyCurrent.p_newname, Len(casepropertyCurrent.p_newname), 1) = "," Then
   casepropertyCurrent.p_newname = Mid(casepropertyCurrent.p_newname, 1, Len(casepropertyCurrent.p_newname) - 1)
   End If


Exit Sub
err_parseFasades:
casepropertyCurrent.p_FasadesString = storeinputstring
MsgBox ("Ошибка обработки строки фасадов!!!")
End Sub
Sub parseFasadesSHLSH(inputstring As String)

'Dim patCheckTbTBV As String
'patCheckTbTBV = "тб[+]\d+тбв"

Dim patlocalGetOption As String
patlocalGetOption = "\d+(?:[+]\d+)?"

Dim patCheckVP As String
patlocalGetOption = "\d+(?:[+]\d+)?"

 casepropertyCurrent.p_newname = "ШЛШ "
'On Error GoTo err_parseFasades
Dim storeinputstring As String
storeinputstring = inputstring
Dim max As Integer
Dim min As Integer
Dim cf_count As Integer
Dim cf_curItem_index As Integer

min = 9999
max = 0
'Dim AddT As Boolean
'AddT = False
While CaseElements.Count > 0
CaseElements.Remove (1)
Wend
Dim caseElementItem As New caseElement

casepropertyCurrent.p_newsystem = True

Dim caseFur As caseFurniture

While casefasades.Count > 0
casefasades.Remove (1)
Wend
If Len(inputstring) = 0 Then
    GoTo err_parseFasades
End If
Dim cnameNew As String
cnameNew = ""
Dim NaprMain As String
Dim NaprMainLength As String
Dim cc() As String
Dim c As Variant
Dim cc_element As String

Dim cc_count As Integer
Dim cf As casefasade

Dim maxcount As Integer
maxcount = 0
Dim mincount As Integer
mincount = 0

'стартовые элементы каркаса
Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "ДВП"
caseElementItem.qty = 1
CaseElements.Add caseElementItem

Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "цоколь верхний"
caseElementItem.qty = 2
CaseElements.Add caseElementItem

Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "крышка ШЛ"
caseElementItem.qty = 1
CaseElements.Add caseElementItem

Set caseElementItem = New caseElement
caseElementItem.init
caseElementItem.name = "бочок ШЛ"
caseElementItem.qty = 2
CaseElements.Add caseElementItem

NaprMain = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravlMain, inputstring)
NaprMainLength = mRegexp.regexp_ReturnSearch(patNumber, NaprMain)

If Len(NaprMainLength) > 0 Then NaprMain = Replace(NaprMain, NaprMainLength, "")
inputstring = mRegexp.regexp_ReturnSearch(patCaseFasadesOnlyString, inputstring)
inputstring = Mid(inputstring, 2, Len(inputstring) - 2)
cc = Split(inputstring, ",")
cc_count = 0

For Each c In cc
    cc_element = CStr(c)
    cc_count = cc_count + 1
    Set cf = New casefasade
    cf.init
    If mRegexp.regexp_check(patCaseFasadesIsNisha, cc_element) Then
        cf.isNisha = True
    ElseIf mRegexp.regexp_check(patCaseFasadesIsDver, cc_element) Then
        cf.isDveri = True
    ElseIf mRegexp.regexp_check(patCaseFasadesIsShufl, cc_element) Then
        cf.isShuflyada = True
        If InStr(1, cc_element, "имит") > 0 Then
        cf.dopinfo = "имитация"
        End If
    Else
       cf.isShuflyada = True
    End If
    If mRegexp.regexp_check(patCaseFasadesIsVitr, cc_element) Then
        cf.isVitr = True
    End If
    
    If mRegexp.regexp_check(patCaseFasadesQty, cc_element) Then
        cf.qty = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesQty, cc_element))
    Else
        cf.qty = 1
    End If
    If cf.isNisha = False And mRegexp.regexp_check(patCaseFasadesNapravl, cc_element) Then
        cf.isShuflyada = True
       ' cf.napravl = cc_element
        cf.napravl = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravl, cc_element)
        If (cf.napravl <> "") Then
            cf.fCustomerFur = mRegexp.regexp_check(patCaseFasadesNapravlCustomer, cf.napravl)
        End If
        
'        If mRegexp.regexp_check(patlocalGetOption, cf.napravl) Then
'            cf.foption = mRegexp.regexp_ReturnSearch(patNumber, cf.napravl)
'            cf.napravl = mRegexp.regexp_replace(patNumber, cf.napravl, "")
'        End If
        
'        If mRegexp.regexp_check(patCheckTbTBV, cf.napravl) Then
'            cf.napravl = mRegexp.regexp_replace(patNumber, cf.napravl, "")
'        End If
    End If
    If mRegexp.regexp_check(patCaseFasadesWidth, cc_element) Then
        cf.size = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesWidth, cc_element))
'        If cf.size >= 570 Then
'            cf.isShuflyada = False
'            cf.isDveri = True
'        End If
    End If
    
    casefasades.Add cf
Next c

Dim cf_item As casefasade

If casefasades.Count > 0 Then
   
    
    For Each cf_item In casefasades
        If cf_item.isShuflyada Then
            If cf_item.size >= 570 And cf_item.napravl = "" Then
                cf_item.isShuflyada = False
                cf_item.isDveri = True
            End If
        End If
        If cf_item.isDveri Then
            If IsEmpty(casepropertyCurrent.p_DoorCount) Then
                casepropertyCurrent.p_DoorCount = cf_item.qty
                Else
                casepropertyCurrent.p_DoorCount = casepropertyCurrent.p_DoorCount + cf_item.qty
            End If
        End If
        If cf_item.isShuflyada Then
            If cf_item.size > 0 Then
                If cf_item.size > max Then max = cf_item.size
                If cf_item.size < min Then min = cf_item.size
            End If
            casepropertyCurrent.p_ShuflCount = casepropertyCurrent.p_ShuflCount + cf_item.qty
        End If
        If cf_item.isVitr Then
            If IsEmpty(casepropertyCurrent.p_windowcount) Then
                casepropertyCurrent.p_windowcount = cf_item.qty
            ElseIf casepropertyCurrent.p_windowcount = 0 Then
                casepropertyCurrent.p_windowcount = cf_item.qty
            Else
               casepropertyCurrent.p_windowcount = casepropertyCurrent.p_windowcount + cf_item.qty
            End If
        End If
        
        If IsEmpty(casepropertyCurrent.p_FasadesCount) Then
            casepropertyCurrent.p_FasadesCount = cf_item.qty
            Else
            casepropertyCurrent.p_FasadesCount = casepropertyCurrent.p_FasadesCount + cf_item.qty
        End If
       
        
        
    Next cf_item
        If NaprMain = "" Then NaprMain = NaprMainDefault
         cf_curItem_index = 0
        For Each cf_item In casefasades
         cf_curItem_index = cf_curItem_index + 1
        If cf_item.isShuflyada Then
          
          
          If cf_item.dopinfo = "имитация" Then
            casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & "имит" & cf_item.qty & cf_item.napravl & ","
            Set caseElementItem = New caseElement
            caseElementItem.init
            caseElementItem.name = "шуфляда имитация"
            caseElementItem.qty = cf_item.qty
            CaseElements.Add caseElementItem
          Else
            If cf_item.napravl = "" Then cf_item.napravl = NaprMain
            If cf_item.fOption = "" Then cf_item.fOption = NaprMainLength
          Call parserDop.getDrawerMountItem(main.is18(casepropertyCurrent.p_CaseColor), _
                             cf_item.napravl, _
                             cf_item.fOption, _
                             cf_item.qty, _
                             cf_item.size, _
                             casepropertyCurrent.p_cabWidth, _
                             casepropertyCurrent.p_cabDepth, _
                             cf_item.dopinfo _
                             )
          End If
        End If
        If cf_item.isNisha Then
            casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & cf_item.qty & "ниш" & ","
            If cf_curItem_index < casefasades.Count Then
                Set caseElementItem = New caseElement
                caseElementItem.init
                caseElementItem.name = "полик"
                caseElementItem.qty = 1
                CaseElements.Add caseElementItem
            End If
        End If
        If cf_item.isDveri Then
            casepropertyCurrent.p_newname = casepropertyCurrent.p_newname & cf_item.qty & "дв" & ","
            Set caseElementItem = New caseElement
            caseElementItem.init
            caseElementItem.name = "дверка"
            caseElementItem.qty = cf_item.qty
            CaseElements.Add caseElementItem
        End If
    Next cf_item

        
    
    
End If
If IsEmpty(casepropertyCurrent.p_DoorCount) Then casepropertyCurrent.p_DoorCount = 0
If casepropertyCurrent.p_DoorCount > 0 Then casepropertyCurrent.p_Doormount = "110"

   If Mid(casepropertyCurrent.p_newname, Len(casepropertyCurrent.p_newname), 1) = "," Then
   casepropertyCurrent.p_newname = Mid(casepropertyCurrent.p_newname, 1, Len(casepropertyCurrent.p_newname) - 1)
   End If


Exit Sub
err_parseFasades:
casepropertyCurrent.p_FasadesString = storeinputstring
MsgBox ("Ошибка обработки строки фасадов!!!")
End Sub

Sub parseFasadesSHLGP(inputstring As String)
'On Error GoTo err_parseFasades
Dim storeinputstring As String
storeinputstring = inputstring
Dim max As Integer
Dim min As Integer
min = 9999
max = 0
Dim AddT As Boolean
AddT = False
While casefasades.Count > 0
casefasades.Remove (1)
Wend
If Len(inputstring) = 0 Then
    GoTo err_parseFasades
End If

Dim cnameNew As String
cnameNew = ""
Dim NaprMain As String
Dim cc() As String
Dim c As Variant
Dim cc_element As String

Dim cc_count As Integer
Dim cf As casefasade

Dim maxcount As Integer
        maxcount = 0
        Dim mincount As Integer
        mincount = 0
        
NaprMain = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravlMain, inputstring)
inputstring = mRegexp.regexp_replace(patCaseFasadesNapravlMain, inputstring, "")
inputstring = Mid(inputstring, 2, Len(inputstring) - 2)
cc = Split(inputstring, ",")
cc_count = 0

For Each c In cc
    cc_element = CStr(c)
    Set cf = New casefasade
    cf.init
    If mRegexp.regexp_check(patCaseFasadesIsNisha, cc_element) Then
        cf.isNisha = True
    ElseIf mRegexp.regexp_check(patCaseFasadesIsDver, cc_element) Then
        cf.isDveri = True
    ElseIf mRegexp.regexp_check(patCaseFasadesIsShufl, cc_element) Then
        cf.isShuflyada = True
    Else
        cf.isShuflyada = True
    End If
    cc_count = cc_count + 1
    If mRegexp.regexp_check(patCaseFasadesQty, cc_element) Then
        cf.qty = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesQty, cc_element))
    Else
        cf.qty = 1
    End If
    If mRegexp.regexp_check(patCaseSHLGPFasadesNapravl, cc_element) Then
        cf.napravl = mRegexp.regexp_ReturnSearch(patCaseSHLGPFasadesNapravl, cc_element)
    End If
    
    
    If mRegexp.regexp_check(patCaseFasadesWidth, cc_element) Then
        cf.size = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesWidth, cc_element))
        
    End If
    casefasades.Add cf
Next c
'Set casepropertyCurrent = New caseProperty
Dim cf_item As casefasade

If casefasades.Count > 0 Then
    
    For Each cf_item In casefasades
        
        If cf_item.isShuflyada Then
            If cf_item.size > 56 Then
                If cf_item.size > max Then max = cf_item.size
                If cf_item.size < min Then min = cf_item.size
                If cf.napravl = "" And NaprMain <> "" Then cf.napravl = NaprMain
                If cf_item.napravl <> "" Then casepropertyCurrent.p_Drawermount = cf_item.napravl
                
                If IsEmpty(casepropertyCurrent.p_ShuflCount) Then
                    casepropertyCurrent.p_ShuflCount = cf_item.qty
                Else
                    casepropertyCurrent.p_ShuflCount = casepropertyCurrent.p_ShuflCount + cf_item.qty
            End If

            End If
        End If
        If cf_item.isNisha Then
            AddT = True
            
                casepropertyCurrent.p_haveNisha = True

        Else
            If IsEmpty(casepropertyCurrent.p_FasadesCount) Then
                casepropertyCurrent.p_FasadesCount = cf_item.qty
                Else
                casepropertyCurrent.p_FasadesCount = casepropertyCurrent.p_FasadesCount + cf_item.qty
            End If
        End If
        
        
    Next cf_item
    
    If casepropertyCurrent.p_haveNisha Then
        If casepropertyCurrent.p_Drawermount <> "" Then
            Select Case Mid(casepropertyCurrent.p_Drawermount, 1, 2)
                Case "мб", "тб"
                    If max <= 139 Then
                        cnameNew = "/" & casepropertyCurrent.p_Drawermount
                    ElseIf max > 139 And max <= 213 Then
                        cnameNew = "/" & casepropertyCurrent.p_Drawermount & "м"
                    ElseIf max > 213 And max <= 713 Then
                        cnameNew = "/" & casepropertyCurrent.p_Drawermount & "б"
                    End If
                Case Else
                    cnameNew = "/" & casepropertyCurrent.p_Drawermount
            End Select
         Else
            cnameNew = ""
         End If
    End If
       
    
    
End If
casepropertyCurrent.p_FasadesString = cnameNew
'MsgBox (storeinputstring + vbCrLf + cnameNew + vbCrLf + "дверей: " + CStr(casepropertyCurrent.p_DoorCount))
If IsEmpty(casepropertyCurrent.p_DoorCount) Then casepropertyCurrent.p_DoorCount = 0
Exit Sub
err_parseFasades:
casepropertyCurrent.p_FasadesString = storeinputstring
MsgBox ("Ошибка обработки строки фасадов!!!")
End Sub
Sub parseFasadesMain(inputstring As String, defaultDoorType As Integer)

'defaultDoorType по умолчанию 1 - шуфляда 2 - дверь

'On Error GoTo err_parseFasades
Dim storeinputstring As String
storeinputstring = inputstring
Dim max As Integer
Dim min As Integer
min = 9999
max = 0


If Len(inputstring) = 0 Then
    GoTo err_parseFasadesMain
End If

Dim cnameNew As String
cnameNew = ""
Dim NaprMain As String
Dim cc() As String
Dim ccplus() As String
Dim c As Variant
Dim cp As Variant
Dim cc_element As String
Dim ccplus_element As String

Dim cc_count As Integer
Dim cf As casefasade
Dim cz As caseZone

Dim maxcount As Integer
maxcount = 0
Dim mincount As Integer
mincount = 0
If mRegexp.regexp_check(patCaseFasadesVVERH, inputstring) Then
    casepropertyCurrent.p_haveVVerh = True
    inputstring = mRegexp.regexp_replace(patCaseFasadesVVERH, inputstring, "")
End If

NaprMain = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravlMain, inputstring)
inputstring = mRegexp.regexp_replace(patCaseFasadesNapravlMain, inputstring, "")
inputstring = Mid(inputstring, 2, Len(inputstring) - 2)
cc = Split(inputstring, ",")
cc_count = 0

For Each c In cc
    cc_element = CStr(c)
    ccplus = Split(cc_element, "+")
    Set cz = New caseZone
    cz.init
    cz.z_rawstring = cc_element
    For Each cp In ccplus
        ccplus_element = CStr(cp)
        Set cf = New casefasade
        cf.init
        cf.frawstring = ccplus_element
        If mRegexp.regexp_check(patCaseFasadesIsNisha, ccplus_element) Then
            cf.isNisha = True
        ElseIf mRegexp.regexp_check(patCaseFasadesIsDver, ccplus_element) Then
            cf.isDveri = True
        ElseIf mRegexp.regexp_check(patCaseFasadesIsShufl, ccplus_element) Then
            cf.isShuflyada = True
        Else
            If defaultDoorType = 1 Then cf.isShuflyada = True
            If defaultDoorType = 2 Then cf.isDveri = True
        End If
        
         If mRegexp.regexp_check(patCaseFasadesIsVitr, ccplus_element) Then
            cf.isVitr = True
        End If
    
        If mRegexp.regexp_check(patCaseFasadesQty, ccplus_element) Then
            cf.qty = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesQty, ccplus_element))
        Else
            cf.qty = 1
        End If
        If mRegexp.regexp_check(patCaseFasadesNapravl, cc_element) Then
            cf.napravl = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravl, ccplus_element)
        End If
        If mRegexp.regexp_check(patCaseFasadesWidth, cc_element) Then
            cf.size = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesWidth, ccplus_element))
        End If
        
        cz.casefasades.Add cf
    Next cp
   
    
    
    casezones.Add cz
Next c


Dim cz_item As caseZone
Dim cf_item As casefasade

If casezones.Count > 0 Then
    For Each cz_item In casezones
        For Each cf_item In cz_item.casefasades
            If cf_item.isDveri Then
                cz.z_doorQty = cz.z_doorQty + cf_item.qty
                If IsEmpty(casepropertyCurrent.p_DoorCount) Then
                    casepropertyCurrent.p_DoorCount = cf_item.qty
                    Else
                    casepropertyCurrent.p_DoorCount = casepropertyCurrent.p_DoorCount + cf_item.qty
                End If
            End If
            If cf_item.isShuflyada Then
                cz.z_shuflQty = cz.z_shuflQty + cf_item.qty
                If cf_item.size > 0 Then
                    If cf_item.size > max Then max = cf_item.size
                    If cf_item.size < min Then min = cf_item.size
                End If
                If IsEmpty(casepropertyCurrent.p_ShuflCount) Then
                    casepropertyCurrent.p_ShuflCount = cf_item.qty
                    Else
                    casepropertyCurrent.p_ShuflCount = casepropertyCurrent.p_ShuflCount + cf_item.qty
                End If
            End If
        
            If cf_item.isVitr Then
                cz.z_windowsQty = cz.z_windowsQty + cf_item.qty
                If IsEmpty(casepropertyCurrent.p_windowcount) Then
                    casepropertyCurrent.p_windowcount = cf_item.qty
                ElseIf casepropertyCurrent.p_windowcount = 0 Then
                    casepropertyCurrent.p_windowcount = cf_item.qty
                Else
                   casepropertyCurrent.p_windowcount = casepropertyCurrent.p_windowcount + cf_item.qty
                End If
            End If
        
            If cf_item.isNisha Then
                cz_item.z_isNisha = True
                casepropertyCurrent.p_haveNisha = True
            Else
                casepropertyCurrent.p_FasadesCount = casepropertyCurrent.p_FasadesCount + cf_item.qty
            End If
        Next cf_item
    Next cz_item
End If

Exit Sub
err_parseFasadesMain:
casepropertyCurrent.p_FasadesString = storeinputstring
MsgBox ("Ошибка обработки строки фасадов!!!")
End Sub


Sub parseFasadesSHN(inputstring As String)
Dim nishQty As Integer
nishQty = 0
'On Error GoTo err_parseFasades
Dim storeinputstring As String
storeinputstring = inputstring
Dim max As Integer
Dim min As Integer
min = 9999
max = 0
While CaseElements.Count > 0
CaseElements.Remove (1)
Wend
Dim caseElementItem As New caseElement
While casefasades.Count > 0
casefasades.Remove (1)
Wend
While casezones.Count > 0
casezones.Remove (1)
Wend
If Len(inputstring) = 0 Then
    GoTo err_parseFasades
End If

Dim cnameNew As String
cnameNew = ""
Dim NaprMain As String
Dim cc() As String
Dim ccplus() As String
Dim c As Variant
Dim cp As Variant
Dim cc_element As String
Dim ccplus_element As String

Dim cc_count As Integer
Dim cf As casefasade
Dim cz As caseZone

Dim maxcount As Integer
maxcount = 0
Dim mincount As Integer
mincount = 0
If mRegexp.regexp_check(patCaseFasadesVVERH, inputstring) Then
    casepropertyCurrent.p_haveVVerh = True
    inputstring = mRegexp.regexp_replace(patCaseFasadesVVERH, inputstring, "")
End If
NaprMain = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravlMain, inputstring)
inputstring = mRegexp.regexp_ReturnSearch(patCaseFasadesOnlyString, inputstring)
inputstring = Mid(inputstring, 2, Len(inputstring) - 2)
cc = Split(inputstring, ",")
cc_count = 0

For Each c In cc
    cc_element = CStr(c)
    ccplus = Split(cc_element, "+")
    Set cz = New caseZone
    cz.init
    cz.z_rawstring = cc_element
    For Each cp In ccplus
        ccplus_element = CStr(cp)
        Set cf = New casefasade
        cf.init
        cf.frawstring = ccplus_element
        If mRegexp.regexp_check(patCaseFasadesIsNisha, ccplus_element) Then
            cf.isNisha = True
            nishQty = nishQty + 1
        ElseIf mRegexp.regexp_check(patCaseFasadesIsDver, ccplus_element) Then
            cf.isDveri = True
        ElseIf mRegexp.regexp_check(patCaseFasadesIsShufl, ccplus_element) Then
            cf.isShuflyada = True
        Else
            cf.isDveri = True
        End If
        
         If mRegexp.regexp_check(patCaseFasadesIsVitr, ccplus_element) Then
            cf.isVitr = True
        End If
    
        If mRegexp.regexp_check(patCaseFasadesQty, ccplus_element) Then
            cf.qty = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesQty, ccplus_element))
        Else
            cf.qty = 1
        End If
        If mRegexp.regexp_check(patCaseFasadesNapravl, cc_element) Then
            cf.napravl = mRegexp.regexp_ReturnSearch(patCaseFasadesNapravl, ccplus_element)
        End If
        If mRegexp.regexp_check(patCaseFasadesWidth, cc_element) Then
            cf.size = CInt(mRegexp.regexp_ReturnSearch(patCaseFasadesWidth, ccplus_element))
        End If
        
        cz.casefasades.Add cf
    Next cp
   
    
    
    casezones.Add cz
Next c


Dim cz_item As caseZone
Dim cf_item As casefasade

If casezones.Count > 0 Then
    For Each cz_item In casezones
        For Each cf_item In cz_item.casefasades
            If cf_item.isDveri Then
                cz.z_doorQty = cz.z_doorQty + cf_item.qty
                If IsEmpty(casepropertyCurrent.p_DoorCount) Then
                    casepropertyCurrent.p_DoorCount = cf_item.qty
                    Else
                    casepropertyCurrent.p_DoorCount = casepropertyCurrent.p_DoorCount + cf_item.qty
                End If
            End If
            If cf_item.isShuflyada Then
                cz.z_shuflQty = cz.z_shuflQty + cf_item.qty
                If cf_item.size > 0 Then
                    If cf_item.size > max Then max = cf_item.size
                    If cf_item.size < min Then min = cf_item.size
                End If
                If IsEmpty(casepropertyCurrent.p_ShuflCount) Then
                    casepropertyCurrent.p_ShuflCount = cf_item.qty
                    Else
                    casepropertyCurrent.p_ShuflCount = casepropertyCurrent.p_ShuflCount + cf_item.qty
                End If
            End If
        
            If cf_item.isVitr Then
                cz.z_windowsQty = cz.z_windowsQty + cf_item.qty
                If IsEmpty(casepropertyCurrent.p_windowcount) Then
                    casepropertyCurrent.p_windowcount = cf_item.qty
                ElseIf casepropertyCurrent.p_windowcount = 0 Then
                    casepropertyCurrent.p_windowcount = cf_item.qty
                Else
                   casepropertyCurrent.p_windowcount = casepropertyCurrent.p_windowcount + cf_item.qty
                End If
            End If
        
            If cf_item.isNisha Then
                casepropertyCurrent.p_NishaQty = casepropertyCurrent.p_NishaQty + 1
                casepropertyCurrent.p_haveNisha = True
            Else
                casepropertyCurrent.p_FasadesCount = casepropertyCurrent.p_FasadesCount + cf_item.qty
            End If
    
        Next cf_item
    Next cz_item
    If casepropertyCurrent.p_haveVVerh And casepropertyCurrent.p_NishaQty = 1 Then
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "полик"
    caseElementItem.qty = 1
    CaseElements.Add caseElementItem
    End If
    
    If casepropertyCurrent.p_NishaQty > 1 Then
    Set caseElementItem = New caseElement
    caseElementItem.init
    caseElementItem.name = "полик"
    caseElementItem.qty = casepropertyCurrent.p_NishaQty - 1
    CaseElements.Add caseElementItem
    End If

    If casepropertyCurrent.p_haveNisha And casepropertyCurrent.p_haveVVerh Then
        If casepropertyCurrent.p_DoorCount > 0 Then cnameNew = cnameNew & "/" & CStr(casepropertyCurrent.p_DoorCount)
        cnameNew = cnameNew & "Т"
    ElseIf casepropertyCurrent.p_haveNisha And casepropertyCurrent.p_haveVVerh = False Then
        cnameNew = cnameNew & "Т"
    ElseIf casepropertyCurrent.p_haveNisha = False And casepropertyCurrent.p_haveVVerh = True Then
        If casepropertyCurrent.p_DoorCount > 0 Then cnameNew = cnameNew & "/" & CStr(casepropertyCurrent.p_DoorCount) 'Else cnameNew = "Т" & cnameNew
        cnameNew = cnameNew & "Т" & "вверх"
    Else
        If casepropertyCurrent.p_DoorCount > 0 Then cnameNew = cnameNew & "/" & CStr(casepropertyCurrent.p_DoorCount)
    End If

End If


casepropertyCurrent.p_FasadesString = cnameNew
'MsgBox (storeinputstring + vbCrLf + cnameNew + vbCrLf + "дверей: " + CStr(casepropertyCurrent.p_DoorCount))
If IsEmpty(casepropertyCurrent.p_DoorCount) Then casepropertyCurrent.p_DoorCount = 0
Exit Sub
err_parseFasades:
casepropertyCurrent.p_FasadesString = storeinputstring
MsgBox ("Ошибка обработки строки фасадов!!!")
End Sub

Private Sub pattest()
'Dim res As String
'parseFasades (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ40(140/1,дв570/1лев)мб глуб48 "))
'parseFasades (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ57(ниш,355/1)мб"))
'parseFasades (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ60(355/1,355/1)кв"))
'parseFasades (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ80(140/1,дв570/2)тб"))
'parseFasades (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ80(140/2,дв570/2) глуб57"))
'parseFasades (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ60(140/1,570/2)"))
MsgBox regexp_ReturnSearch(patGetShtQtyfromStringBegin, "!!! Стяжка межсекционная - 5шт.")
parseFasadesSHN (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ40(140/1мб,дв570/1лев)"))
parseFasadesSHN (mRegexp.regexp_ReturnSearch(patCaseFasades, "ШЛШ30(140/1,284/1,284/1)кв ДНО ДСП"))
'MsgBox (TestRegExp(patCaseFasadesIsDver, "140/1"))
'ШЛК80(713/1прав+176/1,176/1,176/1,176/1)        ШЛК80/4
'MsgBox regexp_ReturnSearch("([0-9]+)", "ш328мб")
'MsgBox parseSKOBOCHKI("ШЛК80(713/1прав+176/1,176/1,176/1,176/1)")
'MsgBox parse_SHLSH("ШЛШ60(б.610)(ниша ,ш283мб) глуб35см")
'MsgBox parse_SHLSH("ШЛШ60(915)(193/1,355/1,355/1)тб")
'MsgBox Replace("ШН90(355/1,355/1)вверх", regexp_ReturnSearch(patGetSkobochki, "ШН90(355/1,355/1)вверх"), "/" & parseGetSumFromSKOBOCHKI(regexp_ReturnSearch(patGetSkobochki, "ШН90(355/1,355/1)вверх")))
End Sub



