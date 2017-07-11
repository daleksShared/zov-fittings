Attribute VB_Name = "main"
Option Explicit
Option Compare Text

Public Sub ProcessRetailSheet()
On Error GoTo err_ProcessRetailSheet
'+    Init_rsOrderFittings False
'+    Init_rsCases False
'+    Init_rsOrderCases False
'+    Init_rsOrderElements False

'    Init_rsHandle False
'    Init_rsLeg False
    
    
    Set FormFitting = New AddFitting
    Set FormElement = New AddElement

    Dim ShipID As Long
'        ShipID = 1


    'Dim TasksForm As MainForm
    Dim TaskID As Long
    'Set TasksForm = New MainForm
    'TasksForm.Show
    'ShipID = TasksForm.ShipID
    MainForm.Show
    ShipID = MainForm.ShipID
    
    'Set TasksForm = Nothing
    If ShipID = 0 Then Exit Sub
    On Error GoTo 0
      
    Dim SubOrder As Boolean, Customer As String, NewCust As String
    Dim L As Long
    Dim EmptyLines As Long
    EmptyLines = 0
       
       
    'Application.ScreenUpdating = False
    
    
    For L = 1 To 10000

        'If EmptyLines > 100 Then Exit Sub
        
        If Rows(L).Hidden = False Then
            If Not (Trim(Cells(L, 1)) = "" And Trim(Cells(L, 2)) = "" And Trim(Cells(L, 3)) = "" And Trim(Cells(L, 4)) = "" _
                     And Trim(Cells(L, 5)) = "" And Trim(Cells(L, 6)) = "" And Trim(Cells(L, 7)) = "" And Trim(Cells(L, 8)) = "" And Trim(Cells(L, 9)) = "") Then
                EmptyLines = 0
                
                'ищем клиента
                If Trim(Cells(L, 2)) <> "" Then
                    If Cells(L, 2).Borders.LineStyle < 0 Then NewCust = Cells(L, 2): Cells(L, 2).Activate
                ElseIf Trim(Cells(L, 3)) <> "" Then
                    If Cells(L, 3).Borders.LineStyle < 0 Then NewCust = Cells(L, 3): Cells(L, 3).Activate
                End If
                If NewCust <> Customer Then
                    If InStr(1, NewCust, "дозак", vbTextCompare) > 0 Or InStr(1, NewCust, "додел", vbTextCompare) > 0 Then
                        SubOrder = True
                        ActiveCell.Interior.ColorIndex = 44
                    Else
                        Dim tCust As String
                        
                        
                        tCust = Trim(InputBox("Введите имя клиента (предыдущий - """ & Customer & """)" & _
                            "(ОТМЕНА - продолжить выборку для клиента " & Customer & ")", "Клиент", Trim(Replace(NewCust, ".", ""))))
                            
                        If tCust <> "" Then
                            Customer = tCust
                            SubOrder = False
                            ActiveCell.Interior.ColorIndex = 39
                        End If
                    End If
                    NewCust = Customer
                End If
                
                If Customer <> "" Then
                
                    'ищем кухню
                    If Trim(Cells(L, 1)) <> "" Then
                        If Left(Trim(Cells(L, 1)), 1) = "№" Then
                            Dim FirstOrderRow As Long, LastOrderRow As Long
                            Dim CasesPreambleRow As Long, FCol As Long
                            
                            
                            EmptyLines = 0
                            FirstOrderRow = L
                            LastOrderRow = L
                            CasesPreambleRow = 0
                            FCol = 0
                            
    'If Cells(Row + 1, 1).Borders(xlEdgeTop).LineStyle > 0 And Cells(Row + 1, 1).Borders(xlEdgeBottom).LineStyle > 0 Then GoTo skipSHPK
    
                            While L <= 10000 And Left(Trim(Cells(L + 1, 1)), 1) <> "№"
                                 If (EmptyLines = 0 And Not Trim(Cells(L, 1)) = "") Or _
                                 Not (Trim(Cells(L, 1)) = "" And Trim(Cells(L, 4)) = "" And _
                                        Trim(Cells(L, 2)) = "" And Trim(Cells(L, 3)) = "" And _
                                        Trim(Cells(L, 5)) = "" And Trim(Cells(L, 6)) = "" And _
                                        Trim(Cells(L, 7)) = "" And Trim(Cells(L, 8)) = "" Or EmptyLines = 1) Then
                                        
                                    LastOrderRow = L
                                Else
                                    If CasesPreambleRow > 0 Then
                                        GoTo ExitWhile
                                    Else
                                        EmptyLines = EmptyLines + 1
                                        If EmptyLines > 1 Then L = L - 1: GoTo ExitWhile
                                    End If
                                End If
                                'ElseIf StrComp(Trim(Cells(l, 2)), "бочки", vbTextCompare) = 0 Or StrComp(Trim(Cells(l, 3)), "бочки", vbTextCompare) = 0 Then
                                If CasesPreambleRow = 0 Then
                                    If Cells(L, 2).Borders(xlEdgeTop).LineStyle > 0 _
                                                Or Cells(L, 2).Borders(xlEdgeLeft).LineStyle > 0 _
                                                Or Cells(L, 2).Borders(xlEdgeRight).LineStyle > 0 _
                                                Or Cells(L, 2).Borders(xlEdgeTop).LineStyle > 0 _
                                                Or Cells(L, 3).Borders(xlEdgeTop).LineStyle > 0 _
                                                Or Cells(L, 3).Borders(xlEdgeLeft).LineStyle > 0 _
                                                Or Cells(L, 3).Borders(xlEdgeRight).LineStyle > 0 _
                                                Or Cells(L, 3).Borders(xlEdgeTop).LineStyle > 0 Then
                                    
                                        Dim cell As Range
                                        For Each cell In Range(Cells(L, 2), Cells(L, 8))
                                            If InStr(1, cell.Value, "ф-ра", vbTextCompare) > 0 Or _
                                                InStr(1, cell.Value, "фур", vbTextCompare) > 0 Then
                                                
                                                FCol = cell.Column
                                                cell.Interior.ColorIndex = 3
                                                Exit For
                                            ElseIf InStr(1, cell.Value, "ф", vbTextCompare) = 1 And Right(CStr(cell.Value), 1) = "а" Then
                                                If MsgBox("Есть фурнитура каркасов?", vbCritical + vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                                                    FCol = cell.Column
                                                    cell.Interior.ColorIndex = 3
                                                    Exit For
                                                End If
                                            End If
                                             
                                             
                                        Next cell
                                        
                                        If CasesPreambleRow = 0 Then CasesPreambleRow = L
                                      
                                        
                                   'Else
                                        'GoTo ExitWhile
                                    End If
                                End If
                                L = L + 1
                            Wend
ExitWhile:
'                            If Not (CasesPreambleRow = 0 And LastOrderRow > FirstOrderRow) Then
'                            Else
'                                l = l - 1
'                                'ActiveSheet.Range(Cells(FirstOrderRow, 1), Cells(LastOrderRow, 1)).Select
'                                'Selection.Interior.Color = RGB(150, 150, 50)
'                            End If
                            
                            
                            
                            If Not AddOrderToShip(ShipID, _
                                                    Customer, _
                                                    SubOrder, _
                                                    FirstOrderRow, _
                                                    CasesPreambleRow, _
                                                    FCol, _
                                                    LastOrderRow) Then Exit Sub
    
                        End If
                    End If
                End If 'if customer<> ?
            Else
                EmptyLines = EmptyLines + 1
                If EmptyLines > 140 Then Exit For
            End If
        End If
    Next L
   
    'Application.ScreenUpdating = True
    
    Exit Sub
err_ProcessRetailSheet:
    'Application.ScreenUpdating = True
    MsgBox Error, vbCritical
    Application.Cursor = xlDefault
End Sub


Private Function AddOrderToShip(ByVal ShipID As Long, _
                                ByVal Customer As String, _
                                ByVal SubOrder As Boolean, _
                                FirstOrderRow As Long, _
                                CasesPreampleRow As Long, _
                                FCol As Long, _
                                LastOrderRow As Long) As Boolean
    Set kitchenPropertyCurrent = New kitchenProperty
    Set casepropertyCurrent = New caseProperty
   ' On Error GoTo err_AddOrderToShip
   Dim paramIterator As Integer
Dim tempString As String
   Dim ExcelCaseName As String
    Dim changeCaseZaves As Integer
    
    changeCaseZaves = 0

    AddOrderToShip = False
    If ShipID = 0 Then Exit Function
    
    If CasesPreampleRow > 0 Then
        ActiveSheet.Range(Cells(FirstOrderRow, 1), Cells(CasesPreampleRow, 1)).Select
        Selection.Interior.Color = RGB(173, 255, 47)
    Else
        ActiveSheet.Range(Cells(FirstOrderRow, 1), Cells(LastOrderRow, 1)).Select
        Selection.Interior.Color = RGB(173, 255, 47)
    End If
                                  
                                  
    Dim OrderN As String
   
   If InStr(1, Cells(FirstOrderRow, 1), ".") > 0 Then
        OrderN = Trim(Mid(Cells(FirstOrderRow, 1), 2, InStr(1, Cells(FirstOrderRow, 1), ".") - 1 - 1))
   Else
        OrderN = Trim(Cells(FirstOrderRow, 1))
   End If
    If InStr(2, OrderN, " ", vbTextCompare) > 0 Then
        OrderN = InputBox("Исправьте номер заказа (возможно нет точки)", "Номер заказа", OrderN)
    End If
    kitchenPropertyCurrent.OrderN = OrderN
'    While (Len(OrderN) - InStr(1, OrderN, "-", vbTextCompare)) > 15 Or _
'            Len(OrderN) = 0 Or _
'                Len(OrderN) < 3
    While Len(OrderN) > 32
        'OrderN = InputBox("Исправьте номер заказа (слишком длинное/недопустимое значение)", "Номер заказа", OrderN)
        OrderN = Left(OrderN, 32)
    Wend
       
'    If SubOrder Then OrderN = "Доз" & OrderN
       
    Dim sKitchen As String
    
    ' ВОЗЬМЕМ ЦВЕТ БОЧКОВ
    Dim CaseColor
    CaseColor = Null
    sKitchen = UCase("Бочки ")
    Dim iTmpKitch As Integer, pBorder As Boolean
    iTmpKitch = InStr(1, UCase(Cells(FirstOrderRow, 1)), sKitchen, vbTextCompare)
    
    If iTmpKitch = 0 Then pBorder = False Else pBorder = True
    If pBorder Then
        Cells(FirstOrderRow, 1) = Replace(Cells(FirstOrderRow, 1) & ".", "..", ".")
        
        If InStr(iTmpKitch, Cells(FirstOrderRow, 1), ".") = 0 Then
            pBorder = False
            
            
            'CaseColor = InputBox("Введите цвет бочков", "Цвет бочков", Trim(UCase(Mid(Cells(FirstOrderRow, 1), iTmpKitch + Len(sKitchen), Len(Cells(FirstOrderRow, 1)) - (iTmpKitch + Len(sKitchen))))))
            'CaseColor = Left(CaseColor, 20)
        End If
    End If
    
    If (pBorder) Then
        
        CaseColor = Trim(UCase(Mid(Cells(FirstOrderRow, 1), iTmpKitch + Len(sKitchen), _
                                   InStr(iTmpKitch, Cells(FirstOrderRow, 1), ".") - (iTmpKitch + Len(sKitchen)))))
        CaseColor = Replace(CaseColor, "мм", "")
        CaseColor = Replace(CaseColor, "  ", " ")
        CaseColor = Left(CaseColor, 20)
   
    End If
    If IsNull(CaseColor) = False Then kitchenPropertyCurrent.dspColor = CaseColor
    Dim ColorId As Integer

'    If FormColor Is Nothing Then Set FormColor = New ColorForm
'    colorid = GetColorID(CaseColor)
'    If colorid = 0 Then
'        FormColor.Show
'        'colorid = FormColor.colorid
'        CaseColor = Left(FormColor.ColorName, 20)
'    End If
    
    
    
    ' ВОЗЬМЕМ ЦВЕТ И ТИП ФАСАДОВ
    Dim face
    face = Null
    sKitchen = UCase("Фасад ")
    iTmpKitch = InStr(1, UCase(Cells(FirstOrderRow, 1)), sKitchen, vbTextCompare)
    
    If iTmpKitch = 0 Then pBorder = False Else pBorder = True
    If pBorder Then If InStr(iTmpKitch, Cells(FirstOrderRow, 1), ".") = 0 Then pBorder = False
    
    If (pBorder) Then
        face = Trim(Mid(Cells(FirstOrderRow, 1), iTmpKitch + Len(sKitchen), _
                                   InStr(iTmpKitch, Cells(FirstOrderRow, 1), ".") - (iTmpKitch + Len(sKitchen))))
        face = Left(face, 50)
    End If
    If IsNull(face) = False Then kitchenPropertyCurrent.fasadColor = face
    ' ВОЗЬМЕМ ЦВЕТ И ТОЛЩИНУ СТОЛЕШНИЦЫ
    Dim TableTopColor, PlankSize, PlankColor, OtbColor
    TableTopColor = Null
    
    sKitchen = UCase("Столеш")
    iTmpKitch = InStr(1, UCase(Cells(FirstOrderRow, 1)), sKitchen, vbTextCompare)
    
    If iTmpKitch = 0 Then ' может популярная описка?
        sKitchen = UCase("Стоелш")
        iTmpKitch = InStr(1, UCase(Cells(FirstOrderRow, 1)), sKitchen, vbTextCompare)
    End If
    
    If iTmpKitch > 0 Then iTmpKitch = InStr(iTmpKitch, Cells(FirstOrderRow, 1), " ")
    
    If iTmpKitch = 0 Then pBorder = False Else pBorder = True
    
    Dim en As Integer
    If pBorder Then en = InStr(iTmpKitch, Cells(FirstOrderRow, 1), ".")
    If en = 0 Then en = Len(Cells(FirstOrderRow, 1))
    
    If pBorder Then If en = 0 Then pBorder = False
    
    If (pBorder) Then
        
        TableTopColor = Trim(Mid(Cells(FirstOrderRow, 1), iTmpKitch + 1, _
                                   en - iTmpKitch))
                                   
        If Right(TableTopColor, 1) = "." Then _
            TableTopColor = Left(TableTopColor, Len(TableTopColor) - 1)
            
        TableTopColor = Trim(TableTopColor)
                                   
        If InStr(1, TableTopColor, "28") Then
            PlankSize = 28
            TableTopColor = Trim(Left(TableTopColor, InStr(1, TableTopColor, "28") - 1))
        ElseIf InStr(1, TableTopColor, "38") Then
            PlankSize = 38
            TableTopColor = Trim(Left(TableTopColor, InStr(1, TableTopColor, "38") - 1))
        End If
        
        PlankColor = GetPlankColor(TableTopColor)
        OtbColor = GetOtbColor(TableTopColor)
        
        TableTopColor = Left(TableTopColor, 25)
    End If

    
    ' ПОЛУЧИМ ИДЕНТИФИКАТОР ЗАКАЗА
    Dim OrderId As Long
    OrderId = AddOrder(ShipID, FirstOrderRow, Customer, OrderN)
    kitchenPropertyCurrent.OrderId = OrderId
    If kitchenPropertyCurrent.dspColor <> "" Then UpdateOrder OrderId, , , , , , kitchenPropertyCurrent.dspColor
    
    If Not IsNull(face) And Not IsEmpty(face) Then
        UpdateOrder OrderId, , , , , face
        Cells(FirstOrderRow, 13).Value = face
        
'        If InStr(1, face, "акрил", vbTextCompare) > 0 Then
'            FormFitting.AddFittingToOrder OrderID, "полироль", Empty, , , , , CasesPreampleRow
'        End If
    End If
   
    Dim row As Long
    
'================================================
'====== РАЗБОР ШАПКИ =======НАЧАЛО===============
'================================================

    Dim comm As ADODB.Command
    Set comm = New ADODB.Command
    comm.ActiveConnection = GetConnection
    comm.CommandType = adCmdStoredProc
    comm.CommandText = "AddPattern"

    Dim parPatt As ADODB.Parameter
    Set parPatt = New ADODB.Parameter
    parPatt.name = "@Pattern"
    parPatt.Direction = adParamInput
    parPatt.Type = adVarChar
    parPatt.size = 150

    Dim parOrderID As ADODB.Parameter
    Set parOrderID = New ADODB.Parameter
    parOrderID.name = "@OrderID"
    parOrderID.Direction = adParamInput
    parOrderID.Type = adInteger
    parOrderID.size = 4
    
    Dim parRow As ADODB.Parameter
    Set parRow = New ADODB.Parameter
    parRow.name = "@Row"
    parRow.Direction = adParamInput
    parRow.Type = adInteger
    parRow.size = 4
    
    comm.Parameters.Append parPatt
    comm.Parameters.Append parOrderID
    comm.Parameters.Append parRow
    
    Dim qty As Integer, Opt
    Dim Handle, Leg
    Dim HandleScrew, HangColor, CaseHang
    
    Dim bPackShelvingWithFittingsKit
    bPackShelvingWithFittingsKit = Null
    
    
    Dim EndRow As Integer
    If CasesPreampleRow > 0 Then EndRow = CasesPreampleRow Else EndRow = LastOrderRow
    
    For row = FirstOrderRow To EndRow 'CasesPreampleRow - 1
        Dim k As Integer, t As Integer, p As Integer, st As Integer
        Dim tstr As String
        
        If Not Rows(row).Hidden Then
            OrderCaseID = 0
            Cells(row, 1).Select
            Cells(row, 1).Activate
            'ActiveCell = Replace(ActiveCell, "  ", " ")
            tstr = ActiveCell
            
            'tstr = Replace(tstr, "  ", " ")
                       
            parPatt.Value = Left(tstr, 150)
            parOrderID.Value = OrderId
            parRow.Value = row
            comm.Execute
            
                        
            
            If Not IsNull(Handle) Then
                Dim Handle_
                Handle_ = FindFittings(OrderId, row, cHandle, tstr, , , , face, HandleScrew)
                'If IsEmpty(handle_) Then handle_ = FindFittings(OrderID, "Р.", tstr)
                If Not IsEmpty(Handle_) Then
                    If IsNull(Handle_) Then
                        Handle = Null
                    ElseIf IsEmpty(Handle) Then
                        If InStr(1, Handle_, "гориз") = 0 And InStr(1, Handle_, "вертик") = 0 Then Handle = Handle_
                    ElseIf InStr(1, ActiveCell.Value, "горизонтально") = 0 And InStr(1, ActiveCell.Value, "вертикально") = 0 Then
                        MsgBox "Внимание! В заказе дважды указывается тип ручек!"
                        If MsgBox("Принять новый?", vbCritical + vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            Handle = Handle_
                        End If
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                End If
            End If
            
            Dim Leg_
            Leg_ = FindFittings(OrderId, row, cLeg, tstr)
            If Not IsEmpty(Leg_) Then
                If IsNull(Leg_) Then
                    Leg = Null
                ElseIf IsEmpty(Leg) Then
                    Leg = Leg_
                    Leg = Replace(Leg, "-", "")
                    Leg = Replace(Leg, "№", "")
                Else
                    MsgBox "Внимание! В заказе дважды указывается тип ножек!"
                    If MsgBox("Принять новый?", vbCritical + vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        Leg = Leg_
                    End If
                    ActiveCell.Interior.ColorIndex = 3
                End If
            End If
            If Not kitchenPropertyCurrent.PProfil Then
                If (InStr(1, tstr, "п-про") > 0 And InStr(1, tstr, "витр") > 0) Then
                    kitchenPropertyCurrent.PProfil = True
                ElseIf InStr(1, tstr, "проф") > InStr(1, tstr, "п") Then
                    kitchenPropertyCurrent.PProfil = True
                ElseIf InStr(1, tstr, "Р-проф") > 0 Then
                    kitchenPropertyCurrent.PProfil = True
                End If
            End If
            FindFittings OrderId, row, "Конф", tstr
            FindFittings OrderId, row, "шкант", tstr
            FindFittings OrderId, row, " VS ", tstr
            FindFittings OrderId, row, " VS-", tstr
            FindFittings OrderId, row, "Экспоз", tstr
            FindFittings OrderId, row, "штанга", tstr
            FindFittings OrderId, row, "Дугов", tstr
            FindFittings OrderId, row, "стенд", tstr
             FindFittings OrderId, row, "серви", tstr
            
            FindFittings OrderId, row, "подпят", tstr
            FindFittings OrderId, row, "потпят", tstr
            FindFittings OrderId, row, "клипсы к цок", tstr
            FindFittings OrderId, row, "клипса к цок", tstr
            FindFittings OrderId, row, "клипсы цок", tstr
            FindFittings OrderId, row, "клипса цок", tstr
            FindFittings OrderId, row, "Клипсы на цок", tstr
            FindFittings OrderId, row, "Клипса на цок", tstr
            FindFittings OrderId, row, "Клипсы для цок", tstr
            FindFittings OrderId, row, "Клипса для цок", tstr
            
            FindFittings OrderId, row, "дюбел", tstr
            FindFittings OrderId, row, "эксц", tstr
               FindFittings OrderId, row, "опора нижн", tstr, , , , , , changeCaseZaves
               FindFittings OrderId, row, "ПМ-3", tstr
             FindFittings OrderId, row, "Уголки метал", tstr
            FindFittings OrderId, row, "Угол метал", tstr
            FindFittings OrderId, row, "Уголок метал", tstr
            FindFittings OrderId, row, "Волпато", tstr
            FindFittings OrderId, row, "Валпато", tstr
            FindFittings OrderId, row, "Вольпато", tstr
            FindFittings OrderId, row, "Вальпато", tstr
            FindFittings OrderId, row, "Volpato", tstr
            FindFittings OrderId, row, "мойка", tstr
            FindFittings OrderId, row, cPlank, tstr, , PlankSize, PlankColor     ' планки
            FindFittings OrderId, row, cGalog, tstr ' галогенки
            FindFittings OrderId, row, "полир", tstr
            FindFittings OrderId, row, "палир", tstr
            
            FindFittings OrderId, row, "VS -", tstr ' карго
            FindFittings OrderId, row, "карг", tstr ' карго
            FindFittings OrderId, row, "карусель", tstr ' карго
            FindFittings OrderId, row, "сушк", tstr ' сушки
            FindFittings OrderId, row, "лоток", tstr ' лотки
            FindFittings OrderId, row, "вклад", tstr ' лотки
            FindFittings OrderId, row, "лифт", tstr ' лифт
            FindFittings OrderId, row, "полка", tstr ' полка оборотная
            FindFittings OrderId, row, "завеш", tstr
            FindFittings OrderId, row, "площадк", tstr
            FindFittings OrderId, row, "завес Sens", tstr ' завес
            FindFittings OrderId, row, "петли Sens", tstr ' завес
            FindFittings OrderId, row, "завес Сенс", tstr ' завес
            FindFittings OrderId, row, "завес Сэнс", tstr ' завес
            FindFittings OrderId, row, "петли Сэнс", tstr ' завес
            FindFittings OrderId, row, "петли Сенс", tstr ' завес
            FindFittings OrderId, row, "завесы Sens", tstr ' завес
            FindFittings OrderId, row, "петля Sens", tstr ' завес
            FindFittings OrderId, row, "завесы Сенс", tstr ' завес
            FindFittings OrderId, row, "завесы Сэнс", tstr ' завес
            FindFittings OrderId, row, "петля Сэнс", tstr ' завес
            FindFittings OrderId, row, "петля Сенс", tstr ' завес
            FindFittings OrderId, row, "Завес CLIP top", tstr ' завес
            FindFittings OrderId, row, "завес", tstr ' завес
            FindFittings OrderId, row, "петл", tstr ' завес
            FindFittings OrderId, row, "ухо", tstr ' петля
            FindFittings OrderId, row, "уши", tstr ' петля

            'FindFittings OrderID, Row, "зав", tstr ' завешки
            FindFittings OrderId, row, "труба", tstr '
            'FindFittings OrderID, Row, cSink, tstr ' Мойки
            'FindFittings OrderID, Row, cStol, tstr ' Столы
            'FindFittings OrderID, Row, cStul, tstr ' Стулья
            'FindFittings OrderID, Row, cStool, tstr ' Стулья
            FindFittings OrderId, row, cNogi, tstr ' ноги
            FindFittings OrderId, row, "нога", tstr ' ноги
            FindFittings OrderId, row, "стул", tstr ' стулья Ибица/Феликс Женева Юлия
            
            'FindFittings OrderID, Row, "цоколь пласт", tstr '030910
            FindFittings OrderId, row, "ведро", tstr ' ведро мусорное
            FindFittings OrderId, row, "поддон", tstr ' поддон алюминиевый
            FindFittings OrderId, row, "диод", tstr
            'FindFittings OrderID, Row, "аморт", tstr
            FindFittings OrderId, row, "отбойник", tstr
            FindFittings OrderId, row, "монт", tstr
'            FindFittings OrderID, Row, "монтажн", tstr
            
            FindFittings OrderId, row, "push to open", tstr
            FindFittings OrderId, row, "push-to-open", tstr
            FindFittings OrderId, row, "пуш то опен", tstr
            FindFittings OrderId, row, "пуш ту опен", tstr
            
            FindFittings OrderId, row, "надставка", tstr
            FindFittings OrderId, row, "удлинитель", tstr
            FindFittings OrderId, row, "система перегородок", tstr
            
            FindFittings OrderId, row, "стеклодерж", tstr
            FindFittings OrderId, row, "полкодерж", tstr
            FindFittings OrderId, row, "пеликан", tstr
            FindFittings OrderId, row, "направл", tstr
'
            FindFittings OrderId, row, "тандембокс", tstr
            
'            FindFittings OrderID, Row, "доводчик к м", tstr
'            FindFittings OrderID, Row, "доводчик на м", tstr
'            FindFittings OrderID, Row, "доводчик для м", tstr
'            FindFittings OrderID, Row, "доводчики к м", tstr
'            FindFittings OrderID, Row, "доводчики на м", tstr
'            FindFittings OrderID, Row, "доводчики для м", tstr
            
             FindFittings OrderId, row, "Решетка вентиляционная", tstr
             FindFittings OrderId, row, "вент", tstr
            FindFittings OrderId, row, "банк", tstr
            FindFittings OrderId, row, "меджик лайт", tstr
           FindFittings OrderId, row, "мейджик лайт", tstr
           FindFittings OrderId, row, "мЭйджик лайт", tstr
           FindFittings OrderId, row, "блок ", tstr
            FindFittings OrderId, row, "magic light", tstr
            FindFittings OrderId, row, "электроблок", tstr
            FindFittings OrderId, row, "трансформатор", tstr
            FindFittings OrderId, row, "коврик", tstr
            FindFittings OrderId, row, "волшеб", tstr
            FindFittings OrderId, row, "Вытяжка", tstr
            FindFittings OrderId, row, "подъемник", tstr
            FindFittings OrderId, row, "подъёмник", tstr
            FindFittings OrderId, row, "дэмпфер", tstr
            FindFittings OrderId, row, "демпфер", tstr
            FindFittings OrderId, row, "магнит", tstr
            
'            If InStr(1, tstr, "На заказ завес") > 0 And (InStr(1, tstr, "Сенсис") > 0 Or InStr(1, tstr, "Sens") > 0) Then
'               changeCaseZaves = 1
'               ActiveCell.Interior.Color = vbRed
'            'Else
'            '   changeCaseZaves = -1
'            End If

            Dim newFittingQty


            newFittingQty = Empty
            
             If (InStr(1, tstr, "довод") > 1) And (InStr(1, tstr, "мб") > 1 Or InStr(1, tstr, "мет") > 1) Then
                If InStr(InStr(1, tstr, "довод"), tstr, "м") > InStr(1, tstr, "довод") Then
                    FindFittings OrderId, row, Mid(tstr, InStr(1, tstr, "довод"), InStr(InStr(1, tstr, "довод"), tstr, "м") - InStr(1, tstr, "довод") + 1), tstr
                End If
            ElseIf InStr(1, tstr, "метаб") > 1 Then
                FindFittings OrderId, row, "метабокс", tstr
            End If


            If (InStr(1, tstr, "сортир") > 1) Or (InStr(1, tstr, "организ") > 1) Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "сортир")))
                    
                FormFitting.AddFittingToOrder OrderId, "???", newFittingQty, , , , , row
            End If
             If (InStr(1, tstr, "пол") > 1) And (InStr(1, tstr, "подсв") > 1) Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "пол")))
                FormFitting.AddFittingToOrder OrderId, "пол с подсветкой", newFittingQty, , , , , row
            End If

             If (InStr(1, tstr, "СИСО") > 0) Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "СИСО")))
                FormFitting.AddFittingToOrder OrderId, "втулка SISO П+М", newFittingQty, , , , , row
            ElseIf (InStr(1, tstr, "SISO") > 0) Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "SISO")))
                FormFitting.AddFittingToOrder OrderId, "втулка SISO П+М", newFittingQty, , , , , row
            ElseIf (InStr(1, tstr, "втулка") > 0) Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "втулка")))
                FormFitting.AddFittingToOrder OrderId, "втулка SISO П+М", newFittingQty, , , , , row
            End If
            
             If (InStr(1, tstr, "мкость") > 1) Or (InStr(1, tstr, "корзин") > 1) Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "мкость")))
                FormFitting.AddFittingToOrder OrderId, "емкость в тб под мойку", newFittingQty, , , , , row
            End If
                        
            If (InStr(1, tstr, "аморт") > 0 And InStr(1, tstr, "врезн") > 0) Or (InStr(1, tstr, "амморт") > 0 And InStr(1, tstr, "врезн") > 0) Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "врезн")))
                FormFitting.AddFittingToOrder OrderId, "амортизатор врезной", newFittingQty, , , , , row
            End If
            If InStr(1, tstr, "механ") > 0 Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "механ")))
                FormFitting.AddFittingToOrder OrderId, "механизм", newFittingQty, , , , , row
            End If
             If InStr(1, tstr, "шуруп") > 0 Then
                FormFitting.AddFittingToOrder OrderId, "добавь шуруп", Empty, , , , , row
            End If
            If InStr(1, tstr, "комплект") > 0 Then
                FormFitting.AddFittingToOrder OrderId, "", Empty, , , , , row
            End If
            
            If InStr(1, tstr, "стяж") > 0 And InStr(1, tstr, "межсек") > 0 Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "стяж")))
                FormFitting.AddFittingToOrder OrderId, "стяжка колпачковая", newFittingQty, "средняя", , , , row
            ElseIf InStr(1, tstr, "стяж") > 0 And InStr(1, tstr, "пуст") > 0 Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "пуст")))
                FormFitting.AddFittingToOrder OrderId, "стяжка для пустотки", newFittingQty, , , , , row
            ElseIf InStr(1, tstr, "стяж") > 0 Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "стяж")))
                FormFitting.AddFittingToOrder OrderId, "стяжка для столешницы", newFittingQty, , , , , row
            End If
             
            If InStr(1, tstr, "средство") > 0 Then
                newFittingQty = parseShtQtyfromString(Mid(tstr, InStr(tstr, "средство")))
                FormFitting.AddFittingToOrder OrderId, "полироль", newFittingQty, "50мл", , , , CasesPreampleRow
            End If
            
            
            If InStr(1, tstr, "портал") > 0 Then
                'FormFitting.AddFittingToOrder OrderID, "дюбель DU325 Rapid S", Empty, , , , , Row
               
                    If FormElement.AddElementToOrder(OrderId, "", "") Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
  
            End If
            
            
            If InStr(1, tstr, "полос") > 0 And (InStr(1, tstr, "цокол") > 0 Or InStr(1, tstr, "прозр") > 0) Then
               FindFittings OrderId, row, "полос", tstr ' полоса прозрачная к цоколю
           
            End If
            
'            If InStr(1, tstr, "аморт") > 0 Then
'                If InStr(1, tstr, "врезн") > 0 Then
'
'                    FormFitting.AddFittingToOrder OrderID, "врезной амортизатор", Empty, , , caseID, , Row
'
'                ElseIf InStr(1, tstr, "blum") > 0 Or InStr(1, tstr, "блюм") > 0 Then
'                    Cells(Row, FCol).Interior.Color = vbYellow
'                    FormFitting.AddFittingToOrder OrderID, "амортизатор BLUM", Empty, , , caseID, , Row
'
'                ElseIf InStr(1, tstr, "врезн") > 0 Then
'                    Cells(Row, FCol).Interior.Color = vbYellow
'                    FormFitting.AddFittingToOrder OrderID, "амортизатор FGV", Empty, , , caseID, , Row
'                Else
'                    Cells(Row, FCol).Interior.Color = vbYellow
'                    FormFitting.AddFittingToOrder OrderID, "амортизатор", Empty, , , caseID, , Row
'
'                End If
'            End If
            
            
            Select Case True
                Case InStr(1, tstr, "труб", vbTextCompare)
                    ActiveCell.Interior.Color = vbRed
                Case InStr(1, tstr, "метаб", vbTextCompare)
                    ActiveCell.Interior.Color = vbRed
                Case InStr(1, tstr, "тандем", vbTextCompare)
                    ActiveCell.Interior.Color = vbRed
                Case InStr(1, tstr, "цоколь пвх", vbTextCompare)
                    ActiveCell.Interior.Color = vbRed
            End Select
    
    
    'ищем цоколь пласт*************************
            If mRegexp.regexp_check("^(.*цок.+пласт.*)$", tstr) Then
                tstr = ""
                mRegexp.regexp_ReturnSearchCollection patSplitPattern, ActiveCell.Value
                parseTsokol OrderId, row
            End If
            
            
    'ищем отбортовку 2.0
'            If mRegexp.regexp_check(".*?([^\.\!]*тборт.*)", tstr) Then
'                Cells(row, 9).Value = getColorForArrayFromString(tstr, OtbGorbColors)
'                'mRegexp.regexp_ReturnSearchCollection patSplitPattern, mRegexp.regexp_ReturnSearch(".*?([^\.\!]*тборт.*)", tstr)
'                'parseOtbort orderid, row
'            End If






'              Opt = Null
'            Qty = 0
'            k = InStr(1, tstr, "цоколь пласт", vbTextCompare)
'            If k = 0 Then k = InStr(1, tstr, "цоколя пласт", vbTextCompare)
'            If k = 0 Then k = InStr(1, tstr, "цоколи пласт", vbTextCompare)
'            If k Then
'
'
'                tstr = ActiveCell.Value
'                'k = InStr(1, tstr, "отб", vbTextCompare)
'                ActiveCell.Characters(k, 12).Font.Color = vbBlue
'                If InStr(1, tstr, "h=15 ", vbTextCompare) > 0 Then
'                    tstr = Replace(tstr, "h=15", "150", 1, 1, vbTextCompare)
'
'                End If
'
'                t = InStr(1, tstr, ".")
'                p = InStr(1, tstr, ",")
'                If t - k < 7 Then t = InStr(t + 1, tstr, ".")
'                If (t < p And p <> 0 And t <> 0) Or (t <> 0 And p = 0) Then
'                    en = t
'                ElseIf p <> 0 Then
'                        en = p
'                    Else
'                        en = Len(tstr)
'                End If
'
'
'                'разбор планок по кускам в см через +
'                Dim part As String, thing As String
'                part = ""
'                tstr = Mid(tstr, k + 12) '!!!
'                p = k '!!!
'                While tstr <> ""
'                    thing = ""
'
'
'                    While Not IsNumeric(Left(tstr, 1)) And Left(tstr, 1) <> "+" And _
'                            Left(tstr, 1) <> " " And Left(tstr, 1) <> "(" And tstr <> ""
'
'                        tstr = Mid(tstr, 2)
'                        p = p + 1
'                    Wend
'
'                    While Left(tstr, 1) = " "
'                        tstr = Mid(tstr, 2)
'                        p = p + 1
'                    Wend
'
'                    If IsNull(Opt) Then
'                        While Not IsNumeric(Left(tstr, 1)) And Left(tstr, 1) <> "+" And tstr <> ""
'                            Opt = Opt & Left(tstr, 1)
'                            tstr = Mid(tstr, 2)
'                            p = p + 1
'                        Wend
'                        Opt = Trim(Opt)
'                        If (InStr(Opt, "чёр") > 0 Or InStr(Opt, "чeр") > 0) And InStr(Opt, "гл") > 3 Then
'                        Opt = "ЧЁРНЫЙГЛ"
'                        End If
'                        If InStr(Opt, "бел") > 0 And InStr(Opt, "гл") > 3 Then
'                        Opt = "БЕЛЫЙГЛ"
'                        End If
'
'                        Dim CokolInd As Integer
'                        Dim checkCokolInd As Integer
'                        checkCokolInd = -1
'                        For CokolInd = 0 To UBound(Цоколь) - 1
'                            If Цоколь(CokolInd) = Opt Then
'                                checkCokolInd = CokolInd
'                                Exit For
'                            End If
'                        Next CokolInd
'
'                        If checkCokolInd = -1 Then
'                            For CokolInd = 0 To UBound(Цоколь) - 1
'                                If InStr(1, Цоколь(CokolInd), Opt) = 1 Then
'                                    checkCokolInd = CokolInd
'                                    Exit For
'                                End If
'                            Next CokolInd
'
'
'                            If checkCokolInd > -1 Then
'
'                                If InStr(1, Left(LTrim(tstr), 4), "150", vbTextCompare) > 0 Or InStr(1, Left(tstr, 5), "15см", vbTextCompare) > 0 Or InStr(1, Left(tstr, 5), "15см", vbTextCompare) > 0 Then
'                                    Opt = Opt & "150"
'                                    tstr = Replace(tstr, "150", "", 1, 1, vbTextCompare)
'                                    tstr = Replace(tstr, "15см", "", 1, 1, vbTextCompare)
'                                    p = p + 3
'                                ElseIf InStr(1, Left(LTrim(tstr), 4), "100", vbTextCompare) > 0 Or InStr(1, Left(tstr, 5), "10см", vbTextCompare) > 0 Then
'                                    Opt = Opt & "100"
'                                    tstr = Replace(tstr, "100", "", 1, 1, vbTextCompare)
'                                    tstr = Replace(tstr, "10см", "", 1, 1, vbTextCompare)
'                                    p = p + 3
'                                End If
'                            End If
'
'
'
''                            If Opt = "ХРОМ" Then
''                                If InStr(1, Left(tstr, 4), "150", vbTextCompare) > 0 Or InStr(1, Left(tstr, 5), "15см", vbTextCompare) > 0 Then
''                                    Opt = "ХРОМ150"
''                                    tstr = Replace(tstr, "150", "", 1, 1, vbTextCompare)
''                                    tstr = Replace(tstr, "15см", "", 1, 1, vbTextCompare)
''                                    p = p + 3
''                                End If
''                            End If
'                        End If
'                    Else
'                        While Not IsNumeric(Left(tstr, 1)) And Left(tstr, 1) <> "+" And tstr <> ""
'                            tstr = Mid(tstr, 2)
'                            p = p + 1
'                        Wend
'                    End If
'
'
'                    While Left(tstr, 1) = "+" Or Left(tstr, 1) = " "
'                        tstr = Mid(tstr, 2)
'                        p = p + 1
'                    Wend
'
'                    While IsNumeric(Left(tstr, 1)) Or Left(tstr, 1) = " "
'                        thing = thing & Left(tstr, 1)
'                        DelSymbol tstr, 1
'                        ActiveCell.Characters(p + 12, 1).Font.Color = vbRed
'                        p = p + 1
'                    Wend
'
'                    If thing = "" Then thing = "1"
'
'                    Dim isTsokol As Boolean
'                    If InStr(1, Left(tstr, 2), "м", vbTextCompare) Then isTsokol = True Else isTsokol = False
'
'                    If isTsokol Then
'                        part = thing
'                        If CInt(part) Mod 3 = 0 Then
'                            Qty = CInt(part) \ 3  'трехметровая отбортовка
'                            part = "3м"
'                            If IsNull(Opt) Then Opt = OtbColor
'
''                        ElseIf CInt(part) Mod 4 = 0 Then 'четырехметровая отбортовка
''                            Qty = CInt(part) \ 4
''                            If InStr(1, Opt, "горб", vbTextCompare) > 0 Then
''                                part = "горб-4"
''                                'If IsNull(Opt) Then Opt = TableTopColor
''                                'Opt = TableTopColor
''                            Else
''                                part = "4м"
''                                If IsNull(Opt) Then Opt = TableTopColor
''                            End If
''
''                        ElseIf CInt(part) Mod 5 = 0 Then 'пятиметровая горбатая отбортовка
''                            Qty = CInt(part) \ 5
''                            If InStr(1, Opt, "горб", vbTextCompare) > 0 Then
''                                part = "горб-5"
''                                If IsNull(Opt) Then Opt = TableTopColor
''                                Opt = TableTopColor
'''                            Else
'''                                part = "5м"
'''                                If IsNull(Opt) Then Opt = TableTopColor
''
''                            ElseIf InStr(1, Opt, "top", vbTextCompare) > 0 Or InStr(1, Opt, "LINE", vbTextCompare) > 0 Then
''                                part = "TOP-Line"
''                                'Opt = TOPLine
''                            End If
'
'                        End If
'                    Else
'                        Qty = CInt(thing)
'                    End If
'
'                    If tstr <> "" Then
'                        If isTsokol Then
'
'                            ActiveCell.Characters(p + 12, 1).Font.Bold = True
'                            If Not FormFitting.AddFittingToOrder(orderid, "цоколь пластик", Qty, Opt, , , , row) Then Exit Function
'                        ElseIf InStr(1, Left(tstr, 2), "з", vbTextCompare) Then
'
'                                'If IsNull(Opt) Then Opt = OtbColor
'
'                                ActiveCell.Characters(p, 1).Font.Bold = True
''                                If InStr(1, Left(tstr, 10), "пр", vbTextCompare) Then
''                                    If Not FormFitting.AddFittingToOrder(OrderID, "Загл. прав к отб. " & part, Qty, Opt, , , , Row) Then Exit Function
''                                ElseIf InStr(1, Left(tstr, 10), "ле", vbTextCompare) Then
''                                    If Not FormFitting.AddFittingToOrder(OrderID, "Загл. лев к отб. " & part, Qty, Opt, , , , Row) Then Exit Function
''                                Else
'                                 If Not FormFitting.AddFittingToOrder(orderid, "заглушка к цоколю", Qty, Opt, , , , row) Then Exit Function
''                                End If
'                        ElseIf InStr(1, Left(tstr, 2), "у", vbTextCompare) Then
'
'                            ActiveCell.Characters(p + 12, 1).Font.Bold = True
'
'                            'If IsNull(Opt) Then Opt = OtbColor
'
'                            If InStr(1, Left(tstr, 8), "90", vbTextCompare) Then
'                                If Not FormFitting.AddFittingToOrder(orderid, "угол90* к цоколю", Qty, Opt, , , , row) Then Exit Function
'                            ElseIf InStr(1, Left(tstr, 9), "135", vbTextCompare) Then
'                                If Not FormFitting.AddFittingToOrder(orderid, "угол135* к цоколю", Qty, Opt, , , , row) Then Exit Function
'                            Else
'                                If Not FormFitting.AddFittingToOrder(orderid, "угол90* к цоколю", Qty, Opt, , , , row) Then Exit Function
'                            End If
'
''                        ElseIf InStr(1, Left(tstr, 2), "к", vbTextCompare) Then
''
''                            ActiveCell.Characters(p + 12, 1).Font.Bold = True
''
''                            'If IsNull(Opt) Then Opt = OtbColor
''                            '20130117 крепление к цоколю не выдаю
''                            If Not FormFitting.AddFittingToOrder(OrderID, "крепление к цоколю", Null, Opt, , , , Row) Then Exit Function
'                        ElseIf InStr(1, Left(tstr, 4), "соед", vbTextCompare) Then
'
'                            ActiveCell.Characters(p + 4, 1).Font.Bold = True
'
'                            'If IsNull(Opt) Then Opt = OtbColor
'
'                            If Not FormFitting.AddFittingToOrder(orderid, "соединитель цоколя", Qty, Opt, , , , row) Then Exit Function
'                        ElseIf InStr(1, Left(tstr, 4), "стык", vbTextCompare) Then
'
'                            ActiveCell.Characters(p + 4, 1).Font.Bold = True
'
'                            'If IsNull(Opt) Then Opt = OtbColor
'
'                            If Not FormFitting.AddFittingToOrder(orderid, "соединитель цоколя", Qty, Opt, , , , row) Then Exit Function
'
'                        End If
'
'                    End If
'                Wend
'            End If
'
    
    
    'ищем отбортовку*************************
            ' разбор планок по кускам в см через +
            Dim part As String, thing As String
            Opt = Null
            qty = 0
            k = InStr(1, tstr, "отб", vbTextCompare)
            If k Then
                tstr = ActiveCell.Value
                k = InStr(1, tstr, "отб", vbTextCompare)
                ActiveCell.Characters(k, 7).Font.Color = vbBlue


                t = InStr(1, tstr, ".")
                p = InStr(1, tstr, ",")
                If t - k < 7 Then t = InStr(t + 1, tstr, ".")
                If (t < p And p <> 0 And t <> 0) Or (t <> 0 And p = 0) Then
                    en = t
                ElseIf p <> 0 Then
                        en = p
                    Else
                        en = Len(tstr)
                End If


                'разбор планок по кускам в см через +
                'Dim part As String, thing As String
                part = ""
                tstr = Mid(tstr, k) '!!!
                p = k '!!!
                While tstr <> ""
                    thing = ""

                    While Not IsNumeric(Left(tstr, 1)) And Left(tstr, 1) <> "+" And _
                            Left(tstr, 1) <> " " And Left(tstr, 1) <> "(" And tstr <> ""

                        tstr = Mid(tstr, 2)
                        p = p + 1
                    Wend

                    While Left(tstr, 1) = " "
                        tstr = Mid(tstr, 2)
                        p = p + 1
                    Wend

                    If IsNull(Opt) Then
                        While Not IsNumeric(Left(tstr, 1)) And Left(tstr, 1) <> "+" And tstr <> ""
                            Opt = Opt & Left(tstr, 1)
                            tstr = Mid(tstr, 2)
                            p = p + 1
                        Wend

                    Else
                        While Not IsNumeric(Left(tstr, 1)) And Left(tstr, 1) <> "+" And tstr <> ""
                            tstr = Mid(tstr, 2)
                            p = p + 1
                        Wend
                    End If


                    While Left(tstr, 1) = "+" Or Left(tstr, 1) = " "
                        tstr = Mid(tstr, 2)
                        p = p + 1
                    Wend

                    While IsNumeric(Left(tstr, 1)) Or Left(tstr, 1) = " " Or Left(tstr, 1) = ","
                        thing = thing & Left(tstr, 1)
                        DelSymbol tstr, 1
                        ActiveCell.Characters(p, 1).Font.Color = vbRed
                        p = p + 1
                    Wend

                    If thing = "" Then thing = "1"

                    Dim IsOtbortovka As Boolean
                    If InStr(1, Left(tstr, 2), "м", vbTextCompare) Then IsOtbortovka = True Else IsOtbortovka = False

                    If IsOtbortovka Then
                        part = thing
                        
                        If CInt(part) Mod 4 = 0 Then 'четырехметровая отбортовка
                            qty = CInt(part) \ 4
                            If InStr(1, Opt, "горб", vbTextCompare) > 0 Then
                                part = "горб-4"
                                'If IsNull(Opt) Then Opt = TableTopColor
                                'Opt = TableTopColor
                            Else
                                part = "4м"
                                If IsNull(Opt) Then Opt = TableTopColor
                            End If
                            
                        ElseIf CInt(part) Mod 3 = 0 Then
                            qty = CInt(part) \ 3  'трехметровая отбортовка
                            part = "3м"
                            If IsNull(Opt) Then Opt = OtbColor

                        

                        ElseIf CInt(part) Mod 5 = 0 Then 'пятиметровая горбатая отбортовка
                            qty = CInt(part) \ 5
                            If InStr(1, Opt, "горб", vbTextCompare) > 0 Then
                                part = "горб-5"
                                If IsNull(Opt) Then Opt = TableTopColor
                                Opt = TableTopColor
'                            Else
'                                part = "5м"
'                                If IsNull(Opt) Then Opt = TableTopColor

                            ElseIf InStr(1, Opt, "top", vbTextCompare) > 0 Or InStr(1, Opt, "LINE", vbTextCompare) > 0 Then
                                part = "TOP-Line"
                                'Opt = TOPLine
                            End If

                        End If
                    Else
                        qty = CInt(thing)
                    End If

                    If tstr <> "" Then
                        If IsOtbortovka Then

                            ActiveCell.Characters(p, 1).Font.Bold = True
                            If Not FormFitting.AddFittingToOrder(OrderId, "Отбортовка " & part, qty, Opt, , , , row) Then Exit Function
                        ElseIf InStr(1, Left(tstr, 2), "з", vbTextCompare) Then

                                If IsNull(Opt) Then Opt = OtbColor

                                ActiveCell.Characters(p, 1).Font.Bold = True
                                If InStr(1, Left(tstr, 12), "пр", vbTextCompare) > 0 And InStr(1, Left(tstr, 12), "лев", vbTextCompare) = 0 Then
                                    If Not FormFitting.AddFittingToOrder(OrderId, "Загл. прав к отб. " & part, qty, Opt, , , , row) Then Exit Function
                                ElseIf InStr(1, Left(tstr, 12), "ле", vbTextCompare) > 0 And InStr(1, Left(tstr, 12), "пр", vbTextCompare) = 0 Then
                                    If Not FormFitting.AddFittingToOrder(OrderId, "Загл. лев к отб. " & part, qty, Opt, , , , row) Then Exit Function
                                Else
                                    If Not FormFitting.AddFittingToOrder(OrderId, "Загл. к отб. " & part, qty, Opt, , , , row) Then Exit Function
                                End If
                            ElseIf InStr(1, Left(tstr, 2), "уг", vbTextCompare) Or _
                                    InStr(1, Left(tstr, 8), "вн", vbTextCompare) Or _
                                    InStr(1, Left(tstr, 10), "нар", vbTextCompare) Then

                                ActiveCell.Characters(p, 1).Font.Bold = True

                                If IsNull(Opt) Then Opt = OtbColor

                                If InStr(1, Left(tstr, 8), "внеш", vbTextCompare) Or _
                                    InStr(1, Left(tstr, 10), "нар", vbTextCompare) Then

                                    If part = "Top-Line" Then
                                        If Not FormFitting.AddFittingToOrder(OrderId, "Угол внеш. к отб " & part, qty, Opt, , , , row) Then Exit Function
                                    Else
                                        If Not FormFitting.AddFittingToOrder(OrderId, "Угол внешн. к отб. " & part, qty, Opt, , , , row) Then Exit Function
                                    End If

                                Else
                                    If Not FormFitting.AddFittingToOrder(OrderId, "Угол к отб. " & part, qty, Opt, , , , row) Then Exit Function
                                End If
                        End If
                    End If
                Wend
            End If
  
        
        
    'разбираемся с реллингом*********
    
    
            Opt = Null
            qty = 0
            Dim ug As Integer, ts As String
            k = InStr(1, tstr, "рел", vbTextCompare)
            
            If k Then If MsgBox("Реллинг?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Фурнитура") = vbNo Then k = 0
            
            If k Then
                tstr = ActiveCell.Value
                k = InStr(1, tstr, "рел", vbTextCompare)
                ActiveCell.Characters(k, 8).Font.Color = vbBlue
                
                'пытаемся выделить слово после "рел*"
                t = InStr(k, tstr, ".")
                p = InStr(k, tstr, " ")
                If (t < p And p <> 0 And t <> 0) Or (t <> 0 And p = 0) Then
                    st = t
                ElseIf p <> 0 Then
                        st = p
                    Else
                        st = k
                End If
                If t - k < 9 Then t = InStr(t + 1, tstr, ".")
                If p <> 0 And p - k < 9 Then p = InStr(p + 1, tstr, " ")
                If (t < p And p <> 0 And t <> 0) Or (t <> 0 And p = 0) Then
                    en = t - 1
                ElseIf p <> 0 Then
                        en = p - 1
                    Else
                        en = Len(tstr)
                End If
                                            
                Opt = Mid(tstr, st + 1, en - st) 'выделили
                Cells(row, 1).Characters(k, st - k).Font.Color = vbBlue
                Cells(row, 1).Characters(st + 1, en - st).Font.Color = vbGreen  'выделяем обработанную часть
                
                
                'разбор по кускам в см и элементам через +
                Dim tn As Integer, i As Integer
                p = en + 1
                tstr = Mid(tstr, p)
                While tstr <> ""
                    t = InStr(1, tstr, "+")
                    If t = 0 Then t = Len(tstr)
                    ts = Mid(tstr, 1, t)
                    tstr = Mid(tstr, t + 1)
                    p = p + t
                    
                    Dim trn As String
                    trn = ""
                    part = ""
                    ug = 0
                    tn = 0
                    For i = 1 To Len(ts)
                        If IsNumeric(Mid(ts, i, 1)) Then
                            trn = trn & Mid(ts, i, 1)
                            Cells(row, 1).Characters(p - t + i - 1, 1).Font.Color = vbRed
                        Else
                            If trn <> "" Then Exit For
                        End If
                    Next i
                    If trn = "" Then
                        tn = 1
                    Else
                        If InStr(1, Mid(ts, i, 3), "см", vbTextCompare) Then
                            part = trn 'кусок
                            Cells(row, 1).Characters(p - t + i + InStr(1, Mid(ts, i, 3), "см", vbTextCompare) - 2, 2).Font.Color = vbBlue
                        ElseIf InStr(1, Mid(ts, i, 3), "гр", vbTextCompare) Or CInt(trn) >= 90 Then
                               ug = CInt(trn) 'угол в градусах
                               Cells(row, 1).Characters(p - t + i + InStr(1, Mid(ts, i, 3), "гр", vbTextCompare) - 2, 2).Font.Color = vbBlue
                            Else
                                tn = CInt(trn) ' штуки
                                qty = tn
                        End If
                    End If
                    
                    
                    trn = ""
                    For i = i To Len(ts)
                        If IsNumeric(Mid(ts, i, 1)) Then
                            trn = trn & Mid(ts, i, 1)
                            Cells(row, 1).Characters(p - t + i - 1, 1).Font.Color = vbRed
                        End If
                    Next i
                    
                    If trn <> "" Then
                        If ug = 0 Then
                            ug = CInt(trn)
                        End If
                        If Not (trn = "90" Or trn = "120") Then
                            qty = CInt(trn)
                        Else
                            qty = tn
                        End If
                    End If
                    

                    
                    If InStr(1, ts, "см", vbTextCompare) Then
                        If qty = 0 Then qty = 1
                        'FAdd 21, Relling, part & "см"
                        If part = 60 Then
                            If Not FormFitting.AddFittingToOrder(OrderId, "Реллинг 60", qty, Opt, , , , row) Then Exit Function
                        ElseIf part = 100 Then
                            If Not FormFitting.AddFittingToOrder(OrderId, "Реллинг 100", qty, Opt, , , , row) Then Exit Function
                        Else
                            MsgBox "Ошибка"
                        End If
                    ElseIf InStr(1, ts, "з", vbTextCompare) Then
                            
                            Cells(row, 1).Characters(p - t + InStr(1, ts, "з", vbTextCompare) - 1, 1).Font.Color = vbBlue
                            'FAdd 22, Relling, tn
                            If Not FormFitting.AddFittingToOrder(OrderId, "Заглушка к реллингу", qty, Opt, , , , row) Then Exit Function
                            ElseIf InStr(1, ts, "уг", vbTextCompare) Then
                            
                                If ug = 90 Then
                                    Cells(row, 1).Characters(p - t + InStr(1, ts, "уг", vbTextCompare) - 1, 2).Font.Color = vbBlue
                                    'FAdd 24, Relling, tn
                                If Not FormFitting.AddFittingToOrder(OrderId, "Угол-90 к реллингу", qty, Opt, , , , row) Then Exit Function
                                ElseIf ug = 120 Then
                                
                                    Cells(row, 1).Characters(p - t + InStr(1, ts, "уг", vbTextCompare) - 1, 2).Font.Color = vbBlue
                                    'FAdd 25, Relling, tn
                                If Not FormFitting.AddFittingToOrder(OrderId, "Угол-120 к реллингу", qty, Opt, , , , row) Then Exit Function
                                End If
                            ElseIf InStr(1, ts, "крю", vbTextCompare) Then
                            
                                Cells(row, 1).Characters(p - t + InStr(1, ts, "крю", vbTextCompare) - 1, 3).Font.Color = vbBlue
                                'FAdd 26, Relling, tn
                                If Not FormFitting.AddFittingToOrder(OrderId, "Крючок к реллингу", qty, Opt, , , , row) Then Exit Function
                                ElseIf InStr(1, ts, "д", vbTextCompare) Then
                                    
                                    Cells(row, 1).Characters(p - t + InStr(1, ts, "д", vbTextCompare) - 1, 1).Font.Color = vbBlue
                                    'FAdd 23, Relling, tn
                                    If Not FormFitting.AddFittingToOrder(OrderId, "Держатель к реллингу", qty, Opt, , , , row) Then Exit Function
                    End If
                Wend
            End If
            'If InStr(1, tstr, "конф", vbTextCompare) > 0 Then
            FindFittings OrderId, row, "заглуш", tstr 'заглушки для конфирмантов
        End If
       
 
        
    Next row
    
'================================================
'====== РАЗБОР ШАПКИ ======КОНЕЦ=================
'================================================


    
    
    '****************************
    '*** РАЗБЕРЕМ ШКАФЫ ЗАКАЗА***
    '****************************
    
    
    If FCol > 0 Then ' если есть столбец "ф-ра"
        Dim check_deleteZaveshki As Boolean
        
        
        Dim BibbColor
        BibbColor = Empty

        Dim CamBibbColor
        CamBibbColor = Empty



        ' проверим тип ножек, если указан
        If Not IsEmpty(Leg) And Not IsNull(Leg) Then
            CheckLeg Leg
        End If

        'ActiveSheet.Range(Cells(CasesPreampleRow, 1), Cells(LastOrderRow, 1)).Select
        'Selection.Interior.Color = RGB(173, 255, 47)
        Dim SetQty
        Dim bBreakOrder As Boolean
        bBreakOrder = False

        For row = CasesPreampleRow + 1 To LastOrderRow

           
            Dim caseType As Integer
            
            ' 0 - зов
            ' 1 - zov-modul
            caseType = 0
          
            
            Dim caseID As Integer, DoorCount, windowcount, Drawermount, Doormount, NoFace As Boolean, HandleExtra, ShelfQty
            Dim localCaseHang
            localCaseHang = CaseHang '!!
            Dim caseglub As Integer
            Dim casename As String, CaseQty As Integer

            Dim FCell As String, bHandleCheck


            FCell = Trim(Cells(row, FCol))



            Cells(row, 1).Activate
            Cells(row, 1).Select
            casename = Trim(ActiveCell.Value)
            
            
            
' casename_old = ""
'            casename_old = Trim(Cells(Row, 15))
'
'            If IsEmpty(casename_old) = False Then
'                If Len(casename_old) > 3 Then
'                    If Mid(casename, 1, 2) = Mid(casename_old, 1, 2) Then
'                    casename = casename_old
'                    Cells(Row, 15).Interior.ColorIndex = 4
'                    Cells(Row, 1).Interior.ColorIndex = 3
'
'                    End If
'                End If
'            End If

            caseID = 0
            OrderCaseID = 0
            DoorCount = Empty
            windowcount = Empty
            Drawermount = Empty
            Doormount = Empty
            NoFace = Empty
            HandleExtra = Empty
            ShelfQty = Empty
            Set CaseElementsCollection = New Collection
            Set CaseFittingsCollection = New Collection
            Set params = New Collection
            OrderCaseID = 0
            bHandleCheck = False
            check_deleteZaveshki = False
            CaseQty = 1
            If InStr(1, FCell, "нет", vbTextCompare) > 0 Then
                CaseQty = InputBox("Укажите кол-во шкафов", "Количество шкафов", 0)
            ElseIf Not IsEmpty(ActiveCell.Offset(, 1)) Then
                If IsNumeric(ActiveCell.Offset(, 1)) Then
                    CaseQty = ActiveCell.Offset(, 1)
                ElseIf InStr(1, ActiveCell.Offset(, 1), "нет", vbTextCompare) > 0 Then
                    CaseQty = InputBox("Укажите кол-во шкафов", "Количество шкафов", ActiveCell.Offset(, 1))
                End If
            End If
            
              '***************************************
                            If FormColor Is Nothing Then Set FormColor = New ColorForm
                        
                           
                                ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                                If ColorId = 0 Then
                                    FormColor.Show
                                    'colorid = FormColor.colorid
                                    CaseColor = Left(FormColor.ColorName, 20)
                                    ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                                    kitchenPropertyCurrent.dspColor = CaseColor
                                    kitchenPropertyCurrent.dspColorId = ColorId
                                    kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                                End If
    
                                If Not IsNull(CaseColor) Then UpdateOrder OrderId, , , , , , CaseColor
                                If ColorId > 0 Then UpdateOrder OrderId, , , , , , , ColorId
                                
                                '***** заглушки ************************
                                If IsEmpty(BibbColor) Then
                                    BibbColor = GetBibbColor(CaseColor)
                                End If
                                If Not IsNull(BibbColor) Then UpdateOrder OrderId, , , BibbColor
    
                                If IsEmpty(CamBibbColor) Then
                                    CamBibbColor = GetCamBibbColor(CaseColor)
                                    kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                                End If
                                If Not IsNull(CamBibbColor) Then
                                    UpdateOrder OrderId, , , , , , , , CamBibbColor
                                End If
            
            If bBreakOrder Then
'                If CaseQty = 1 Then
'                    OrderID = AddOrder(ShipID, Customer, OrderN)
'                Else
                    OrderId = AddOrder(ShipID, FirstOrderRow, Customer, Left(OrderN & "/" & casename, 32), CaseQty)
                    UpdateOrder OrderId, HandleScrew, HangColor, BibbColor, , face, , , CamBibbColor
'                End If
            ElseIf CaseQty > 1 And IsEmpty(SetQty) Then
                If MsgBox("Обработать заказ как оптовый 'по шкафам'?", vbQuestion + vbDefaultButton3 + vbYesNo, "Тип заказа") = vbYes Then
                    OrderId = AddOrder(ShipID, FirstOrderRow, Customer, Left(OrderN & "/" & casename, 32), CaseQty)
                    UpdateOrder OrderId, HandleScrew, HangColor, BibbColor, , face, , , CamBibbColor
                    bBreakOrder = True
                    bPackShelvingWithFittingsKit = False
                Else
                    If MsgBox("Установить кол-во комплектов заказа равным " & CaseQty & "?", vbQuestion + vbDefaultButton3 + vbYesNo, "Количество комплектов") = vbYes Then
                        SetQty = CaseQty
                        UpdateOrder OrderId, , , , SetQty
                    End If
                End If
            End If
            
            If bBreakOrder Then
              '***************************************
                            If FormColor Is Nothing Then Set FormColor = New ColorForm
                        
                           
                                ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                                If ColorId = 0 Then
                                    FormColor.Show
                                    'colorid = FormColor.colorid
                                    CaseColor = Left(FormColor.ColorName, 20)
                                    ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                                    kitchenPropertyCurrent.dspColor = CaseColor
                                    kitchenPropertyCurrent.dspColorId = ColorId
                                    kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                                End If
    
                                If Not IsNull(CaseColor) Then UpdateOrder OrderId, , , , , , CaseColor
                                If ColorId > 0 Then UpdateOrder OrderId, , , , , , , ColorId
                                
                                '***** заглушки ************************
                                If IsEmpty(BibbColor) Then
                                    BibbColor = GetBibbColor(CaseColor)
                                End If
                                If Not IsNull(BibbColor) Then UpdateOrder OrderId, , , BibbColor
    
                                If IsEmpty(CamBibbColor) Then
                                    CamBibbColor = GetCamBibbColor(CaseColor)
                                    kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                                End If
                                If Not IsNull(CamBibbColor) Then
                                    UpdateOrder OrderId, , , , , , , , CamBibbColor
                                End If
            
            End If
                          

            If CaseQty > 0 Then

                ' проверим пилястру

                 If InStr(1, casename, "пилястра 7,5", vbTextCompare) > 0 Or InStr(1, casename, "пилястра 7.5", vbTextCompare) > 0 Or InStr(1, casename, "пилястра7,5", vbTextCompare) > 0 Or InStr(1, casename, "пилястра7.5", vbTextCompare) > 0 Then
                    If MsgBox("Пилястра 16 - ДА, или 18-НЕТ", vbYesNo, "Тип пилястры?") = vbYes Then

                    If FormElement.AddElementToOrder(OrderId, "пилястра 7,5 16", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                    Else
                    If FormElement.AddElementToOrder(OrderId, "пилястра 7,5 18", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                    End If
                ElseIf InStr(1, casename, "пилястр", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "пилястра", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П-14", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П14", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П-9", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П9", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П14", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П14", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П9", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П9", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ", vbTextCompare) > 0 And InStr(1, casename, "12", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П12(1900)", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П10", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П10", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П15", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П15", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П11", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П11", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                 ElseIf InStr(1, casename, "ПОРТАЛ П-10", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П10", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П-15", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П15", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                ElseIf InStr(1, casename, "ПОРТАЛ П-11", vbTextCompare) > 0 Then
                    If FormElement.AddElementToOrder(OrderId, "ПОРТАЛ П11", CaseQty) Then
                        ActiveCell.Interior.ColorIndex = 3
                    End If
                
                Else
                    
                    Select Case Left(casename, 1)
    '***************
    '*** шкафы *****
    '***************
                        Case "!"
                        Case "D", "K", "T", "V", "Y", "L", "A"
                        Dim almataName As String
                            almataName = casename
                            If InStr(1, casename, " ", vbTextCompare) > 1 Then
                                almataName = Mid(casename, 1, InStr(1, casename, " ", vbTextCompare) - 1)
                            End If
                                                    FormFitting.AddFittingToOrder OrderId, "ф-ра комплект Алмата", Empty, almataName, , , , row

                        Case "О", "П", "Ш", "Р" 'шкафы
                            If Left(casename, 2) = "ПС" And Cells(CasesPreampleRow, 2) = "" Then
                                GoTo stenki
                            End If
                            If Left(casename, 3) = "ШТВ" And Cells(CasesPreampleRow, 2) = "" Then
                                GoTo stenki
                            End If

                            If Left(casename, 2) = "ШВ" And Cells(CasesPreampleRow, 2) = "" Then
                                GoTo stenki
                            End If

                            If _
                            (Left(casename, 4) = "ПЛНД" Or _
                            Left(casename, 4) = "ПННД" _
                            ) _
                            And Cells(CasesPreampleRow, 2) = "" Then
                                GoTo stenki
                            End If
'                            Select Case CaseName
'                                Case "ПСН6", "ПСТ7", "ПСП1", "ПСП3", "ПСТ1", "ПСТ3", "ПСТ2", "ПСП2", _
'                                "ПСН1", "ПСЖ1", "ПСК1", "ПСК2", "ПСК3", "ПСК4", "ПСК5", "ПСН2", "ПСН3", "ПСН4", "ПСН5", "ПСШ1", "ПСШ2", "ПСШ3", "ПСШ4", "ПСТ6", "ПСТ5", "ПСН7", "ПСН8", "ПСТ4", _
'                                "ПСК(578)/4(203-4)", "ПСК(877)/4(176-4)", "ПСК(877)/4(223-4)", "ПСК(978)/4(176-4)", "ПСК(978)/4(223-4)", _
'                                "ПСК(578)/4(176-3,283)", "ПСК(1277)/4(296-4)полки стекло", "ПСК(1277)/5(713,176-4)", "ПСК(1277)/5(901,223-4)", "ПСК(1876)/6(713,176-4,713)", "ПСК(1876)/6(901,223-4,901)", _
'                                "ПСК(578)/2(640,176)", "ПСК(1227)/5(484-2,223,484-2)", "ПСК(578)/1(818)", _
'                                "ПСК(1277)/3(596-2,1196)", "ПСК(1277)/4(596-4)", "ПСК(1277)/6(396-6)", "ПСТ(1277)/2(223-2)", "ПСТ(1277)/2(396-2)", "ПСТ(678)/1(223)", "ПСТ(678)/1(396)", _
'                                "ПСТ(678)/2(296-2)", "ПСТ(1876)/3(223-3)", "ПСТ(1876)/3(396,223,396)", "ПСТ(1876)/3(396-3)", "ПСТ(2476)/3(396,223,396)", "ПСТ(1876)/4(396,197-2,396)", _
'                                "ПСТ(2476)/4(396,197-2,396)", "ПСТ(1876)/2(223-2)полка", "ПСТ(1876)/2(396-2)полка", _
'                                "ПСТ(1876)/2(596-2)полка", "ПСТ(678)/1(596)", "ПСТ(1277)/2(596-2)", "ПСТ(1876)/3(596-3)", _
'                                "ПСТ(1876)/3(596,296,596)", "ПСТ(2476)/3(596,296,596)", "ПСТ(1876)/4(596,296-2,596)", "ПСТ(2476)/4(596,296-2,596)", _
'                                "ПСН(478)/1(596)", "ПСН(478)/1(896)", "ПСН(478)/2(596-2)", "ПСН(678)/2(596-2)", _
'                                "ПСН(478)/1(1196)", "ПСН(678)/1(396)", "ПСН(678)/1(396)бар", "ПСН(678)/1(596)", _
'                                "ПСН(978)/1(396)", "ПСН(978)/1(396)бар", "ПСН(1277)/1(396)", "ПСН(1876)/1(396)полки", _
'                                "ПСН(1876)/1(596)полки", "ПСН(2176)/1(396)полки", "ПСН(678)/2(396-2)", "ПСН(877)/2(1196-2)", _
'                                "ПСН(1277)/2(396-2)", "ПСН(1277)/2(596-2)", "ПСН(1876)/2(396-2)полка", "ПСН(1876)/2(596-2)полка", _
'                                "ПСН(1876)/3(396-3)", "ПСН(1876)/3(596-3)", "ПСН(1277)/4(396-4)", "ПСН(1277)/4(596-4)", "ПСН(376)/1(996)витрина", _
'                                "ПСНТ(678-1200)/полки", "ПСНТ(678-1400)/полки", "ПСНТ(678-1573)/полки", "ПСШ(678)/1(1796)", "ПСШ(678)/3(596-3)", "ПСШ(1475)/4(1596-4)", "ПСШ(877)/6(396-2,1000-2,396-2)", "ПСШ(678)/2(596-2)полка", "ПСШ(877)/2(1796-2)", "ПСШ(1475)/4(1870-4)", "ПСШ(1475)/4(2074-4)", "ПСШ(1076)/6(496-2,1074-2,496-2)", _
'                                "ПСШ(877)/5(897-2,1346,223-2)", "ПСШ(678)/3(748,296,748)", "ПСШ(678)/4(596-2,296-2)", "ПСШ(678)/2(296-2)полки", "ПСШ(877)/4(1400-2,196-2)", "ПСШ(678)/полки"
'
'd
'                                    GoTo stenki
'
'                            End Select

                            

                            ' проверим тип ручек по умолчанию
                            If Not bHandleCheck Then
                                If IsEmpty(Handle) Then
                                    MsgBox "Не указан тип ручек!!!", vbCritical
                                    Handle = Null
                                ElseIf Not IsNull(Handle) Then
                                    CheckHandle Handle
                                End If
                                bHandleCheck = True
                            End If


                            Dim NQty
                            
                            Dim Width
                            Dim caseHeight As Integer
                            Set casepropertyCurrent = New caseProperty
                            casepropertyCurrent.init
                            casepropertyCurrent.p_fullcn = casename
                            ExcelCaseName = casepropertyCurrent.p_fullcn
                            casename = casepropertyCurrent.p_fullcn
                            
                            
                            ' определю признаки из названия шкафа
                            
                            
                            If casepropertyCurrent.p_cabType = 3 Then
                                 check_deleteZaveshki = True
                            ElseIf InStr(1, casepropertyCurrent.p_fullcn, "без завеш", vbTextCompare) > 0 Or InStr(1, casepropertyCurrent.p_fullcn, "б/св завеш", vbTextCompare) > 0 Then
                                If MsgBox("Шкаф без завешек?", vbQuestion + vbYesNo + vbDefaultButton1, "Завешки на каркас") = vbYes Then
                                    check_deleteZaveshki = True
                                End If
                            End If
                            
                            Select Case kitchenPropertyCurrent.changeCaseZaves
                                    
                                    Case 1:
                                    Cells(row, 9).Value = "Sensys"

                                    Cells(row, 9).Interior.ThemeColor = xlThemeColorLight2
                                    Cells(row, 9).Font.ThemeColor = xlThemeColorDark1
                                    Cells(row, 9).Font.Bold = True
                                    Case 2:
                                    Cells(row, 9).Value = "BluMot"
                                    Cells(row, 9).Interior.ThemeColor = xlThemeColorLight2
                                    Cells(row, 9).Font.ThemeColor = xlThemeColorDark1
                                    Cells(row, 9).Font.Bold = True
                                    Case 0:
                                    If kitchenPropertyCurrent.dspWidth = 16 And casepropertyCurrent.p_cabType <> 3 Then 'And InStr(1, fullCN, "П", vbTextCompare) = 1 Then
                                        Cells(row, 9).Value = "SlideOn"
                                        Cells(row, 9).Interior.ThemeColor = xlThemeColorLight2
                                        Cells(row, 9).Font.ThemeColor = xlThemeColorDark1
                                        Cells(row, 9).Font.Bold = True
                                    End If
                                    If (mRegexp.regexp_check(patCaseIsZovModul, casename)) Then
                                        Cells(row, 9).Value = Cells(row, 9).Value & "ZovMod"
                                        casename = casepropertyCurrent.p_casename
                                    End If
                            End Select
                            If casepropertyCurrent.p_cabType = 3 Then
                                Cells(row, 9).Interior.ThemeColor = xlThemeColorLight2
                                Cells(row, 9).Font.ThemeColor = xlThemeColorDark1
                                Cells(row, 9).Value = Cells(row, 9).Value & "Optima"
                            End If
                            
                           
                            
                            Do

                                
                                While caseFurnCollection.Count > 0
                                    caseFurnCollection.Remove (1)
                                Wend
                                While CaseElements.Count > 0
                                    CaseElements.Remove (1)
                                Wend
                                While casefasades.Count > 0
                                    casefasades.Remove (1)
                                Wend
                                While casezones.Count > 0
                                    casezones.Remove (1)
                                Wend

                                casepropertyCurrent.p_CaseColor = CaseColor
                                casepropertyCurrent.p_changeZaves = changeCaseZaves
                                
                                If mRegexp.regexp_check(patSHL_check2, casename) Then
                                    casename = parser.parse_case(casename)
                                ElseIf mRegexp.regexp_check(patNewName, casename) Then
                                    casename = parser.parse_case(casename)
                                End If
                                If casepropertyCurrent.p_newname <> "" Then
                                    ActiveCell.ClearComments
                                    ActiveCell.AddComment "разобран как " & Chr(10) & casepropertyCurrent.p_newname
                                End If
                                ParseCase casename, caseID, DoorCount, windowcount, Drawermount, Doormount, NoFace, Handle, HandleExtra, ShelfQty, Width, NQty, CaseColor, caseglub, caseHeight
                                If caseID = 0 Then casename = InputBox("введите наименование шкафа", "Идентификация шкафа", casename) 'parser.parse_case(fullCN))

                            Loop Until casename = "" Or caseID > 0

                            
                        
                            
                                   'комплекты фурнитуры
                            If InStr(1, casename, "ВЛШВ2", vbTextCompare) > 0 Then FormFitting.AddFittingToOrder OrderId, "ф-ра комплект ВЛШВ2", CaseQty, CaseColor, , caseID, , row
                           
                            If casepropertyCurrent.p_changeCaseKonfirmant = 1 Then
                                      Cells(row, FCol).Characters(k, 5).Font.Color = vbRed
                                      Cells(row, 33).Value = 1
                            End If

                            If Not IsNull(Handle) And Not IsEmpty(Handle) Then
                                Dim he
                                he = CheckHandleExtra(Handle)
                                If Not IsEmpty(he) Then HandleExtra = he
                            Else
                                HandleExtra = 0
                            End If

                            Dim isKarg As Boolean
                            isKarg = False

                        If caseID > 0 Then
                            If kitchenPropertyCurrent.dspWidth >= 18 And IsEmpty(HangColor) Then
                               HangColor = "камар806"
                               UpdateOrder OrderId, , HangColor
                            End If
                            If check_deleteZaveshki = False And casepropertyCurrent.p_z_st_dsp = False And casepropertyCurrent.p_dvpNahlest = False Then
                                If (Left(casename, 3) = "ШНУ" Or Left(casename, 4) = "ШНЗУ") Then
                                    If kitchenPropertyCurrent.dspWidth >= 18 Then
                                    additem2caseFittings OrderId, "завешка CAMAR 806Лев.", 1, , , caseID, , row
                                    End If
                                    localCaseHang = "завешка"
                                    If IsEmpty(HangColor) Then
                                        HangColor = GetHangColor(CaseColor)
                                        If HangColor = "КАМАР807" Then addItem2param "ножки удалить"
                                        UpdateOrder OrderId, , HangColor
                                    End If
                                End If
                                If InStr(1, casepropertyCurrent.p_fullcn, "камар807", vbTextCompare) > 0 Or InStr(1, casepropertyCurrent.p_fullcn, "camar807", vbTextCompare) > 0 Then
                                    'добавлю параметр к шкафу
                                    localCaseHang = "завешка"
                                    addItem2param "цвет завешки", "КАМАР807"
                                    addItem2param "ножки удалить"
                                    If IsEmpty(HangColor) Then
                                        HangColor = GetHangColor(CaseColor)
                                        UpdateOrder OrderId, , HangColor
                                    End If
                                    '-----
                                ElseIf InStr(1, casepropertyCurrent.p_fullcn, "камар808", vbTextCompare) > 0 Or InStr(1, casepropertyCurrent.p_fullcn, "camar808", vbTextCompare) > 0 Then
                                    localCaseHang = "завешка"
                                    'добавлю параметр к шкафу
                                    addItem2param "цвет завешки", "КАМАР808"
                                    '-----
                                    If IsEmpty(HangColor) Then
                                        HangColor = GetHangColor(CaseColor)
                                        UpdateOrder OrderId, , HangColor
                                    End If
                                ElseIf InStr(1, casepropertyCurrent.p_fullcn, "камар806", vbTextCompare) > 0 Or InStr(1, casepropertyCurrent.p_fullcn, "camar806", vbTextCompare) > 0 Then
                                    localCaseHang = "завешка"
                                    'добавлю параметр к шкафу
                                    addItem2param "цвет завешки", "КАМАР806"
                                    '-----
                                    If IsEmpty(HangColor) Then
                                        HangColor = GetHangColor(CaseColor)
                                        UpdateOrder OrderId, , HangColor
                                    End If
                                ElseIf is18(CaseColor) And localCaseHang = Empty Then
                                    If _
                                     (Left(casepropertyCurrent.p_fullcn, 3) = "ШНЗ" And Left(casepropertyCurrent.p_fullcn, 4) <> "ШНЗУ") _
                                        Or Left(casepropertyCurrent.p_fullcn, 3) = "ШНП" _
                                        Or (Left(casepropertyCurrent.p_fullcn, 2) = "ШН" And InStr(1, casepropertyCurrent.p_fullcn, "скос", vbTextCompare) > 4 And casepropertyCurrent.p_cabDepth = casepropertyCurrent.p_cabWidth) _
                                    Then
                                            localCaseHang = Empty
'                                        End If
'                                    ElseIf Then
'                                        localCaseHang = Empty
                                    Else
                                        localCaseHang = "завешка"
                                        If IsEmpty(HangColor) Then
                                            HangColor = "КАМАР806"
                                            UpdateOrder OrderId, , HangColor
                                        End If
                                    End If
                                ElseIf is18(CaseColor) = False And casepropertyCurrent.p_cabLevel = 2 And localCaseHang = Empty Then
                                    If _
                                     (Left(casepropertyCurrent.p_fullcn, 3) = "ШНЗ" And Left(casepropertyCurrent.p_fullcn, 4) <> "ШНЗУ") _
                                        Or Left(casepropertyCurrent.p_fullcn, 3) = "ШНП" _
                                        Or (Left(casepropertyCurrent.p_fullcn, 2) = "ШН" And InStr(1, casepropertyCurrent.p_fullcn, "скос", vbTextCompare) > 4 And casepropertyCurrent.p_cabDepth = casepropertyCurrent.p_cabWidth) _
                                    Then
                                            localCaseHang = Empty
'
                                    Else
                                        localCaseHang = "завешка" '!!!

                                        ' определим цвет для завешек
                                        If IsEmpty(HangColor) Then
                                            HangColor = GetHangColor(CaseColor)
                                            If HangColor = "КАМАР807" Then addItem2param "ножки удалить"
                                            UpdateOrder OrderId, , HangColor
                                        End If
                                    End If
                                End If
                            Else
                               'добавлю параметр к шкафу
                               addItem2param "без завешек", ""
                               
                            End If




                                If bBreakOrder And Not IsEmpty(CaseHang) Then
                                    UpdateOrder OrderId, , HangColor
                                End If

                                ' проверим столбец ф-ра
                                If Len(FCell) > 1 Then


                                    k = InStr(1, FCell, "клиент")
                                    If k Then
                                        If MsgBox("орг/мб/тб/напр клиента?", vbQuestion + vbYesNo + vbDefaultButton1, "Направлюящие на каркас") = vbYes Then
                                            Cells(row, FCol).Characters(k, 6).Font.Color = vbRed
                                            Drawermount = Null
                                            'добавлю параметр к шкафу
                                            addItem2param "направляюшие клиента", ""
                                            
                                            '-----
                                        End If
                                    End If

                                    k = InStr(1, FCell, "клиент")
                                    If k > 0 Then
                                        If InStr(1, FCell, "петл") > 0 And InStr(1, FCell, "петл") < k Then
                                            If MsgBox("петли клиента?", vbQuestion + vbYesNo + vbDefaultButton1, "Петли на каркас") = vbYes Then
                                                Cells(row, FCol).Characters(k, 6).Font.Color = vbRed
                                                Drawermount = Null
                                                'добавлю параметр к шкафу
                                                addItem2param "завесы удалить", ""
                                                '-----

                                            End If
                                        End If
                                    End If

                                    If IsEmpty(localCaseHang) And check_deleteZaveshki = False Then
                                        k = InStr(1, FCell, "зав", vbTextCompare)
                                        If k > 0 Then

                                            Cells(row, FCol).Characters(k, 3).Font.Color = vbRed
                                            If MsgBox("Завешка?", vbDefaultButton1 + vbYesNo + vbQuestion, "Крепление навесного шкафа") = vbYes Then
                                                localCaseHang = "завешка" '!!!

                                                If MsgBox("Установить ЗАВЕШКИ по умолчанию для всего заказа?", vbDefaultButton1 + vbYesNo + vbQuestion, "Крепление навесных шкафов") = vbYes Then
                                                    CaseHang = localCaseHang
                                                End If

                                                ' определим цвет для завешек
                                                If IsEmpty(HangColor) Then
                                                    HangColor = GetHangColor(CaseColor)
                                                    UpdateOrder OrderId, , HangColor
                                                End If
                                            End If



                                        End If

'                                    ElseIf bBreakOrder Then
'                                        UpdateOrder OrderID, , HangColor
                                    End If

                                    k = InStr(1, FCell, "sens")
                                    If k = 0 Then k = InStr(1, FCell, "сенси")

                                    If k > 0 Then
                                        If kitchenPropertyCurrent.changeCaseZaves <> 1 Then
                                            If MsgBox("Заменить завесы на СЕНСИС?", vbDefaultButton1 + vbYesNo + vbQuestion, "Крепление навесного шкафа") = vbYes Then
                                                Cells(row, FCol).Characters(k, 6).Font.Color = vbRed
                                                'добавлю параметр к шкафу
                                                casepropertyCurrent.p_changeZaves = 1
                                                'addItem2param "смена завесов", "Sensis"
                                            End If
                                        Else
                                            Cells(row, FCol).Characters(k, 6).Font.Color = vbRed
                                            'добавлю параметр к шкафу
                                            casepropertyCurrent.p_changeZaves = 1
                                            'addItem2param "смена завесов", "Sensis"
                                        End If
                                    End If
                                    
                                    k = InStr(1, FCell, "интермат")
                                    If k = 0 Then k = InStr(1, FCell, "intermat")

                                    If k > 0 Then
                                                Cells(row, FCol).Characters(k, 6).Font.Color = vbRed
                                                'добавлю параметр к шкафу
                                                casepropertyCurrent.p_changeZaves = 0
                                    End If
                                    

                                    k = InStr(1, FCell, "175")
                                    If k = 0 Then k = InStr(1, FCell, "180")
                                    If k Then

                                            If MsgBox("Удалить завесы из ШКАФА?", vbDefaultButton1 + vbYesNo + vbQuestion, "Крепление шкафа") = vbYes Then
                                                addItem2param "завесы удалить", ""
                                            End If

                                        Cells(row, FCol).Characters(k, 3).Font.Color = vbRed

                                        additem2caseFittings OrderId, "завес", Null, "175", , caseID, , row


                                    Else
                                    k = InStr(1, FCell, "165")
                                    If k Then

                                            If MsgBox("Удалить завесы из ШКАФА?", vbDefaultButton1 + vbYesNo + vbQuestion, "Крепление шкафа") = vbYes Then
                                                addItem2param "завесы удалить"
                                            End If
                                        Cells(row, FCol).Characters(k, 3).Font.Color = vbRed
                                        
                                        
                                        additem2caseFittings OrderId, "завес", Empty, "165", , caseID, , row
'

                                    Else

                                        k = InStr(1, FCell, "лифт")
                                        If k Then
                                            Cells(row, FCol).Characters(k, 4).Font.Color = vbRed
                                            
                                            Set caseFittingCurrent = New caseOrderFitting
                                            caseFittingCurrent.fName = "лифт"
                                            caseFittingCurrent.fQty = Empty
                                            CaseFittingsCollection.Add caseFittingCurrent
                                            
                                            additem2caseFittings OrderId, "лифт", Empty, , , caseID, , row

                                        Else

                                            k = InStr(1, FCell, "дов")
                                            If k Then
                                                Cells(row, FCol).Characters(k, 4).Font.Color = vbRed
                                                additem2caseFittings OrderId, "доводчик на метабокс", Empty, , , caseID, , row

                                            Else

                                                    k = InStr(1, FCell, "карг")
                                                    If k Then
                                                        Cells(row, FCol).Characters(k, 5).Font.Color = vbRed
                                                        additem2caseFittings OrderId, "карго", CaseQty, , , caseID, , row
                                                        Doormount = Null
                                                        isKarg = True

                                                    Else
                                                        k = InStr(1, FCell, "карусель")
                                                        If k Then
                                                            Cells(row, FCol).Characters(k, 5).Font.Color = vbRed
                                                            
                                                            additem2caseFittings OrderId, "карго", CaseQty, , , caseID, , row
                                                            Doormount = Null
                                                            isKarg = True
                                                    Else
                                                        k = InStr(1, FCell, "гарм")
                                                        If k Then
                                                            Cells(row, FCol).Interior.Color = vbRed
                                                        Else

                                                            k = InStr(1, FCell, "аморт")
                                                            If k > 0 Then
                                                                If InStr(1, FCell, "врезн") > 0 Then
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "врезной амортизатор", Empty, , , caseID, , row

                                                                ElseIf InStr(1, FCell, "blum") > 0 Or InStr(1, FCell, "блюм") > 0 Then
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "амортизатор BLUM", Empty, , , caseID, , row

                                                                ElseIf InStr(1, FCell, "врезн") > 0 Then
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "амортизатор FGV", Empty, , , caseID, , row
                                                                Else
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "амортизатор", Empty, , , caseID, , row

                                                                End If

                                                            Else
                                                            k = InStr(1, FCell, "амморт")
                                                            If k > 0 Then
                                                                If InStr(1, FCell, "врезн") > 0 Then
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "врезной амортизатор", Empty, , , caseID, , row

                                                                ElseIf InStr(1, FCell, "blum") > 0 Or InStr(1, FCell, "блюм") > 0 Then
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "амортизатор BLUM", Empty, , , caseID, , row

                                                                ElseIf InStr(1, FCell, "врезн") > 0 Then
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "амортизатор FGV", Empty, , , caseID, , row
                                                                Else
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    additem2caseFittings OrderId, "амортизатор", Empty, , , caseID, , row
                                                                End If

                                                            Else
                                                                k = InStr(1, FCell, "HK")
                                                                If k = 0 Then k = InStr(1, FCell, "НК")
                                                                If k = 0 Then k = InStr(1, FCell, "НK")
                                                                If k = 0 Then k = InStr(1, FCell, "HК")

                                                                If k > 0 Then
                                                                    Cells(row, FCol).Interior.Color = vbYellow
                                                                    If InStr(1, FCell, "XS") > 1 Then
                                                                        
                                                                        additem2caseFittings OrderId, "завес HK-XS", Empty, , , caseID, , row
                                                                        
                                                                    Else
                                                                        'добавлю параметр к шкафу
                                                                        addItem2param "завесы удалить"
                                                                        '-----
                                                                        If InStr(1, FCell, "27") > 1 And InStr(1, FCell, "TIP") > 1 Then
                                                                            Doormount = Null
                                                                            additem2caseFittings OrderId, "завес HK27 (TIP-ON)", Null, "Хром", , caseID, , row
                                                                            
                                                                        ElseIf InStr(1, FCell, "25") > 1 And InStr(1, FCell, "TIP") > 1 Then
                                                                            additem2caseFittings OrderId, "завес HK25 (TIP-ON)", Null, "Хром", , caseID, , row
                                                                            Doormount = Null

                                                                        ElseIf InStr(1, FCell, "25") > 1 Then
                                                                           additem2caseFittings OrderId, "завес HK25", Null, "Хром", , caseID, , row
                                                                            Doormount = Null
                                                                            
                                                                        ElseIf InStr(1, FCell, "27") > 1 Then
                                                                            additem2caseFittings OrderId, "завес HK27", Null, "Хром", , caseID, , row
                                                                            Doormount = Null
                                                                            
                                                                         ElseIf InStr(1, FCell, "29") > 1 And InStr(1, FCell, "TIP") > 1 Then
                                                                            additem2caseFittings OrderId, "завес HK29 (TIP-ON)", Null, "Хром", , caseID, , row
                                                                            Doormount = Null
                                                                            
                                                                        ElseIf InStr(1, FCell, "29") > 1 Then
                                                                             additem2caseFittings OrderId, "завес HK29", Null, "Хром", , caseID, , row
                                                                            Doormount = Null
                                                                            
                                                                        ElseIf InStr(1, FCell, "hk-s") > 0 Or InStr(1, FCell, "hks") > 0 Then
                                                                            additem2caseFittings OrderId, "завес HK-S", Null, "Хром", , caseID, , row
                                                                            Doormount = Null
                                                                         
                                                                        ElseIf InStr(1, FCell, "hk") > 0 Then
                                                                            additem2caseFittings OrderId, "завес HK27", Null, "Хром", , caseID, , row
                                                                            Doormount = Null
                                                                            
                                                                        Else
                                                                            additem2caseFittings OrderId, "завес", Null, "Хром", , caseID, , row
                                                                            Doormount = Null
                                                                            
                                                                        End If
                                                                    End If

                                                                Else

                                                                k = InStr(1, FCell, "HL")

                                                                If k = 0 Then k = InStr(1, FCell, "HL23/35")
                                                                If k = 0 Then k = InStr(1, FCell, "HL23/38")
                                                                If k = 0 Then k = InStr(1, FCell, "HL25/35")
                                                                If k = 0 Then k = InStr(1, FCell, "HL25/38")
                                                                If k = 0 Then k = InStr(1, FCell, "HL27/35")
                                                                If k = 0 Then k = InStr(1, FCell, "HL27/38")
                                                                If k = 0 Then k = InStr(1, FCell, "HL23/39")
                                                                If k = 0 Then k = InStr(1, FCell, "HL25/39")
                                                                If k = 0 Then k = InStr(1, FCell, "HL27/39")
                                                                If k = 0 Then k = InStr(1, FCell, "HL29/39")

                                                                If k > 0 Then
                                                                    'добавлю параметр к шкафу
                                                                    addItem2param "завесы удалить"
                                                                   
                                                                    Cells(row, FCol).Interior.Color = vbYellow

                                                                    additem2caseFittings OrderId, "завес", CaseQty, Empty, , caseID, , row
                                                                    Doormount = Null
                                                                   
                                                               Else

                                                                k = InStr(1, FCell, "HS")
                                                                If k = 0 Then k = InStr(1, FCell, "HS A")
                                                                If k = 0 Then k = InStr(1, FCell, "HS B")
                                                                If k = 0 Then k = InStr(1, FCell, "HS D")
                                                                If k = 0 Then k = InStr(1, FCell, "HS E")
                                                                If k = 0 Then k = InStr(1, FCell, "HS G")
                                                                If k = 0 Then k = InStr(1, FCell, "HS H")
                                                                If k = 0 Then k = InStr(1, FCell, "HS I")
                                                                If k > 0 Then
                                                                    'добавлю параметр к шкафу
                                                                    addItem2param "завесы удалить"
                                                                    '-----
                                                                    Cells(row, FCol).Interior.Color = vbYellow

                                                                    additem2caseFittings OrderId, "завес", CaseQty, Empty, , caseID, , row
                                                                    Doormount = Null
                                                                    


                                                                Else
                                                                    k = InStr(1, FCell, "FB-1")
                                                                    If k = 0 Then k = InStr(1, Trim(ActiveCell.Value), "FB-1")
                                                                    If k = 0 Then k = InStr(1, FCell, "FВ-1")
                                                                    If k = 0 Then k = InStr(1, Trim(ActiveCell.Value), "FВ-1")


                                                                    If k > 0 Then

                                                                        Cells(row, FCol).Interior.Color = vbYellow

                                                                        additem2caseFittings OrderId, "завес", Empty, "FВ-1", , caseID, , row
                                                                        Doormount = Null
                                                                        

                                                                        If DoorCount = 2 Then DoorCount = 1

                                                                    Else

                                                                        If k = 0 Then k = InStr(1, FCell, "HF22")
                                                                        If k = 0 Then k = InStr(1, ActiveCell.Value, "HF22")
                                                                        If k = 0 Then k = InStr(1, FCell, "НF22")
                                                                        If k = 0 Then k = InStr(1, ActiveCell.Value, "НF22")

                                                                        If k > 0 Then
                                                                            'добавлю параметр к шкафу
                                                                            addItem2param "завесы удалить"
                                                                           
                                                                            Cells(row, FCol).Interior.Color = vbYellow

                                                                            additem2caseFittings OrderId, "завес HF22", Empty, "ХРОМ", , caseID, , row
                                                                            Doormount = Null
                                                                            

                                                                            If DoorCount = 2 Then DoorCount = 1

                                                                        Else

                                                                            If k = 0 Then k = InStr(1, FCell, "НF25")
                                                                            If k = 0 Then k = InStr(1, ActiveCell.Value, "НF25")
                                                                            If k = 0 Then k = InStr(1, FCell, "HF25")
                                                                            If k = 0 Then k = InStr(1, ActiveCell.Value, "HF25")

                                                                            If k > 0 Then
                                                                                'добавлю параметр к шкафу
                                                                                addItem2param "завесы удалить"
                                                                                Cells(row, FCol).Interior.Color = vbYellow

                                                                                additem2caseFittings OrderId, "завес HF25", Empty, "ХРОМ", , caseID, , row
                                                                                Doormount = Null
                                                                                

                                                                                If DoorCount = 2 Then DoorCount = 1

                                                                            Else

                                                                                If k = 0 Then k = InStr(1, FCell, "HF28")
                                                                                If k = 0 Then k = InStr(1, ActiveCell.Value, "HF28")
                                                                                If k = 0 Then k = InStr(1, FCell, "НF28")
                                                                                If k = 0 Then k = InStr(1, ActiveCell.Value, "НF28")

                                                                                If k > 0 Then
                                                                                    'добавлю параметр к шкафу
                                                                                    addItem2param "завесы удалить"
                                                                                    
                                                                                    Cells(row, FCol).Interior.Color = vbYellow

                                                                                    additem2caseFittings OrderId, "завес HF28", Empty, "Хром", , caseID, , row
                                                                                    Doormount = Null
                                                                                    

                                                                                    If DoorCount = 2 Then DoorCount = 1

                                                                                Else

                                                                                            If k = 0 Then k = InStr(1, FCell, "HF")
                                                                                            If k = 0 Then k = InStr(1, ActiveCell.Value, "HF")
                                                                                            If k = 0 Then k = InStr(1, FCell, "НF")
                                                                                            If k = 0 Then k = InStr(1, ActiveCell.Value, "НF")

                                                                                            If k > 0 Then
                                                                                                'добавлю параметр к шкафу
                                                                                               addItem2param "завесы удалить"
                                                                                                
                                                                                                Cells(row, FCol).Interior.Color = vbYellow

                                                                                               additem2caseFittings OrderId, "завес HF", Empty, "ХРОМ", , caseID, , row
                                                                                                Doormount = Null
                                                                                                

                                                                                                If DoorCount = 2 Then DoorCount = 1
                                                                                            Else
                                                                                                Cells(row, 33).Value = 0
                                                                                                If k = 0 Then k = InStr(1, FCell, "стяж")
                                                                                                    If k > 0 Then

                                                                                                        Cells(row, FCol).Characters(k, 5).Font.Color = vbRed
                                                                                                        Cells(row, 33).Value = 1
                                                                                                    Else
                                                                                                        If (FCell Like "*push*open*") Then k = InStr(1, FCell, "push")
                                                                                                        If k = 0 Then If (FCell Like "*пуш*опен*") Then k = InStr(1, FCell, "пуш")
                                                                                                        If k = 0 Then k = InStr(1, FCell, "p2o")
                                                                                                        If k = 0 Then k = InStr(1, FCell, "р2о")
                                                                                                        If k > 0 Then
                                                                                                            additem2caseFittings OrderId, "нажимной м-м Push-To-Open", Empty, , , caseID, , row
                                                                                                            Cells(row, FCol).Characters(k, 5).Font.Color = vbYellow
                                                                                                        End If
                                                                                                  End If
                                                                                            End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                            End If
                                                              End If
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                            End If
                                        End If
                                    End If

'                                    If k = 0 Then
'                                        FormFitting.AddFittingToOrder OrderID, "", Empty, , , caseID, , Row
'                                    End If

                                End If ' If Len(FCell) > 2 Then



                                k = InStr(1, Cells(row, 3), "Z")
                                If k = 0 Then
                                    If InStr(1, Cells(row, 3), "A") + InStr(1, Cells(row, 3), "А") > 0 Then

                                        k = InStr(1, Cells(row, 3), "A1") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A2") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A3") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A4") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A5") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A6") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A7") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A8") 'англ A
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "A9") 'англ A

                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А1") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А2") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А3") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А4") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А5") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А6") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А7") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А8") ' рус А
                                        If k = 0 Then k = InStr(1, Cells(row, 3), "А9") ' рус А
                                    End If
                                End If
                                If k Then
                                    Cells(row, 3).Characters(k, 1).Font.Color = vbRed
                                    additem2caseFittings OrderId, "саморез", Empty, , , caseID, , row
                                End If

                                If Not isKarg Then
                                    If InStr(1, Trim(ActiveCell.Value), "карг") > 0 Then
                                        additem2caseFittings OrderId, "карго", CaseQty, , , caseID, , row
                                        Doormount = Null
                                    End If
                                End If
                                If InStr(1, Trim(ActiveCell.Value), "стяж") > 0 Then
                                      Cells(row, FCol).Characters(k, 5).Font.Color = vbRed
                                      Cells(row, 33).Value = 1
                               End If

                                
                                If InStr(1, Trim(ActiveCell.Value), "полк") And InStr(1, Trim(ActiveCell.Value), "стекл") > 0 Then
                                    addItem2param "полкодержатели удалить"
                                    If xInArray(casename, Array("ШНП", "ШНП915", "ШНЗ915", "ШНЗУ", "ШНЗУ915", "ШЛЗУ", "ШЛП", "ШНЗВ", "ШСП", "ШНЗ", "ШЛЗ", "ШСЗ")) Then
                                        additem2caseFittings OrderId, "полкодержатель", Null, "тип C", , caseID, , row
                                    ElseIf InStr(1, casepropertyCurrent.p_fullcn, "шн", vbTextCompare) = 1 And InStr(1, casepropertyCurrent.p_fullcn, "скос", vbTextCompare) > 4 And xInArray(CStr(casepropertyCurrent.p_cabWidth), Array("300", "400")) And casepropertyCurrent.p_cabDepth = 300 Then
                                        additem2caseFittings OrderId, "полкодержатель", Null, "тип C", , caseID, , row
                                    ElseIf InStr(1, casepropertyCurrent.p_fullcn, "шл", vbTextCompare) = 1 And InStr(1, casepropertyCurrent.p_fullcn, "скос", vbTextCompare) > 4 And xInArray(CStr(casepropertyCurrent.p_cabWidth), Array("300")) And casepropertyCurrent.p_cabDepth = 300 Then
                                        additem2caseFittings OrderId, "полкодержатель", Null, "тип C", , caseID, , row
                                    Else
                                        additem2caseFittings OrderId, "полкодержатель", Null, "Sekura 8 (для стекла)", , caseID, , row
                                    End If
                                ElseIf InStr(1, Trim(ActiveCell.Value), "полк") > 0 Then
                                    additem2caseFittings OrderId, "полкодержатель", Null, "5", , caseID, , row
                                End If

                               

                                Select Case casename
                                    Case "ШЛ скос", "ШЛ скос915"
                                        If IsNull(Doormount) Then
                                            additem2caseFittings OrderId, "Завес", 2 * CaseQty, , , caseID, , row
                                        End If
                                    Case "ШН скос", "ШН скос915"
                                        If IsNull(Doormount) Then
                                            additem2caseFittings OrderId, "Завес", 2 * CaseQty, , , caseID, , row
                                        ElseIf Doormount = "+20" Then
                                            additem2caseFittings OrderId, "угловой адаптер 10гр", CaseQty * 2 * DoorCount, , , caseID, , row
                                        ElseIf Width = "25" Then
                                            additem2caseFittings OrderId, "Завес CLIP top", CaseQty * 2 * DoorCount, "BLUMOTION +45", , caseID, , row
                                            addItem2param "завесы удалить"
                                            
                                        End If

                                    Case "ШЛП", "ШСП", "ШНП", "ШНП915" ' если шкафы с гнутыми фасадами, добавим завесы
                                        If InStr(1, Trim(ActiveCell.Value), "ДГГ") > 0 Then

                                            additem2caseFittings OrderId, "завес", CaseQty * 2, "полусофт", , caseID, , row
                                            additem2caseFittings OrderId, cHandle, CaseQty, Handle, , caseID, , row

                                        ElseIf InStr(1, Trim(ActiveCell.Value), "ДГВ") > 0 Then

                                            additem2caseFittings OrderId, "стеклодержатель", CaseQty * 8, , , caseID, , row
                                            additem2caseFittings OrderId, "завес", CaseQty * 2, "110", , caseID, , row
                                            additem2caseFittings OrderId, cHandle, CaseQty, Handle, , caseID, , row
                                        End If

                                    Case "ШНЗ", "ШСЗ", "ШЛЗ"
                                        If InStr(1, Trim(ActiveCell.Value), "ДГВ") > 0 Then
                                            additem2caseFittings OrderId, "стеклодержатель", CaseQty * 8, , , caseID, , row
                                        End If
                                End Select

                                Select Case Left(casename, 1)
                                    Case "П"
                                        If CDec(DoorCount) + CDec(NQty) > 2 Then
                                            additem2caseElements OrderId, "полик", (CDec(DoorCount) + NQty - 2), caseID
                                            'FormElement.AddElementToOrder orderid, "полик", CaseQty * (CDec(DoorCount) + NQty - 2), caseID
                                        End If
                                        If (DoorCount) > 0 Then
                                            If ActiveCell.Font.Bold Then
                                                additem2caseFittings OrderId, "завес", CaseQty, , , caseID, , row
                                            Else ' если стандартный пенал, лишних завесов не надо
                                                additem2caseFittings OrderId, "завес", CaseQty, , , caseID, True, row
                                            End If
                                        End If
                                End Select
                                
                                Dim dvpQty As Integer
                                dvpQty = 0
                                
                                If casepropertyCurrent.p_dvpNahlest Then
                                    addItem2param "зад стенка", "двп в нахл"
                                ElseIf casepropertyCurrent.p_z_st_dsp Then
                                    addItem2param "зад стенка", "дсп"
                                   
'                                    If Mid(casename, 2, 1) = "Н" Then
'                                        dvpQty = 2 * Round((casepropertyCurrent.p_cabHeigth / 200), 0) + 2 * Round((casepropertyCurrent.p_cabWidth / 200), 0)
'                                        additem2caseFittings OrderID, "крепление ДВП в паз RV-8", CaseQty * dvpQty, , , caseID, , Row
'                                    ElseIf Mid(casename, 2, 1) = "Л" Then
'                                        dvpQty = 2 * Round((casepropertyCurrent.p_cabHeigth / 200), 0)
'                                        additem2caseFittings OrderID, "крепление ДВП в паз RV-8", CaseQty * dvpQty, , , caseID, , Row
'                                    End If
                                Else
                                     addItem2param "зад стенка", "двп в паз"
                                End If
                               

'                                'двп паз для 16-ки
'                                 If Not is18(CaseColor) And InStr(1, fullCN, "паз", vbTextCompare) > 0 Then
'
'                                        Select Case casename
'                                           ' Case "ШНЗ", "ШНЗ915", "ШНП", "ШНП915", "ШН скос", "ШН скос915"
'
'                                            Case "ШНУ", "ШНУГ", "ШНЗУ", "ШНУР"
'                                                FormFitting.AddFittingToOrder OrderID, "крепление ДВП в паз RV-8", CaseQty * 14, , , caseID, , Row
'
'                                            Case "ШНУ915", "ШНУГ915", "ШНЗУ915", "ШНУР915"
'                                                FormFitting.AddFittingToOrder OrderID, "крепление ДВП в паз RV-8", CaseQty * 16, , , caseID, , Row
'
'                                            Case Else
'                                                If Left(casename, 2) = "ПН" Then
'                                                    FormFitting.AddFittingToOrder OrderID, "крепление ДВП в паз RV-8", CaseQty * 10, , , caseID, , Row
'                                                ElseIf InStr(1, casename, "915") > 0 Then
'                                                    FormFitting.AddFittingToOrder OrderID, "крепление ДВП в паз RV-8", CaseQty * 10, , , caseID, , Row
'                                                Else
'                                                    FormFitting.AddFittingToOrder OrderID, "крепление ДВП в паз RV-8", CaseQty * 8, , , caseID, , Row
'                                                End If
'                                        End Select
'                                End If

                                ' получим длину шурупов для ручек
                                If IsEmpty(HandleScrew) Then
                                    HandleScrew = GetHandleScrew(Handle, face)
                                ElseIf bBreakOrder Then
                                    UpdateOrder OrderId, HandleScrew
                                End If
                                If Not IsEmpty(HandleScrew) Then UpdateOrder OrderId, HandleScrew

                                '***************************************************************************************

                                If Not IsEmpty(windowcount) Then
                                    If (kitchenPropertyCurrent.addGlassHolders And casepropertyCurrent.p_addGlassHolders) = False Then
                                        windowcount = Empty
                                    End If
'                                    If Not IsNull(face) And Not IsEmpty(face) Then
'                                        If InStr(1, face, "прим", vbTextCompare) > 0 And InStr(1, face, "женева", vbTextCompare) = 0 Then
'                                            windowcount = Empty
'
'                                        ElseIf InStr(1, face, "софт", vbTextCompare) > 0 Or _
'                                                InStr(1, face, "ДСП", vbTextCompare) > 0 Then
'
'                                            If MsgBox("Давать стеклодержатели?", vbQuestion + vbYesNo + vbDefaultButton1, face) = vbNo Then
'                                                windowcount = Empty
'                                            End If
'                                        End If
'                                    End If

                                    If casename = "ШНВ А" Or casename = "ШНВ З" Then windowcount = Empty

                                End If

                                 If InStr(1, ExcelCaseName, "ШЛК", vbTextCompare) = 1 Then
                                If mRegexp.regexp_check(patSHLK_check2, ExcelCaseName) Then
                                    'If casepropertyCurrent.p_changeZaves = 0 Then
'                                        If casepropertyCurrent.p_changeZaves = 2 Then
'
'                                        Else
'
'                                        End If
                                        If casepropertyCurrent.p_changeZaves = 2 Then
                                       additem2caseFittings OrderId, "Завес CLIP top", 2 * CaseQty, "+155", , caseID, , row
                                        Else
                                            additem2caseFittings OrderId, "Завес", 2 * CaseQty, "175", , caseID, , row
                                       End If
                                        addItem2param "завесы удалить"
                                   ' ElseIf casepropertyCurrent.p_changeZaves = 1 Then
                                   '     additem2caseFittings OrderID, "завес Sensys", 2 * CaseQty, "165", , caseID, , Row
                                        'FormFitting.AddFittingToOrder OrderID, "амморт. Sensys 165", 2 * CaseQty, , , caseID, , Row
                                    'End If
                                    additem2caseFittings OrderId, "Завес CLIP top", 2 * CaseQty, "BLUMOTION +90 под фп", , caseID, , row
                                Else
                                    If (InStr(1, ExcelCaseName, "лев", vbTextCompare) > 0 And InStr(1, ExcelCaseName, "прав", vbTextCompare) > 0) Then
                                    
                                    additem2caseFittings OrderId, "Завес CLIP top", 2 * CaseQty, "BLUMOTION +90 под фп", , caseID, , row
                                        ElseIf InStr(1, ExcelCaseName, "лев", vbTextCompare) > 0 Or InStr(1, ExcelCaseName, "прав", vbTextCompare) > 0 Then
                                    
                                       ' If casepropertyCurrent.p_changeZaves = 0 Then
                                       
                                       If casepropertyCurrent.p_changeZaves = 2 Then
                                       additem2caseFittings OrderId, "Завес CLIP top", 2 * CaseQty, "+155", , caseID, , row
                                        Else
                                            additem2caseFittings OrderId, "Завес", 2 * CaseQty, "175", , caseID, , row
                                       End If
                                            addItem2param "завесы удалить"
                                      '  ElseIf casepropertyCurrent.p_changeZaves = 1 Then
                                        '    additem2caseFittings OrderID, "завес Sensys", 2 * CaseQty, "165", , caseID, , Row
                                            'FormFitting.AddFittingToOrder OrderID, "амморт. Sensys 165", 2 * CaseQty, , , caseID, , Row
                                        'End If
                                    End If
                                End If
                            End If
                            
                            For paramIterator = 1 To params.Count
                                If params(paramIterator).paramName = "завесы удалить" Then
                                    Exit For
                                End If
                            Next paramIterator
                            If paramIterator > params.Count Then
                                paramIterator = params.Count
                            End If
                            If params(paramIterator).paramName <> "завесы удалить" Then
                            tempString = casepropertyCurrent.p_fullcn
                                k = 0
                                If k = 0 Then k = InStr(1, tempString, "HKXS")
                                If k = 0 Then k = InStr(1, tempString, "HK-XS")
                                If k > 0 Then
                                    'добавлю параметр к шкафу
                                    additem2caseFittings OrderId, "завес HK-XS", Empty, "ХРОМ", , caseID, , row
                                    tempString = Replace(tempString, "HKXS", "", , , vbTextCompare)
                                    tempString = Replace(tempString, "HK-XS", "", , , vbTextCompare)
                                End If
                                k = 0
                                If k = 0 Then k = InStr(1, tempString, "HF")
                                If k = 0 Then k = InStr(1, tempString, "HF")
                                If k = 0 Then k = InStr(1, tempString, "НF")
                                If k = 0 Then k = InStr(1, tempString, "НF")
                                If k > 0 Then
                                    'добавлю параметр к шкафу
                                   addItem2param "завесы удалить"
                                    
                                    additem2caseFittings OrderId, "завес HF", Empty, "ХРОМ", , caseID, , row
                                    Doormount = Null
                                    tempString = Replace(tempString, "HF", "", , , vbTextCompare)
    
                                    If DoorCount = 2 Then DoorCount = 1
                                End If
                                 k = 0
                                If k = 0 Then k = InStr(1, tempString, "HS")
                                If k > 0 Then
                                    'добавлю параметр к шкафу
                                    addItem2param "завесы удалить"
                                    additem2caseFittings OrderId, "завес HS", Empty, "ХРОМ", , caseID, , row
                                    tempString = Replace(tempString, "HS", "", , , vbTextCompare)
                                    Doormount = Null
                                End If
                                k = 0
                                If k = 0 Then k = InStr(1, tempString, "HL")
                                If k > 0 Then
                                    'добавлю параметр к шкафу
                                    addItem2param "завесы удалить"
                                    additem2caseFittings OrderId, "завес HL", Empty, "ХРОМ", , caseID, , row
                                    tempString = Replace(tempString, "HL", "", , , vbTextCompare)
                                    Doormount = Null
                                End If
                                k = 0
                                If k = 0 Then k = InStr(1, tempString, "HK-s")
                                If k = 0 Then k = InStr(1, tempString, "HKs")
                                
                                If k > 0 Then
                                    'добавлю параметр к шкафу
                                    addItem2param "завесы удалить"
                                    additem2caseFittings OrderId, "завес HK-S", Empty, "ХРОМ", , caseID, , row
                                    tempString = Replace(tempString, "HK-s", "", , , vbTextCompare)
                                    tempString = Replace(tempString, "HKs", "", , , vbTextCompare)
                                    Doormount = Null
                                Else
                                    If k = 0 Then k = InStr(1, tempString, "HK")
                                    If k = 0 And InStr(5, tempString, "НК") > 5 Then k = InStr(5, tempString, "НК")
                                    If k > 0 Then
                                        'добавлю параметр к шкафу
                                        
                                        addItem2param "завесы удалить"

                                        additem2caseFittings OrderId, "завес HK27", Empty, "ХРОМ", , caseID, , row
                                        tempString = Replace(tempString, "HK", "", , , vbTextCompare)
                                        Doormount = Null
                                    End If
                                End If
                                
                                
                                
                            End If
                            
                            If ShelfQty >= 2 And Left(casename, 1) <> "П" Then
                                Select Case casename
                                    Case "ШН 2Т"
                                        additem2caseElements OrderId, "полик", qty, caseID
                                    Case "ШН915", "ШНП915", "ШНУ915", "ШН скос915", "ШНЗ915", "ШНЗУ915", "ШНУГ915", "ШНС915"
                                    Case "ШЛП", "ШСП"   '"ШНП", "ШН скос"
                                        additem2caseElements OrderId, "полик", qty, caseID
                                    Case "ШНВ А"
                                        additem2caseFittings OrderId, "крестик", qty, , , caseID
                                    Case Else
                                        additem2caseElements OrderId, "полка", qty, caseID
                                End Select
                            End If


                            
                            
                            If InStr(1, casepropertyCurrent.p_fullcn, "б/ф", vbTextCompare) > 0 Then
                                    'добавлю параметр к шкафу
                                    addItem2param "завесы удалить"
                                    
                                    Doormount = Null
                                    DoorCount = 0
                                    windowcount = 0
                                    
                                    Cells(row, 17).Value = 0
                                    Cells(row, 18).Value = 0
                            End If
                            If InStr(1, casepropertyCurrent.p_fullcn, "ШНК", vbTextCompare) = 1 Then
                                    additem2caseFittings OrderId, "Завес CLIP top", 2 * CaseQty, , , caseID, , row
                                    additem2caseFittings OrderId, "Завес", 2 * CaseQty * DoorCount, , , caseID, , row
                                    addItem2param "завесы удалить"
                            End If


                                '***************************************************************************************
                                Dim qqqqqqqqq
                                qqqqqqqqq = (Cells(row, 32).Value)
                                


                               ' AddCase OrderID, caseID, CaseName, ActiveCell.Value, CaseQty, localCaseHang, Handle, HandleExtra, Leg, DoorCount, windowcount, Drawermount, Doormount, 1, caseglub, ShelfQty, ActiveCell.Font.Bold, Row, NoFace, changeCaseZaves, changeCaseKonfirmanttemp
                                'шкаф всегда теперь нестовый
                                
                               
                                OrderCaseID = AddCaseBySp(OrderId, caseID, casename, ActiveCell.Value, CaseQty, localCaseHang, Handle, HandleExtra, Leg, DoorCount, windowcount, Drawermount, Doormount, 1, qqqqqqqqq, ShelfQty, True, row, NoFace, casepropertyCurrent.p_cabTypeName)
'                                OrderCasesID = AddCaseBySp(OrderID, caseID, casename, ActiveCell.Value, CaseQty, localCaseHang, Handle, HandleExtra, Leg, DoorCount, windowcount, Drawermount, Doormount, 1, qqqqqqqqq, ShelfQty, ActiveCell.Font.Bold, Row, NoFace, changeCaseZaves, changeCaseKonfirmanttemp, dspbottom, caseHeight)
                                                                                                
                                
                                
                                
                                If OrderCaseID > 0 And params.Count > 0 Then

                                    For paramIterator = 1 To params.Count
                                        If params(paramIterator).paramName = "направляюшие клиента" Then
                                            casepropertyCurrent.p_CustomerDrawermount = True
                                        End If
                                        AddCaseParamsbySp OrderCaseID, params(paramIterator).paramName, params(paramIterator).paramValue
                                    Next paramIterator

                                End If

                                If OrderCaseID > 0 Then
                                    
                                    For i = 1 To CaseFittingsCollection.Count
                                        If CaseFittingsCollection(i).ismissingfQty Then
                                            FormFitting.AddFittingToOrder OrderId, CaseFittingsCollection(i).fName, Null, CaseFittingsCollection(i).fOption, CaseFittingsCollection(i).fLength, caseID, , row
                                        Else
                                            FormFitting.AddFittingToOrder OrderId, CaseFittingsCollection(i).fName, CaseFittingsCollection(i).fQty, CaseFittingsCollection(i).fOption, CaseFittingsCollection(i).fLength, caseID, , row
                                        End If
                                    Next i
                                    
                                    For i = 1 To CaseElementsCollection.Count
                                        FormElement.AddElementToOrder OrderId, CaseElementsCollection(i).name, CaseElementsCollection(i).qty, caseID
                                    Next i
                                End If

                                If Not (casepropertyCurrent Is Nothing) Then
                                    If casepropertyCurrent.p_newsystem Then
                                    casename = casepropertyCurrent.p_casename
                                    End If
                                    If Left(casepropertyCurrent.p_fullcn, 2) = "ПЛ" And InStr(1, casepropertyCurrent.p_fullcn, "нд", vbTextCompare) And InStr(1, casepropertyCurrent.p_fullcn, "599", vbTextCompare) > 0 Then
                                         FormFitting.AddFittingToOrder OrderId, "дюбель DU325 Rapid S", 4, , , caseID, , row
                                    End If
                                    
                                    If ((Left(casepropertyCurrent.p_fullcn, 3) <> "ШНП" And Left(casepropertyCurrent.p_fullcn, 2) = "ШН") Or Left(casepropertyCurrent.p_fullcn, 2) = "ШЛ") And casepropertyCurrent.p_haveNisha And casepropertyCurrent.p_NishaQty >= 1 Then
                                        
                                        Dim cfcElem As caseElement
                                        For Each cfcElem In CaseElements
                                            FormElement.AddElementToOrder OrderId, cfcElem.name, cfcElem.qty, caseID
                                        Next cfcElem
                                        
                                    End If
                                    If caseID > 0 And caseFurnCollection.Count > 0 Then
                                        Dim cfcItem As caseFurniture
                                        For Each cfcItem In caseFurnCollection
                                            If (cfcItem.fType = "drawermount" And casepropertyCurrent.p_CustomerDrawermount = False) Or cfcItem.fType = "" _
                                            Then
                                                If cfcItem.fOption <> "не давать" And cfcItem.fLength <> "не давать" Then
                                                    If cfcItem.fOption <> "" And cfcItem.fLength <> "" Then
                                                        FormFitting.AddFittingToOrder OrderId, cfcItem.fName, cfcItem.qty, cfcItem.fOption, cfcItem.fLength, caseID, , row
                                                    ElseIf cfcItem.fOption <> "" And cfcItem.fLength = "" Then
                                                        FormFitting.AddFittingToOrder OrderId, cfcItem.fName, cfcItem.qty, cfcItem.fOption, , caseID, , row
                                                    ElseIf cfcItem.fOption = "" And cfcItem.fLength <> "" Then
                                                        FormFitting.AddFittingToOrder OrderId, cfcItem.fName, cfcItem.qty, , cfcItem.fLength, caseID, , row
                                                    Else
                                                        FormFitting.AddFittingToOrder OrderId, cfcItem.fName, cfcItem.qty, , , caseID, , row
                                                    End If
                                                End If
                                            End If
                                        Next cfcItem
                                    End If
                                    If OrderCaseID > 0 And casepropertyCurrent.p_cabTypeName = "MODUL" Then
                                        If casepropertyCurrent.p_caseLetters = "ШНУ" Then
                                            FormFitting.AddFittingToOrder OrderId, "планка монтажная 100мм", 3, , , caseID, , row
                                        ElseIf casepropertyCurrent.p_cabLevel = 2 Then
                                            FormFitting.AddFittingToOrder OrderId, "планка монтажная 100мм", 2, , , caseID, , row
                                        End If
                                    End If
                                   ' планка монтажная 100мм
                                End If

                                Selection.Interior.Color = RGB(173, 255, 47)
                            Else
                                Selection.Interior.Color = vbRed
                            End If

    '***************
    '*** стенки ****
    '***************
                        Case Else ' стенки
stenki:
                            ' цвет бочков
                            k = InStr(1, Cells(FirstOrderRow, 1), ".")
                            If k Then
                                p = InStr(k + 1, Cells(FirstOrderRow, 1), ".")
                                If p Then
                                    CaseColor = Trim(Mid(Cells(FirstOrderRow, 1), k + 1, p - k - 1))
                                End If
                            End If
                            '*** цвет для виол
                    Dim sssss As String
    sssss = "Набор корп-ой мебели "
        iTmpKitch = InStr(1, UCase(Cells(FirstOrderRow, 1)), sssss, vbTextCompare)
        If iTmpKitch > 0 Then iTmpKitch = InStr(1, UCase(Cells(FirstOrderRow, 1)), "Виола", vbTextCompare)

        If iTmpKitch > 0 Then

                Dim ms As Integer
                Dim mc1 As String
                Dim mc As Integer
                Dim mc2 As String
                Dim mn As Integer
                Dim mc3 As String
                ms = 0
                mc = 0
                mn = 0
                mc1 = ""
                mc2 = ""
                mc3 = ""
                ms = InStr(iTmpKitch, UCase(Cells(FirstOrderRow, 1)), "массив", vbTextCompare)
                If ms > 0 Then
                    mc1 = "массив"
                    mc = InStr(iTmpKitch + Len("массив"), UCase(Cells(FirstOrderRow, 1)), "ясен", vbTextCompare)
                    If mc > 0 Then
                        mc2 = "ясень"
                        Else
                        mc = InStr(iTmpKitch + Len("массив"), UCase(Cells(FirstOrderRow, 1)), "ольх", vbTextCompare)
                        If mc > 0 Then mc2 = "ольха"
                    End If
                End If
                If mc > 0 Then
                mn = mc + 1
                Dim searchn As Integer
                searchn = 1
                Dim tempStr As String
                tempStr = UCase(Cells(FirstOrderRow, 1))
                While searchn = 1 And mn <= Len(tempStr)
                If InStr(1, "1234567890", Mid(tempStr, mn, 1)) > 0 Then searchn = 0 Else mn = mn + 1
                Wend

                While InStr(1, "1234567890", Mid(tempStr, mn, 1)) > 0 And mn <= Len(tempStr)
                    mc3 = mc3 & Mid(tempStr, mn, 1)
                    mn = mn + 1
                Wend

                End If
                If mn > 0 Then CaseColor = mc1 & " " & mc2 & " " & mc3

            End If

                            '**********************************

                            If IsNull(CaseColor) Then

                                'Dim colorid As Integer
                                If FormColor Is Nothing Then Set FormColor = New ColorForm
                                ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                                If ColorId = 0 Then
                                    FormColor.Show
                                    'colorid = FormColor.colorid
                                    CaseColor = Left(FormColor.ColorName, 20)
                                    ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                                    kitchenPropertyCurrent.dspColor = CaseColor
                                    kitchenPropertyCurrent.dspColorId = ColorId
                                    kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                                    
                                End If

'                                CaseColor = InputBox("Введите цвет бочков", "Цвет бочков", Cells(FirstOrderRow, 1))
'                                CaseColor = Left(CaseColor, 20)

                            End If
                            If CaseColor <> "" Then
                                UpdateOrder OrderId, , , , , , CaseColor
                                If ColorId > 0 Then UpdateOrder OrderId, , , , , , , ColorId
                            End If
                            '***** заглушки *******************
                            If IsEmpty(BibbColor) Then
                                BibbColor = GetBibbColor(CaseColor)
                            End If
                            If Not IsNull(BibbColor) Then UpdateOrder OrderId, , , BibbColor

                            If IsEmpty(CamBibbColor) Then
                                CamBibbColor = GetCamBibbColor(CaseColor)
                                kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                            End If
                            If Not IsNull(CamBibbColor) Then UpdateOrder OrderId, , , , , , , , CamBibbColor

                            ' тип/цвет фасадов
                            If IsNull(face) Then
                                k = InStr(1, Cells(FirstOrderRow, 1), ".")
                                If k Then
                                    k = InStr(k + 1, Cells(FirstOrderRow, 1), ".")
                                    If k Then
                                        p = InStr(k + 1, Cells(FirstOrderRow, 1), ".")
                                        If p = 0 Then p = Len(Cells(FirstOrderRow, 1))
                                        If p > k Then
                                            face = Trim(Mid(Cells(FirstOrderRow, 1), k + 1, p - k - 1))
                                            UpdateOrder OrderId, , , , , face


'                                            If InStr(1, face, "акрил", vbTextCompare) > 0 Then
'                                                FormFitting.AddFittingToOrder OrderID, "полироль", Empty, , , , , CasesPreampleRow
'                                            End If

                                            Cells(FirstOrderRow, 13).Value = face
                                        End If
                                    End If
                                End If

    '                            While Len(Face) < 6
    '                                Face = InputBox("Введите фасад стенки", "Фасад стенки", Face)
    '                            Wend
                            End If

                             k = InStr(1, casename, "карг")
                                                    If k Then
                                                        Cells(row, FCol).Characters(k, 5).Font.Color = vbRed

                                                        If FormFitting.AddFittingToOrder(OrderId, "карго", CaseQty, , , caseID, , row) Then
                                                            Doormount = Null
                                                            isKarg = True
                                                        End If
                            End If
                            ' проверим тип ручек по умолчанию
                            If IsEmpty(Handle) Then
                                k = InStr(1, Cells(FirstOrderRow, 1), "р.", vbTextCompare)
                                If k Then
                                    Handle = Trim(Mid(Cells(FirstOrderRow, 1), k + 2, 12))
                                End If
                            End If

                            '230311
'                            If IsEmpty(Leg) Then Leg = GetLegShelving(CaseColor)

                            Dim DefHandle
                            If IsNull(Handle) Then
                                DefHandle = Null
                            ElseIf IsEmpty(Handle) Then
                                DefHandle = Empty
                            Else
                                DefHandle = Handle
                            End If
                            k = InStr(1, casename, "вит", vbTextCompare)
                            If k > 0 Then
                                casename = Trim(Left(casename, k - 1))
                                windowcount = 1
                            End If

                            Dim tWindowCount, bStandart As Boolean

                            While Not ParseShelving(casename, caseID, DefHandle, Leg, Drawermount, Doormount, tWindowCount, bStandart, CaseColor, face) And casename <> ""
                                Dim fIB As fInputBox
                                Set fIB = New fInputBox

                                Dim commC As ADODB.Command
                                Set commC = New ADODB.Command
                                commC.ActiveConnection = GetConnection
                                commC.CommandType = adCmdText
                                commC.CommandText = "select * from [case] where contains(name,'""" & casename & """')"

                                Dim rs As ADODB.Recordset
                                Set rs = New ADODB.Recordset
                                rs.CursorLocation = adUseClient
                                rs.LockType = adLockReadOnly
                                rs.Open commC, , adOpenStatic, adLockReadOnly

                                fIB.cbList.Clear
                                If rs.RecordCount > 0 Then
                                    Dim Arr()
                                    ReDim Arr(rs.RecordCount - 1)

                                    Dim ii As Long
                                    rs.MoveFirst
                                    For ii = 0 To rs.RecordCount - 1
                                        Arr(ii) = rs!name
                                        rs.MoveNext
                                    Next ii

                                    fIB.cbList.List = Arr
                                Else
                                    fIB.cbList.AddItem casename
                                End If

                                fIB.cbList.MatchRequired = False
                                fIB.lblCaption = casename
                                fIB.Caption = "Введите наименование секции"
                                If fIB.cbList.ListCount = 1 Then fIB.cbList.ListIndex = 0

                                fIB.Show

                                If Not fIB.result Then GoTo skipshelving

                                casename = fIB.cbList.Value

                                'CaseName = InputBox("Введите наименование секции", "Ошибка определения секции", CaseName)
                                Rows(row).Interior.ColorIndex = 6
                            Wend

                            If IsEmpty(Leg) Or Leg = "бочка" Then Leg = GetLegShelving(CaseColor)


                            k = InStr(1, FCell, "шар")
                            If k Then
                                Cells(row, FCol).Characters(k, 3).Font.Color = vbRed
                                If InStr(1, Drawermount, "шарик ", vbTextCompare) <> 1 Then
                                If IsEmpty(Drawermount) Then
                                    Drawermount = "шарик 50"
                                End If
                                End If
                                ActiveCell.Offset(, 18).Value = Drawermount

                            End If
                           k = InStr(1, FCell, "рол")
                            If k Then
                                Cells(row, FCol).Characters(k, 3).Font.Color = vbRed
                                If InStr(1, Drawermount, "шарик ", vbTextCompare) <> 1 Then
                                    If IsEmpty(Drawermount) Then
                                        Drawermount = "ролик 50"
                                    Else
                                       Drawermount = "ролик " & Drawermount
                                    End If
                                End If
                                ActiveCell.Offset(, 18).Value = Drawermount
                            End If
                            'If IsNull(bPackShelvingWithFittingsKit) Then
                                If Left(casename, 1) = "С" Then
                                    Select Case casename
                                        Case "С1", "С2", "С3", "С4", "С5"
                                        Case Else ' если системы - даем ф-ру в пакетах
                                            bPackShelvingWithFittingsKit = True
                                    End Select
                                End If
                            'End If

                            If bStandart And IsNull(bPackShelvingWithFittingsKit) Then
                                If MsgBox("Дать на заказ комплект с фурнитурой?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Комплектация") = vbYes Then
                                    bPackShelvingWithFittingsKit = True
                                Else
                                    bPackShelvingWithFittingsKit = False
                                End If
                            End If

                            If Not IsNull(bPackShelvingWithFittingsKit) Then
                                bStandart = bStandart And bPackShelvingWithFittingsKit
                            End If

                            If Not IsEmpty(windowcount) Then
                                ' если указано "витр" (см. выше), даем витрины в любом случае
                                If windowcount = 1 Then tWindowCount = 1
                            End If

                            windowcount = tWindowCount

                            If IsEmpty(Handle) Then Handle = DefHandle

                            If Not bHandleCheck Then
                                If Not IsNull(Handle) Then
                                    CheckHandle Handle
                                End If
                                bHandleCheck = True
                            End If

                            If caseID > 0 Then

                                '********************************************************

                                If Not IsEmpty(windowcount) Then
                                    If Not IsNull(face) And Not IsEmpty(face) Then
                                        If InStr(1, face, "прим", vbTextCompare) > 0 Then
                                            windowcount = Empty

                                        ElseIf InStr(1, face, "софт", vbTextCompare) > 0 Or _
                                                InStr(1, face, "ДСП", vbTextCompare) > 0 Then

                                            If MsgBox("Давать стеклодержатели?", vbQuestion + vbYesNo + vbDefaultButton1, face) = vbNo Then
                                                windowcount = Empty
                                            End If
                                        End If
                                    End If

                                    If casename = "ШНВ А" Or casename = "ШНВ З" Then windowcount = Empty

                                End If

                                '********************************************************

                                If IsEmpty(HandleScrew) Then HandleScrew = GetHandleScrew(Handle, face)
                                If Not IsNull(HandleScrew) Then UpdateOrder OrderId, HandleScrew

                                If Left(casename, 2) = "ВЛ" Or Left(casename, 2) = "АЛ" Then

                                FormFitting.AddFittingToOrder OrderId, "ф-ра комплект ВИОЛА", Empty, casename, CaseColor, , , CasesPreampleRow



                                End If
                                HandleExtra = GetHandleExtra(Handle) '!

                                he = CheckHandleExtra(Handle)
                                If Not IsEmpty(he) Then HandleExtra = he

                                If Left(casename, 1) = "У" Then
                                    ' ножки черные в угол виктория

                                    OrderCaseID = AddCaseBySp(OrderId, caseID, casename, ActiveCell.Value, CaseQty, CaseHang, Handle, HandleExtra, Empty, 1, windowcount, Drawermount, Doormount, 1, caseglub, , Not bStandart, row, NoFace)
                                Else
                                    OrderCaseID = AddCaseBySp(OrderId, caseID, casename, ActiveCell.Value, CaseQty, CaseHang, Handle, HandleExtra, Leg, 1, windowcount, Drawermount, Doormount, 1, caseglub, , Not bStandart, row, NoFace)
                                End If

                                Selection.Interior.Color = RGB(173, 255, 47)
                            Else
skipshelving:
                                 ActiveCell.Interior.ColorIndex = 3
                            End If
                    End Select
                End If
            Else
                ActiveCell.Interior.ColorIndex = 3
            End If ' Qty > 0

        Next row
    End If
   
    If IsEmpty(Handle) Or IsNull(Handle) Then
        Cells(FirstOrderRow, 11).Value = "Без ручек"
    Else
        Cells(FirstOrderRow, 11).Value = "Ручки " & Handle
    End If

    If Not IsEmpty(Leg) Then
        If IsNull(Leg) Then
            Cells(FirstOrderRow, 12).Value = "Без ножек"
        Else
            Cells(FirstOrderRow, 12).Value = "Ножки " & Leg
        End If
    Else
        Cells(FirstOrderRow, 12).Value = "Ножки черн"
    End If
   


    ' теперь, разобравшись с заказом до конца, сохраним изменения
    If Not rsOrderFittings Is Nothing Then rsOrderFittings.UpdateBatch
'    If Not rsOrderCases Is Nothing Then rsOrderCases.UpdateBatch
    If Not rsOrderElements Is Nothing Then rsOrderElements.UpdateBatch
              
    AddOrderToShip = True
    Exit Function
err_AddOrderToShip:
    MsgBox Error, vbCritical
    Application.Cursor = xlDefault
    AddOrderToShip = False
End Function






Public Sub ProcessWHSheet()
On Error GoTo err_ProcessWHSheet
'    Init_rsOrderFittings False
'    Init_rsCases False
'    Init_rsOrderCases False
'    Init_rsOrderElements False
    
 '   Init_rsHandle False
 '   Init_rsLeg False
    
    
    Set FormFitting = New AddFitting
    Set FormElement = New AddElement

    Dim ShipID As Long
        'ShipID = 1


    'Dim TasksForm As MainForm
    Dim TaskID As Long
    'Set TasksForm = New MainForm
    'TasksForm.Show
    'ShipID = TasksForm.ShipID
    
    MainForm.Show
    ShipID = MainForm.ShipID
    
'    Set TasksForm = Nothing
    If ShipID = 0 Then Exit Sub
    On Error GoTo 0
      
    Dim EmptyLines As Long
    EmptyLines = 0
    
    Dim Customer As String
    Customer = InputBox("Введите имя клиента", "Клиент", Cells(1, 1).Value)
    
    If Trim(Customer) = "" Then Exit Sub
    
    Dim NewKitch As Boolean
    NewKitch = True
    
    Dim row As Integer
    Dim FirstOrderPreambleRow As Long, CaseColorRow As Long, FaceRow As Long

    For row = 1 To 1000
        If Not IsEmpty(Cells(row, 1)) Then
        
            If Cells(row, 1).Borders(xlEdgeTop).LineStyle > 0 And _
                    Cells(row, 1).Borders(xlEdgeBottom).LineStyle > 0 Then
                
                Select Case Cells(row, 1).Value
                    Case "Кухня"
                        If NewKitch Then
                            FirstOrderPreambleRow = row
                            NewKitch = False
                        End If
                    Case "бочок", "бочки", "каркас"
                        CaseColorRow = row
                    Case "фасад"
                        FaceRow = row
                End Select
            ElseIf Not NewKitch Then
                NewKitch = True
            End If
            
        ElseIf FirstOrderPreambleRow > 0 Then
                Dim r As Integer
                    
                r = FirstOrderPreambleRow
                
                While Left(Cells(r, 1).Value, 1) <> "Ш" And Left(Cells(r, 1).Value, 1) <> "П" And Not IsEmpty(Cells(r, 1).Value)
                    r = r + 1
                Wend
                
                ' бегаем по столбцам
                Dim col As Range
                For Each col In Range(Cells(FirstOrderPreambleRow, 2), Cells(FirstOrderPreambleRow, 50))
                    If Not (col.Borders(xlEdgeTop).LineStyle > 0 And col.Borders(xlEdgeBottom).LineStyle > 0) Then Exit For
                        
                     col.Activate
                     col.Select
                     
                     Dim OrderId As Long
                     
                     Dim CaseColor, face, isWindow As Boolean, Handle, Leg, CaseHang, SetQty, isExtraOrder As Boolean
                     Dim HandleScrew, HangColor, BibbColor, CamBibbColor
                     HandleScrew = Empty
                     HangColor = Empty
                     BibbColor = Empty
                     CamBibbColor = Empty
                     CaseHang = Empty
                     
                     
                     Dim fOrderParams As WHOrderParamsForm
                    
                     Do
                         Set fOrderParams = New WHOrderParamsForm
                         fOrderParams.Show 1
                         
                         If fOrderParams.result Then
                             Handle = fOrderParams.cbHandle.Text
                             If Handle = "клиента" Then Handle = Null
                             Leg = fOrderParams.cbLeg.Text
                             If Leg = "клиента" Then Leg = Null
                             
                             Dim OrderN As String
                             If fOrderParams.cbExtraOrder.Value Then
                                OrderN = "дозагруз"
                             Else
                                OrderN = ""
                             End If
                             
                             If IsNumeric(fOrderParams.tbSetQty.Text) Then SetQty = CInt(fOrderParams.tbSetQty.Text)
                             
                             OrderId = AddOrder(ShipID, FirstOrderPreambleRow, Customer, OrderN, SetQty)
                             
                             isExtraOrder = fOrderParams.cbExtraOrder.Value
                             If fOrderParams.cbCaseHang Then
                                 CaseHang = "завешка"
                                 
                                 ' определим цвет для завешек
                                 If IsEmpty(HangColor) Then
                                     HangColor = GetHangColor(CaseColor)
                                     
                                     UpdateOrder OrderId, , HangColor
                                 End If
                             Else
                                 CaseHang = "петля"
                             End If
                         Else
                            Exit Sub
                         End If ' If fOrderParams.Result
                     Loop Until fOrderParams.result
                     
                     If CaseColorRow > 0 Then
                         CaseColor = Trim(Cells(CaseColorRow, col.Column))
                        
                        Dim ColorId As Integer
                        If FormColor Is Nothing Then Set FormColor = New ColorForm
                        ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                        If ColorId = 0 Then
                            FormColor.Show
                            'colorid = FormColor.colorid
                            CaseColor = Left(FormColor.ColorName, 20)
                            ColorId = GetColorID(CaseColor, BibbColor, CamBibbColor)
                            kitchenPropertyCurrent.dspColor = CaseColor
                            kitchenPropertyCurrent.dspColorId = ColorId
                            kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                        End If
                         
                         UpdateOrder OrderId, , , , , , CaseColor
                         If ColorId > 0 Then UpdateOrder OrderId, , , , , , , ColorId
                     End If
                     
                     If IsEmpty(BibbColor) Then
                         BibbColor = GetBibbColor(CaseColor)
                     End If
                    If Not IsNull(BibbColor) Then UpdateOrder OrderId, , , BibbColor
                    
                     If IsEmpty(CamBibbColor) Then
                         CamBibbColor = GetCamBibbColor(CaseColor)
                         kitchenPropertyCurrent.CamBibbColor = CamBibbColor
                     End If
                         If Not IsNull(CamBibbColor) Then UpdateOrder OrderId, , , , , , , , CamBibbColor
                     
                     If FaceRow > 0 Then
                         face = Trim(Cells(FaceRow, col.Column))
                         UpdateOrder OrderId, , , , , face
        
'                        If InStr(1, face, "акрил", vbTextCompare) > 0 Then
'                            FormFitting.AddFittingToOrder OrderID, "полироль", Empty, , , , , FirstOrderPreambleRow
'                        End If
                        
                     End If ' If FaceRow > 0
                     
                     
                     If IsEmpty(HandleScrew) Then HandleScrew = GetHandleScrew(Handle, face)
                     If Not IsNull(HandleScrew) Then UpdateOrder OrderId, HandleScrew
                     

                     Dim c As Integer
                     c = r
                     While c < row ' Not IsEmpty(Cells(c, 1)) And
                         Dim caseID As Integer, DoorCount, windowcount, Drawermount, Doormount, NoFace As Boolean, HandleExtra, ShelfQty
                         Dim casename As String, CaseQty
                         
                         If Not IsEmpty(Cells(c, col.Column).Value) Then
                             Cells(c, col.Column).Activate
                             Cells(c, col.Column).Select
                             
                             CaseQty = Cells(c, col.Column).Value
                             If Not IsNumeric(CaseQty) Then
                                 Do
                                     CaseQty = InputBox("Введите кол-во шкафов", "кол-во шкафов", Cells(c, col.Column).Value)
                                 Loop Until IsNumeric(CaseQty)
                             End If
                             
                             casename = Cells(c, 1).Value
                             
                             caseID = 0
                             DoorCount = Empty
                             windowcount = Empty
                             Drawermount = Empty
                             Doormount = Empty
                             NoFace = Empty
                             HandleExtra = Empty
                             ShelfQty = Empty
                             Dim caseglub As Integer
                             caseglub = 0
                             Dim Width, NQty
                             Dim caseHeight As Integer
                             Do
                                ParseCase casename, caseID, DoorCount, windowcount, Drawermount, Doormount, NoFace, Handle, HandleExtra, ShelfQty, Width, NQty, CaseColor, caseglub, caseHeight
                                If caseID = 0 Then casename = InputBox("введите наименование шкафа", "Идентификация шкафа")
                             Loop Until casename = "" Or caseID > 0
                            'комплекты фурнитуры
                            If InStr(1, casename, "ВЛШВ2", vbTextCompare) > 0 Then FormFitting.AddFittingToOrder OrderId, "ф-ра комплект ВЛШВ2", CaseQty, CaseColor, , caseID, , row
         
                             Dim he
                             he = CheckHandleExtra(Handle)
                             If Not IsEmpty(he) Then HandleExtra = he
                             
                             
                             If caseID > 0 Then
                                 'Dim tHandle
                                 'If NoFace Then tHandle = Null Else tHandle = Handle
                                    
                                
                                 Dim k As Integer, VH
                                 If IsEmpty(ShelfQty) Then
                                    
                                    k = InStr(1, col.Value, "/")
                                    
                                    If k Then
                                        VH = Left(col.Value, k - 1)
                                        'NH = Mid(col.Value, k + 1)
'                                        If InStr(1, VH, "фр", vbTextCompare) Then
'                                            VH = Replace(VH, "фр", "", 1, 1, vbTextCompare)
'                                            Name = Name & " фр"
'                                        End If
                                    End If
                                    
                                    If IsNumeric(VH) Then
                                        If VH > 800 Then
                                            ShelfQty = 2
                                        ElseIf VH > 500 Then
                                            ShelfQty = 1
                                        Else
                                            ShelfQty = 0
                                        End If
                                    End If
                                End If
                                
                                If InStr(1, Cells(c, 1).Value, "карг") > 0 Then
                                    Doormount = Null
                                    FormFitting.AddFittingToOrder OrderId, "карго", CaseQty, , , caseID, , row
                                End If
                                
                                Select Case Left(casename, 1)
                                    Case "П"
                                        If CDec(DoorCount) + CDec(NQty) > 2 Then
                                            FormElement.AddElementToOrder OrderId, "полик", CaseQty * (CDec(DoorCount) + NQty - 2), caseID
                                        ElseIf (DoorCount) > 0 Then
                                            If Cells(c, 1).Font.Bold Then
                                                FormFitting.AddFittingToOrder OrderId, "завес", CaseQty, , , caseID, False, row
                                            Else ' если пенал стандартный
                                                FormFitting.AddFittingToOrder OrderId, "завес", CaseQty, , , caseID, True, row
                                            End If
                                        End If
                                End Select
                                
                                Select Case casename
                                '140311
'                                    Case "ШЛК/1", "ШСК/1"
'                                        If Width >= 80 And Not IsNull(Leg) Then
'                                            Dim tLeg
'                                            If IsEmpty(Leg) Then tLeg = "черная 100" Else tLeg = Leg
'                                            FormFitting.AddFittingToOrder OrderID, "ножка", CaseQty, tLeg, , CaseID, , Row
'                                        End If
                                
                                    Case "ШЛП", "ШСП", "ШНП", "ШНП915" ' если шкафы с гнутыми фасадами, добавим завесы
                                        If InStr(1, Trim(ActiveCell.Value), "ДГГ") > 0 Then
                                                                                    
                                            FormFitting.AddFittingToOrder OrderId, "завес", CaseQty * 2, "полусофт", , caseID, , row
                                            FormFitting.AddFittingToOrder OrderId, cHandle, CaseQty, Handle, , caseID, , row
                                        
                                        ElseIf InStr(1, Trim(ActiveCell.Value), "ДГВ") > 0 Then
                                            
                                            FormFitting.AddFittingToOrder OrderId, "стеклодержатель", CaseQty * 8, , , caseID, , row
                                            FormFitting.AddFittingToOrder OrderId, "завес", CaseQty * 2, "110", , caseID, , row
                                            FormFitting.AddFittingToOrder OrderId, cHandle, CaseQty, Handle, , caseID, , row
                                        End If
                                
                                    Case "ШНВ А", "ШНВ З"
                                        windowcount = Empty
                                 End Select
                                
                                '***************************************
                                
                                If Not IsEmpty(windowcount) Then
                                    If Not IsNull(face) And Not IsEmpty(face) Then
                                        If InStr(1, face, "прим", vbTextCompare) > 0 Then
                                            windowcount = Empty
                                            
                                        ElseIf InStr(1, face, "софт", vbTextCompare) > 0 Or _
                                                InStr(1, face, "ДСП", vbTextCompare) > 0 Then
                                                
                                            If MsgBox("Давать стеклодержатели?", vbQuestion + vbYesNo + vbDefaultButton1, face) = vbNo Then
                                                windowcount = Empty
                                            End If
                                        End If
                                    End If
                                    
                                    If casename = "ШНВ А" Or casename = "ШНВ З" Then windowcount = Empty
    
                                End If
                                
                                
'                                If InStr(1, CaseColor, "акция", vbTextCompare) > 0 Then
'                                    If IsEmpty(DoorMount) Or DoorMount = "110" Then
'                                        If MsgBox("Завесы акция?", vbDefaultButton1 Or vbYesNo Or vbQuestion, "АКЦИЯ") = vbYes Then
'                                            DoorMount = "110 акция"
'                                        End If
'                                    End If
'                                End If
                                

                                '*****************************************
                                OrderCaseID = AddCaseBySp(OrderId, caseID, casename, Cells(c, 1).Value, CaseQty, CaseHang, Handle, HandleExtra, Leg, DoorCount, windowcount, Drawermount, Doormount, 1, caseglub, ShelfQty, 1, , NoFace)
                               ' OrderCasesID = AddCaseBySp(OrderID, caseID, casename, Cells(c, 1).Value, CaseQty, CaseHang, Handle, HandleExtra, Leg, DoorCount, windowcount, Drawermount, Doormount, 1, caseglub, ShelfQty, Cells(c, 1).Font.Bold Or Cells(c, col.Column).Font.Bold, , NoFace)
                                 Cells(c, col.Column).Interior.Color = RGB(173, 255, 47)
                             Else
                                 Cells(c, col.Column).Interior.Color = vbRed
                             End If ' If CaseID >0
                         End If ' If Not IsEmpty(col.Value)
                         
                         c = c + 1
                     Wend ' While Not IsEmpty(Cells(c, col.Column))
                     
                    ' теперь, разобравшись с заказом до конца, сохраним изменения
                    If Not rsOrderFittings Is Nothing Then rsOrderFittings.UpdateBatch
                    'If Not rsOrderCases Is Nothing Then rsOrderCases.UpdateBatch
                    If Not rsOrderElements Is Nothing Then rsOrderElements.UpdateBatch
                    Init_rsOrderElements False
                Next col
            NewKitch = True
            FirstOrderPreambleRow = 0
        End If
    Next row
    
    Exit Sub
err_ProcessWHSheet:
    MsgBox Error, vbCritical, "обработка оптовых заказов"
End Sub



Public Sub ДобавитьФурнитуру()
   AddFitting.AddFitting
End Sub
Public Sub AutoReplace()
  On Error GoTo err_AutoReplace
'+    Init_rsOrderFittings False
'+    Init_rsCases False
'+    Init_rsOrderCases False
'+    Init_rsOrderElements False

'    Init_rsHandle False
'    Init_rsLeg False
    Application.Cursor = xlWait
     Init_rsOrderReplaces
        
    
    Set FormSearchReplace = New frmSearchReplace
   
    Dim L As Long
    Dim EmptyLines As Long
    EmptyLines = 0
       
       
    'Application.ScreenUpdating = False
    
    
    For L = 1 To 10000

        'If EmptyLines > 100 Then Exit Sub
        
        If Rows(L).Hidden = False Then
            If Not (Trim(Cells(L, 1)) = "" And Trim(Cells(L, 2)) = "" And Trim(Cells(L, 3)) = "" And Trim(Cells(L, 4)) = "" _
                     And Trim(Cells(L, 5)) = "" And Trim(Cells(L, 6)) = "" And Trim(Cells(L, 7)) = "" And Trim(Cells(L, 8)) = "" And Trim(Cells(L, 9)) = "") Then
                EmptyLines = 0
                
                    rsOrderReplacements.MoveFirst
                    While Not rsOrderReplacements.EOF
                    If rsOrderReplacements!isfullStringSearch = False And rsOrderReplacements!isRegExp = False Then
                        If InStr(1, Cells(L, 1), "!!") = 1 Or InStr(1, Cells(L, 1), " !!") = 1 Or InStr(1, Cells(L, 1), "ручк", vbTextCompare) = 1 _
                          Or InStr(1, Cells(L, 1), "Ш", vbTextCompare) = 1 _
                          Or InStr(1, Cells(L, 1), "П", vbTextCompare) = 1 _
                        Then
                            If InStr(1, Cells(L, 1), rsOrderReplacements!FindString, vbTextCompare) > 0 Then
                                If rsOrderReplacements!AskOnFind = 0 Then
                                    Cells(L, 1) = Replace(Cells(L, 1), rsOrderReplacements!FindString, rsOrderReplacements!ReplaceString, 1, , vbTextCompare)
                                    Cells(L, 1).Interior.Color = vbGreen
                                End If
                            End If
                        End If
                    ElseIf rsOrderReplacements!isfullStringSearch = True And rsOrderReplacements!isRegExp = False Then
                        If InStr(1, Cells(L, 1).Text, rsOrderReplacements!FindString, vbTextCompare) > 0 Then
                            Cells(L, 1) = rsOrderReplacements!ReplaceString
                            Cells(L, 1).Interior.Color = vbGreen
                        End If
                    End If
                        rsOrderReplacements.MoveNext
                    Wend
                End If
                
        Else
                EmptyLines = EmptyLines + 1
                If EmptyLines > 140 Then Exit For
        End If
    Next L
   
    'Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    MsgBox "Автозамены прогнал"
    Exit Sub
err_AutoReplace:
   'Application.ScreenUpdating = True
    MsgBox Error, vbCritical
    Application.Cursor = xlDefault
End Sub

Public Sub ДобавитьЭлементы()
   AddElement.AddElement
End Sub
Function xInArray(x As String, StringArray) As Boolean
    Dim i As Integer
    xInArray = False
    For i = 0 To UBound(StringArray)
        If LCase(x) = LCase(StringArray(i)) Then
            xInArray = True
            Exit For
        End If
    Next i
End Function
Public Function is18(ByVal CaseColor) As Boolean
    If InStr(1, CaseColor, "легно") > 0 Or _
         InStr(1, CaseColor, "лаванда") > 0 Or _
         InStr(1, CaseColor, "платина") > 0 Or _
         InStr(1, CaseColor, "черешня") > 0 Or _
         InStr(1, CaseColor, "золото") > 0 Or _
         InStr(1, CaseColor, "марино") > 0 Or _
         InStr(1, CaseColor, "магия") > 0 Or _
         InStr(1, CaseColor, "листвен") > 0 Or _
         InStr(1, CaseColor, "авиньон") > 0 Or _
         InStr(1, CaseColor, "шпон") > 0 Or _
         InStr(1, CaseColor, "массив") > 0 Or _
         InStr(1, CaseColor, "18") > 0 Then
         is18 = True
    Else
        is18 = False
    End If
    
End Function


