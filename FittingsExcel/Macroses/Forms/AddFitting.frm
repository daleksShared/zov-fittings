VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddFitting 
   Caption         =   "Добавить фурнитуру"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   OleObjectBlob   =   "AddFitting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddFitting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private TempMinus As Integer
Private TempMinusSingle As Single

Private result As Boolean
Private resultElement As Boolean

Private rsFittings As ADODB.Recordset
Private tstr, i

Public FittingOption, FittingLength
Private binit As Boolean

Private Krepl As String


Private Sub cb1_Click()
    tbQty.Text = 1
End Sub

Private Sub cb2_Click()
    tbQty.Text = 2
End Sub

Private Sub cb3_Click()
    tbQty.Text = 3
End Sub

Private Sub cb4_Click()
    tbQty.Text = 4
End Sub

Private Sub cbAdd_Click()
    result = False
    
    Dim bSpecified As Boolean
    bSpecified = True
    
    If cbFittingName.Enabled And cbFittingName.Text = "" Then bSpecified = False
    If bSpecified Then _
        If cbOpt.Enabled And cbOpt.Text = "" Then bSpecified = False
    If bSpecified Then _
        If cbLength.Enabled And cbLength.Text = "" Then bSpecified = False
    If bSpecified Then
        If Not IsNumeric(tbQty.Text) Then
            bSpecified = False
        Else
            If CDec(tbQty.Text) <= 0 Then bSpecified = False
        End If
    End If
    
    If bSpecified Then
        result = True
        Me.Hide
    Else
        result = False
        MsgBox "Фурнитура не добавлена." & vbCrLf & "Не все значения определены", vbExclamation, "Добавление фурнитуры"
    End If
End Sub




Private Sub cbCancel_Click()
    result = False 'True '!
    Me.Hide
End Sub




Private Sub UserForm_Activate()
    result = False
End Sub

Public Sub addFittArraysInit()
 ' цвета отбортовки
    GetOtbColors OtbColors
    GetOtbGorbColors OtbGorbColors
    
    ReDim vytyazhka_perfim(12)
    vytyazhka_perfim(0) = "IRIS угловая"
    vytyazhka_perfim(1) = "IRIS 60"
    vytyazhka_perfim(2) = "IRIS 90"
    vytyazhka_perfim(3) = "Egizia 60"
    vytyazhka_perfim(4) = "Egizia 90"
    vytyazhka_perfim(5) = "Colalto 60"
    vytyazhka_perfim(6) = "Colalto 90"
    vytyazhka_perfim(7) = "Tirolese 60"
    vytyazhka_perfim(8) = "Tirolese 90"
    vytyazhka_perfim(9) = "Isabella 90"
    vytyazhka_perfim(10) = "Sirius 99 SL(ан.Isabella)"
    vytyazhka_perfim(11) = "Sirius 903P-900 SL(ан.INN)"
    vytyazhka_perfim(12) = "Sirius 903-700 SL(ан.INN)"
    
    
    
    
    ReDim Направляющие(14) '(11)
    Направляющие(0) = "ролик 25"
    Направляющие(1) = "ролик 30"
    Направляющие(2) = "ролик 35"
    Направляющие(3) = "ролик 40"
    Направляющие(4) = "ролик 45"
    Направляющие(5) = "ролик 50"
    Направляющие(6) = "шарик 25"
    Направляющие(7) = "шарик 30"
    Направляющие(8) = "шарик 35"
    Направляющие(9) = "шарик 40"
    Направляющие(10) = "шарик 45"
    Направляющие(11) = "шарик 50"
    Направляющие(12) = "шарик с доводч 35"
    Направляющие(13) = "шарик с доводч 40"
    Направляющие(14) = "шарик с доводч 50"
    'Направляющие(11) = "quadro с доводч."
    
    ReDim Полкодержатель(8)
    Полкодержатель(0) = "3"
    Полкодержатель(1) = "5"
    Полкодержатель(2) = "Sekura 2-1"
    Полкодержатель(3) = "Sekura 8 (для стекла)"
    Полкодержатель(4) = "Пеликан ХРОМ большой"
    Полкодержатель(5) = "Пеликан ХРОМ макси"
    Полкодержатель(6) = "GS-3"
    Полкодержатель(7) = "тип C"
    Полкодержатель(8) = "PP-LUK-00-01"
    'глубины тандем боксов
    ReDim tbLength(4)
    tbLength(0) = "470"
    tbLength(1) = "420"
    tbLength(2) = "350"
    tbLength(3) = "260"
   
   'глубины коврика тандем боксов
    ReDim tbkovrLength(4)
    tbkovrLength(0) = "464"
    tbkovrLength(1) = "414"
    tbkovrLength(2) = "344"
    tbkovrLength(3) = "254"
    
    ReDim tbkovrOpt(39)
    tbkovrOpt(0) = "300\16\194"
    tbkovrOpt(1) = "350\16\244"
    tbkovrOpt(2) = "400\16\294"
    tbkovrOpt(3) = "450\16\344"
    tbkovrOpt(4) = "500\16\394"
    tbkovrOpt(5) = "550\16\444"
    tbkovrOpt(6) = "600\16\494"
    tbkovrOpt(7) = "650\16\544"
    tbkovrOpt(8) = "700\16\594"
    tbkovrOpt(9) = "750\16\644"
    tbkovrOpt(10) = "800\16\694"
    tbkovrOpt(11) = "850\16\744"
    tbkovrOpt(12) = "900\16\794"
    tbkovrOpt(13) = "950\16\844"
    tbkovrOpt(14) = "1000\16\894"
    tbkovrOpt(15) = "1050\16\944"
    tbkovrOpt(16) = "1100\16\994"
    tbkovrOpt(17) = "1150\16\1044"
    tbkovrOpt(18) = "1200\16\1094"
    tbkovrOpt(19) = "300\18\190"
    tbkovrOpt(20) = "350\18\240"
    tbkovrOpt(21) = "400\18\290"
    tbkovrOpt(22) = "450\18\340"
    tbkovrOpt(23) = "500\18\390"
    tbkovrOpt(24) = "550\18\440"
    tbkovrOpt(25) = "600\18\490"
    tbkovrOpt(26) = "650\18\540"
    tbkovrOpt(27) = "700\18\590"
    tbkovrOpt(28) = "750\18\640"
    tbkovrOpt(29) = "800\18\690"
    tbkovrOpt(30) = "850\18\740"
    tbkovrOpt(31) = "900\18\790"
    tbkovrOpt(32) = "950\18\840"
    tbkovrOpt(33) = "1000\18\890"
    tbkovrOpt(34) = "1050\18\940"
    tbkovrOpt(35) = "1100\18\990"
    tbkovrOpt(36) = "1150\18\1040"
    tbkovrOpt(37) = "1200\18\1090"
    tbkovrOpt(38) = "мп"
    tbkovrOpt(39) = "см"
   
    
    



    ' типы завесов
    ReDim Doormount(44)
    Doormount(0) = "110"
    Doormount(1) = "SlideOn 110"
    Doormount(2) = "+45"
    Doormount(3) = "175"
    Doormount(4) = "-45"
    Doormount(5) = "гармошка"
    Doormount(6) = "полусофт"
    Doormount(7) = "под амортизатор BLUM"
    Doormount(8) = "под амортизатор FGV"
    Doormount(9) = "равносторонний"
    Doormount(10) = "HK-S"
    Doormount(11) = "HF22"
    Doormount(12) = "HF25"
    Doormount(13) = "подъемник SK-105"
    Doormount(14) = "+30"
    Doormount(15) = "HETTICH софт"
    Doormount(16) = "+45 акция"
    Doormount(17) = "175 акция"
    Doormount(18) = "софт с обр. пружиной"
    Doormount(19) = "+20"
    Doormount(20) = "FGV софт"
    Doormount(21) = "Clip top 120 без пружинок"
    Doormount(22) = "Clip top средняя"
    Doormount(23) = "HF28"
    Doormount(24) = "HL23/35"
    Doormount(25) = "HL23/38"
    Doormount(26) = "HL25/35"
    Doormount(27) = "HL25/38"
    Doormount(28) = "HL27/35"
    Doormount(29) = "HL27/38"
    Doormount(30) = "HL23/39"
    Doormount(31) = "HL25/39"
    Doormount(32) = "HL27/39"
    Doormount(33) = "HL29/39"
    Doormount(34) = "HS A"
    Doormount(35) = "HS B"
    Doormount(36) = "HS D"
    Doormount(37) = "HS E"
    Doormount(38) = "HS G"
    Doormount(39) = "HS H"
    Doormount(40) = "HS I"
    Doormount(41) = "HK25"
    Doormount(42) = "HK27"
    Doormount(43) = "HK25 (Tip-On)"
    Doormount(44) = "HK27 (Tip-On)"

    

    ' типы лифтов
    ReDim Лифт(4)
    Лифт(0) = "50"
    Лифт(1) = "60"
    Лифт(2) = "80"
    Лифт(3) = "100"
    Лифт(4) = "120"
    
    ' каркасы
    ReDim StulNogi(2)
    StulNogi(0) = "КОНУС"
    StulNogi(1) = "ТРУБЫ 72"
    StulNogi(2) = "ТРУБЫ 82"
    
'    ReDim НогиСтол(10)
'    НогиСтол(0) = "Ахилл"
'    НогиСтол(1) = "TG(Зевс) с рогаткой"
'    НогиСтол(2) = "Аполло (с рамкой+2 палки)"
'    НогиСтол(3) = "Минос"
'    НогиСтол(4) = "Дедал"
'    НогиСтол(5) = "Калисто"
'    НогиСтол(6) = "Гермес"
'    НогиСтол(7) = "Максимус"
'    НогиСтол(8) = "Посейдон"
'    НогиСтол(9) = "Орфей"
'    НогиСтол(10) = "Альгео дуо"
    
    ReDim Стекло(14)
    Стекло(0) = "Ахилл"
    Стекло(1) = "TG(Зевс) большое"
    Стекло(2) = "TG(Зевс) маленькое"
    Стекло(3) = "Аполло"
    Стекло(4) = "Минос"
    Стекло(5) = "Минос маленькое"
    Стекло(6) = "Дедал"
    Стекло(7) = "Калисто"
    Стекло(8) = "Гермес"
    Стекло(9) = "Максимус"
    Стекло(10) = "Посейдон"
    Стекло(11) = "Орфей"
    Стекло(12) = "Клио"
    Стекло(13) = "Эрис"
    Стекло(14) = "Олимпиа"
    
'    ReDim Палки(0)
'    Палки(0) = "Аполло (5шт)"
    
    ' цвета планок
    ReDim Plank(5)
    Plank(0) = "ХРОМ"
    Plank(1) = "БЕЛЫЙ"
    Plank(2) = "БУК"
    Plank(3) = "ГРАНИТ"
    Plank(4) = "ЗЕЛЕНЫЙ"
    Plank(5) = "СИНИЙ"
    
    
    ' цвета галогенок
    ReDim Galog(1)
    Galog(0) = "ХРОМ"
    Galog(1) = "ЗОЛОТО"
    
    ' ширина сушек
    ReDim SW(4)
    SW(0) = "50"
    SW(1) = "60"
    SW(2) = "70"
    SW(3) = "80"
    SW(4) = "90"
    
    ReDim SW_bel(2)
    SW_bel(0) = "50"
    SW_bel(2) = "80"
    
    
    ' цвет сушек
    ReDim Sushk(3)
    Sushk(0) = "белая"
    Sushk(1) = "хром"
    Sushk(2) = "одноуровневая хром"
    Sushk(3) = "боярд"
 
    
    ' ширина лотков
    ReDim LW(11)
    LW(0) = "30"
    LW(1) = "40"
    LW(2) = "50"
    LW(3) = "60"
    LW(4) = "70"
    LW(5) = "80"
    LW(6) = "90"
    LW(7) = "ORGALINE 45"
    LW(8) = "ORGALINE 50"
    LW(9) = "ORGALINE 60"
    LW(10) = "ORGALINE 80"
    LW(11) = "ORGALINE 90"
    
    
    ' поддон алюминиевый
    ReDim PA(2)
    PA(0) = "50"
    PA(1) = "60"
    PA(2) = "80"
    
    ' цвета реллинга
    ReDim Rell(2)
    Rell(0) = "ХРОМ"
    Rell(1) = "ЗОЛОТО"
    Rell(2) = "БРОНЗА"
'    ' мойки
'    ReDim Sink(36)
'    Sink(0) = "PIX610"
'    Sink(1) = "BLX710"
'    Sink(2) = "PMN610"
'    Sink(3) = "PMN610 3,5"
'    Sink(4) = "PML610 3,5 декор"
'    Sink(5) = "S45 прав"
'    Sink(6) = "S45 лев"
'    Sink(7) = "SL45 3,5 декор прав"
'    Sink(8) = "SL45 3,5 декор лев"
'    Sink(9) = "NORM45 прав"
'    Sink(10) = "NORM45 лев"
'    Sink(11) = "NORM45 уни"
'    Sink(12) = "NORM45 декор уни"
'    Sink(13) = "NORM45 3,5 прав"
'    Sink(14) = "NORM45 3,5 лев"
'    Sink(15) = "NORM45 3,5 декор прав"
'    Sink(16) = "NORM45 3,5 декор лев"
'    Sink(17) = "BLN710-60 прав"
'    Sink(18) = "BLN710-60 лев"
'    Sink(19) = "BLL710-60 декор прав"
'    Sink(20) = "BLL710-60 декор лев"
'    Sink(21) = "BLN711 прав"
'    Sink(22) = "BLN711 лев"
'    Sink(23) = "BLL711 декор прав"
'    Sink(24) = "BLL711 декор лев"
'    Sink(25) = "COM лев"
'    Sink(26) = "COM прав"
'    Sink(27) = "COM 3,5 лев"
'    Sink(28) = "COM 3,5 прав"
'    Sink(29) = "COL 3,5 лев"
'    Sink(30) = "COL 3,5 прав"
'    Sink(31) = "FAM лев"
'    Sink(32) = "FAM прав"
'    Sink(33) = "FAM 3,5 лев"
'    Sink(34) = "FAM 3,5 прав"
'    Sink(35) = "FAL 3,5 лев"
'    Sink(36) = "FAL 3,5 прав"
    
'    ' столы
'    ReDim Stol(13)
'    Stol(0) = "АПОЛЛО махонь"
'    Stol(1) = "АХИЛЛ"
'    Stol(2) = "ГЕРМЕС"
'    Stol(3) = "ДЕДАЛ вишня"
'    Stol(4) = "TG вишня"
'    Stol(5) = "КАЛИСТО вишня"
'    Stol(6) = "КАЛИСТО махонь"
'    Stol(7) = "МИНОС вишня"
'    Stol(8) = "МАКСИМУС"
'    Stol(9) = "ОРФЕЙ"
'    Stol(10) = "ПОСЕЙДОН"
'    Stol(11) = "Клио"
'    Stol(12) = "Эрис"
'    Stol(13) = "Олимпиа"
'
'
'    ReDim Спинка(0)
'    Спинка(0) = "Колибер"
'
'    ReDim Крышка(0)
'    Крышка(0) = "Альгео Дуо МАРС"
    
'    ' стулья
'    ReDim Stul(14) ' порядок не менять!!!! используется обращение по индексу!!!
'    Stul(0) = "ВЕНУС"
'    Stul(1) = "ФОСКА"
'    Stul(2) = "ХАРПО"
'    Stul(3) = "ЧИКО"
'    Stul(4) = "НЕРОН"
'    Stul(5) = "КЛЕО махонь"
'    Stul(6) = "TC вишня"
'    Stul(7) = "TC махонь"
'    Stul(8) = "СФИНКС"
'    Stul(9) = "ЗЕВС"
'    Stul(10) = "ГЕКТОР"
'    Stul(11) = "КОЛИБЕР ХРОМ"
'    Stul(12) = "КОЛИБЕР АЛЮМ"
'    Stul(13) = "МАРКОС"
'    Stul(14) = "ЦЕЗАРЬ"
'
'    ' цвета сидений к кит. стульям - сфинкс, зевс, гектор, цезарь
'    ReDim SitK(4)
'    SitK(0) = "фисташковый"
'    SitK(1) = "оранжевый"
'    SitK(2) = "бежевый"
'    SitK(3) = "красный"
'    SitK(4) = "зеленый"
    
'    ' цвета сидений к КОЛИБЕР
'    ReDim SitKolib(2)
'    SitKolib(0) = "ольха"
'    SitKolib(1) = "св.серый"
'    SitKolib(2) = "бежевый"
'
'    ' цвета спинок к КОЛИБЕР
'    ReDim BackKolib(2)
'    BackKolib(0) = "махонь"
'    BackKolib(1) = "бук"
'    BackKolib(2) = "вишня"
    
'    ReDim Sit(3) ' порядок не менять!!!! используется обращение по индексу!!!
'    Sit(0) = "D390 (В,Ф)"
'    Sit(1) = "D340 (Ч,Х)"
'    Sit(2) = "Нерон"
'    Sit(3) = "Колибер"
    
'    ' сидушки цвета (к венусу, фоске, харпо, чико, нерон)
'    ReDim SitColors(9)
'    SitColors(0) = "ольха"
'    SitColors(1) = "св.-коричневая"
'    SitColors(2) = "зелёная"
'    SitColors(3) = "св.-зелёная"
'    SitColors(4) = "пепел"
'    SitColors(5) = "ваниль"
'    SitColors(6) = "темно-синяя"
'    SitColors(7) = "синий контраст"
'    SitColors(8) = "бежевая"
'    SitColors(9) = "жёлтая"
    
    ' цвета заглушек
    GetBibbColors Заглушки
    ' цвета заглушек
    GetCamBibbColors ЗаглЭксц

    ' цвета завешек
    GetHangColors Завешки
    
    ' Карго
    ReDim Карго(19)
    Карго(0) = "15лев"
    Карго(1) = "15прав"
    Карго(2) = "20лев"
    Карго(3) = "20прав"
    Карго(4) = "30"
    Карго(5) = "40"
    Карго(6) = "45"
    Карго(7) = "50"
    Карго(8) = "миди"
    Карго(9) = "миди VIBO"
    Карго(10) = "макси"
    Карго(11) = "макси VIBO"
    Карго(12) = "угол 45 лев"
    Карго(13) = "угол 45 прав"
    Карго(14) = "сушка 40"
    Карго(15) = "сушка 50"
    Карго(16) = "корзина 40"
    Карго(17) = "в угловой шкаф ХРОМ VIBO"
    Карго(18) = "распашн в нижн ХРОМ VIBO"
    Карго(19) = "распашн в верхн ХРОМ VIBO"
    
'    Карго(22) = "VS-Корзина под мойку 80"
'    Карго(23) = "VS-Корзина 60 внутр."
'    Карго(24) = "VS-Корзина 80 внутр."
'    Карго(25) = "VS-Карго расп. 45"
'    Карго(26) = "VS-Карго расп. 60"
'    Карго(27) = "VS-Верх. кар. 4/4"
'    Карго(28) = "VS-Верх. кар. 4/4 Пл-к"
'    Карго(29) = "VS-Карго 30"
'    Карго(30) = "VS-Раздел. в карго 30"
    
    ReDim MoikaColors(7)
    
   MoikaColors(0) = "беж"
    MoikaColors(1) = "графит"
    MoikaColors(2) = "звездный"
    MoikaColors(3) = "иней"
    MoikaColors(4) = "кофе"
    MoikaColors(5) = "серая крошка"
    MoikaColors(6) = "снег"
    MoikaColors(7) = "темный беж"


    
    ' Полка
    ReDim Полка(7)
    Полка(0) = "оборотная 1/2"
    Полка(1) = "оборотная 1/2 VIBO"
    Полка(2) = "оборотная 3/4"
    Полка(3) = "оборотная 3/4 VIBO"
    Полка(4) = "Р-08"
    Полка(5) = "Р-11"
    Полка(6) = "Р-12"
    Полка(7) = "Р-37"
    
    ReDim Вставка(96)
    Вставка(0) = "Алюминиевая полоса"
    Вставка(1) = "Антрацитовый перламутр"
    Вставка(2) = "Арктик"
    Вставка(3) = "Базальт"
    Вставка(4) = "Беж гранит глянец"
    Вставка(5) = "Беж гранит матовый"
    Вставка(6) = "Бежевый монохром"
    Вставка(7) = "Белая крошка"
    Вставка(8) = "Бук матовый"
    Вставка(9) = "Весенний перламутр"
    Вставка(10) = "Груша"
    Вставка(11) = "Делфи"
    Вставка(12) = "Дуб полосатый"
    Вставка(13) = "Желтый камень"
    Вставка(14) = "Зеленый глянец"
    Вставка(15) = "Зеленый матовый"
    Вставка(16) = "Золото"
    Вставка(17) = "Известняк"
    Вставка(18) = "Иней"
    Вставка(19) = "Камень"
    Вставка(20) = "Камушки"
    Вставка(21) = "Каппучино"
    Вставка(22) = "Кафель"
    Вставка(23) = "Корень глянец"
    Вставка(24) = "Красный монохром"
    Вставка(25) = "Лазурный глянец"
    Вставка(26) = "Лазурный матовый"
    Вставка(27) = "Латунь"
    Вставка(28) = "Лимонный глянец"
    Вставка(29) = "Лимонный матовый"
    Вставка(30) = "Лотос белый"
    Вставка(31) = "Лотос черный"
    Вставка(32) = "Лунный металл"
    Вставка(33) = "Малахитовая полоса"
    Вставка(34) = "Марс глянец"
    Вставка(35) = "Махонь"
    Вставка(36) = "Медный глянец"
    Вставка(37) = "Медный матовый"
    Вставка(38) = "Марокко камень"
    Вставка(39) = "Милано глянец"
    Вставка(40) = "Мрамор коричневый"
    Вставка(41) = "Мрамор черный"
    Вставка(42) = "Оникс"
    Вставка(43) = "Песчаник"
    Вставка(44) = "Платина"
    Вставка(45) = "Ровенна"
    Вставка(46) = "Рубиновая полоса"
    Вставка(47) = "Салатовый глянец"
    Вставка(48) = "Салатовый матовый"
    Вставка(49) = "Серая крошка"
    Вставка(50) = "Серебристый перламутр"
    Вставка(51) = "Сизый камень"
    Вставка(52) = "Синий глянец"
    Вставка(53) = "Синий матовый"
    Вставка(54) = "Терракот"
    Вставка(55) = "Темная крошка"
    Вставка(56) = "Туринский гранит"
    Вставка(57) = "Цитрусовый глянец"
    Вставка(58) = "Цитрусовый матовый"
    Вставка(59) = "Черная бронза"
    Вставка(60) = "Ясень темный"
    Вставка(61) = "Ясень светлый"
    Вставка(62) = "Яшма"
    Вставка(63) = "Красный иней"
    Вставка(64) = "Рыжий иней"
    Вставка(65) = "Серый иней"
    Вставка(66) = "Кварц"
    Вставка(67) = "Кремовый перламутр"
    Вставка(68) = "Вишня"
    Вставка(69) = "Магма"
    Вставка(70) = "Металлик"
    Вставка(71) = "Ракушечник"
    Вставка(72) = "Снежный"
    Вставка(73) = "Шерл"
    Вставка(74) = "Янтарь"
    Вставка(75) = "Коралл" 'кор
    Вставка(76) = "Мореный дуб" 'кор
    Вставка(77) = "Туя" 'кор
    Вставка(78) = "Песочный иней"
    Вставка(79) = "Морион"
    Вставка(80) = "СНОУ БЛЭК"
    Вставка(81) = "СНОУ УАЙТ"
    Вставка(82) = "СНОУ МИЛКИ"
    Вставка(83) = "Мрамор"
    Вставка(84) = "Накарадо"
    Вставка(85) = "Галициа"
    Вставка(86) = "Брешиа"
    
    Вставка(87) = "Аргиллит белый"
    Вставка(88) = "Хромикс серебро"
    Вставка(89) = "Хромикс антрацит"
    Вставка(90) = "Дуб Аризона серый"
    Вставка(91) = "Дуб Канзас коричневый"
    Вставка(92) = "Дуб Давос трюфель"
    Вставка(93) = "Сосна Касцина"
    Вставка(94) = "Керамика антрацит"
    Вставка(95) = "Мрамор Вальмасино светло-серый"
    Вставка(96) = "Мрамор Гиада голубой"
    
    
    SortArray Вставка
    
    
    ReDim Цоколь(22)
    Цоколь(0) = "ХРОМ100"
    Цоколь(1) = "ХРОМ150"
    Цоколь(2) = "БЕЛЫЙгл100"
    Цоколь(3) = "БЕЛЫЙгл150"
    Цоколь(4) = "БЕЛЫЙмат100"
    Цоколь(5) = "БЕЛЫЙмат150"
    Цоколь(6) = "БУК100"
    Цоколь(7) = "ВЕНГЕ100"
    Цоколь(8) = "ВЕНГЕ150"
    Цоколь(9) = "ГРУША100"
    Цоколь(10) = "КЛЕН100"
    Цоколь(11) = "КЛЕН150"
    Цоколь(12) = "ОЛЬХА100"
    Цоколь(13) = "ОРЕХ100"
    Цоколь(14) = "ОРЕХ150"
    Цоколь(15) = "МАХОНЬ100"
    Цоколь(16) = "КРЕМ100"
    Цоколь(17) = "КРЕМ150"
    Цоколь(18) = "ЧЁРНЫЙгл100"
    Цоколь(19) = "ЧЁРНЫЙгл150"
    Цоколь(20) = "РУСТИК100"
    Цоколь(21) = "ЯСЕНЬ100"
    Цоколь(22) = "СОСНАсветл100"
    
    
    
    
    
    ReDim СоединительЦоколя(19)
    СоединительЦоколя(0) = "ХРОМ100"
    СоединительЦоколя(1) = "ХРОМ150"
    СоединительЦоколя(2) = "БУК100"
    СоединительЦоколя(3) = "ВЕНГЕ100"
    СоединительЦоколя(4) = "ВЕНГЕ150"
    СоединительЦоколя(5) = "ГРУША100"
    СоединительЦоколя(6) = "КЛЕН100"
    СоединительЦоколя(7) = "КЛЕН150"
    СоединительЦоколя(8) = "ОЛЬХА100"
    СоединительЦоколя(9) = "МАХОНЬ100"
    СоединительЦоколя(10) = "КРЕМ100"
    СоединительЦоколя(11) = "КРЕМ150"
    СоединительЦоколя(12) = "БЕЛЫЙгл100"
    СоединительЦоколя(13) = "БЕЛЫЙгл150"
    СоединительЦоколя(14) = "ЧЁРНЫЙгл100"
    СоединительЦоколя(15) = "ЧЁРНЫЙгл150"
    СоединительЦоколя(16) = "ОРЕХ100"
    СоединительЦоколя(17) = "ОРЕХ150"
    СоединительЦоколя(18) = "ЯСЕНЬ100"
    СоединительЦоколя(19) = "СОСНАсветл100"

    
    ReDim Отбойники(2)
    Отбойники(0) = "ПВХ"
   ' Отбойники(1) = "GTV"
   ' Отбойники(2) = "FBV"
    
    ' не менять порядок
    ReDim ГорбатаяМелочь(9)
    ГорбатаяМелочь(0) = "Беж"
    ГорбатаяМелочь(1) = "С-Сер"
    ГорбатаяМелочь(2) = "Рыжий"
    ГорбатаяМелочь(3) = "Т-кор"
    ГорбатаяМелочь(4) = "Т-сер"
    ГорбатаяМелочь(5) = "Черн"
    ГорбатаяМелочь(6) = "Бел"
    ГорбатаяМелочь(7) = "Крем"
    ГорбатаяМелочь(8) = "Т-беж"
    ГорбатаяМелочь(9) = "Накар"
    
    
    
    ReDim МелочьОтб4м(6)
    МелочьОтб4м(0) = "Сер"
    МелочьОтб4м(1) = "Беж"
    МелочьОтб4м(2) = "Кор"
    МелочьОтб4м(3) = "Черн"
    МелочьОтб4м(4) = "Зел"
    МелочьОтб4м(5) = "Бел"
    МелочьОтб4м(6) = "Син"
    SortArray МелочьОтб4м
    
    
    ReDim TOPLine(2)
    TOPLine(0) = "Алюминий"
    TOPLine(1) = "Черная"
    TOPLine(2) = "Белая"
    
    ReDim zavesHL(29)
    zavesHL(0) = "500(18/342)"
    zavesHL(1) = "550(18/392)"
    zavesHL(2) = "600(18/442)"
    zavesHL(3) = "650(18/492)"
    zavesHL(4) = "700(18/542)"
    zavesHL(5) = "750(18/592)"
    zavesHL(6) = "800(18/642)"
    zavesHL(7) = "850(18/692)"
    zavesHL(8) = "900(18/742)"
    zavesHL(9) = "950(18/792)"
    zavesHL(10) = "1000(18/842)"
    zavesHL(11) = "1050(18/892)"
    zavesHL(12) = "1100(18/942)"
    zavesHL(13) = "1150(18/992)"
    zavesHL(14) = "1200(18/1042)"
    zavesHL(15) = "500(16/346)"
    zavesHL(16) = "550(16/396)"
    zavesHL(17) = "600(16/446)"
    zavesHL(18) = "650(16/496)"
    zavesHL(19) = "700(16/546)"
    zavesHL(20) = "750(16/596)"
    zavesHL(21) = "800(16/646)"
    zavesHL(22) = "850(16/696)"
    zavesHL(23) = "900(16/746)"
    zavesHL(24) = "950(16/796)"
    zavesHL(25) = "1000(16/846)"
    zavesHL(26) = "1050(16/896)"
    zavesHL(27) = "1100(16/946)"
    zavesHL(28) = "1150(16/996)"
    zavesHL(29) = "1200(16/1046)"
    
    ReDim zavesSensys(9)
    zavesSensys(0) = "110"
    zavesSensys(1) = "165"
    zavesSensys(2) = "+30"
    zavesSensys(3) = "+45"
    zavesSensys(4) = "Равностор-й"
    zavesSensys(5) = "110 узк AL проф"
    zavesSensys(6) = "софт"
    zavesSensys(7) = "полусофт"
    zavesSensys(8) = "гармошка"
    
    ReDim zavesClipTop(8)
    zavesClipTop(0) = "BLUMOTION +110"
    zavesClipTop(1) = "BLUMOTION Полусофт"
    zavesClipTop(2) = "BLUMOTION +45"
    zavesClipTop(3) = "BLUMOTION -45"
    zavesClipTop(4) = "+155"
    zavesClipTop(5) = "BLUMOTION равносторонний"
    zavesClipTop(6) = "BLUMOTION +90 под фп"
    zavesClipTop(7) = "110 без пружины"
    zavesClipTop(8) = "полусофт без пружины"
    
    
    
    ReDim ploschadkaSensys(6)
    ploschadkaSensys(0) = "D-0"
    ploschadkaSensys(1) = "D-0.5"
    ploschadkaSensys(2) = "D-1.5"
    ploschadkaSensys(3) = "W45 D-1.5"
    ploschadkaSensys(4) = "D-3"
    ploschadkaSensys(5) = "D-5"
    ploschadkaSensys(6) = "W45 D-3"

End Sub

Private Sub UserForm_Initialize()
    Dim comm As ADODB.Command
    Set comm = New ADODB.Command
    comm.ActiveConnection = GetConnection
    comm.CommandType = adCmdText
    comm.CommandText = "SELECT * FROM TopFittings ORDER BY Name"

    Set rsFittings = New ADODB.Recordset
    rsFittings.CursorLocation = adUseClient
    rsFittings.LockType = adLockBatchOptimistic
    rsFittings.Open comm, , adOpenDynamic, adLockBatchOptimistic
    
    Init_rsHandle
    Init_rsLeg
    'Init_rs
  
    ReDim FittingArray(0)
    ReDim HandleArray(0)
    ReDim LegArray(0)
    
    Dim i As Integer
    
    cbFittingName.Clear
    If rsFittings.RecordCount > 0 Then
        ReDim FittingArray(rsFittings.RecordCount - 1)
        For i = 0 To rsFittings.RecordCount - 1
            FittingArray(i) = rsFittings!name
            rsFittings.MoveNext
        Next i
        
        cbFittingName.List = FittingArray
    End If

    If rsHandle.RecordCount > 0 Then
        ReDim HandleArray(rsHandle.RecordCount - 1)
        rsHandle.MoveFirst
        For i = 0 To rsHandle.RecordCount - 1
            HandleArray(i) = rsHandle!Handle
            rsHandle.MoveNext
        Next i
    End If

    If rsLeg.RecordCount > 0 Then
        ReDim LegArray(rsLeg.RecordCount - 1)
        rsLeg.MoveFirst
        For i = 0 To rsLeg.RecordCount - 1
            LegArray(i) = rsLeg!Leg
            rsLeg.MoveNext
        Next i
    End If
    
   addFittArraysInit
    
    
    get_st_par

End Sub

Public Sub get_st_par()
ReDim Stul_color_no(16)
Stul_color_no(1) = "101"
Stul_color_no(2) = "102"
Stul_color_no(3) = "103"
Stul_color_no(4) = "104"
Stul_color_no(5) = "105"
Stul_color_no(6) = "106"
Stul_color_no(7) = "107"
Stul_color_no(8) = "108"
Stul_color_no(9) = "109"
Stul_color_no(10) = "110"
Stul_color_no(11) = "111"
Stul_color_no(12) = "112"
Stul_color_no(13) = "113"
Stul_color_no(14) = "114"
Stul_color_no(15) = "115"
Stul_color_no(16) = "116"

ReDim Stul_color_1(4)
Stul_color_1(1) = " пат."
Stul_color_1(2) = " п.зол."
Stul_color_1(3) = " п.сер."
Stul_color_1(4) = " "

ReDim Stul_color_2(9)
Stul_color_2(1) = " (ж.бордо п-са)"
Stul_color_2(2) = " (ж.бордо)"
Stul_color_2(3) = " (ж.синий)"
Stul_color_2(4) = " (ж.ромб крас)"
Stul_color_2(5) = " (ж.зел. п-са)"
Stul_color_2(6) = " (ж.желтый)"
Stul_color_2(7) = " (чапин 1000)"
Stul_color_2(8) = " (Оригон 114)"
Stul_color_2(9) = " (ж.зелён.)"



End Sub




Public Function AddFittingToOrder(ByVal OrderId As Long, _
                                    ByVal name As String, _
                                    ByVal qty, _
                                    Optional ByRef Opt = Empty, _
                                    Optional ByRef length = Empty, _
                                    Optional ByVal caseID, _
                                    Optional ByVal Standart As Boolean = False, _
                                    Optional ByVal RowN As Integer = 0) As Boolean
                                    
    AddFittingToOrder = False
    Me.Caption = "Добавить фурнитуру"
    If Not kitchenPropertyCurrent Is Nothing Then
        If kitchenPropertyCurrent.dspColor <> "" Then
            Me.Caption = Me.Caption & " б:" & kitchenPropertyCurrent.dspColor
        End If
    End If
    'затычка
    If name = cHandle Then name = "ручка"
    
    binit = True
    cbOpt.Text = ""
    cbFittingName.Text = ""
    binit = False
    
    Dim bSpecified As Boolean
    bSpecified = False
    
    If IsEmpty(qty) Then
        If name = cNogi Then
            qty = 4
            tbQty.Text = qty
        Else
            'Qty = 1
            tbQty.Text = ""
            bSpecified = False
        End If
    'End If
    
    ElseIf IsNumeric(qty) Then
        If qty = 0 Then
            AddFittingToOrder = True
            Exit Function
        End If
        tbQty.Text = qty
        'tbQty.Enabled = False
    Else
        'tbQty.Enabled = True
        tbQty.Text = ""
        bSpecified = False
    End If
                                    
                                    
    cbAddNext.Value = False
    If Not IsMissing(Opt) Then
    FittingOption = Opt
    Else
    FittingOption = Empty
    End If
    If Not IsMissing(length) Then
    FittingLength = length
    Else
    FittingLength = Empty
    End If
    
    
    
    'Debug.Print ("Fitting=" & Name & RTrim(" " & Opt) & ", QTY=" & Qty)
    Dim i As Integer

    cbFittingName.Enabled = True
    
    cbFittingName.Text = ""
    If name <> "" Then
        For i = 0 To cbFittingName.ListCount - 1
            If InStr(1, cbFittingName.List(i), name) = 1 Then
                cbFittingName.Text = cbFittingName.List(i)
                cbFittingName.Enabled = False
                bSpecified = True
                Exit For
            End If
        Next i
    End If
    
    If cbFittingName.Text = "" Then cbFittingName.Text = name
    
    If Not bSpecified Then
        MsgBox "ОШИБКА!!! НЕИЗВЕСТНАЯ ФУРНИТУРА", vbCritical
        'Exit Function
    End If
              
    If IsMissing(caseID) Then
        FormRutin OrderId, RowN, Opt, length, , Standart
    Else
        FormRutin OrderId, RowN, Opt, length, caseID, Standart
    End If

    AddFittingToOrder = result
End Function


Private Sub FormRutin(ByVal OrderId As Long, _
                       Optional ByVal RowN As Integer, _
                       Optional ByRef Opt, _
                       Optional ByRef length, _
                       Optional ByVal caseID, _
                       Optional ByVal Standart As Boolean = False)
    
    Dim bSpecified As Boolean
    bSpecified = True
    
    Do
        If bSpecified Then _
            If cbFittingName.Enabled And cbFittingName.Text = "" Then bSpecified = False
        If bSpecified Then _
            If cbOpt.Enabled And cbOpt.Text = "" Then bSpecified = False
        If bSpecified Then _
            If cbLength.Enabled And cbLength.Text = "" Then bSpecified = False
        If bSpecified Then
            If Not IsNumeric(tbQty.Text) Then
                bSpecified = False
            Else
                If CDec(tbQty.Text) <= 0 Then bSpecified = False
            End If
        End If
        
        Select Case cbFittingName.Text
            Case cNogi, "Отбортовка 3м", "Отбортовка 4м", "Отбортовка горб-4", "Отбортовка горб-5", "Отбортовка TOP-Line", "вставка в отбортовку", "площадка для завеса", "завешка"
                bSpecified = False
            Case "ф-ра комплект ВИОЛА"
            cbOpt.Text = Opt
            cbLength.Text = length
'            Case "завес", "завес Sensys"
'                If IsMissing(caseID) Then
'                cbAddNext.Value = True
'                bSpecified = False
'                End If
            Case "клиновая планка Sensys"
            cbOpt.Text = Opt
            bSpecified = True
        End Select
        
        'если все элементы определены, то форму показывать не будем
        If bSpecified Then
            result = True
        Else
            
            If cbAddNext.Value Then
                Select Case cbFittingName.Text
                    Case cNogi, "Отбортовка 3м", "Отбортовка 4м", "Отбортовка горб-4", "Отбортовка горб-5", "Отбортовка TOP-Line", "дюбель DU325 Rapid S", "лиц панель в АНТ вн" ', "завес", "завес Sensys"
                    Case Else
                        cbAddNext.Value = False
                End Select
                
                Me.Show 1
                
                Select Case cbFittingName.Text
                     Case "вставка в отбортовку"
                        If result And cbOpt.Enabled Then Opt = cbOpt.Text: length = cbLength.Text
                End Select
            ElseIf cbAddNextElement.Value Then
                cbAddNextElement.Value = False
                
                resultElement = FormElement.AddElementToOrder(OrderId, "", "")
            Else
                Me.Show 1
                'If IsEmpty(Opt) Then
                    If result And cbOpt.Enabled Then Opt = cbOpt.Text Else Opt = ""
                    If result And cbLength.Enabled Then length = cbLength.Text Else length = ""
                'End If
            End If
        End If
    
        
        If result Then
            If IsMissing(caseID) Then
                
                result = AddFitting2Order(OrderId, RowN, , Standart)
            Else
                result = AddFitting2Order(OrderId, RowN, caseID, Standart)
            End If
        Else
            result = True 'была нажата отмена
        End If
            
         If cbAddNext.Value Then
         
             Select Case cbFittingName.Text
                Case "дюбель DU325 Rapid S"
                    cbFittingName.Text = "эксцентрик 18мм"
                    
                Case "отбортовка Волпато 4м Ал"
                    cbFittingName.Text = "угл+загл к отб Волпато"
                    
'                Case "завешка"
'                    If cbOpt.Text = "CAMAR Л+П" Then
'                        cbFittingName.Text = "накладка CAMAR  Л+П"
'
'                    End If
               
'               Case "завес Sensys"
'                    cbFittingName.Enabled = True
'                    cbFittingName.Text = "площадка для завеса"
               Case "амморт CLIP TOP +155"
                cbAddNext.Value = False
               Case "Завес CLIP top"
                    If cbOpt.Text = "+155" Then
                        cbFittingName.Text = "амморт CLIP TOP +155"
                        If tbQty.Text = "2" Then tbQty.Text = "1"
                        
                    Else
                        tbQty.Text = ""
                        cbFittingName.Enabled = True
                        cbFittingName.Text = ""
                    End If
                    
               'Case "амморт. Sensys 165"
               '     If cbOpt.Text = "165" Then
               '        cbFittingName.Text = "площадка зав.Sensys"
               'End If
                   
'                Case cStul, cStool, "спинка" ' после стула добавим сидушки
'
'                    If cbOpt.Text = Stul(11) Or cbOpt.Text = Stul(12) Then     ' "КОЛИБЕР ХРОМ"  "КОЛИБЕР АЛЮМ"
'                        cbFittingName.Text = "спинка"
'                    Else
'                        cbFittingName.Text = cSit
'                    End If
                    
                    
'                    For i = 0 To cbOpt.ListCount - 1
'                        If InStr(1, cbFittingName.List(i), cSit, vbTextCompare) Then
'                            cbFittingName.Text = cbFittingName.List(i)
'                            Exit For
'                        End If
'                    Next

'            Case "мойка"
'                Krepl = cbOpt.Text
'                cbFittingName.Text = "крепление к мойке"
'                bSpecified = True
'
'            Case "крепление к мойке"
'                cbFittingName.Text = "сифон к мойке"
                
            Case "Отбортовка 4м"
            
                cbFittingName.Text = "вставка в отбортовку"
                Dim tqty As Single
                tqty = CDec(tbQty.Text) * 4 / 3
                If tqty > Round(tqty) Then
                    tbQty.Text = Round(tqty) + 1
                Else
                    tbQty.Text = Round(tqty)
                End If
            
            Case "завес"
                Select Case cbOpt.Text
                    Case "под амортизатор BLUM"
                        cbFittingName.Text = "амортизатор BLUM"
                    Case "под амортизатор FGV"
                        cbFittingName.Text = "амортизатор FGV"
                    Case Else
                        'tbQty.Text = ""
                        cbFittingName.Enabled = True
                        cbFittingName.Text = ""
                End Select
                
             Case "завес под амортиз-р BLUM"
                cbFittingName.Text = "амортизатор BLUM"
              cbLength.Enabled = False
                cbOpt.Enabled = False
            Case "дюбель VB15"
                cbFittingName.Text = "соед планка 200мм"
              cbLength.Enabled = False
                cbOpt.Enabled = False
             Case "завес под амортизатор FGV"
                cbFittingName.Text = "амортизатор FGV"
                 cbLength.Enabled = False
             Case "завес HL23/35", "завес HL23/38", "завес HL25/35", "завес HL25/38", "завес HL27/35", "завес HL27/38", _
             "завес HL23/39", "завес HL25/39", "завес HL27/39", "завес HL29/39"
                cbLength.Enabled = False
                cbOpt.Enabled = False
                cbFittingName.Text = "штанга HL овальная"
              Case "трансформатор", "трансф-р LED 30w"
              cbLength.Enabled = False
              cbOpt.Enabled = False
              cbFittingName.Text = "кабель+вилка+выкл 220V"
              Case "завес HS I", "завес HS A", "завес HS B", "завес HS D", "завес HS E", "завес HS G", "завес HS H", "завес HS F"
                cbLength.Enabled = False
                cbOpt.Enabled = False
                cbFittingName.Text = "штанга HS круглая"
                
'            Case "Реллинг 60", "Реллинг 100"
'                cbFittingName.Text = "переходник к реллингу"
'                cbOpt.Enabled = False
'                If CDec(tbQty.Text) > 1 Then tbQty.Text = CDec(tbQty.Text) - 1
'            Case "Тб Ант500/C бел внут", "Тб Ант500/M бел внутр", "Тб Ант500/D бел внутр"
'                cbFittingName.Text = "лиц панель в АНТ вн"
'                cbLength.Enabled = True
'                If Not IsMissing(Opt) Then cbOpt.Text = Opt
            Case "Тб Ант бел внут"
                
                cbFittingName.Text = "лиц панель в АНТ вн"
                cbLength.Enabled = True
                If Not IsMissing(Opt) Then If Not IsEmpty(Opt) Then cbOpt.Text = Opt
            Case "лиц панель в АНТ вн"
                cbFittingName.Text = "попер реллинг в АНТ вн"
                If Not IsMissing(Opt) Then If Not IsEmpty(Opt) Then cbOpt.Text = Opt
                If Not IsMissing(length) Then If Not IsEmpty(length) Then cbLength.Text = length
                
             Case "ТБ Архитех внутр"
                
                cbFittingName.Text = "лиц панель Архитех внутр"
                'cbLength.Enabled = True
                If Not IsMissing(Opt) Then cbOpt.Text = Opt
            Case "лиц панель Архитех внутр"
                cbFittingName.Text = "попер рел Архитех внутр"
                tbQty.Text = ""
            
            Case Else
                 tbQty.Text = ""
                 'tbQty.Enabled = True
                 
                 cbFittingName.Enabled = True
                 cbFittingName.Text = ""
             End Select
             
         ElseIf cbAddNextElement.Value Then
                cbAddNextElement.Value = False
                Me.Hide
                If (FormElement Is Nothing) Then Set FormElement = New AddElement
                resultElement = FormElement.AddElementToOrder(OrderId, "", "")
                
                
            Else
             Me.Hide
         End If
        
    Loop While cbAddNext.Value
End Sub



Private Function AddFitting2Order(ByVal OrderId As Long, _
                                    ByVal RowN As Integer, _
                                    Optional ByVal caseID, _
                                    Optional ByVal Standart As Boolean = False) As Boolean
                                    
    'On Error GoTo err_AddFitting2Order
    
    AddFitting2Order = False
    Application.Cursor = xlWait
    
    Dim Opt, FittingID As Integer
    If rsFittings.RecordCount > 0 Then rsFittings.MoveFirst
    rsFittings.Find "Name='" & cbFittingName.Text & "'"
    If Not rsFittings.EOF Then
        FittingID = rsFittings!FittingID
    Else
        AddFitting2Order = False
        MsgBox "Неизвестный тип фурнитуры '" & cbFittingName.Text & "'", vbCritical
        Exit Function
    End If
    
    If Len(cbOpt.Text) > 0 Then
        Opt = Trim(cbOpt.Text)
    End If
    If Len(cbLength.Text) > 0 Then
        If InStr(1, Opt, "ВЛ", vbTextCompare) = 1 And InStr(1, cbLength.Text, "массив", vbTextCompare) = 1 Then
            Opt = LTrim(Opt & " " & Replace(Replace(cbLength.Text, "ассив", "."), " ", ""))
            Else
            Opt = LTrim(Opt & " " & cbLength.Text)
        End If
    End If
    
    Init_rsOrderFittings
    
    rsOrderFittings.AddNew
    
    rsOrderFittings!OrderId = OrderId
    rsOrderFittings!FittingID = FittingID
    rsOrderFittings!qty = CDec(tbQty.Text)
    
    If Not IsMissing(caseID) Then
        rsOrderFittings!caseID = caseID
        rsOrderFittings!Standart = Standart
    End If
    If Not IsMissing(OrderCaseID) Then
        If OrderCaseID > 0 Then
        rsOrderFittings!ocid = OrderCaseID
        End If
    End If
    If RowN > 0 Then
        rsOrderFittings!row = RowN
    End If
    
    
    If Not IsEmpty(Opt) Then rsOrderFittings!Option = Opt
    
    If Cells(ActiveCell.row, 10).Value <> "" Then
        Dim t As String
        t = Cells(ActiveCell.row, 10).Value
        
        Cells(ActiveCell.row, 10).Value = t & "; " & "Ф=" & cbFittingName.Text & RTrim(" " & cbOpt.Text) & ", QTY=" & tbQty.Text & ", L=" & cbLength.Text
    Else
        Cells(ActiveCell.row, 10).Value = "Ф=" & cbFittingName.Text & RTrim(" " & cbOpt.Text) & ", QTY=" & tbQty.Text & ", L=" & cbLength.Text
    End If
    
    
    AddFitting2Order = True
    Application.Cursor = xlDefault
    Exit Function
err_AddFitting2Order:
    MsgBox "Ошибка добавления фурнитуры", vbCritical
    AddFitting2Order = False
    Application.Cursor = xlDefault
End Function



Private Sub cbFittingName_Change()
    If binit Then Exit Sub
    
    cbOpt.Enabled = True
    cbLength.Enabled = True
    
    binit = True
    cbOpt.Text = ""
    cbLength.Text = ""
    cbOpt.Clear
    cbLength.Clear
    binit = False
    
    'cbOpt.Clear
    'cbLength.Clear
    
    Dim i As Integer
 
    Select Case cbFittingName.Text
     
         

        Case "Вытяжка"
            cbOpt.Text = "PERFIM"
            cbLength.List = vytyazhka_perfim
        Case ""
        Case "ручка", "Завес", _
             "ножка", _
             "галогенки 3", "галогенки 5", _
             "Планка в угол", "Планка м/у столешницами", _
             "планка к газ. плите", _
             "Ноги", _
             "Реллинг 60", "Реллинг 100", "Заглушка к реллингу", "Держатель к реллингу", "Крючок к реллингу", "Угол-90 к реллингу", "Угол-120 к реллингу", "Угол-135 к реллингу", "переходник к реллингу", _
             "лифт", _
             "завешка", _
             "заглушка", _
             "карго", _
             "полка", _
             "крестик", "кромка с клеем", "кромка без клея", _
             "цоколь пластик", "заглушка к цоколю", "угол90* к цоколю", "угол135* к цоколю", _
             "шуруп к pучке", "винт к pучке", "стекло", "палки", "каркас", _
             "полкодержатель", "соединитель цоколя", "планка для завеш 807"
             
            
               ' cbOpt.Enabled = False
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            
            Dim bSkip As Boolean
            bSkip = False
            
            ' цвет
            cbOpt.Text = ""
            Select Case cbFittingName.Text
                Case "полкодержатель"
               
                    cbOpt.List = Полкодержатель
            
'                Case "каркас"
'                    cbOpt.List = Stul
            
                Case "ручка"
                    cbOpt.List = HandleArray
                
                Case "ножка"
                 
                    cbOpt.List = LegArray
                    
                
                Case "шуруп к pучке", "винт к pучке"
                    cbOpt.AddItem "22"
                    cbOpt.AddItem "25"
                    cbOpt.AddItem "28"
                    cbOpt.AddItem "35"
                    cbOpt.AddItem "40"
                    
                Case "крестик"
                    cbOpt.AddItem "золотой"
                    bSkip = True
                
                Case "цоколь пластик", _
                     "заглушка к цоколю", _
                     "угол90* к цоколю", _
                     "угол135* к цоколю", _
                     "соединитель цоколя"
                     cbOpt.List = Цоколь
                     'cbAddNext.Value = True
'                Case "соединитель цоколя"
'                     cbOpt.List = Цоколь
'                     cbLength.Enabled = False
                     
                Case "планка в угол", _
                        "планка монтажная", _
                        "планка м/у столешницами", _
                        "планка к газ. плите"
                       
'                        Select Case FittingLength
'                            Case "28"
'                                cbOpt.List = Plank
'
'                                If Not IsEmpty(FittingOption) And Len(FittingOption) >= 2 Then
'                                    For i = 0 To cbOpt.ListCount - 1
'                                        If FittingOption = cbOpt.List(i) Then
'                                            cbOpt.Text = cbOpt.List(i)
'                                            Exit For
'                                        End If
'                                    Next
'                                End If
'
'                                cbLength.Text = "28"
                          If FittingLength = "38" Then
                        
                                cbOpt.Text = "ХРОМ"
                                FittingOption = "ХРОМ"
                                cbLength.Text = "38"
                         
                         Else
                         cbOpt.Enabled = True
                                cbOpt.List = Plank
                                
                                cbLength.Text = ""
                                cbLength.Enabled = True
                                cbLength.AddItem "28"
                                cbLength.AddItem "38"
                        
                         End If
                         
                    
                Case "галогенки 3", "галогенки 5"
                    cbOpt.List = Galog
                ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False

                
                Case "завес"
                ' длина не нужна
            cbLength.Enabled = False
                    cbOpt.List = Doormount
                                    
                    If FittingOption <> "110" Then bSkip = True
                    If Not IsEmpty(FittingOption) Then If FittingOption = "175" Then cbOpt.Text = FittingOption
'                    If Not IsEmpty(FittingOption) And Len(FittingOption) >= 2 Then
'                        For i = 0 To cbOpt.ListCount - 1
'                            If FittingOption = cbOpt.List(i) Then
'                                cbOpt.Text = cbOpt.List(i)
'                                Exit For
'                            End If
'                        Next
'                    End If
                    
                
                Case "кромка с клеем", "кромка без клея"
                
                cbOpt.Enabled = False
                Case "лифт"
                    cbOpt.List = Лифт
'                Case "мойка"
'                    cbOpt.List = Sink
'                    bSkip = True
'                    'cbAddNext.Value = True
                    
                Case "карго"
                    cbOpt.List = Карго
                    bSkip = True
                    
                Case "полка"
                    cbOpt.List = Полка
                    bSkip = True
                                
'                Case cStool, cStul
'                    cbOpt.List = Stul
'                    cbAddNext.Value = True
'                    'cbAddNext.Enabled = False
'                    bSkip = True
                    
'                Case "стол китайский"
'                    cbOpt.List = Stol
'                    bSkip = True
                
                Case "стекло"
                    cbOpt.List = Стекло
                    bSkip = True
'                Case "палки"
'                    cbOpt.List = Палки
'                    cbOpt.Text = Палки(0)
'                    bSkip = True
'                Case "крышка"
'                    cbOpt.List = Крышка
'                    cbOpt.Text = Крышка(0)
'                    bSkip = True
'                Case "ноги к столам"
'                    cbOpt.List = НогиСтол
'                    bSkip = True
                    
                Case "ноги"
                ' длина не нужна
            cbLength.Enabled = False
                    cbOpt.List = StulNogi
                    bSkip = True
                    If IsNumeric(tbQty.Text) Then
                        If CInt(tbQty.Text) Mod 4 <> 0 Then
                            MsgBox "Проверьте правильность кол-ва ног!", vbExclamation
                        End If
                    End If
                    
                    
                Case "Реллинг 60", "Реллинг 100", _
                    "Заглушка к реллингу", _
                    "Держатель к реллингу", _
                    "Крючок к реллингу", _
                    "Угол-90 к реллингу", _
                    "Угол-120 к реллингу", _
                    "Угол-135 к реллингу"
                   ' длина не нужна
            cbLength.Enabled = False
                    
                    cbOpt.List = Rell
                
'                Case cSit
'                    cbOpt.List = Sit
'                    bSkip = True
                
                Case "завешка"
                    cbOpt.List = Завешки
                ' длина не нужна
            cbLength.Enabled = False
                Case "заглушка"
                    cbOpt.List = Заглушки
                ' длина не нужна
            cbLength.Enabled = False
                Case Else
                    cbOpt.Clear
            ' длина не нужна
            cbLength.Enabled = False
            End Select
            
            If Not bSkip Then
                If Not IsEmpty(FittingOption) And Len(FittingOption) >= 3 Then
                    For i = 0 To cbOpt.ListCount - 1
                        If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                            InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                            cbOpt.Text = cbOpt.List(i)
                            Exit For
                        End If
                    Next
                End If
            End If
        
        
        Case "клипса для подсветки 2К"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
         Case "крепление к цоколю", "крепл. цоколя универсал", "уголок под запресовку", "ответная планка под шуруп"
          cbOpt.Enabled = False
            cbLength.Enabled = False
            cbAddNext.Value = False
            
        Case "стяжка колпачковая"
        cbOpt.AddItem "короткая"
        cbOpt.AddItem "средняя"
        cbOpt.AddItem "длинная"
        If Not IsEmpty(FittingOption) Then
            If FittingOption = "средняя" Then cbOpt.Text = FittingOption
        End If
        cbLength.Enabled = False
        Case "подъемник SK-105"
        cbOpt.Enabled = True
        cbLength.Enabled = False
        Case "держатель к штанге"
        cbOpt.AddItem "БЕЛЫЙ"
        cbOpt.AddItem "ХРОМ"
        cbOpt.AddItem "крепл к стене d-25"
        cbOpt.Enabled = True
        cbLength.Enabled = False
        
        Case "крючок большой"
            cbOpt.AddItem "БЕЛЫЙ"
            cbOpt.AddItem "ХРОМ"
            cbOpt.Text = "ХРОМ"
            cbLength.Enabled = False
        
        Case "полоска к цоколю"
        cbOpt.AddItem "16"
        cbOpt.AddItem "18"
        cbLength.Enabled = False
        Case "планка монтажная 100мм", "планка монтажная", _
             "метабокс ШЛГП", _
             "бародержатель", "бародержатель HT", "бародержатель мет.", _
             "гвозди", "еврошуруп", "завешка", "заглушка", "замок", "каретка для верхн. напр.", "каретка для нижн. напр.", _
             "конфирмант", "ключ к конфирманту", "крестик", "к-т к стеклу", "магнит", "магнит коричневый", _
             "направляющая верхняя", "направляющая нижняя", _
             "петля", "петля мебельная(СШ5)", "пластинка мет. для магн.", _
             "подпятник", "полкодержатель 5", "стеклодержатель", "стяжка Ж6", "стяжка колпачковая", "уголок мет. пенал", "уголок мет.", "часы", _
             "шуруп 3*30", "шуруп 3,5*16", "шуруп 4*16", "шуруп 5*30", "шкант", "клипса", "ноги трубы 82", _
             "крепление к стол тр. бар.", "крепление к трубе бар.", "крепление к кровати", _
             "завес полусофт", "завес софт", "полкодержаетль Secura 8", _
             "уголок золотой", "крестик зол. плоский ", "кретик зол. выпуклый", "полоска Gold", "полоска White/Gold", "полоска L хром", "полоска L золото", _
             "винт кор. к стеклу", "крепление к сушке", "С-профиль металлик(см)", "крепление к мб 110", "крепление к мб 010", _
              "держатель к штанге", "шайба пластиковая", "винт к ручке верона", "винт к ручке модена", "крепл к реллингу мб задн", "переходник к реллингу", _
             "стяжка для столешницы", "амортизатор BLUM", "амортизатор FGV", "амортизатор врезной", "шуруп 4*20", "муфта №41", _
              "крючок малый", "крючок 08", "ведро мусорное", "угловой адаптер 10гр", _
             "доводчик на метабокс", "стяжка для пустотки", "удлинитель крепл. мойки", "вставка в пустотку 3м", _
             "дюбель DU860 двойной шарн", "дюбель DU868 двойной 16мм", "опора колесная", "фиксатор стенки ДВП RV8", "дюбель DU321", "стяжка VB 36 HT", _
             "накладка CAMAR Л+П", "завешка CAMAR Л+П", "стяжка VB 35D/16", _
             "саморез", "м-зм Push-To-Open Magnet", "подвеска шкафа SAH-5 п/з", "полкодержатель тип C", "завешка CAMAR 806лев.", "завешка CAMAR 806прав.", _
             "завешка 807", "завешка 808", "тандембокс под мойку", "емкость в тб под мойку", _
             "т/б оргабокс XXL", "бародержатель АКВИЛА", "Клипса для подсветки 2К", "крепление фас. тандембокс", "крепление к дов. мб L+R", "втулка SISO П+М", _
             "ключ золотой для замка", "сервис-пакет", "крепл к релл мб перед"
             '/*"завес Ecomat", "завес FGV180", _*/
          cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "Подпятник регулируемый"
            cbOpt.Text = "малый M8 L-30"
            cbOpt.AddItem "малый M8 L-30"
            cbLength.Enabled = False
            cbLength.Text = ""
        Case "метабокс малый", "метабокс большой"
            cbOpt.Enabled = False
            cbLength.Text = ""
            cbLength.Enabled = True
            cbLength.AddItem "500"
            cbLength.AddItem "450"
            cbLength.AddItem "оптима500"
             If Not IsEmpty(FittingLength) Then
                If Not IsEmpty(FittingLength) Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next i
                End If
            End If
        
        Case "подст верт проф 81G19A10", "педаль для выдв ящика"
          cbOpt.Enabled = False
            cbLength.Enabled = False
        Case "ServoDrive"
            cbOpt.AddItem "Блок питания"
            cbOpt.AddItem "Кабель(мп)"
            cbOpt.AddItem "на 1 шуфл"
            cbOpt.AddItem "на 2 шуфл"
            cbOpt.AddItem "на 3 шуфл"
            cbOpt.AddItem "на 4 шуфл"
            cbOpt.AddItem "UNO"
            cbOpt.Text = ""
            cbLength.Enabled = False
        
        Case "ServoDrive БП"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "ServoDrive кабель"
            cbOpt.Enabled = False
            cbLength.Enabled = True
            
        Case "Komandor"
            cbOpt.AddItem "планка д/крепл рельс"
            cbOpt.AddItem "проф AGAT верт ALU"
            cbOpt.AddItem "проф аллюм верх+низ"
            cbOpt.AddItem "проф аллюм гор верх"
            cbOpt.AddItem "проф аллюм гор нижн"
            cbOpt.AddItem "ролик верхн"
            cbOpt.AddItem "ролик нижн"
            cbOpt.AddItem "стопор самокл белый"
            cbOpt.AddItem "уплотнит H-4 прозр"
            cbOpt.AddItem "щетка длин буферн"
            cbOpt.AddItem "щетка коротк буферн"

            cbOpt.Enabled = True
            cbLength.Enabled = True
        Case "Hafele SD"
            cbOpt.AddItem "Ходовая шина каретки"
            cbOpt.AddItem "Ходовой ролик"
            cbOpt.AddItem "Концевая загл без паза"
            cbOpt.AddItem "Двойная ход шина нижн"
            cbOpt.AddItem "Двойная напр шина верх"

            cbOpt.Enabled = True
            cbLength.Enabled = True
        Case "Astin", "фур", "ArciTech"
            cbOpt.Enabled = True
            cbLength.Enabled = True
            
        Case "завес HK-S (TIP-ON)", "завес HK27 (TIP-ON)", "завес HK25", "завес HK25 (TIP-ON)", "завес HK29 (TIP-ON)", "завес HK29", "завес HK27", "завес HK-S", "завес HF22", "завес HF25", "завес HF28"
          'cbOpt.Enabled = False
           cbOpt.AddItem "ХРОМ"
           cbOpt.AddItem "БЕЛЫЙ"
           cbOpt.AddItem "БЕЛЫЙ серво драйв"
           cbOpt.Text = "ХРОМ"
            cbLength.Enabled = False
             
            
            Case "к-т крепл под ф/п"
                cbOpt.Enabled = False
                cbLength.Enabled = False

            Case "завес HK-XS"
                cbOpt.Enabled = False
                cbLength.Enabled = False
            
          Case "завес HL23/35", "завес HL23/38", "завес HL25/35", "завес HL25/38", "завес HL27/35", "завес HL27/38", _
          "завес HL23/39", "завес HL25/39", "завес HL27/39", "завес HL29/39", _
           "завес HS I", "завес HS A", "завес HS B", "завес HS D", "завес HS E", "завес HS G", "завес HS H", "завес HS F"
            'cbOpt.Enabled = False
           cbOpt.AddItem "ХРОМ"
           cbOpt.AddItem "БЕЛЫЙ"
           cbOpt.AddItem "БЕЛЫЙ серво драйв"
           'cbOpt.AddItem "ХРОМ серво драйв"

           cbOpt.Text = "ХРОМ"
            cbLength.Enabled = False
            cbAddNext.Value = True
        Case "поддон ПЛХ+тр.углы"
            cbLength.Enabled = False
            cbOpt.Enabled = False
            cbAddNext.Value = False
        Case "транспорт углы ПЛХ"
            cbLength.Enabled = False
            cbOpt.Enabled = False
            cbAddNext.Value = False
            
'         Case "Тб Ант500/C под мойку", "Тб Ант500/C бел", "Тб Ант500/D бел", "Тб Ант500/M бел", "Тб Ант500/N бел", "Тб Ант500/D бел под мойку"
'          cbOpt.Enabled = False'            cbLength.Enabled = False
        Case "VS - VSA"
            cbOpt.AddItem "5 - 600"
            cbOpt.Text = "5 - 600"
            cbLength.Enabled = False
            
        Case "VS - Экспозитор прав+лев", "VS - Экспозитор прав", "VS - Экспозитор лев"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "VS - Дуг держ лев+прав", "VS - Дуг держ лев", "VS - Дуг держ прав"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "VS - HSA"
            cbOpt.AddItem "2 - 300"
            cbOpt.AddItem "2 - 450"
            cbOpt.AddItem "2 - 600"
            cbOpt.AddItem "3 - 300"
            cbOpt.AddItem "3 - 450"
            cbOpt.AddItem "3 - 600"
            cbOpt.AddItem "5 - 300"
            cbOpt.AddItem "5 - 450"
            cbLength.Enabled = False
            
        Case "VS - DSA"
            cbOpt.AddItem "3 - 150"
            cbOpt.AddItem "3 - 150 с пол.держ."
            cbOpt.AddItem "3 - 200"
            cbOpt.AddItem "3 - 300"
            cbOpt.AddItem "3 - 400"
            cbOpt.AddItem "8 - 150"
            cbOpt.AddItem "8 - 200"
            cbLength.Enabled = False
            
        Case "VS - Верхняя карусель"
            cbOpt.AddItem "600-3/4"
            cbOpt.AddItem "4/4"
            
            cbLength.Enabled = False
        
        Case "VS - Нижняя карусель"
            cbOpt.AddItem "900-3/4"
            cbOpt.AddItem "90*90-4/4"
            cbLength.Enabled = False
            
        Case "VS - Wari Corner"
            cbOpt.AddItem "900 - L"
            cbOpt.AddItem "900 - R"
            cbOpt.AddItem "1000 - L"
            cbOpt.AddItem "1000 - R"
            cbLength.Enabled = False

         Case "VS - Карго распашная"
            cbOpt.AddItem "DUSA 1 - 450"
            cbOpt.AddItem "DUSA 1 - 600"
            cbOpt.AddItem "DUSA 3 - 450"
            cbOpt.AddItem "DUSA 3 - 600"
            cbOpt.AddItem "DUSA 5 - 450"
            cbOpt.AddItem "DUSA 5 - 600"
            cbOpt.AddItem "DUSA 6 - 450"
            cbOpt.AddItem "DUSA 6 - 600"
            cbLength.Enabled = False
            
         Case "VS - Twin Corner"
            cbOpt.AddItem "450 - L"
            cbOpt.AddItem "450 - R"
            cbOpt.AddItem "500 - L"
            cbOpt.AddItem "500 - R"
            cbOpt.AddItem "600 - L"
            cbOpt.AddItem "600 - R"
            cbLength.Enabled = False
        
         Case "VS - Eco Center"
            cbOpt.AddItem "2"
            cbOpt.AddItem "3"
            cbOpt.AddItem "4"
            cbLength.Enabled = False
            
        Case "VS - Система Eco Liner"
         cbOpt.AddItem "450"
            cbOpt.AddItem "600"
            cbOpt.Text = "600"
            cbLength.Enabled = False
        Case "VS - термо планка"
            cbOpt.AddItem "16мм"
            cbOpt.Text = "16мм"
            cbLength.Enabled = False
            
            
        Case "VS - Eco flex liner"
            cbOpt.AddItem "СИСО 450 2в"
            cbOpt.AddItem "СИСО 600 3в"
            cbOpt.AddItem "СИСО 600 2в"
            
            cbLength.Enabled = False
        
        Case "VS - CornerStone"
            cbOpt.AddItem "450 -R"
            cbOpt.AddItem "450 -L"
            cbOpt.AddItem "500 -R"
            cbOpt.AddItem "500 -L"
            cbOpt.AddItem "600 -R"
            cbOpt.AddItem "600 -L"
            cbOpt.AddItem "1000 -R"
            cbOpt.AddItem "1000 -L"
            
            cbLength.Enabled = False
            
        Case "VS - OSA"
            cbOpt.AddItem "1 - 450"
            cbOpt.AddItem "1 - 600"
            cbLength.Enabled = False
            
        Case "VS - Base Liner"
            cbOpt.AddItem "450"
            cbOpt.AddItem "500"
            cbOpt.AddItem "600"
            cbLength.Enabled = False
            
        Case "VS - Выдвижная корзина"
            cbOpt.AddItem "450 для фр.кр. без ф.кр."
            cbOpt.AddItem "600 для фр.кр. без ф.кр."
            cbOpt.AddItem "900 для фр.кр. без ф.кр."
            cbOpt.AddItem "450 на расп дв с планкой"
            cbOpt.AddItem "600 на расп дв с планкой"
            cbOpt.AddItem "900 на расп дв с планкой"
            cbLength.Enabled = False
            If Not IsEmpty(FittingOption) Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next i
            End If
        Case "VS - Фронт крепл вдв крз"
            cbLength.Enabled = False
            cbOpt.Enabled = False
            
        Case "VS - Корзина под мойку"
            cbOpt.AddItem "800"
            cbOpt.AddItem "900"
            cbLength.Enabled = False
            If Not IsEmpty(FittingOption) Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next i
            End If
        Case "VS - Сетчатая полка"
            cbOpt.AddItem "600"
            cbOpt.AddItem "900"
            cbLength.Enabled = False
                
        Case "Тб Ант бел"
           cbOpt.Enabled = False
           cbLength.AddItem "500/N"
           cbLength.AddItem "500/M"
           cbLength.AddItem "500/C"
           cbLength.AddItem "500/D"
           cbLength.AddItem "300/N"
           cbLength.AddItem "300/M"
           cbLength.AddItem "300/C"
           cbLength.AddItem "300/D"
           If Not IsEmpty(FittingLength) Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next i
            End If
        Case "Тб Ант бел под мойку"
           cbOpt.Enabled = False
           cbLength.AddItem "500/C"
           cbLength.AddItem "500/D"
           cbLength.AddItem "300/C"
           cbLength.AddItem "300/D"
            If Not IsEmpty(FittingLength) Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next i
            End If
         Case "Тб Ант бел внут"
            cbLength.AddItem "500/M"
            cbLength.AddItem "500/C"
            cbLength.AddItem "500/D"
            cbLength.AddItem "300/M"
            cbLength.AddItem "300/C"
            cbLength.AddItem "300/D"
            cbOpt.AddItem "16"
            cbOpt.AddItem "18"
            cbAddNext.Value = True
            If Not IsEmpty(FittingLength) Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next i
            End If
        If Not IsEmpty(FittingOption) Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next i
            End If
        Case "лиц панель в АНТ вн"
            cbOpt.AddItem "16"
            cbOpt.AddItem "18"
            cbLength.Text = ""
            cbAddNext.Value = True
           
           
         Case "ТБ Архитех внутр"
            cbOpt.AddItem "Белый"
            cbOpt.AddItem "Антрацит"
            If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
               For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
             
            cbLength.AddItem "500/94 мал"
            cbLength.AddItem "500/186 1релл"
            cbLength.AddItem "500/186 стекло"
            cbLength.AddItem "300/94 мал"
            cbLength.AddItem "300/186 1релл"
           
              If Not IsEmpty(FittingLength) And Len(FittingLength) >= 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
           
            cbAddNext.Value = True
            
            
             Case "ТБ Архитех"
             cbOpt.AddItem "Белый"
             cbOpt.AddItem "Антрацит"
             If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
             
            cbLength.AddItem "500/78 ШЛГП"
            cbLength.AddItem "500/94 мал"
            cbLength.AddItem "500/186 1релл"
            cbLength.AddItem "500/186 стекло"
            cbLength.AddItem "500/250 2релл"
            cbLength.AddItem "300/94 мал"
            cbLength.AddItem "300/186 1релл"
            cbLength.AddItem "300/250 2релл"
            
            If Not IsEmpty(FittingLength) And Len(FittingLength) >= 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
            
            
             Case "лоток в Архитех"
             cbOpt.AddItem "Белый"
             cbOpt.AddItem "Антрацит"
              If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
            
             cbLength.AddItem "60см"
             cbLength.AddItem "90см"
            If Not IsEmpty(FittingLength) And Len(FittingLength) >= 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
                  
             Case "лоток ORGALINE"
            cbOpt.Enabled = False
            cbLength.AddItem "40"
            cbLength.AddItem "45"
            cbLength.AddItem "50"
            cbLength.AddItem "60"
            cbLength.AddItem "80"
            cbLength.AddItem "90"
           
                  
            Case "лиц панель Архитех внутр"
'                cbOpt.AddItem "16"
'                cbOpt.AddItem "18"
                cbOpt.AddItem "Белый"
                cbOpt.AddItem "Антрацит"
                  If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
                    For i = 0 To cbOpt.ListCount - 1
                        If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                            InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                            cbOpt.Text = cbOpt.List(i)
                            Exit For
                        End If
                    Next
                End If
                
                
                cbLength.AddItem "16 30(179)"
                cbLength.AddItem "16 35(229)"
                cbLength.AddItem "16 40(279)"
                cbLength.AddItem "16 45(329)"
                cbLength.AddItem "16 50(379)"
                cbLength.AddItem "16 55(429)"
                cbLength.AddItem "16 60(479)"
                cbLength.AddItem "16 65(529)"
                cbLength.AddItem "16 70(579)"
                cbLength.AddItem "16 75(629)"
                cbLength.AddItem "16 80(679)"
                cbLength.AddItem "16 85(729)"
                cbLength.AddItem "16 90(779)"
                cbLength.AddItem "18 30(175)"
                cbLength.AddItem "18 35(225)"
                cbLength.AddItem "18 40(275)"
                cbLength.AddItem "18 45(325)"
                cbLength.AddItem "18 50(375)"
                cbLength.AddItem "18 55(425)"
                cbLength.AddItem "18 60(475)"
                cbLength.AddItem "18 65(525)"
                cbLength.AddItem "18 70(575)"
                cbLength.AddItem "18 75(625)"
                cbLength.AddItem "18 80(675)"
                cbLength.AddItem "18 85(725)"
                cbLength.AddItem "18 90(775)"
                 If Not kitchenPropertyCurrent Is Nothing Then
                    If kitchenPropertyCurrent.dspWidth > 0 Then
                        If Not casepropertyCurrent Is Nothing Then
                            If casepropertyCurrent.p_cabWidth > 0 Then
                               
                                If kitchenPropertyCurrent.dspWidth < 18 Then TempMinus = 121 Else TempMinus = 125
                                cbLength.Text = CStr(kitchenPropertyCurrent.dspWidth) & " " & Mid(CStr(casepropertyCurrent.p_cabWidth), 1, 2) & "(" & CStr(casepropertyCurrent.p_cabWidth - TempMinus) & ")"
                            End If
                        End If
                 End If
                End If
                cbAddNext.Value = True
             Case "попер рел Архитех внутр"
                 cbOpt.AddItem "Белый"
                cbOpt.AddItem "Антрацит"
                  If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
                    For i = 0 To cbOpt.ListCount - 1
                        If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                            InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                            cbOpt.Text = cbOpt.List(i)
                            Exit For
                        End If
                    Next
                End If
                
                cbLength.AddItem "16 30(174,5)"
                cbLength.AddItem "16 35(224,5)"
                cbLength.AddItem "16 40(274,5)"
                cbLength.AddItem "16 45(324,5)"
                cbLength.AddItem "16 50(374,5)"
                cbLength.AddItem "16 55(424,5)"
                cbLength.AddItem "16 60(474,5)"
                cbLength.AddItem "16 65(524,5)"
                cbLength.AddItem "16 70(574,5)"
                cbLength.AddItem "16 75(624,5)"
                cbLength.AddItem "16 80(674,5)"
                cbLength.AddItem "16 85(724,5)"
                cbLength.AddItem "16 90(774,5)"
                cbLength.AddItem "18 30(170,5)"
                cbLength.AddItem "18 35(220,5)"
                cbLength.AddItem "18 40(270,5)"
                cbLength.AddItem "18 45(320,5)"
                cbLength.AddItem "18 50(370,5)"
                cbLength.AddItem "18 55(420,5)"
                cbLength.AddItem "18 60(470,5)"
                cbLength.AddItem "18 65(520,5)"
                cbLength.AddItem "18 70(570,5)"
                cbLength.AddItem "18 75(620,5)"
                cbLength.AddItem "18 80(670,5)"
                cbLength.AddItem "18 85(720,5)"
                cbLength.AddItem "18 90(770,5)"
                If Not kitchenPropertyCurrent Is Nothing Then
                    If kitchenPropertyCurrent.dspWidth > 0 Then
                        If Not casepropertyCurrent Is Nothing Then
                            If casepropertyCurrent.p_cabWidth > 0 Then
                                If kitchenPropertyCurrent.dspWidth < 18 Then TempMinusSingle = 125.5 Else TempMinusSingle = 129.5
                                cbLength.Text = CStr(kitchenPropertyCurrent.dspWidth) & " " & Mid(CStr(casepropertyCurrent.p_cabWidth), 1, 2) & "(" & CStr(CSng(casepropertyCurrent.p_cabWidth) - TempMinusSingle) & ")"
                            End If
                        End If
                 End If
                End If
                
        Case "реллинг на тандембокс"
            cbOpt.AddItem "470"
            cbOpt.AddItem "420"
            cbOpt.AddItem "350"
            cbOpt.AddItem "260"
            cbLength.Enabled = False
        Case "реллинг на метабокс"
            cbOpt.AddItem "500"
            cbOpt.AddItem "450"
            cbOpt.AddItem "оптима500"
            cbLength.Enabled = False
          Case "тандембокс ШЛГП", "тандембокс малый", "тандембокс большой"
            cbOpt.Enabled = False
          
            cbLength.List = tbLength
            cbLength.Text = "470"
            If Not IsEmpty(FittingLength) Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next i
            End If
           
        
          Case "тандембокс БЛЮМ"
            cbOpt.AddItem "малый"
            cbOpt.AddItem "большой"
            cbLength.Text = "50"
          
          Case "ручка UKW-7"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
         Case "соед. штанги BLUM"
            cbOpt.Enabled = True
            cbOpt.AddItem "HS"
            cbOpt.AddItem "HL"
            cbOpt.Text = ""
            
            cbLength.Enabled = False
         
            
 
          Case "штанга HL овальная", "штанга HS круглая"


            cbOpt.Enabled = True
            cbOpt.Text = ""
            cbOpt.List = zavesHL
            cbLength.Enabled = False
           ' cbAddNext.Value = False
            
            

        Case "болт мет. 100мм", "вставка для кровати SISO", "дюбель 8*40", "дюбель шарнирный DU650", _
                "еврошуруп 6,3*16", "завес Волна1000", "завес рояльный", "клей в тюбике", _
                "лифт 60", "лифт 80", "лифт 100", "лифт 50", "лифт 120", "муфта для стяжки DU650", _
                "накладка CAMAR  Л+П", "накладка CAMAR Лев", "накладка CAMAR Прав", _
                "пластина крепежная мет.", "подвеска шкафа SAH130", "подвеска шкафа SAH130 ЛЕВ", "подвеска шкафа SAH130 ПР", "подвесной крюк 6*50", _
                "стяжка колпачковая длинн.", "стяжка колпачковая кор.", "стяжка кроватная miniluna", _
                "стяжка-винт DU232", "дюбель DU232", "стяжка-винт DU634", "дюбель DU634", _
                "фиксатор стенки ДВП RV8", "фиксатор стенки ДВП RV1", "шуруп 3,5*35", "шуруп 4*35", _
                "лицевая панель д/тб внутр", "крепление ДВП в паз RV-8", _
                "накладка CAMAR  Л+П", "накладка CAMAR Лев", "накладка CAMAR Прав", "Тб Ант Ручка с повод. бел"
                ' "клиновая планка 5гр"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "эксцентрик 18мм", "эксцентрик 16мм", "эксцентрик 22мм"
            cbOpt.Enabled = False
                If kitchenPropertyCurrent.CamBibbColor = "" Then
                kitchenPropertyCurrent.CamBibbColor = GetCamBibbColor(kitchenPropertyCurrent.dspColor)
                If kitchenPropertyCurrent.CamBibbColor <> "" Then
                    UpdateOrder kitchenPropertyCurrent.OrderId, , , , , , , , kitchenPropertyCurrent.CamBibbColor
                End If
            End If
            cbLength.Enabled = False
        
        Case "заглушка для эксцентрика"
            cbOpt.Text = ""
            cbOpt.List = ЗаглЭксц
            cbLength.Text = ""
            cbLength.Enabled = False
            
        Case "клиновая планка Sensys"
            cbOpt.AddItem "5гр"
            cbOpt.AddItem "10гр"
            cbLength.Enabled = False
            
        Case "штанга"
            cbOpt.AddItem "765"
            cbOpt.AddItem "865"
            cbOpt.AddItem "815"
            cbOpt.AddItem "415"
            cbOpt.AddItem "425"
            cbOpt.AddItem "575"
            cbOpt.AddItem "773мм"
            cbOpt.AddItem "800мм"
            cbOpt.AddItem "688мм"
            cbOpt.AddItem "798мм"
            cbOpt.AddItem "998мм"
            cbOpt.AddItem "847мм"
            cbOpt.AddItem "ХРОМ 762мм"
            cbOpt.AddItem "ХРОМ 756мм"
            cbOpt.AddItem "ХРОМ 665мм"
            cbOpt.AddItem "ХРОМ 956мм"
            cbOpt.AddItem "труба хром d-25"
            
            ' длина не нужна
            cbLength.Enabled = False
        Case "ведро"
            cbOpt.Enabled = True
            cbOpt.AddItem "мусорное"
            cbOpt.AddItem "встроенное 5л"
            cbOpt.AddItem "встроенное 11л"
            
            If Not IsEmpty(FittingOption) And Len(FittingOption) >= 3 Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
                Else
            cbOpt.Value = "мусорное"
            End If
            
            cbLength.Enabled = False
        
        Case "сист. сорт. отх. тб"
            cbOpt.Enabled = False
            cbLength.Enabled = True
            cbLength.AddItem "90см"
            cbLength.Value = "90см"
        
        Case "замазка в тюбиках"
            cbOpt.Enabled = True
            cbOpt.AddItem "белая"
            cbOpt.AddItem "бук"
            cbOpt.AddItem "дуб"
            cbOpt.AddItem "махонь"
            cbOpt.AddItem "ольха"
            cbOpt.AddItem "орех"
            cbOpt.AddItem "чёрная"
            
            cbLength.Enabled = False
        
'        Case "направляющая верхняя", "направляющая нижняя"
'            cbOpt.AddItem "1567мм"
'            cbOpt.AddItem "1567мм"
'            cbOpt.AddItem "1400мм"
'            cbOpt.AddItem "1400мм"
'            cbOpt.AddItem "1362мм"
'            cbOpt.AddItem "1362мм"
'            'длина не нужна
'            cbLength.Text = ""
'            cbLength.Enabled = False
        Case "труба барная"
            cbOpt.AddItem "3м"
            cbOpt.AddItem "1,14м"
            cbOpt.AddItem "258мм"
            cbOpt.AddItem "274мм"
            cbOpt.AddItem "435мм"
            cbOpt.AddItem "526мм"
            cbOpt.AddItem "664мм"
            cbOpt.AddItem "816мм"
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        Case "направляющие Квадро"
            cbOpt.AddItem "25"
            cbOpt.AddItem "30"
            cbOpt.AddItem "35"
            cbOpt.AddItem "40"
            cbOpt.AddItem "45"
            cbOpt.AddItem "50"
            cbOpt.AddItem "50+PTO"
            
            If Not IsEmpty(FittingOption) Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next i
            End If
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        Case "вешалка выдвижная"
            cbOpt.AddItem "30"
            cbOpt.AddItem "45"
            cbOpt.AddItem "50"
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        Case "площадка для завеса"
            cbOpt.AddItem "Intermat D-0"
            cbOpt.AddItem "Intermat D-1.5"
            cbOpt.AddItem "Intermat D-3"
            cbOpt.AddItem "Sensys D-0"
            cbOpt.AddItem "Sensys D-1.5"
            cbOpt.AddItem "Sensys W45 D-1.5"
            cbOpt.AddItem "Sensys D-3"
            cbOpt.AddItem "Sens/Hett W45 D-3"
            cbOpt.AddItem "Clip накладная"
            cbOpt.AddItem "Clip прямая"
            cbOpt.AddItem "CLIP TOP D-3"
            cbOpt.AddItem "FGV H-4"
            cbOpt.AddItem "Intermat D-5"
            cbOpt.AddItem "Intermat D-8"
            cbOpt.AddItem "BLUM H-8,5"
            cbOpt.AddItem "BLUM H-11,5"
            cbOpt.AddItem "BLUM H-14,5"
            cbOpt.AddItem "SlideOn D-3"

            If Not IsEmpty(FittingOption) Then

            cbOpt.Text = FittingOption
            End If
            ' длина не нужна
            cbLength.Enabled = False
            
        Case "скотч"
            cbOpt.AddItem "ЗОВ"
            cbOpt.AddItem "коричневый"
            cbOpt.AddItem "прозрачный"
            ' длина не нужна
            cbLength.Enabled = False
        Case "амортизатор"
            cbOpt.AddItem "BLUM"
            cbOpt.AddItem "FGV"
            cbOpt.AddItem "врезной"
            ' длина не нужна
            cbLength.Enabled = False
        Case "подсветка диодная"
            cbOpt.AddItem "6400К"
            cbOpt.AddItem "6400К - 5м"
            cbOpt.AddItem "цветная 2м (с пультом)"
            cbOpt.AddItem "2К"
            ' длина не нужна
            cbLength.Enabled = False
            
        Case "кабель+вилка+выкл 220V"
            cbLength.Enabled = False
            cbOpt.Enabled = False
           
            
        Case "трансформатор"
            cbOpt.AddItem "LED 30W"
            cbOpt.AddItem "LED 50W"
            cbOpt.AddItem "гал 60W"
            cbOpt.AddItem "гал 105W"
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            cbAddNext.Value = True
        Case "крепл лиц пан тб внут L", "крепл лиц пан тб внут R"
            cbOpt.AddItem "мал"
            cbOpt.AddItem "бол"
            cbLength.Enabled = False
            'cbAddNext.Value = True
        Case "тандембокс внутр. 16 мал"
            cbOpt.AddItem "400 (289)"
            cbOpt.AddItem "450 (339)"
            cbOpt.AddItem "500 (389)"
            cbOpt.AddItem "550 (439)"
            cbOpt.AddItem "600 (489)"
            cbOpt.AddItem "650 (539)"
            cbOpt.AddItem "700 (589)"
            cbOpt.AddItem "750 (639)"
            cbOpt.AddItem "800 (689)"
            cbOpt.AddItem "850 (739)"
            cbOpt.AddItem "900 (789)"
            cbOpt.AddItem "950 (839)"
            cbOpt.AddItem "1000 (889)"
            cbOpt.AddItem "1050 (939)"
            cbOpt.AddItem "1100 (989)"
            cbOpt.AddItem "1150 (1039)"
            cbOpt.AddItem "1200 (1089)"
            ' cbLength.Enabled = False
            cbLength.List = tbLength
            cbLength.Text = "470"
            If Not IsEmpty(FittingLength) Then
                cbLength.Text = FittingLength
            End If
            If Not IsEmpty(FittingOption) Then
                cbOpt.Text = FittingOption
            End If
            
            Case "лиц пан т/б внутр 16 мал"
            cbOpt.AddItem "400 (289)"
            cbOpt.AddItem "450 (339)"
            cbOpt.AddItem "500 (389)"
            cbOpt.AddItem "550 (439)"
            cbOpt.AddItem "600 (489)"
            cbOpt.AddItem "650 (539)"
            cbOpt.AddItem "700 (589)"
            cbOpt.AddItem "750 (639)"
            cbOpt.AddItem "800 (689)"
            cbOpt.AddItem "850 (739)"
            cbOpt.AddItem "900 (789)"
            cbOpt.AddItem "950 (839)"
            cbOpt.AddItem "1000 (889)"
            cbOpt.AddItem "1050 (939)"
            cbOpt.AddItem "1100 (989)"
            cbOpt.AddItem "1150 (1039)"
            cbOpt.AddItem "1200 (1089)"
         ' длина не нужна
            cbLength.Enabled = False
            
            
        Case "тандембокс внутр. 16 бол"
            cbOpt.AddItem "400 (289/345)"
            cbOpt.AddItem "500 (389/445)"
            cbOpt.AddItem "550 (439/495)"
            cbOpt.AddItem "600 (489/545)"
            cbOpt.AddItem "650 (539/595)"
            cbOpt.AddItem "700 (589/645)"
            cbOpt.AddItem "750 (639/695)"
            cbOpt.AddItem "800 (689/745)"
            cbOpt.AddItem "850 (739/795)"
            cbOpt.AddItem "900 (789/845)"
            cbOpt.AddItem "950 (839/895)"
            cbOpt.AddItem "1000 (889/945)"
            cbOpt.AddItem "1050 (939/995)"
            cbOpt.AddItem "1100 (989/1045)"
            cbOpt.AddItem "1150 (1039/1095)"
            cbOpt.AddItem "1200 (1089/1145)"
            '  cbLength.Enabled = False
           cbLength.Enabled = True
            cbLength.List = tbLength
            cbLength.Text = "470"
            If Not IsEmpty(FittingLength) Then
                cbLength.Text = FittingLength
            End If
            If Not IsEmpty(FittingOption) Then
                cbOpt.Text = FittingOption
            End If
'          If Not IsEmpty(FittingLength) Then
'                For i = 0 To cbLength.ListCount - 1
'                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
'                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
'                        cbLength.Text = cbLength.List(i)
'                        Exit For
'                    End If
'                Next i
'            End If
'            If Not IsEmpty(FittingOption) Then
'                For i = 0 To cbOpt.ListCount - 1
'                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
'                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
'                        cbOpt.Text = cbOpt.List(i)
'                        Exit For
'                    End If
'                Next i
'            End If
      Case "лиц пан т/б внутр 16 бол"
      
            cbOpt.AddItem "400 (289/345)"
            cbOpt.AddItem "450 (339/395)"
            cbOpt.AddItem "500 (389/445)"
            cbOpt.AddItem "550 (439/495)"
            cbOpt.AddItem "600 (489/545)"
            cbOpt.AddItem "650 (539/595)"
            cbOpt.AddItem "700 (589/645)"
            cbOpt.AddItem "750 (639/695)"
            cbOpt.AddItem "800 (689/745)"
            cbOpt.AddItem "850 (739/795)"
            cbOpt.AddItem "900 (789/845)"
            cbOpt.AddItem "950 (839/895)"
            cbOpt.AddItem "1000 (889/945)"
            cbOpt.AddItem "1050 (939/995)"
            cbOpt.AddItem "1100 (989/1045)"
            cbOpt.AddItem "1150 (1039/1095)"
            cbOpt.AddItem "1200 (1089/1145)"
             ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
          
        Case "прод реллинг тб вн 16 бол"
           
            cbOpt.AddItem "400 (345)"
            cbOpt.AddItem "500 (445)"
            cbOpt.AddItem "550 (495)"
            cbOpt.AddItem "600 (545)"
            cbOpt.AddItem "650 (595)"
            cbOpt.AddItem "700 (645)"
            cbOpt.AddItem "750 (695)"
            cbOpt.AddItem "800 (745)"
            cbOpt.AddItem "850 (795)"
            cbOpt.AddItem "900 (845)"
            cbOpt.AddItem "950 (895)"
            cbOpt.AddItem "1000 (945)"
            cbOpt.AddItem "1050 (995)"
            cbOpt.AddItem "1100 (1045)"
            cbOpt.AddItem "1150 (1095)"
            cbOpt.AddItem "1200 (1145)"
             ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            
      
         Case "прод реллинг тб вн 18 бол"
            cbOpt.AddItem "400 (341)"
            cbOpt.AddItem "500 (441)"
            cbOpt.AddItem "550 (491)"
            cbOpt.AddItem "600 (541)"
            cbOpt.AddItem "650 (591)"
            cbOpt.AddItem "700 (641)"
            cbOpt.AddItem "750 (691)"
            cbOpt.AddItem "800 (741)"
            cbOpt.AddItem "850 (791)"
            cbOpt.AddItem "900 (841)"
            cbOpt.AddItem "950 (891)"
            cbOpt.AddItem "1000 (941)"
            cbOpt.AddItem "1050 (991)"
            cbOpt.AddItem "1100 (1041)"
            cbOpt.AddItem "1150 (1091)"
            cbOpt.AddItem "1200 (1141)"
             ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
                 
                 
              
        Case "тандембокс внутр. 18 мал"
            
            cbOpt.AddItem "400 (285)"
            cbOpt.AddItem "500 (385)"
            cbOpt.AddItem "550 (435)"
            cbOpt.AddItem "600 (485)"
            cbOpt.AddItem "650 (535)"
            cbOpt.AddItem "700 (585)"
            cbOpt.AddItem "750 (635)"
            cbOpt.AddItem "800 (685)"
            cbOpt.AddItem "850 (735)"
            cbOpt.AddItem "900 (785)"
            cbOpt.AddItem "950 (835)"
            cbOpt.AddItem "1000 (885)"
            cbOpt.AddItem "1050 (935)"
            cbOpt.AddItem "1100 (985)"
            cbOpt.AddItem "1150 (1035)"
            cbOpt.AddItem "1200 (1085)"
            
            'cbLength.Enabled = False
            cbLength.Enabled = True
            cbLength.List = tbLength
            cbLength.Text = "470"
            
            If Not IsEmpty(FittingLength) Then
                cbLength.Text = FittingLength
            End If
            If Not IsEmpty(FittingOption) Then
                cbOpt.Text = FittingOption
            End If
            
'           If Not IsEmpty(FittingLength) Then
'                For i = 0 To cbLength.ListCount - 1
'                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
'                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
'                        cbLength.Text = cbLength.List(i)
'                        Exit For
'                    End If
'                Next i
'            End If
'            If Not IsEmpty(FittingOption) Then
'                For i = 0 To cbOpt.ListCount - 1
'                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
'                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
'                        cbOpt.Text = cbOpt.List(i)
'                        Exit For
'                    End If
'                Next i
'            End If
       Case "лиц пан т/б внутр 18 мал"
            cbOpt.AddItem "400 (285)"
            cbOpt.AddItem "500 (385)"
            cbOpt.AddItem "550 (435)"
            cbOpt.AddItem "600 (485)"
            cbOpt.AddItem "650 (535)"
            cbOpt.AddItem "700 (585)"
            cbOpt.AddItem "750 (635)"
            cbOpt.AddItem "800 (685)"
            cbOpt.AddItem "850 (735)"
            cbOpt.AddItem "900 (785)"
            cbOpt.AddItem "950 (835)"
            cbOpt.AddItem "1000 (885)"
            cbOpt.AddItem "1050 (935)"
            cbOpt.AddItem "1100 (985)"
            cbOpt.AddItem "1150 (1035)"
            cbOpt.AddItem "1200 (1085)"
           ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
           
        Case "тандембокс внутр. 18 бол"
            cbOpt.AddItem "400 (285/341)"
            cbOpt.AddItem "500 (385/441)"
            cbOpt.AddItem "550 (435/491)"
            cbOpt.AddItem "600 (485/541)"
            cbOpt.AddItem "650 (535/591)"
            cbOpt.AddItem "700 (585/641)"
            cbOpt.AddItem "750 (635/691)"
            cbOpt.AddItem "800 (685/741)"
            cbOpt.AddItem "850 (735/791)"
            cbOpt.AddItem "900 (785/841)"
            cbOpt.AddItem "950 (835/891)"
            cbOpt.AddItem "1000 (885/941)"
            cbOpt.AddItem "1050 (935/991)"
            cbOpt.AddItem "1100 (985/1041)"
            cbOpt.AddItem "1150 (1035/1091)"
            cbOpt.AddItem "1200 (1085/1141)"
            ' cbLength.Enabled = False
              cbLength.Enabled = True
                     cbLength.List = tbLength
         cbLength.Text = "470"
            If Not IsEmpty(FittingLength) Then
                cbLength.Text = FittingLength
            End If
            If Not IsEmpty(FittingOption) Then
                cbOpt.Text = FittingOption
            End If
'         If Not IsEmpty(FittingLength) Then
'                For i = 0 To cbLength.ListCount - 1
'                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) = 1 Or _
'                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
'                        cbLength.Text = cbLength.List(i)
'                        Exit For
'                    End If
'                Next i
'            End If
'            If Not IsEmpty(FittingOption) Then
'                For i = 0 To cbOpt.ListCount - 1
'                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
'                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
'                        cbOpt.Text = cbOpt.List(i)
'                        Exit For
'                    End If
'                Next i
'            End If
         Case "лиц пан т/б внутр 18 бол"
            cbOpt.AddItem "400 (285/341)"
            cbOpt.AddItem "500 (385/441)"
            cbOpt.AddItem "550 (435/491)"
            cbOpt.AddItem "600 (485/541)"
            cbOpt.AddItem "650 (535/591)"
            cbOpt.AddItem "700 (585/641)"
            cbOpt.AddItem "750 (635/691)"
            cbOpt.AddItem "800 (685/741)"
            cbOpt.AddItem "850 (735/791)"
            cbOpt.AddItem "900 (785/841)"
            cbOpt.AddItem "950 (835/891)"
            cbOpt.AddItem "1000 (885/941)"
            cbOpt.AddItem "1050 (935/991)"
            cbOpt.AddItem "1100 (985/1041)"
            cbOpt.AddItem "1150 (1035/1091)"
            cbOpt.AddItem "1200 (1085/1141)"
             ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
    
        
        
        Case "попер реллинг в АНТ вн"
            cbOpt.AddItem "16"
            cbOpt.AddItem "18"
            cbLength.Text = ""


        Case "вешалка выдвижная"
            cbOpt.AddItem "30"
            cbOpt.AddItem "35"
            cbOpt.AddItem "40"
            cbOpt.AddItem "45"
            cbOpt.AddItem "50"
            
        Case "полироль"
            cbOpt.AddItem "50мл"
            cbOpt.Text = "50мл"
            'cbOpt.AddItem "250мл"
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
             
        Case "пол с подсветкой"
            cbOpt.AddItem "HLT45"
            cbOpt.AddItem "HLT60"
            cbOpt.AddItem "HLT90"
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        
        Case "соед планка 200мм"
           cbOpt.Enabled = False
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        
        
        Case "завес под амортиз-р BLUM", "завес под амортизатор FGV", "дюбель DU325 Rapid S", "дюбель VB15"
                    
            cbOpt.Enabled = False
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            cbAddNext.Value = True
            
   '     Case "поддон алюминиевый"
        Case "поддон"
            cbOpt.AddItem "алюминиевый"
            cbOpt.AddItem "пластик"
    '       cbOpt.Text = ""
    '       cbOpt.Enabled = False
            cbLength.Enabled = True
            cbLength.Text = ""
            cbLength.List = PA
        
'        Case "сифон к мойке"
'            cbOpt.Enabled = True
'            cbLength.Enabled = False
'            cbOpt.AddItem "1/2"
'            cbOpt.AddItem "3,5"
'            cbAddNext.Value = True
            
'        Case "крепление к мойке"
'            cbOpt.Enabled = True
'            cbLength.Enabled = False
'            cbOpt.List = Sink
'            cbOpt.Text = Krepl
'            cbAddNext.Value = True
            
'        Case "Отбортовка 4м"
'            cbOpt.Enabled = False
'            cbLength.Enabled = False
'            cbAddNext.Value = True
            
        Case "сушка"
           ' ширина
           cbLength.Enabled = True
            cbLength.Text = ""
            cbLength.List = SW

            If Not IsEmpty(FittingLength) And Len(FittingLength) > 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
'
            ' цвет
            cbOpt.Text = ""
            cbOpt.List = Sushk
            
            If Not IsEmpty(FittingOption) And Len(FittingOption) >= 3 Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) = 1 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
        
        Case "лоток"
    
            ' ширина
            cbLength.Enabled = True
            cbLength.Text = ""
            cbLength.List = LW
            
            If Not IsEmpty(FittingLength) And Len(FittingLength) > 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
            
            ' цвет не нужен
            cbOpt.Text = ""
            cbOpt.Enabled = False
            
'        Case "спинка"
'            cbOpt.List = Спинка
'            cbLength.List = BackKolib
'            cbAddNext.Value = True
            
        Case "отбойник"
            cbOpt.Text = "ПВХ"
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            
        Case "направляющие"
            cbOpt.List = Направляющие
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
       
            Dim fo As String
             If IsEmpty(FittingOption) Then fo = "шарик" Else fo = FittingOption
            If Not IsEmpty(FittingLength) Then fo = fo & " " & FittingLength
            fo = Trim(fo)
                For i = 0 To cbOpt.ListCount - 1
                    If fo = cbOpt.List(i) Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next i
        Case "Отбортовка 3м", "Отбортовка горб-4", "Отбортовка горб-5", "Отбортовка TOP-Line", _
             "Угол к отб. 3м", "Загл. к отб. 3м", "Угол внешн. к отб. 3м", _
             "Угол к отб. 4м", "Загл. к отб. 4м", "Угол внешн. к отб. 4м", _
             "Угол к отб. горб-4", "Загл. к отб. горб-4", "Загл. лев к отб. горб-4", "Загл. прав к отб. горб-4", "Угол внешн. к отб. горб-4", _
             "Угол к отб. горб-5", "Загл. к отб. горб-5", "Загл. лев к отб. горб-5", "Загл. прав к отб. горб-5", "Угол внешн. к отб. горб-5", _
             "Угол к отб. TOP-Line", "Загл. к отб. TOP-Line", "Угол внеш. к отб TOP-Line", _
             "вставка в отбортовку", "вставка-40 в отбортовку", "Отбортовка 4м"
             
            
            
            cbLength.Text = ""
            cbLength.Enabled = False
            
            cbOpt.Text = ""
            
            Select Case cbFittingName.Text
            
                Case "Отбортовка 4м"
                    cbOpt.Enabled = False
                    ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
                    cbAddNext.Value = True
                    
                    
                
                Case "Отбортовка горб-4", "Отбортовка горб-5"
                    cbOpt.List = OtbGorbColors
'
'                Case "Отбортовка 4м"
'                    'bSkip = True
'                    cbOpt.List = Вставка
                    
                Case "вставка в отбортовку", "вставка-40 в отбортовку"
                    'bSkip = True
                    cbOpt.List = Вставка
                    
                Case "Угол к отб. 4м", "Загл. к отб. 4м", "Угол внешн. к отб. 4м"
                
                    cbOpt.List = МелочьОтб4м
                    
                    If Not IsEmpty(FittingOption) Then
                        Select Case FittingOption
                            Case МелочьОтб4м(0), МелочьОтб4м(1), МелочьОтб4м(2), МелочьОтб4м(3), МелочьОтб4м(4), МелочьОтб4м(5), МелочьОтб4м(6)
                            Case Else
                            Select Case FittingOption
                            
                                Case "Алюминиевая полоса", "Антрацитовый перламутр", "Арктик", "Белая крошка", "Весенний перламутр", _
                                    "Иней", "Каппучино", "Кварц", "Лазурный глянец", "Лазурный матовый", "Латунь", "Лунный металл", "Магма", "Малахитовая полоса", _
                                    "Металлик", "Песчаник", "Платина", "Ровенна", "Рубиновая полоса", "Серая крошка", "Серебристый перламутр", "Сизый камень", "Серый иней", "Мрамор"
                                
                                    FittingOption = "СЕР"
                                    
                                Case "Беж гранит глянец", "Беж гранит матовый", "Бежевый монохром", "Дуб полосатый", "Желтый камень", "Янтарь", "Песочный иней", _
                                    "Известняк", "Камушки", "Кремовый перламутр", "Лимонный глянец", "Лимонный матовый", "Медный матовый", "Медный глянец", _
                                    "Оникс", "Салатовый глянец", "Салатовый матовый", "Туринский гранит", "Цитрусовый глянец", "Цитрусовый матовый", "Ясень светлый", "Ракушечник", "Туя"
                                                                    
                                    FittingOption = "БЕЖ"
                                    
                                Case "Накарадо"
                                    
                                    FittingOption = "БЕЖ"
                                
                                Case "СНОУ МИЛКИ"
                                                                    
                                    FittingOption = "БЕЖ"
                                    
                                Case "Базальт", "Груша", "Вишня", "Золото", "Красный иней", "Рыжий иней", "Брешиа", _
                                "Кафель", "Корень глянец", "Мрамор коричневый", "Красный монохром", "Марс глянец", "Терракот", "Яшма", "Коралл", "Мореный дуб"
                                
                                    FittingOption = "Кор"
                                    
                                Case "Лотос черный", "Махонь", "Темная крошка", "Черная бронза", "Мрамор черный", "Ясень темный", "Шерл", "Морион", "СНОУ БЛЭК", "Галициа"
                                
                                    FittingOption = "Черн"
                                    
                                Case "Желтый камень", "Зеленый матовый"
                                
                                    FittingOption = "Зел"
                                    
                                Case "СНОУ УАЙТ", "Лотос белый", "Марокко камень"
                                
                                    FittingOption = "Бел"
                                
                                Case "Милано глянец", "Синий глянец", "Синий матовый"
                                    
                                    FittingOption = "Син"
                                    
                            End Select
                        End Select
                     End If
                    
                Case "Угол к отб. горб-4", _
                     "Загл. к отб. горб-4", _
                     "Загл. лев к отб. горб-4", _
                     "Загл. прав к отб. горб-4", _
                     "Угол внешн. к отб. горб-4", _
                     "Угол к отб. горб-5", _
                     "Загл. к отб. горб-5", _
                     "Загл. лев к отб. горб-5", _
                     "Загл. прав к отб. горб-5", _
                     "Угол внешн. к отб. горб-5"
                     
                     cbOpt.List = ГорбатаяМелочь
                     ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
                     
                     If Not IsEmpty(FittingOption) Then
                        Select Case FittingOption
                            Case ГорбатаяМелочь(0), ГорбатаяМелочь(1), ГорбатаяМелочь(2), ГорбатаяМелочь(3), ГорбатаяМелочь(4), ГорбатаяМелочь(5)
                            Case Else
                            Select Case FittingOption
                            
                                                                   
                                Case "БЕЛ КРОШКА", "ДЕЛФИ", "СИЗ КАМ", "СЕРАЯ КРОШКА", "ИНЕЙ БЕЛЫЙ", "АРКТИК", "ПЕСОЧНЫЙ ИНЕЙ", "ИЗВЕСТНЯК", "МРАМОР"
                                    FittingOption = "С-СЕР"
                                Case "БЕЖ ГР", "БЕЖ ГР ГЛЯНЕЦ", "КАМУШКИ", "ЯСЕНЬ СВЕТЛЫЙ", _
                                    "МАРС ГЛЯНЕЦ", "КАФЕЛЬ", "ОНИКС", "РАКУШЕЧНИК", "ЯНТАРЬ", "БЕЖ ГР", "ТУРИН ГРАНИТ"
                                    FittingOption = "БЕЖ"
                                Case "ТУЯ"
                                    FittingOption = "Т-беж"
                                Case "ЯСЕНЬ ТЕМНЫЙ", "ЯШМА", "СЕРЫЙ ИНЕЙ", "РЫЖИЙ ИНЕЙ", "КРАСНЫЙ ИНЕЙ"
                                    FittingOption = "Т-кор"
                                Case "КОРЕНЬ ГЛЯНЕЦ"
                                    FittingOption = "Рыжий"
                                Case "ЛУННЫЙ КАМЕНЬ", "АЛЮМИНИЙ", "КВАРЦ", "БРЕШИА"
                                    FittingOption = "Т-сер"
                                Case "КОРАЛЛ"
                                    FittingOption = "Т-кор"
                                Case "ЧЕРН МРАМОР ГЛ", "Черная бронза", "Темная крошка", "СНОУ БЛЭК", "ГАЛИЦИА"
                                    FittingOption = "Черн"
                                Case "СНОУ МИЛКИ", "ЖЕЛТ КАМ"
                                    FittingOption = "Крем"
                                Case "НАКАРАДО"
                                    FittingOption = "Накар"
                                Case "СНОУ УАЙТ"
                                    FittingOption = "Бел"
                            End Select
                        End Select
                     End If
                     
                Case "Отбортовка TOP-Line", "Угол к отб. TOP-Line", "Загл. к отб. TOP-Line", "Угол внеш. к отб TOP-Line"
                    cbOpt.List = TOPLine
                  ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
                Case Else
                    cbOpt.List = OtbColors
                    ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            End Select
                                   
            
            If Not IsEmpty(FittingOption) And Len(FittingOption) >= 3 Then
                For i = 0 To cbOpt.ListCount - 1
                    If FittingOption = cbOpt.List(i) Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
            
        Case "нажимной м-м Push-To-Open"
            cbOpt.Enabled = True
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            
            cbOpt.AddItem "универсальный"
            cbOpt.AddItem "с магнитом"
            
        Case "разделитель д/тб с флажк."
            cbOpt.Enabled = False
            cbLength.Enabled = True
            
            cbLength.AddItem "60см"
            cbLength.AddItem "80см"
            cbLength.AddItem "90см"
        Case "лоток д/тб"
            cbOpt.Enabled = False
            cbLength.Enabled = True
            cbLength.AddItem "40см"
            cbLength.AddItem "50см"
            cbLength.AddItem "60см"
            cbLength.AddItem "80см"
            cbLength.AddItem "90см"
            cbLength.AddItem "90см с банками"
        Case "надставка д/тб"
            cbOpt.Enabled = True
            cbLength.Enabled = False
            
            cbOpt.AddItem "серая"
            cbOpt.AddItem "прозрачная"
            
        Case "крепление т/б з.с. 70мм"
            cbOpt.Enabled = True
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "прав"
            cbOpt.AddItem "лев"
            
        Case "крепление т/б з.с.144мм"
            cbOpt.Enabled = True
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "прав"
            cbOpt.AddItem "лев"
         Case "организация под мойку"
            cbOpt.Enabled = True
            cbOpt.Text = "OrgalFlex"
            cbLength.Text = ""
            cbLength.Enabled = False
             
        Case "крепление к квадро"
            cbOpt.Enabled = True
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "заднее"
            cbOpt.AddItem "переднее"
           
        Case "Решетка вентиляционная"
            cbOpt.Enabled = True
            cbOpt.AddItem "стандарт"
            cbOpt.AddItem "Вальпато"
            cbOpt.Text = "стандарт"
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        
        Case "Подсветка диодная 2K", "пл. монт. clip HF,HKS", "крепление к рел. тб", "ограничитель угла 83 HF", "завес Blum110 п/с узAlпро", "завес Blum110 п/а узAlпро", "ограничитель угла 75 HK"
            cbOpt.Enabled = False
           ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
      Case "трансф-р LED 30w"
            cbOpt.Enabled = False
           ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            cbAddNext.Value = True
      
        Case "Светильник диодный"
            cbOpt.Enabled = True
            ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "Barri (3шт)"
         
        Case "стул И/Ф", "стул Женева", "стул Юлия", "стул Браун"
        
         cbOpt.Enabled = True
         ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
         get_st_par
         cbOpt.List = Stul_color_no
         
         Case "ф-ра комплект Алмата"
            cbOpt.AddItem "DK-1"
            cbOpt.AddItem "DP-1"
            cbOpt.AddItem "DP-2"
            cbOpt.AddItem "DP-3/1"
            cbOpt.AddItem "DP-3/2"
            cbOpt.AddItem "DP-4"
            cbOpt.AddItem "DP-5"
            cbOpt.AddItem "DP-6"
            cbOpt.AddItem "DP-9"
            cbOpt.AddItem "DP-10"
            cbOpt.AddItem "DP-11"
            cbOpt.AddItem "KOM-1"
            cbOpt.AddItem "KOM-2"
            cbOpt.AddItem "TV-SP-1"
            cbOpt.AddItem "TV-SP-1/1"
            cbOpt.AddItem "TV-SP-2"
            cbOpt.AddItem "TV-SP-3"
            cbOpt.AddItem "TV-SP-4"
            cbOpt.AddItem "TV-SP-5"
            cbOpt.AddItem "TV-SP-6"
            cbOpt.AddItem "V-DP-1"
            cbOpt.AddItem "V-DP-2"
            cbOpt.AddItem "V-DP-3"
            cbOpt.AddItem "V-DP-4"
            cbOpt.AddItem "V-DP-5"
            cbOpt.AddItem "V-DP-6"
            cbOpt.AddItem "V-DP-8"
            cbOpt.AddItem "VTR-1"
            cbOpt.AddItem "VTR-2"
            cbOpt.AddItem "YT-1"
            cbOpt.AddItem "YT-2"
            cbOpt.AddItem "LAV-1"
            cbOpt.AddItem "LAV-2"
            cbOpt.AddItem "LAV-3"
            cbOpt.AddItem "LAV-4"
            cbOpt.AddItem "LAV-5"
            cbOpt.AddItem "LAV-6"
            cbOpt.AddItem "LAV-7"
            cbOpt.AddItem "LAV-8"
            cbOpt.AddItem "LAV-9"
            cbOpt.AddItem "LAV-10"
            cbOpt.AddItem "LAV-11"
            cbOpt.AddItem "AY-1"
            cbOpt.AddItem "AY-2"
            cbOpt.AddItem "AY-3"
            cbOpt.AddItem "AY-4"
            cbOpt.AddItem "AY-5"
            cbOpt.AddItem "AY-6"
            cbOpt.AddItem "AY-7"
            cbOpt.AddItem "AY-8"
            cbOpt.AddItem "AY-9"
            cbOpt.AddItem "AY-10"
            cbOpt.AddItem "AY-11"

            If Not IsEmpty(FittingOption) Then
            For i = 0 To cbOpt.ListCount - 1
                If FittingOption = cbOpt.List(i) Then
                    cbOpt.Text = cbOpt.List(i)
                    Exit For
                End If
            Next
            Else
                    cbOpt.Text = ""
            End If
            cbLength.Enabled = False
            
        Case "ф-ра комплект ВЛШВ2"
          cbOpt.Enabled = True
            cbLength.Text = ""
            cbLength.Enabled = False
        
        Case "стул Zebra"
         cbOpt.AddItem "2273(100)прозр."
         cbOpt.AddItem "2273(130)прозр.оранж."
         cbOpt.AddItem "2273(140)прозр.красный"
         cbOpt.AddItem "2273(183)прозр."
         cbOpt.AddItem "2273(310)гл.белый"
         cbOpt.AddItem "2273(380)гл.чёрный"
         cbLength.Text = ""
         cbLength.Enabled = False
        
        Case "Завес CLIP top"
            'cbAddNext.Value = False
        cbOpt.Enabled = True
        cbOpt.Text = ""
        cbOpt.List = zavesClipTop
        cbLength.Text = ""
        cbLength.Enabled = False
        If Not IsEmpty(FittingOption) Then
            If FittingOption = "BLUMOTION +90 под фп" Then cbOpt.Text = FittingOption
            If FittingOption = "BLUMOTION +45" Then cbOpt.Text = FittingOption
            If FittingOption = "+155" Then
            cbOpt.Text = FittingOption
            cbAddNext.Value = True
            End If
        End If
        Case "завес Sensys"
            
        cbOpt.Enabled = True
        
        cbOpt.List = zavesSensys
        
        If Not IsEmpty(FittingOption) Then
            For i = 0 To cbOpt.ListCount - 1
                If FittingOption = cbOpt.List(i) Then
                    cbOpt.Text = cbOpt.List(i)
                    Exit For
                End If
            Next
        Else
                cbOpt.Text = ""
        End If
        ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
'        If casepropertyCurrent Is Nothing Then
'        cbAddNext.Value = True
'        ElseIf casepropertyCurrent.p_fullcn = "" Then
'        cbAddNext.Value = True
'        End If
        
            
        Case "площадка зав.Sensys"
        cbOpt.Enabled = True
        cbOpt.Text = ""
        cbOpt.List = ploschadkaSensys
        ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
         
        Case "амморт CLIP TOP +155"
        cbOpt.Enabled = False
        cbLength.Text = ""
        cbLength.Enabled = False
        
        Case "амморт. Sensys 165"
        cbOpt.Enabled = False
        ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        'cbAddNext.Value = False
         
        Case "огр. угла Sensys 110-85 ", "загл на чашку зав.Sensys", "загл на плечо зав.Sensys", "боковина тб лев", "боковина тб прав"
        cbOpt.Enabled = False
        ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
      
       Case "крепление фасада BLUM"
        cbOpt.Enabled = True
        ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False
        cbOpt.AddItem "HK"
        cbOpt.AddItem "HL"
        cbOpt.AddItem "HS"
        cbOpt.AddItem "HK-S"
        
        Case "SL56 алюмм профиль"
        cbOpt.Enabled = False
        cbLength.Enabled = True
        cbLength.Text = "3000 мм"
        cbLength.AddItem "3000 мм"
        cbLength.AddItem "1638 мм"
        Case "SL56  ходовой элемент", "SL56 направляющий элемент", "SL56 распорный держатель", "SL56 упор лев-прав двери", "SL56 алюмм профиль 1638мм", "SL56 сред стоп пер дв"
        cbOpt.Enabled = False
        ' длина не нужна
            cbLength.Text = ""
            cbLength.Enabled = False

        Case "крепл. фасадов HK,HL,HS"
        cbLength.Enabled = False
        cbOpt.Enabled = False

        Case "соединитель цоколя"
        cbOpt.List = Цоколь
        cbLength.Enabled = False
        
        Case "коврик антискольжения", "коврик антискольж БЕЛ"
        cbLength.Enabled = True
        cbLength.List = tbkovrLength
        cbOpt.Enabled = True
        cbOpt.List = tbkovrOpt
        
        Case "Сушка в нижний шкаф"
            cbOpt.Enabled = True
            cbOpt.AddItem "60"
            cbOpt.AddItem "80"
            cbOpt.AddItem "90"
            cbLength.Enabled = False
            If Not IsEmpty(FittingOption) Then
                cbOpt.Text = FittingOption
            End If
        
        Case "SIGE"
            cbOpt.Enabled = True
            cbOpt.AddItem "070i - полка обор 1/2"
            cbOpt.AddItem "361i - полка обор 3/4"
            cbOpt.AddItem "370L - Nuvola L"
            cbOpt.AddItem "370R - Nuvola R"
            cbOpt.AddItem "258A 20 - Карго клн 20"
            cbOpt.AddItem "258A 15 - Карго клн 15"
            cbOpt.AddItem "575 - ССО"
            cbOpt.AddItem "сушка в нижний шкаф 60"
            cbOpt.AddItem "сушка в нижний шкаф 90"
            cbOpt.AddItem "230A 450 - Карг расп Maxi"
            cbOpt.AddItem "230A 600 - Карг расп Maxi"
            cbOpt.AddItem "230B 450 - Карг расп MIDI"
            cbOpt.AddItem "230B 600 - Карг расп MIDI"
            cbLength.Enabled = False
            If Not IsEmpty(FittingOption) Then
                cbOpt.Text = FittingOption
            End If
        
        Case "Tip-ON"
        cbOpt.Enabled = True
        cbOpt.AddItem "955 комплект"
        cbOpt.AddItem "955 д/пет без пруж"
        cbOpt.AddItem "955А комплект"
        cbOpt.AddItem "955А усил д/пет без пруж"
        cbOpt.AddItem "планка на клею"
        cbOpt.AddItem "держатель крестообразный"
        cbLength.Enabled = False
        
        Case "ф-ра комплект ВИОЛА"
        cbOpt.Enabled = True
        cbOpt.AddItem "АЛШВ1"
        cbOpt.AddItem "АЛТВ1"
        cbOpt.AddItem "ВЛШБ1"
        cbOpt.AddItem "ВЛШБ2"
        cbOpt.AddItem "ВЛШБ3"
        cbOpt.AddItem "ВЛШБК1"
        cbOpt.AddItem "ВЛШБК2"
        cbOpt.AddItem "ВЛШБК3"
        cbOpt.AddItem "ВЛШБК4"
        cbOpt.AddItem "ВЛШКМ1"
        cbOpt.AddItem "ВЛШСП"
        cbOpt.AddItem "ВЛШСПМ"
        cbOpt.AddItem "ВЛТВ1"
        cbOpt.AddItem "ВЛТВ2"
        cbOpt.AddItem "ВЛШВ1"
        cbOpt.AddItem "ВЛШВ2"
        cbOpt.AddItem "ВЛШВ3"
        cbOpt.AddItem "ВЛШК1"
        cbOpt.AddItem "ВЛШК2"
        cbOpt.AddItem "ВЛШН1"
        cbOpt.AddItem "ВЛШТВ1"
        cbOpt.AddItem "ВЛШТВ2"
        cbOpt.AddItem "ВЛШТГ1"
        cbOpt.AddItem "ВЛШТГ2"
        cbOpt.AddItem "ВЛШУ1"
        cbOpt.AddItem "ВЛТСТ1"
        cbOpt.AddItem "ВЛТСТ2"
        cbOpt.AddItem "ВЛТП1"
        cbOpt.AddItem "ВЛШТГ4"
        cbOpt.AddItem "ВЛШТГ5"
        cbOpt.AddItem "ВЛШТ4"
        cbOpt.AddItem "ВЛШТ7"
        cbOpt.AddItem "ВЛШПР1"
        cbOpt.AddItem "ВЛШПР2"
        cbOpt.AddItem "ВЛШПР3"
        cbOpt.AddItem "ВЛШПР4"
        cbOpt.AddItem "ВЛШПР5"
        cbOpt.AddItem "ВЛШТГ6комби"
        cbOpt.AddItem "ВЛШТГ6полки"
        cbOpt.AddItem "ВЛШТГ6штанга"
        cbOpt.AddItem "ВЛШСПП"
        
        cbLength.Enabled = True
        cbLength.AddItem "Массив ясень 101"
        cbLength.AddItem "Массив ясень 110"
        cbLength.AddItem "Массив ясень 111"
        cbLength.AddItem "Массив ясень 112"
        cbLength.AddItem "Массив ясень 113"
        cbLength.AddItem "Массив ясень 119"
        cbLength.AddItem "Массив ясень 120"
        cbLength.AddItem "Массив ясень 122"
        cbLength.AddItem "Массив ольха 115"
        cbLength.AddItem "Массив ольха 113"
        cbLength.AddItem "Массив ольха 120"

        
        Case "ф-ра комплект"
        cbOpt.Enabled = True
        cbOpt.AddItem "стенд под мойки"
        cbLength.Enabled = False
        
        Case "цоколь Волпато100 Ал", "цоколь Волпато150 Ал"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "отбортовка Волпато 4м Ал"
            cbAddNext.Value = True
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "угол90 к цоколю Волпато", "клипса к цоколю Волпато", "угол90 к цок Волпато150", "угл+загл к отб Волпато"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "профиль Алюм"
            cbOpt.AddItem "гор 1закр G12AL07"
            cbOpt.AddItem "гор 2закр G13AL07"
            cbOpt.AddItem "верт 1закр G16AL07"
            cbOpt.AddItem "верт 2закр G15AL07"
            cbOpt.AddItem "с упл 1закр G14AL07"
            cbOpt.AddItem "подстав верт закр (к-т)"
            cbOpt.AddItem "загл мал"
            cbOpt.AddItem "загл бол"
            cbOpt.AddItem "угол 90внут"
            cbOpt.AddItem "угол 90наруж"
            cbOpt.AddItem "к-т крепл"
            
            cbLength.Enabled = True
            
        Case "крепление профиля Алюм"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "заглушка профиля Волпато"
            cbOpt.AddItem "81/G1.1AT2"
            cbOpt.Text = "81/G1.1AT2"
            cbLength.Enabled = False

        
        Case "Мойка"
            cbOpt.AddItem "Аурис"
            cbOpt.AddItem "АурисЭко"
            cbOpt.AddItem "Квадро"
            cbOpt.AddItem "Квадро+"
            cbOpt.AddItem "КвадроЭко"
            cbOpt.AddItem "Циркум"
            cbOpt.AddItem "Циркум+"
            cbOpt.AddItem "ЦиркумЭко"
            cbOpt.AddItem "Рабис"
            cbOpt.AddItem "Гравис"
            cbOpt.AddItem "КМ-1"
            cbOpt.AddItem "КМ-2"
            cbOpt.AddItem "КМ-3"
            cbOpt.AddItem "КМ-4"
            cbOpt.AddItem "КМ-5"
            cbOpt.AddItem "КМ-6"
        Case "Мойка Аурис", "Мойка АурисЭко", "Мойка Гравис", "Мойка Квадро", "Мойка Квадро+", "Мойка КвадроЭко", "Мойка КМ-1", "Мойка КМ-2", "Мойка КМ-3", "Мойка КМ-4", "Мойка КМ-5", _
            "Мойка КМ-6", "Мойка Рабис", "Мойка Циркум", "Мойка Циркум+", "Мойка ЦиркумЭко"
            
            cbLength.Clear
            cbLength.AddItem "слева(чаша)"
            cbLength.AddItem "справа(чаша)"
            cbOpt.Clear
            cbOpt.List = MoikaColors
        
        Case "пробка ANODA натур", "дюбель 8*60", "стяжка RAFIX TAB20", "шуруп 5*80", "уголок мет. 60*60*50"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        
        Case "загл д/отверстия под провод"
            cbOpt.AddItem "по умолч."
            cbOpt.AddItem "серый"
            cbOpt.Text = "серый"
            cbLength.Enabled = False
            
       Case Else
            If cbFittingName.ListIndex > -1 Then
            cbOpt.Text = ""
            cbOpt.Enabled = False
            cbLength.Text = ""
            cbLength.Enabled = False
            Else
            
            MsgBox "Я такой фурнитуры не знаю..."
            End If
            'Exit Function
    
    
    
    
    End Select
    

End Sub
    
Private Sub cbOpt_Change()
    If binit Then Exit Sub
    
    Select Case cbFittingName.Text
    
        Case "ручка"
            If cbOpt.Text = "неликвид" Then
                cbLength.Enabled = True
            Else
                cbLength.Enabled = False
            End If
        Case "коврик антискольжения", "коврик антискольж БЕЛ"
            If cbOpt.Text = "мп" Then
                cbLength.Clear
                cbLength.Text = ""
                Else
                cbLength.Text = ""
                cbLength.List = tbkovrLength
                
            End If
'        Case "завешка"
'            If cbOpt.Text = "CAMAR Л+П" Then
'                cbAddNext.Value = True
'            End If
        Case "Мойка"
            If cbOpt.Text <> "" Then
                cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
            End If
        
        Case "стул И/Ф", "стул Женева", "стул Юлия", "стул Браун"
            If Len(cbOpt.Text) = 3 Then
            i = 0
            For Each tstr In Stul_color_1
            If i > 0 Then Stul_color_1(i) = cbOpt.Text & Stul_color_1(i)
            i = i + 1
            Next
            cbOpt.List = Stul_color_1
            cbAddNext.Value = False
            ElseIf Len(cbOpt.Text) = 0 Then
                get_st_par
                cbOpt.List = Stul_color_no
                cbAddNext.Value = False
            ElseIf Len(cbOpt.Text) = 4 Then
                i = 0
                For Each tstr In Stul_color_2
                If i > 0 Then Stul_color_2(i) = Trim(cbOpt.Text) & Stul_color_2(i)
                i = i + 1
                Next
                cbOpt.List = Stul_color_2
                cbAddNext.Value = False
            ElseIf Len(cbOpt.Text) > 4 And Len(cbOpt.Text) < 12 Then
                i = 0
                For Each tstr In Stul_color_2
                If i > 0 Then Stul_color_2(i) = (cbOpt.Text) & Stul_color_2(i)
                i = i + 1
                Next
                cbOpt.List = Stul_color_2
                cbAddNext.Value = False
            End If
'        Case cStool, cStul
'            Select Case cbOpt.Text
'                Case Stul(5), Stul(6) ' КЛЕО, ТС
'                    cbLength.Text = "белый"
'                    cbLength.Enabled = False
'                    cbAddNext.Value = False
'                Case Stul(8), Stul(9), Stul(10), Stul(14) ' СФИНКС, ЗЕВС, ГЕКТОР, ЦЕЗАРЬ
'                    cbLength.List = SitK
'                    cbLength.Enabled = True
'                    cbAddNext.Value = False
'                Case Stul(11), Stul(12) ' "КОЛИБЕР ХРОМ" "КОЛИБЕР АЛЮМ"
'                    cbAddNext.Value = True
'                Case Stul(0), Stul(1), Stul(2), Stul(3), Stul(4), Stul(13) ' "ВЕНУС" "ФОСКА" "ХАРПО" "ЧИКО" "НЕРОН","МАРКОС"
'                    cbLength.List = SitColors
'                    cbLength.Enabled = True
'                    cbAddNext.Value = False
'                Case Else
'                    cbLength.Clear
'                    cbLength.Enabled = False
'                    cbAddNext.Value = True
'            End Select
        
'        Case cSit
'            Select Case cbOpt.Text
'                Case Sit(0), Sit(1), Sit(2) ' "D390" "D340" "Нерон"
'                    cbLength.List = SitColors
'                    cbLength.Enabled = True
'                Case Sit(3) ' "Колибер"
'                    cbLength.List = SitKolib
'                    cbLength.Enabled = True
'            End Select
'
'        Case "спинка"
'            cbLength.Enabled = True
'            cbAddNext.Value = True
        
'        Case "мойка"
'            cbAddNext.Value = True
            
'        Case "крепление к мойке"
'            cbAddNext.Value = True
        Case "профиль Алюм"
            Select Case cbOpt.Text
                Case "загл мал"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.AddItem "1закр(1AT2GA)"
                cbLength.AddItem "2закр(3AT2GA)"
                cbLength.Text = ""
                Case "загл бол"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.AddItem "1закр(1AT3GA)"
                cbLength.AddItem "2закр(3AT3GA)"
                cbLength.Text = ""
                
                Case "угол 90внут"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.Text = "1закр(1A90B)"
                
                Case "угол 90наруж"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.Text = "1закр(1A90A)"
                
            End Select
       Case "завес Sensys"
'            Select Case cbOpt.Text
'            Case "165"
'                    cbAddNext.Value = True
'            End Select
        Case "завес"
            Select Case cbOpt.Text
                Case "под амортизатор BLUM", "под амортизатор FGV"
                    cbAddNext.Value = True
                Case "HF28", "HF22", "HF25", "HK-S", "полусофт", "под амортизатор"
                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
                Case "HK", "НК", "НK", "HК" ' в разных раскладках! "FGV180"
                    cbFittingName.Text = cbFittingName.Text & " " & "HK27"
                 Case "HK25", "HK27", "HK25 (TIP-ON)", "HK27 (TIP-ON)"
                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
                Case "HL23/35", "HL23/38", "HL25/35", "HL25/38", "HL27/35", "HL27/38", _
                "HL25/39", "HL27/39", "HL29/39", "HL23/39", _
                "HS A", "HS B", "HS D", "HS E", "HS G", "HS H", "HS I"
                
                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
            
            End Select
        
            
        Case "лифт"
'            Select Case cbOpt.Text
'                Case "80"
'                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
'            End Select
            
        Case "сушка"
            If cbOpt.Text = "одноуровневая хром" Then
                cbLength.Clear
                cbLength.AddItem "60"
                cbLength.AddItem "90"
            
            Else
                Select Case cbOpt.Text
                    Case "белая"
                        cbLength.Text = ""
                        cbLength.List = SW_bel
                    Case Else
                        cbLength.Text = ""
                        cbLength.List = SW
                End Select
            End If
            If Not IsEmpty(FittingLength) And Len(FittingLength) > 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
            
        Case "планка в угол", _
                "планка м/у столешницами", _
                "планка к газ. плите"
            Select Case cbOpt.Text
                Case "хром"
                Case Else
                    cbLength.Text = "28"
            End Select
        Case "лиц панель в АНТ вн"
            Select Case cbOpt.Text
                Case "16"
                    cbLength.Enabled = True
                    cbLength.Clear
                    cbLength.AddItem "400(235)"
                    cbLength.AddItem "450(285)"
                    cbLength.AddItem "500(335)"
                    cbLength.AddItem "550(385)"
                    cbLength.AddItem "600(435)"
                    cbLength.AddItem "650(485)"
                    cbLength.AddItem "700(535)"
                    cbLength.AddItem "750(585)"
                    cbLength.AddItem "800(635)"
                    cbLength.AddItem "850(685)"
                    cbLength.AddItem "900(735)"
                    cbLength.AddItem "950(785)"
                    cbLength.AddItem "1000(835)"
                    cbLength.AddItem "1050(885)"
                    cbLength.AddItem "1100(935)"
                    cbLength.AddItem "1150(985)"
                    cbLength.AddItem "1200(1035)"
                Case "18"
                    cbLength.Enabled = True
                    cbLength.Clear
                    cbLength.AddItem "400(231)"
                    cbLength.AddItem "450(281)"
                    cbLength.AddItem "500(331)"
                    cbLength.AddItem "550(381)"
                    cbLength.AddItem "600(431)"
                    cbLength.AddItem "650(481)"
                    cbLength.AddItem "700(531)"
                    cbLength.AddItem "750(581)"
                    cbLength.AddItem "800(631)"
                    cbLength.AddItem "850(681)"
                    cbLength.AddItem "900(731)"
                    cbLength.AddItem "950(781)"
                    cbLength.AddItem "1000(831)"
                    cbLength.AddItem "1050(881)"
                    cbLength.AddItem "1100(931)"
                    cbLength.AddItem "1150(981)"
                    cbLength.AddItem "1200(1031)"
                Case Else
                    cbLength.Text = ""
            End Select
        Case "попер реллинг в АНТ вн"
            Select Case cbOpt.Text
                Case "16"
                    cbLength.Enabled = True
                    cbLength.Clear
                    cbLength.AddItem "400(245)"
                    cbLength.AddItem "450(295)"
                    cbLength.AddItem "500(345)"
                    cbLength.AddItem "550(395)"
                    cbLength.AddItem "600(445)"
                    cbLength.AddItem "650(495)"
                    cbLength.AddItem "700(545)"
                    cbLength.AddItem "750(595)"
                    cbLength.AddItem "800(645)"
                    cbLength.AddItem "850(695)"
                    cbLength.AddItem "900(745)"
                    cbLength.AddItem "950(795)"
                    cbLength.AddItem "1000(845)"
                    cbLength.AddItem "1050(895)"
                    cbLength.AddItem "1100(945)"
                    cbLength.AddItem "1150(995)"
                    cbLength.AddItem "1200(1045)"
                Case "18"
                    cbLength.Enabled = True
                    cbLength.Clear
                    cbLength.AddItem "400(241)"
                    cbLength.AddItem "450(291)"
                    cbLength.AddItem "500(341)"
                    cbLength.AddItem "550(391)"
                    cbLength.AddItem "600(441)"
                    cbLength.AddItem "650(491)"
                    cbLength.AddItem "700(541)"
                    cbLength.AddItem "750(591)"
                    cbLength.AddItem "800(641)"
                    cbLength.AddItem "850(691)"
                    cbLength.AddItem "900(741)"
                    cbLength.AddItem "950(791)"
                    cbLength.AddItem "1000(841)"
                    cbLength.AddItem "1050(891)"
                    cbLength.AddItem "1100(941)"
                    cbLength.AddItem "1150(991)"
                    cbLength.AddItem "1200(1041)"
                Case Else
                    cbLength.Text = ""
            End Select
    
    
    End Select
End Sub

Public Sub AddFitting()
    On Error GoTo err_ДобавитьФурнитуру
    
    ' веберем отгрузку
    'Dim TasksForm As MainForm
    Dim ShipID As Long
    'Set TasksForm = New MainForm
    MainForm.Show
    ShipID = MainForm.ShipID
    
    'Set TasksForm = Nothing
    If ShipID = 0 Then Exit Sub
    Set kitchenPropertyCurrent = New kitchenProperty
    
    Set casepropertyCurrent = Nothing
    
    ' выберем клиента и заказ
    Dim SelectOrder As SelectOrderForm
    Set SelectOrder = New SelectOrderForm
    SelectOrder.ShowForm ShipID
    
    OrderCaseID = 0
    Dim SelectCase As SelectCaseForm
    Set SelectCase = New SelectCaseForm
    If SelectOrder.OrderId > 0 Then
        kitchenPropertyCurrent.OrderId = SelectOrder.OrderId
        SelectCase.ShowForm SelectOrder.OrderId
    End If
    Set SelectCase = Nothing
    
    binit = True
    
    
    
    cbOpt.Text = ""
    cbOpt.Clear
    cbFittingName.Text = ""
    tbQty.Text = ""
    cbAddNext.Value = False
    cbAddNextElement.Value = False
    cbFittingName.Text = ""
    
    binit = False
    
    FormRutin SelectOrder.OrderId
    
    Set SelectOrder = Nothing
    
    If Not rsOrderFittings Is Nothing Then
        rsOrderFittings.UpdateBatch
        MsgBox "Успешно добавлено", vbInformation, "Добавление фурнитуры"
    End If
        
    Exit Sub
    
err_ДобавитьФурнитуру:
    MsgBox Error, vbCritical, "Добавление фурнитуры"
End Sub
   


