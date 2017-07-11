VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddFitting 
   Caption         =   "�������� ���������"
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
        MsgBox "��������� �� ���������." & vbCrLf & "�� ��� �������� ����������", vbExclamation, "���������� ���������"
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
 ' ����� ����������
    GetOtbColors OtbColors
    GetOtbGorbColors OtbGorbColors
    
    ReDim vytyazhka_perfim(12)
    vytyazhka_perfim(0) = "IRIS �������"
    vytyazhka_perfim(1) = "IRIS 60"
    vytyazhka_perfim(2) = "IRIS 90"
    vytyazhka_perfim(3) = "Egizia 60"
    vytyazhka_perfim(4) = "Egizia 90"
    vytyazhka_perfim(5) = "Colalto 60"
    vytyazhka_perfim(6) = "Colalto 90"
    vytyazhka_perfim(7) = "Tirolese 60"
    vytyazhka_perfim(8) = "Tirolese 90"
    vytyazhka_perfim(9) = "Isabella 90"
    vytyazhka_perfim(10) = "Sirius 99 SL(��.Isabella)"
    vytyazhka_perfim(11) = "Sirius 903P-900 SL(��.INN)"
    vytyazhka_perfim(12) = "Sirius 903-700 SL(��.INN)"
    
    
    
    
    ReDim ������������(14) '(11)
    ������������(0) = "����� 25"
    ������������(1) = "����� 30"
    ������������(2) = "����� 35"
    ������������(3) = "����� 40"
    ������������(4) = "����� 45"
    ������������(5) = "����� 50"
    ������������(6) = "����� 25"
    ������������(7) = "����� 30"
    ������������(8) = "����� 35"
    ������������(9) = "����� 40"
    ������������(10) = "����� 45"
    ������������(11) = "����� 50"
    ������������(12) = "����� � ������ 35"
    ������������(13) = "����� � ������ 40"
    ������������(14) = "����� � ������ 50"
    '������������(11) = "quadro � ������."
    
    ReDim ��������������(8)
    ��������������(0) = "3"
    ��������������(1) = "5"
    ��������������(2) = "Sekura 2-1"
    ��������������(3) = "Sekura 8 (��� ������)"
    ��������������(4) = "������� ���� �������"
    ��������������(5) = "������� ���� �����"
    ��������������(6) = "GS-3"
    ��������������(7) = "��� C"
    ��������������(8) = "PP-LUK-00-01"
    '������� ������ ������
    ReDim tbLength(4)
    tbLength(0) = "470"
    tbLength(1) = "420"
    tbLength(2) = "350"
    tbLength(3) = "260"
   
   '������� ������� ������ ������
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
    tbkovrOpt(38) = "��"
    tbkovrOpt(39) = "��"
   
    
    



    ' ���� �������
    ReDim Doormount(44)
    Doormount(0) = "110"
    Doormount(1) = "SlideOn 110"
    Doormount(2) = "+45"
    Doormount(3) = "175"
    Doormount(4) = "-45"
    Doormount(5) = "��������"
    Doormount(6) = "��������"
    Doormount(7) = "��� ����������� BLUM"
    Doormount(8) = "��� ����������� FGV"
    Doormount(9) = "��������������"
    Doormount(10) = "HK-S"
    Doormount(11) = "HF22"
    Doormount(12) = "HF25"
    Doormount(13) = "��������� SK-105"
    Doormount(14) = "+30"
    Doormount(15) = "HETTICH ����"
    Doormount(16) = "+45 �����"
    Doormount(17) = "175 �����"
    Doormount(18) = "���� � ���. ��������"
    Doormount(19) = "+20"
    Doormount(20) = "FGV ����"
    Doormount(21) = "Clip top 120 ��� ��������"
    Doormount(22) = "Clip top �������"
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

    

    ' ���� ������
    ReDim ����(4)
    ����(0) = "50"
    ����(1) = "60"
    ����(2) = "80"
    ����(3) = "100"
    ����(4) = "120"
    
    ' �������
    ReDim StulNogi(2)
    StulNogi(0) = "�����"
    StulNogi(1) = "����� 72"
    StulNogi(2) = "����� 82"
    
'    ReDim ��������(10)
'    ��������(0) = "�����"
'    ��������(1) = "TG(����) � ��������"
'    ��������(2) = "������ (� ������+2 �����)"
'    ��������(3) = "�����"
'    ��������(4) = "�����"
'    ��������(5) = "�������"
'    ��������(6) = "������"
'    ��������(7) = "��������"
'    ��������(8) = "��������"
'    ��������(9) = "�����"
'    ��������(10) = "������ ���"
    
    ReDim ������(14)
    ������(0) = "�����"
    ������(1) = "TG(����) �������"
    ������(2) = "TG(����) ���������"
    ������(3) = "������"
    ������(4) = "�����"
    ������(5) = "����� ���������"
    ������(6) = "�����"
    ������(7) = "�������"
    ������(8) = "������"
    ������(9) = "��������"
    ������(10) = "��������"
    ������(11) = "�����"
    ������(12) = "����"
    ������(13) = "����"
    ������(14) = "�������"
    
'    ReDim �����(0)
'    �����(0) = "������ (5��)"
    
    ' ����� ������
    ReDim Plank(5)
    Plank(0) = "����"
    Plank(1) = "�����"
    Plank(2) = "���"
    Plank(3) = "������"
    Plank(4) = "�������"
    Plank(5) = "�����"
    
    
    ' ����� ���������
    ReDim Galog(1)
    Galog(0) = "����"
    Galog(1) = "������"
    
    ' ������ �����
    ReDim SW(4)
    SW(0) = "50"
    SW(1) = "60"
    SW(2) = "70"
    SW(3) = "80"
    SW(4) = "90"
    
    ReDim SW_bel(2)
    SW_bel(0) = "50"
    SW_bel(2) = "80"
    
    
    ' ���� �����
    ReDim Sushk(3)
    Sushk(0) = "�����"
    Sushk(1) = "����"
    Sushk(2) = "������������� ����"
    Sushk(3) = "�����"
 
    
    ' ������ ������
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
    
    
    ' ������ �����������
    ReDim PA(2)
    PA(0) = "50"
    PA(1) = "60"
    PA(2) = "80"
    
    ' ����� ��������
    ReDim Rell(2)
    Rell(0) = "����"
    Rell(1) = "������"
    Rell(2) = "������"
'    ' �����
'    ReDim Sink(36)
'    Sink(0) = "PIX610"
'    Sink(1) = "BLX710"
'    Sink(2) = "PMN610"
'    Sink(3) = "PMN610 3,5"
'    Sink(4) = "PML610 3,5 �����"
'    Sink(5) = "S45 ����"
'    Sink(6) = "S45 ���"
'    Sink(7) = "SL45 3,5 ����� ����"
'    Sink(8) = "SL45 3,5 ����� ���"
'    Sink(9) = "NORM45 ����"
'    Sink(10) = "NORM45 ���"
'    Sink(11) = "NORM45 ���"
'    Sink(12) = "NORM45 ����� ���"
'    Sink(13) = "NORM45 3,5 ����"
'    Sink(14) = "NORM45 3,5 ���"
'    Sink(15) = "NORM45 3,5 ����� ����"
'    Sink(16) = "NORM45 3,5 ����� ���"
'    Sink(17) = "BLN710-60 ����"
'    Sink(18) = "BLN710-60 ���"
'    Sink(19) = "BLL710-60 ����� ����"
'    Sink(20) = "BLL710-60 ����� ���"
'    Sink(21) = "BLN711 ����"
'    Sink(22) = "BLN711 ���"
'    Sink(23) = "BLL711 ����� ����"
'    Sink(24) = "BLL711 ����� ���"
'    Sink(25) = "COM ���"
'    Sink(26) = "COM ����"
'    Sink(27) = "COM 3,5 ���"
'    Sink(28) = "COM 3,5 ����"
'    Sink(29) = "COL 3,5 ���"
'    Sink(30) = "COL 3,5 ����"
'    Sink(31) = "FAM ���"
'    Sink(32) = "FAM ����"
'    Sink(33) = "FAM 3,5 ���"
'    Sink(34) = "FAM 3,5 ����"
'    Sink(35) = "FAL 3,5 ���"
'    Sink(36) = "FAL 3,5 ����"
    
'    ' �����
'    ReDim Stol(13)
'    Stol(0) = "������ ������"
'    Stol(1) = "�����"
'    Stol(2) = "������"
'    Stol(3) = "����� �����"
'    Stol(4) = "TG �����"
'    Stol(5) = "������� �����"
'    Stol(6) = "������� ������"
'    Stol(7) = "����� �����"
'    Stol(8) = "��������"
'    Stol(9) = "�����"
'    Stol(10) = "��������"
'    Stol(11) = "����"
'    Stol(12) = "����"
'    Stol(13) = "�������"
'
'
'    ReDim ������(0)
'    ������(0) = "�������"
'
'    ReDim ������(0)
'    ������(0) = "������ ��� ����"
    
'    ' ������
'    ReDim Stul(14) ' ������� �� ������!!!! ������������ ��������� �� �������!!!
'    Stul(0) = "�����"
'    Stul(1) = "�����"
'    Stul(2) = "�����"
'    Stul(3) = "����"
'    Stul(4) = "�����"
'    Stul(5) = "���� ������"
'    Stul(6) = "TC �����"
'    Stul(7) = "TC ������"
'    Stul(8) = "������"
'    Stul(9) = "����"
'    Stul(10) = "������"
'    Stul(11) = "������� ����"
'    Stul(12) = "������� ����"
'    Stul(13) = "������"
'    Stul(14) = "������"
'
'    ' ����� ������� � ���. ������� - ������, ����, ������, ������
'    ReDim SitK(4)
'    SitK(0) = "�����������"
'    SitK(1) = "���������"
'    SitK(2) = "�������"
'    SitK(3) = "�������"
'    SitK(4) = "�������"
    
'    ' ����� ������� � �������
'    ReDim SitKolib(2)
'    SitKolib(0) = "�����"
'    SitKolib(1) = "��.�����"
'    SitKolib(2) = "�������"
'
'    ' ����� ������ � �������
'    ReDim BackKolib(2)
'    BackKolib(0) = "������"
'    BackKolib(1) = "���"
'    BackKolib(2) = "�����"
    
'    ReDim Sit(3) ' ������� �� ������!!!! ������������ ��������� �� �������!!!
'    Sit(0) = "D390 (�,�)"
'    Sit(1) = "D340 (�,�)"
'    Sit(2) = "�����"
'    Sit(3) = "�������"
    
'    ' ������� ����� (� ������, �����, �����, ����, �����)
'    ReDim SitColors(9)
'    SitColors(0) = "�����"
'    SitColors(1) = "��.-����������"
'    SitColors(2) = "������"
'    SitColors(3) = "��.-������"
'    SitColors(4) = "�����"
'    SitColors(5) = "������"
'    SitColors(6) = "�����-�����"
'    SitColors(7) = "����� ��������"
'    SitColors(8) = "�������"
'    SitColors(9) = "�����"
    
    ' ����� ��������
    GetBibbColors ��������
    ' ����� ��������
    GetCamBibbColors ��������

    ' ����� �������
    GetHangColors �������
    
    ' �����
    ReDim �����(19)
    �����(0) = "15���"
    �����(1) = "15����"
    �����(2) = "20���"
    �����(3) = "20����"
    �����(4) = "30"
    �����(5) = "40"
    �����(6) = "45"
    �����(7) = "50"
    �����(8) = "����"
    �����(9) = "���� VIBO"
    �����(10) = "�����"
    �����(11) = "����� VIBO"
    �����(12) = "���� 45 ���"
    �����(13) = "���� 45 ����"
    �����(14) = "����� 40"
    �����(15) = "����� 50"
    �����(16) = "������� 40"
    �����(17) = "� ������� ���� ���� VIBO"
    �����(18) = "������� � ���� ���� VIBO"
    �����(19) = "������� � ����� ���� VIBO"
    
'    �����(22) = "VS-������� ��� ����� 80"
'    �����(23) = "VS-������� 60 �����."
'    �����(24) = "VS-������� 80 �����."
'    �����(25) = "VS-����� ����. 45"
'    �����(26) = "VS-����� ����. 60"
'    �����(27) = "VS-����. ���. 4/4"
'    �����(28) = "VS-����. ���. 4/4 ��-�"
'    �����(29) = "VS-����� 30"
'    �����(30) = "VS-������. � ����� 30"
    
    ReDim MoikaColors(7)
    
   MoikaColors(0) = "���"
    MoikaColors(1) = "������"
    MoikaColors(2) = "��������"
    MoikaColors(3) = "����"
    MoikaColors(4) = "����"
    MoikaColors(5) = "����� ������"
    MoikaColors(6) = "����"
    MoikaColors(7) = "������ ���"


    
    ' �����
    ReDim �����(7)
    �����(0) = "��������� 1/2"
    �����(1) = "��������� 1/2 VIBO"
    �����(2) = "��������� 3/4"
    �����(3) = "��������� 3/4 VIBO"
    �����(4) = "�-08"
    �����(5) = "�-11"
    �����(6) = "�-12"
    �����(7) = "�-37"
    
    ReDim �������(96)
    �������(0) = "����������� ������"
    �������(1) = "������������ ���������"
    �������(2) = "������"
    �������(3) = "�������"
    �������(4) = "��� ������ ������"
    �������(5) = "��� ������ �������"
    �������(6) = "������� ��������"
    �������(7) = "����� ������"
    �������(8) = "��� �������"
    �������(9) = "�������� ���������"
    �������(10) = "�����"
    �������(11) = "�����"
    �������(12) = "��� ���������"
    �������(13) = "������ ������"
    �������(14) = "������� ������"
    �������(15) = "������� �������"
    �������(16) = "������"
    �������(17) = "���������"
    �������(18) = "����"
    �������(19) = "������"
    �������(20) = "�������"
    �������(21) = "���������"
    �������(22) = "������"
    �������(23) = "������ ������"
    �������(24) = "������� ��������"
    �������(25) = "�������� ������"
    �������(26) = "�������� �������"
    �������(27) = "������"
    �������(28) = "�������� ������"
    �������(29) = "�������� �������"
    �������(30) = "����� �����"
    �������(31) = "����� ������"
    �������(32) = "������ ������"
    �������(33) = "����������� ������"
    �������(34) = "���� ������"
    �������(35) = "������"
    �������(36) = "������ ������"
    �������(37) = "������ �������"
    �������(38) = "������� ������"
    �������(39) = "������ ������"
    �������(40) = "������ ����������"
    �������(41) = "������ ������"
    �������(42) = "�����"
    �������(43) = "��������"
    �������(44) = "�������"
    �������(45) = "�������"
    �������(46) = "��������� ������"
    �������(47) = "��������� ������"
    �������(48) = "��������� �������"
    �������(49) = "����� ������"
    �������(50) = "����������� ���������"
    �������(51) = "����� ������"
    �������(52) = "����� ������"
    �������(53) = "����� �������"
    �������(54) = "��������"
    �������(55) = "������ ������"
    �������(56) = "��������� ������"
    �������(57) = "���������� ������"
    �������(58) = "���������� �������"
    �������(59) = "������ ������"
    �������(60) = "����� ������"
    �������(61) = "����� �������"
    �������(62) = "����"
    �������(63) = "������� ����"
    �������(64) = "����� ����"
    �������(65) = "����� ����"
    �������(66) = "�����"
    �������(67) = "�������� ���������"
    �������(68) = "�����"
    �������(69) = "�����"
    �������(70) = "��������"
    �������(71) = "����������"
    �������(72) = "�������"
    �������(73) = "����"
    �������(74) = "������"
    �������(75) = "������" '���
    �������(76) = "������� ���" '���
    �������(77) = "���" '���
    �������(78) = "�������� ����"
    �������(79) = "������"
    �������(80) = "���� ����"
    �������(81) = "���� ����"
    �������(82) = "���� �����"
    �������(83) = "������"
    �������(84) = "��������"
    �������(85) = "�������"
    �������(86) = "������"
    
    �������(87) = "�������� �����"
    �������(88) = "������� �������"
    �������(89) = "������� ��������"
    �������(90) = "��� ������� �����"
    �������(91) = "��� ������ ����������"
    �������(92) = "��� ����� �������"
    �������(93) = "����� �������"
    �������(94) = "�������� ��������"
    �������(95) = "������ ���������� ������-�����"
    �������(96) = "������ ����� �������"
    
    
    SortArray �������
    
    
    ReDim ������(22)
    ������(0) = "����100"
    ������(1) = "����150"
    ������(2) = "�������100"
    ������(3) = "�������150"
    ������(4) = "��������100"
    ������(5) = "��������150"
    ������(6) = "���100"
    ������(7) = "�����100"
    ������(8) = "�����150"
    ������(9) = "�����100"
    ������(10) = "����100"
    ������(11) = "����150"
    ������(12) = "�����100"
    ������(13) = "����100"
    ������(14) = "����150"
    ������(15) = "������100"
    ������(16) = "����100"
    ������(17) = "����150"
    ������(18) = "ר������100"
    ������(19) = "ר������150"
    ������(20) = "������100"
    ������(21) = "�����100"
    ������(22) = "����������100"
    
    
    
    
    
    ReDim �����������������(19)
    �����������������(0) = "����100"
    �����������������(1) = "����150"
    �����������������(2) = "���100"
    �����������������(3) = "�����100"
    �����������������(4) = "�����150"
    �����������������(5) = "�����100"
    �����������������(6) = "����100"
    �����������������(7) = "����150"
    �����������������(8) = "�����100"
    �����������������(9) = "������100"
    �����������������(10) = "����100"
    �����������������(11) = "����150"
    �����������������(12) = "�������100"
    �����������������(13) = "�������150"
    �����������������(14) = "ר������100"
    �����������������(15) = "ר������150"
    �����������������(16) = "����100"
    �����������������(17) = "����150"
    �����������������(18) = "�����100"
    �����������������(19) = "����������100"

    
    ReDim ���������(2)
    ���������(0) = "���"
   ' ���������(1) = "GTV"
   ' ���������(2) = "FBV"
    
    ' �� ������ �������
    ReDim ��������������(9)
    ��������������(0) = "���"
    ��������������(1) = "�-���"
    ��������������(2) = "�����"
    ��������������(3) = "�-���"
    ��������������(4) = "�-���"
    ��������������(5) = "����"
    ��������������(6) = "���"
    ��������������(7) = "����"
    ��������������(8) = "�-���"
    ��������������(9) = "�����"
    
    
    
    ReDim ���������4�(6)
    ���������4�(0) = "���"
    ���������4�(1) = "���"
    ���������4�(2) = "���"
    ���������4�(3) = "����"
    ���������4�(4) = "���"
    ���������4�(5) = "���"
    ���������4�(6) = "���"
    SortArray ���������4�
    
    
    ReDim TOPLine(2)
    TOPLine(0) = "��������"
    TOPLine(1) = "������"
    TOPLine(2) = "�����"
    
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
    zavesSensys(4) = "���������-�"
    zavesSensys(5) = "110 ��� AL ����"
    zavesSensys(6) = "����"
    zavesSensys(7) = "��������"
    zavesSensys(8) = "��������"
    
    ReDim zavesClipTop(8)
    zavesClipTop(0) = "BLUMOTION +110"
    zavesClipTop(1) = "BLUMOTION ��������"
    zavesClipTop(2) = "BLUMOTION +45"
    zavesClipTop(3) = "BLUMOTION -45"
    zavesClipTop(4) = "+155"
    zavesClipTop(5) = "BLUMOTION ��������������"
    zavesClipTop(6) = "BLUMOTION +90 ��� ��"
    zavesClipTop(7) = "110 ��� �������"
    zavesClipTop(8) = "�������� ��� �������"
    
    
    
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
Stul_color_1(1) = " ���."
Stul_color_1(2) = " �.���."
Stul_color_1(3) = " �.���."
Stul_color_1(4) = " "

ReDim Stul_color_2(9)
Stul_color_2(1) = " (�.����� �-��)"
Stul_color_2(2) = " (�.�����)"
Stul_color_2(3) = " (�.�����)"
Stul_color_2(4) = " (�.���� ����)"
Stul_color_2(5) = " (�.���. �-��)"
Stul_color_2(6) = " (�.������)"
Stul_color_2(7) = " (����� 1000)"
Stul_color_2(8) = " (������ 114)"
Stul_color_2(9) = " (�.����.)"



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
    Me.Caption = "�������� ���������"
    If Not kitchenPropertyCurrent Is Nothing Then
        If kitchenPropertyCurrent.dspColor <> "" Then
            Me.Caption = Me.Caption & " �:" & kitchenPropertyCurrent.dspColor
        End If
    End If
    '�������
    If name = cHandle Then name = "�����"
    
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
        MsgBox "������!!! ����������� ���������", vbCritical
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
            Case cNogi, "���������� 3�", "���������� 4�", "���������� ����-4", "���������� ����-5", "���������� TOP-Line", "������� � ����������", "�������� ��� ������", "�������"
                bSpecified = False
            Case "�-�� �������� �����"
            cbOpt.Text = Opt
            cbLength.Text = length
'            Case "�����", "����� Sensys"
'                If IsMissing(caseID) Then
'                cbAddNext.Value = True
'                bSpecified = False
'                End If
            Case "�������� ������ Sensys"
            cbOpt.Text = Opt
            bSpecified = True
        End Select
        
        '���� ��� �������� ����������, �� ����� ���������� �� �����
        If bSpecified Then
            result = True
        Else
            
            If cbAddNext.Value Then
                Select Case cbFittingName.Text
                    Case cNogi, "���������� 3�", "���������� 4�", "���������� ����-4", "���������� ����-5", "���������� TOP-Line", "������ DU325 Rapid S", "��� ������ � ��� ��" ', "�����", "����� Sensys"
                    Case Else
                        cbAddNext.Value = False
                End Select
                
                Me.Show 1
                
                Select Case cbFittingName.Text
                     Case "������� � ����������"
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
            result = True '���� ������ ������
        End If
            
         If cbAddNext.Value Then
         
             Select Case cbFittingName.Text
                Case "������ DU325 Rapid S"
                    cbFittingName.Text = "���������� 18��"
                    
                Case "���������� ������� 4� ��"
                    cbFittingName.Text = "���+���� � ��� �������"
                    
'                Case "�������"
'                    If cbOpt.Text = "CAMAR �+�" Then
'                        cbFittingName.Text = "�������� CAMAR  �+�"
'
'                    End If
               
'               Case "����� Sensys"
'                    cbFittingName.Enabled = True
'                    cbFittingName.Text = "�������� ��� ������"
               Case "������ CLIP TOP +155"
                cbAddNext.Value = False
               Case "����� CLIP top"
                    If cbOpt.Text = "+155" Then
                        cbFittingName.Text = "������ CLIP TOP +155"
                        If tbQty.Text = "2" Then tbQty.Text = "1"
                        
                    Else
                        tbQty.Text = ""
                        cbFittingName.Enabled = True
                        cbFittingName.Text = ""
                    End If
                    
               'Case "������. Sensys 165"
               '     If cbOpt.Text = "165" Then
               '        cbFittingName.Text = "�������� ���.Sensys"
               'End If
                   
'                Case cStul, cStool, "������" ' ����� ����� ������� �������
'
'                    If cbOpt.Text = Stul(11) Or cbOpt.Text = Stul(12) Then     ' "������� ����"  "������� ����"
'                        cbFittingName.Text = "������"
'                    Else
'                        cbFittingName.Text = cSit
'                    End If
                    
                    
'                    For i = 0 To cbOpt.ListCount - 1
'                        If InStr(1, cbFittingName.List(i), cSit, vbTextCompare) Then
'                            cbFittingName.Text = cbFittingName.List(i)
'                            Exit For
'                        End If
'                    Next

'            Case "�����"
'                Krepl = cbOpt.Text
'                cbFittingName.Text = "��������� � �����"
'                bSpecified = True
'
'            Case "��������� � �����"
'                cbFittingName.Text = "����� � �����"
                
            Case "���������� 4�"
            
                cbFittingName.Text = "������� � ����������"
                Dim tqty As Single
                tqty = CDec(tbQty.Text) * 4 / 3
                If tqty > Round(tqty) Then
                    tbQty.Text = Round(tqty) + 1
                Else
                    tbQty.Text = Round(tqty)
                End If
            
            Case "�����"
                Select Case cbOpt.Text
                    Case "��� ����������� BLUM"
                        cbFittingName.Text = "����������� BLUM"
                    Case "��� ����������� FGV"
                        cbFittingName.Text = "����������� FGV"
                    Case Else
                        'tbQty.Text = ""
                        cbFittingName.Enabled = True
                        cbFittingName.Text = ""
                End Select
                
             Case "����� ��� �������-� BLUM"
                cbFittingName.Text = "����������� BLUM"
              cbLength.Enabled = False
                cbOpt.Enabled = False
            Case "������ VB15"
                cbFittingName.Text = "���� ������ 200��"
              cbLength.Enabled = False
                cbOpt.Enabled = False
             Case "����� ��� ����������� FGV"
                cbFittingName.Text = "����������� FGV"
                 cbLength.Enabled = False
             Case "����� HL23/35", "����� HL23/38", "����� HL25/35", "����� HL25/38", "����� HL27/35", "����� HL27/38", _
             "����� HL23/39", "����� HL25/39", "����� HL27/39", "����� HL29/39"
                cbLength.Enabled = False
                cbOpt.Enabled = False
                cbFittingName.Text = "������ HL ��������"
              Case "�������������", "������-� LED 30w"
              cbLength.Enabled = False
              cbOpt.Enabled = False
              cbFittingName.Text = "������+�����+���� 220V"
              Case "����� HS I", "����� HS A", "����� HS B", "����� HS D", "����� HS E", "����� HS G", "����� HS H", "����� HS F"
                cbLength.Enabled = False
                cbOpt.Enabled = False
                cbFittingName.Text = "������ HS �������"
                
'            Case "������� 60", "������� 100"
'                cbFittingName.Text = "���������� � ��������"
'                cbOpt.Enabled = False
'                If CDec(tbQty.Text) > 1 Then tbQty.Text = CDec(tbQty.Text) - 1
'            Case "�� ���500/C ��� ����", "�� ���500/M ��� �����", "�� ���500/D ��� �����"
'                cbFittingName.Text = "��� ������ � ��� ��"
'                cbLength.Enabled = True
'                If Not IsMissing(Opt) Then cbOpt.Text = Opt
            Case "�� ��� ��� ����"
                
                cbFittingName.Text = "��� ������ � ��� ��"
                cbLength.Enabled = True
                If Not IsMissing(Opt) Then If Not IsEmpty(Opt) Then cbOpt.Text = Opt
            Case "��� ������ � ��� ��"
                cbFittingName.Text = "����� ������� � ��� ��"
                If Not IsMissing(Opt) Then If Not IsEmpty(Opt) Then cbOpt.Text = Opt
                If Not IsMissing(length) Then If Not IsEmpty(length) Then cbLength.Text = length
                
             Case "�� ������� �����"
                
                cbFittingName.Text = "��� ������ ������� �����"
                'cbLength.Enabled = True
                If Not IsMissing(Opt) Then cbOpt.Text = Opt
            Case "��� ������ ������� �����"
                cbFittingName.Text = "����� ��� ������� �����"
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
        MsgBox "����������� ��� ��������� '" & cbFittingName.Text & "'", vbCritical
        Exit Function
    End If
    
    If Len(cbOpt.Text) > 0 Then
        Opt = Trim(cbOpt.Text)
    End If
    If Len(cbLength.Text) > 0 Then
        If InStr(1, Opt, "��", vbTextCompare) = 1 And InStr(1, cbLength.Text, "������", vbTextCompare) = 1 Then
            Opt = LTrim(Opt & " " & Replace(Replace(cbLength.Text, "�����", "."), " ", ""))
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
        
        Cells(ActiveCell.row, 10).Value = t & "; " & "�=" & cbFittingName.Text & RTrim(" " & cbOpt.Text) & ", QTY=" & tbQty.Text & ", L=" & cbLength.Text
    Else
        Cells(ActiveCell.row, 10).Value = "�=" & cbFittingName.Text & RTrim(" " & cbOpt.Text) & ", QTY=" & tbQty.Text & ", L=" & cbLength.Text
    End If
    
    
    AddFitting2Order = True
    Application.Cursor = xlDefault
    Exit Function
err_AddFitting2Order:
    MsgBox "������ ���������� ���������", vbCritical
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
     
         

        Case "�������"
            cbOpt.Text = "PERFIM"
            cbLength.List = vytyazhka_perfim
        Case ""
        Case "�����", "�����", _
             "�����", _
             "��������� 3", "��������� 5", _
             "������ � ����", "������ �/� ������������", _
             "������ � ���. �����", _
             "����", _
             "������� 60", "������� 100", "�������� � ��������", "��������� � ��������", "������ � ��������", "����-90 � ��������", "����-120 � ��������", "����-135 � ��������", "���������� � ��������", _
             "����", _
             "�������", _
             "��������", _
             "�����", _
             "�����", _
             "�������", "������ � �����", "������ ��� ����", _
             "������ �������", "�������� � ������", "����90* � ������", "����135* � ������", _
             "����� � p����", "���� � p����", "������", "�����", "������", _
             "��������������", "����������� ������", "������ ��� ����� 807"
             
            
               ' cbOpt.Enabled = False
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            
            Dim bSkip As Boolean
            bSkip = False
            
            ' ����
            cbOpt.Text = ""
            Select Case cbFittingName.Text
                Case "��������������"
               
                    cbOpt.List = ��������������
            
'                Case "������"
'                    cbOpt.List = Stul
            
                Case "�����"
                    cbOpt.List = HandleArray
                
                Case "�����"
                 
                    cbOpt.List = LegArray
                    
                
                Case "����� � p����", "���� � p����"
                    cbOpt.AddItem "22"
                    cbOpt.AddItem "25"
                    cbOpt.AddItem "28"
                    cbOpt.AddItem "35"
                    cbOpt.AddItem "40"
                    
                Case "�������"
                    cbOpt.AddItem "�������"
                    bSkip = True
                
                Case "������ �������", _
                     "�������� � ������", _
                     "����90* � ������", _
                     "����135* � ������", _
                     "����������� ������"
                     cbOpt.List = ������
                     'cbAddNext.Value = True
'                Case "����������� ������"
'                     cbOpt.List = ������
'                     cbLength.Enabled = False
                     
                Case "������ � ����", _
                        "������ ���������", _
                        "������ �/� ������������", _
                        "������ � ���. �����"
                       
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
                        
                                cbOpt.Text = "����"
                                FittingOption = "����"
                                cbLength.Text = "38"
                         
                         Else
                         cbOpt.Enabled = True
                                cbOpt.List = Plank
                                
                                cbLength.Text = ""
                                cbLength.Enabled = True
                                cbLength.AddItem "28"
                                cbLength.AddItem "38"
                        
                         End If
                         
                    
                Case "��������� 3", "��������� 5"
                    cbOpt.List = Galog
                ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False

                
                Case "�����"
                ' ����� �� �����
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
                    
                
                Case "������ � �����", "������ ��� ����"
                
                cbOpt.Enabled = False
                Case "����"
                    cbOpt.List = ����
'                Case "�����"
'                    cbOpt.List = Sink
'                    bSkip = True
'                    'cbAddNext.Value = True
                    
                Case "�����"
                    cbOpt.List = �����
                    bSkip = True
                    
                Case "�����"
                    cbOpt.List = �����
                    bSkip = True
                                
'                Case cStool, cStul
'                    cbOpt.List = Stul
'                    cbAddNext.Value = True
'                    'cbAddNext.Enabled = False
'                    bSkip = True
                    
'                Case "���� ���������"
'                    cbOpt.List = Stol
'                    bSkip = True
                
                Case "������"
                    cbOpt.List = ������
                    bSkip = True
'                Case "�����"
'                    cbOpt.List = �����
'                    cbOpt.Text = �����(0)
'                    bSkip = True
'                Case "������"
'                    cbOpt.List = ������
'                    cbOpt.Text = ������(0)
'                    bSkip = True
'                Case "���� � ������"
'                    cbOpt.List = ��������
'                    bSkip = True
                    
                Case "����"
                ' ����� �� �����
            cbLength.Enabled = False
                    cbOpt.List = StulNogi
                    bSkip = True
                    If IsNumeric(tbQty.Text) Then
                        If CInt(tbQty.Text) Mod 4 <> 0 Then
                            MsgBox "��������� ������������ ���-�� ���!", vbExclamation
                        End If
                    End If
                    
                    
                Case "������� 60", "������� 100", _
                    "�������� � ��������", _
                    "��������� � ��������", _
                    "������ � ��������", _
                    "����-90 � ��������", _
                    "����-120 � ��������", _
                    "����-135 � ��������"
                   ' ����� �� �����
            cbLength.Enabled = False
                    
                    cbOpt.List = Rell
                
'                Case cSit
'                    cbOpt.List = Sit
'                    bSkip = True
                
                Case "�������"
                    cbOpt.List = �������
                ' ����� �� �����
            cbLength.Enabled = False
                Case "��������"
                    cbOpt.List = ��������
                ' ����� �� �����
            cbLength.Enabled = False
                Case Else
                    cbOpt.Clear
            ' ����� �� �����
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
        
        
        Case "������ ��� ��������� 2�"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
         Case "��������� � ������", "�����. ������ ���������", "������ ��� ����������", "�������� ������ ��� �����"
          cbOpt.Enabled = False
            cbLength.Enabled = False
            cbAddNext.Value = False
            
        Case "������ �����������"
        cbOpt.AddItem "��������"
        cbOpt.AddItem "�������"
        cbOpt.AddItem "�������"
        If Not IsEmpty(FittingOption) Then
            If FittingOption = "�������" Then cbOpt.Text = FittingOption
        End If
        cbLength.Enabled = False
        Case "��������� SK-105"
        cbOpt.Enabled = True
        cbLength.Enabled = False
        Case "��������� � ������"
        cbOpt.AddItem "�����"
        cbOpt.AddItem "����"
        cbOpt.AddItem "����� � ����� d-25"
        cbOpt.Enabled = True
        cbLength.Enabled = False
        
        Case "������ �������"
            cbOpt.AddItem "�����"
            cbOpt.AddItem "����"
            cbOpt.Text = "����"
            cbLength.Enabled = False
        
        Case "������� � ������"
        cbOpt.AddItem "16"
        cbOpt.AddItem "18"
        cbLength.Enabled = False
        Case "������ ��������� 100��", "������ ���������", _
             "�������� ����", _
             "�������������", "������������� HT", "������������� ���.", _
             "������", "���������", "�������", "��������", "�����", "������� ��� �����. ����.", "������� ��� ����. ����.", _
             "����������", "���� � �����������", "�������", "�-� � ������", "������", "������ ����������", _
             "������������ �������", "������������ ������", _
             "�����", "����� ���������(��5)", "��������� ���. ��� ����.", _
             "���������", "�������������� 5", "���������������", "������ �6", "������ �����������", "������ ���. �����", "������ ���.", "����", _
             "����� 3*30", "����� 3,5*16", "����� 4*16", "����� 5*30", "�����", "������", "���� ����� 82", _
             "��������� � ���� ��. ���.", "��������� � ����� ���.", "��������� � �������", _
             "����� ��������", "����� ����", "�������������� Secura 8", _
             "������ �������", "������� ���. ������� ", "������ ���. ��������", "������� Gold", "������� White/Gold", "������� L ����", "������� L ������", _
             "���� ���. � ������", "��������� � �����", "�-������� ��������(��)", "��������� � �� 110", "��������� � �� 010", _
              "��������� � ������", "����� �����������", "���� � ����� ������", "���� � ����� ������", "����� � �������� �� ����", "���������� � ��������", _
             "������ ��� ����������", "����������� BLUM", "����������� FGV", "����������� �������", "����� 4*20", "����� �41", _
              "������ �����", "������ 08", "����� ��������", "������� ������� 10��", _
             "�������� �� ��������", "������ ��� ��������", "���������� �����. �����", "������� � �������� 3�", _
             "������ DU860 ������� ����", "������ DU868 ������� 16��", "����� ��������", "�������� ������ ��� RV8", "������ DU321", "������ VB 36 HT", _
             "�������� CAMAR �+�", "������� CAMAR �+�", "������ VB 35D/16", _
             "�������", "�-�� Push-To-Open Magnet", "�������� ����� SAH-5 �/�", "�������������� ��� C", "������� CAMAR 806���.", "������� CAMAR 806����.", _
             "������� 807", "������� 808", "���������� ��� �����", "������� � �� ��� �����", _
             "�/� �������� XXL", "������������� ������", "������ ��� ��������� 2�", "��������� ���. ����������", "��������� � ���. �� L+R", "������ SISO �+�", _
             "���� ������� ��� �����", "������-�����", "����� � ���� �� �����"
             '/*"����� Ecomat", "����� FGV180", _*/
          cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "��������� ������������"
            cbOpt.Text = "����� M8 L-30"
            cbOpt.AddItem "����� M8 L-30"
            cbLength.Enabled = False
            cbLength.Text = ""
        Case "�������� �����", "�������� �������"
            cbOpt.Enabled = False
            cbLength.Text = ""
            cbLength.Enabled = True
            cbLength.AddItem "500"
            cbLength.AddItem "450"
            cbLength.AddItem "������500"
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
        
        Case "����� ���� ���� 81G19A10", "������ ��� ���� �����"
          cbOpt.Enabled = False
            cbLength.Enabled = False
        Case "ServoDrive"
            cbOpt.AddItem "���� �������"
            cbOpt.AddItem "������(��)"
            cbOpt.AddItem "�� 1 ����"
            cbOpt.AddItem "�� 2 ����"
            cbOpt.AddItem "�� 3 ����"
            cbOpt.AddItem "�� 4 ����"
            cbOpt.AddItem "UNO"
            cbOpt.Text = ""
            cbLength.Enabled = False
        
        Case "ServoDrive ��"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "ServoDrive ������"
            cbOpt.Enabled = False
            cbLength.Enabled = True
            
        Case "Komandor"
            cbOpt.AddItem "������ �/����� �����"
            cbOpt.AddItem "���� AGAT ���� ALU"
            cbOpt.AddItem "���� ����� ����+���"
            cbOpt.AddItem "���� ����� ��� ����"
            cbOpt.AddItem "���� ����� ��� ����"
            cbOpt.AddItem "����� �����"
            cbOpt.AddItem "����� ����"
            cbOpt.AddItem "������ ������ �����"
            cbOpt.AddItem "�������� H-4 �����"
            cbOpt.AddItem "����� ���� ������"
            cbOpt.AddItem "����� ������ ������"

            cbOpt.Enabled = True
            cbLength.Enabled = True
        Case "Hafele SD"
            cbOpt.AddItem "������� ���� �������"
            cbOpt.AddItem "������� �����"
            cbOpt.AddItem "�������� ���� ��� ����"
            cbOpt.AddItem "������� ��� ���� ����"
            cbOpt.AddItem "������� ���� ���� ����"

            cbOpt.Enabled = True
            cbLength.Enabled = True
        Case "Astin", "���", "ArciTech"
            cbOpt.Enabled = True
            cbLength.Enabled = True
            
        Case "����� HK-S (TIP-ON)", "����� HK27 (TIP-ON)", "����� HK25", "����� HK25 (TIP-ON)", "����� HK29 (TIP-ON)", "����� HK29", "����� HK27", "����� HK-S", "����� HF22", "����� HF25", "����� HF28"
          'cbOpt.Enabled = False
           cbOpt.AddItem "����"
           cbOpt.AddItem "�����"
           cbOpt.AddItem "����� ����� �����"
           cbOpt.Text = "����"
            cbLength.Enabled = False
             
            
            Case "�-� ����� ��� �/�"
                cbOpt.Enabled = False
                cbLength.Enabled = False

            Case "����� HK-XS"
                cbOpt.Enabled = False
                cbLength.Enabled = False
            
          Case "����� HL23/35", "����� HL23/38", "����� HL25/35", "����� HL25/38", "����� HL27/35", "����� HL27/38", _
          "����� HL23/39", "����� HL25/39", "����� HL27/39", "����� HL29/39", _
           "����� HS I", "����� HS A", "����� HS B", "����� HS D", "����� HS E", "����� HS G", "����� HS H", "����� HS F"
            'cbOpt.Enabled = False
           cbOpt.AddItem "����"
           cbOpt.AddItem "�����"
           cbOpt.AddItem "����� ����� �����"
           'cbOpt.AddItem "���� ����� �����"

           cbOpt.Text = "����"
            cbLength.Enabled = False
            cbAddNext.Value = True
        Case "������ ���+��.����"
            cbLength.Enabled = False
            cbOpt.Enabled = False
            cbAddNext.Value = False
        Case "��������� ���� ���"
            cbLength.Enabled = False
            cbOpt.Enabled = False
            cbAddNext.Value = False
            
'         Case "�� ���500/C ��� �����", "�� ���500/C ���", "�� ���500/D ���", "�� ���500/M ���", "�� ���500/N ���", "�� ���500/D ��� ��� �����"
'          cbOpt.Enabled = False'            cbLength.Enabled = False
        Case "VS - VSA"
            cbOpt.AddItem "5 - 600"
            cbOpt.Text = "5 - 600"
            cbLength.Enabled = False
            
        Case "VS - ���������� ����+���", "VS - ���������� ����", "VS - ���������� ���"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "VS - ��� ���� ���+����", "VS - ��� ���� ���", "VS - ��� ���� ����"
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
            cbOpt.AddItem "3 - 150 � ���.����."
            cbOpt.AddItem "3 - 200"
            cbOpt.AddItem "3 - 300"
            cbOpt.AddItem "3 - 400"
            cbOpt.AddItem "8 - 150"
            cbOpt.AddItem "8 - 200"
            cbLength.Enabled = False
            
        Case "VS - ������� ��������"
            cbOpt.AddItem "600-3/4"
            cbOpt.AddItem "4/4"
            
            cbLength.Enabled = False
        
        Case "VS - ������ ��������"
            cbOpt.AddItem "900-3/4"
            cbOpt.AddItem "90*90-4/4"
            cbLength.Enabled = False
            
        Case "VS - Wari Corner"
            cbOpt.AddItem "900 - L"
            cbOpt.AddItem "900 - R"
            cbOpt.AddItem "1000 - L"
            cbOpt.AddItem "1000 - R"
            cbLength.Enabled = False

         Case "VS - ����� ���������"
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
            
        Case "VS - ������� Eco Liner"
         cbOpt.AddItem "450"
            cbOpt.AddItem "600"
            cbOpt.Text = "600"
            cbLength.Enabled = False
        Case "VS - ����� ������"
            cbOpt.AddItem "16��"
            cbOpt.Text = "16��"
            cbLength.Enabled = False
            
            
        Case "VS - Eco flex liner"
            cbOpt.AddItem "���� 450 2�"
            cbOpt.AddItem "���� 600 3�"
            cbOpt.AddItem "���� 600 2�"
            
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
            
        Case "VS - ��������� �������"
            cbOpt.AddItem "450 ��� ��.��. ��� �.��."
            cbOpt.AddItem "600 ��� ��.��. ��� �.��."
            cbOpt.AddItem "900 ��� ��.��. ��� �.��."
            cbOpt.AddItem "450 �� ���� �� � �������"
            cbOpt.AddItem "600 �� ���� �� � �������"
            cbOpt.AddItem "900 �� ���� �� � �������"
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
        Case "VS - ����� ����� ��� ���"
            cbLength.Enabled = False
            cbOpt.Enabled = False
            
        Case "VS - ������� ��� �����"
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
        Case "VS - �������� �����"
            cbOpt.AddItem "600"
            cbOpt.AddItem "900"
            cbLength.Enabled = False
                
        Case "�� ��� ���"
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
        Case "�� ��� ��� ��� �����"
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
         Case "�� ��� ��� ����"
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
        Case "��� ������ � ��� ��"
            cbOpt.AddItem "16"
            cbOpt.AddItem "18"
            cbLength.Text = ""
            cbAddNext.Value = True
           
           
         Case "�� ������� �����"
            cbOpt.AddItem "�����"
            cbOpt.AddItem "��������"
            If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
               For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
             
            cbLength.AddItem "500/94 ���"
            cbLength.AddItem "500/186 1����"
            cbLength.AddItem "500/186 ������"
            cbLength.AddItem "300/94 ���"
            cbLength.AddItem "300/186 1����"
           
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
            
            
             Case "�� �������"
             cbOpt.AddItem "�����"
             cbOpt.AddItem "��������"
             If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
             
            cbLength.AddItem "500/78 ����"
            cbLength.AddItem "500/94 ���"
            cbLength.AddItem "500/186 1����"
            cbLength.AddItem "500/186 ������"
            cbLength.AddItem "500/250 2����"
            cbLength.AddItem "300/94 ���"
            cbLength.AddItem "300/186 1����"
            cbLength.AddItem "300/250 2����"
            
            If Not IsEmpty(FittingLength) And Len(FittingLength) >= 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
            
            
             Case "����� � �������"
             cbOpt.AddItem "�����"
             cbOpt.AddItem "��������"
              If Not IsEmpty(FittingOption) And Len(FittingOption) >= 1 Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
            End If
            
             cbLength.AddItem "60��"
             cbLength.AddItem "90��"
            If Not IsEmpty(FittingLength) And Len(FittingLength) >= 1 Then
                For i = 0 To cbLength.ListCount - 1
                    If InStr(1, FittingLength, cbLength.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbLength.List(i), FittingLength, vbTextCompare) > 0 Then
                        cbLength.Text = cbLength.List(i)
                        Exit For
                    End If
                Next
            End If
                  
             Case "����� ORGALINE"
            cbOpt.Enabled = False
            cbLength.AddItem "40"
            cbLength.AddItem "45"
            cbLength.AddItem "50"
            cbLength.AddItem "60"
            cbLength.AddItem "80"
            cbLength.AddItem "90"
           
                  
            Case "��� ������ ������� �����"
'                cbOpt.AddItem "16"
'                cbOpt.AddItem "18"
                cbOpt.AddItem "�����"
                cbOpt.AddItem "��������"
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
             Case "����� ��� ������� �����"
                 cbOpt.AddItem "�����"
                cbOpt.AddItem "��������"
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
                
        Case "������� �� ����������"
            cbOpt.AddItem "470"
            cbOpt.AddItem "420"
            cbOpt.AddItem "350"
            cbOpt.AddItem "260"
            cbLength.Enabled = False
        Case "������� �� ��������"
            cbOpt.AddItem "500"
            cbOpt.AddItem "450"
            cbOpt.AddItem "������500"
            cbLength.Enabled = False
          Case "���������� ����", "���������� �����", "���������� �������"
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
           
        
          Case "���������� ����"
            cbOpt.AddItem "�����"
            cbOpt.AddItem "�������"
            cbLength.Text = "50"
          
          Case "����� UKW-7"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
         Case "����. ������ BLUM"
            cbOpt.Enabled = True
            cbOpt.AddItem "HS"
            cbOpt.AddItem "HL"
            cbOpt.Text = ""
            
            cbLength.Enabled = False
         
            
 
          Case "������ HL ��������", "������ HS �������"


            cbOpt.Enabled = True
            cbOpt.Text = ""
            cbOpt.List = zavesHL
            cbLength.Enabled = False
           ' cbAddNext.Value = False
            
            

        Case "���� ���. 100��", "������� ��� ������� SISO", "������ 8*40", "������ ��������� DU650", _
                "��������� 6,3*16", "����� �����1000", "����� ��������", "���� � ������", _
                "���� 60", "���� 80", "���� 100", "���� 50", "���� 120", "����� ��� ������ DU650", _
                "�������� CAMAR  �+�", "�������� CAMAR ���", "�������� CAMAR ����", _
                "�������� ��������� ���.", "�������� ����� SAH130", "�������� ����� SAH130 ���", "�������� ����� SAH130 ��", "��������� ���� 6*50", _
                "������ ����������� �����.", "������ ����������� ���.", "������ ��������� miniluna", _
                "������-���� DU232", "������ DU232", "������-���� DU634", "������ DU634", _
                "�������� ������ ��� RV8", "�������� ������ ��� RV1", "����� 3,5*35", "����� 4*35", _
                "������� ������ �/�� �����", "��������� ��� � ��� RV-8", _
                "�������� CAMAR  �+�", "�������� CAMAR ���", "�������� CAMAR ����", "�� ��� ����� � �����. ���"
                ' "�������� ������ 5��"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "���������� 18��", "���������� 16��", "���������� 22��"
            cbOpt.Enabled = False
                If kitchenPropertyCurrent.CamBibbColor = "" Then
                kitchenPropertyCurrent.CamBibbColor = GetCamBibbColor(kitchenPropertyCurrent.dspColor)
                If kitchenPropertyCurrent.CamBibbColor <> "" Then
                    UpdateOrder kitchenPropertyCurrent.OrderId, , , , , , , , kitchenPropertyCurrent.CamBibbColor
                End If
            End If
            cbLength.Enabled = False
        
        Case "�������� ��� �����������"
            cbOpt.Text = ""
            cbOpt.List = ��������
            cbLength.Text = ""
            cbLength.Enabled = False
            
        Case "�������� ������ Sensys"
            cbOpt.AddItem "5��"
            cbOpt.AddItem "10��"
            cbLength.Enabled = False
            
        Case "������"
            cbOpt.AddItem "765"
            cbOpt.AddItem "865"
            cbOpt.AddItem "815"
            cbOpt.AddItem "415"
            cbOpt.AddItem "425"
            cbOpt.AddItem "575"
            cbOpt.AddItem "773��"
            cbOpt.AddItem "800��"
            cbOpt.AddItem "688��"
            cbOpt.AddItem "798��"
            cbOpt.AddItem "998��"
            cbOpt.AddItem "847��"
            cbOpt.AddItem "���� 762��"
            cbOpt.AddItem "���� 756��"
            cbOpt.AddItem "���� 665��"
            cbOpt.AddItem "���� 956��"
            cbOpt.AddItem "����� ���� d-25"
            
            ' ����� �� �����
            cbLength.Enabled = False
        Case "�����"
            cbOpt.Enabled = True
            cbOpt.AddItem "��������"
            cbOpt.AddItem "���������� 5�"
            cbOpt.AddItem "���������� 11�"
            
            If Not IsEmpty(FittingOption) And Len(FittingOption) >= 3 Then
                For i = 0 To cbOpt.ListCount - 1
                    If InStr(1, FittingOption, cbOpt.List(i), vbTextCompare) > 0 Or _
                        InStr(1, cbOpt.List(i), FittingOption, vbTextCompare) > 0 Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next
                Else
            cbOpt.Value = "��������"
            End If
            
            cbLength.Enabled = False
        
        Case "����. ����. ���. ��"
            cbOpt.Enabled = False
            cbLength.Enabled = True
            cbLength.AddItem "90��"
            cbLength.Value = "90��"
        
        Case "������� � �������"
            cbOpt.Enabled = True
            cbOpt.AddItem "�����"
            cbOpt.AddItem "���"
            cbOpt.AddItem "���"
            cbOpt.AddItem "������"
            cbOpt.AddItem "�����"
            cbOpt.AddItem "����"
            cbOpt.AddItem "������"
            
            cbLength.Enabled = False
        
'        Case "������������ �������", "������������ ������"
'            cbOpt.AddItem "1567��"
'            cbOpt.AddItem "1567��"
'            cbOpt.AddItem "1400��"
'            cbOpt.AddItem "1400��"
'            cbOpt.AddItem "1362��"
'            cbOpt.AddItem "1362��"
'            '����� �� �����
'            cbLength.Text = ""
'            cbLength.Enabled = False
        Case "����� ������"
            cbOpt.AddItem "3�"
            cbOpt.AddItem "1,14�"
            cbOpt.AddItem "258��"
            cbOpt.AddItem "274��"
            cbOpt.AddItem "435��"
            cbOpt.AddItem "526��"
            cbOpt.AddItem "664��"
            cbOpt.AddItem "816��"
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        Case "������������ ������"
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
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        Case "������� ���������"
            cbOpt.AddItem "30"
            cbOpt.AddItem "45"
            cbOpt.AddItem "50"
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        Case "�������� ��� ������"
            cbOpt.AddItem "Intermat D-0"
            cbOpt.AddItem "Intermat D-1.5"
            cbOpt.AddItem "Intermat D-3"
            cbOpt.AddItem "Sensys D-0"
            cbOpt.AddItem "Sensys D-1.5"
            cbOpt.AddItem "Sensys W45 D-1.5"
            cbOpt.AddItem "Sensys D-3"
            cbOpt.AddItem "Sens/Hett W45 D-3"
            cbOpt.AddItem "Clip ���������"
            cbOpt.AddItem "Clip ������"
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
            ' ����� �� �����
            cbLength.Enabled = False
            
        Case "�����"
            cbOpt.AddItem "���"
            cbOpt.AddItem "����������"
            cbOpt.AddItem "����������"
            ' ����� �� �����
            cbLength.Enabled = False
        Case "�����������"
            cbOpt.AddItem "BLUM"
            cbOpt.AddItem "FGV"
            cbOpt.AddItem "�������"
            ' ����� �� �����
            cbLength.Enabled = False
        Case "��������� �������"
            cbOpt.AddItem "6400�"
            cbOpt.AddItem "6400� - 5�"
            cbOpt.AddItem "������� 2� (� �������)"
            cbOpt.AddItem "2�"
            ' ����� �� �����
            cbLength.Enabled = False
            
        Case "������+�����+���� 220V"
            cbLength.Enabled = False
            cbOpt.Enabled = False
           
            
        Case "�������������"
            cbOpt.AddItem "LED 30W"
            cbOpt.AddItem "LED 50W"
            cbOpt.AddItem "��� 60W"
            cbOpt.AddItem "��� 105W"
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            cbAddNext.Value = True
        Case "����� ��� ��� �� ���� L", "����� ��� ��� �� ���� R"
            cbOpt.AddItem "���"
            cbOpt.AddItem "���"
            cbLength.Enabled = False
            'cbAddNext.Value = True
        Case "���������� �����. 16 ���"
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
            
            Case "��� ��� �/� ����� 16 ���"
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
         ' ����� �� �����
            cbLength.Enabled = False
            
            
        Case "���������� �����. 16 ���"
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
      Case "��� ��� �/� ����� 16 ���"
      
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
             ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
          
        Case "���� ������� �� �� 16 ���"
           
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
             ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            
      
         Case "���� ������� �� �� 18 ���"
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
             ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
                 
                 
              
        Case "���������� �����. 18 ���"
            
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
       Case "��� ��� �/� ����� 18 ���"
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
           ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
           
        Case "���������� �����. 18 ���"
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
         Case "��� ��� �/� ����� 18 ���"
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
             ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
    
        
        
        Case "����� ������� � ��� ��"
            cbOpt.AddItem "16"
            cbOpt.AddItem "18"
            cbLength.Text = ""


        Case "������� ���������"
            cbOpt.AddItem "30"
            cbOpt.AddItem "35"
            cbOpt.AddItem "40"
            cbOpt.AddItem "45"
            cbOpt.AddItem "50"
            
        Case "��������"
            cbOpt.AddItem "50��"
            cbOpt.Text = "50��"
            'cbOpt.AddItem "250��"
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
             
        Case "��� � ����������"
            cbOpt.AddItem "HLT45"
            cbOpt.AddItem "HLT60"
            cbOpt.AddItem "HLT90"
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        
        Case "���� ������ 200��"
           cbOpt.Enabled = False
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        
        
        Case "����� ��� �������-� BLUM", "����� ��� ����������� FGV", "������ DU325 Rapid S", "������ VB15"
                    
            cbOpt.Enabled = False
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            cbAddNext.Value = True
            
   '     Case "������ �����������"
        Case "������"
            cbOpt.AddItem "�����������"
            cbOpt.AddItem "�������"
    '       cbOpt.Text = ""
    '       cbOpt.Enabled = False
            cbLength.Enabled = True
            cbLength.Text = ""
            cbLength.List = PA
        
'        Case "����� � �����"
'            cbOpt.Enabled = True
'            cbLength.Enabled = False
'            cbOpt.AddItem "1/2"
'            cbOpt.AddItem "3,5"
'            cbAddNext.Value = True
            
'        Case "��������� � �����"
'            cbOpt.Enabled = True
'            cbLength.Enabled = False
'            cbOpt.List = Sink
'            cbOpt.Text = Krepl
'            cbAddNext.Value = True
            
'        Case "���������� 4�"
'            cbOpt.Enabled = False
'            cbLength.Enabled = False
'            cbAddNext.Value = True
            
        Case "�����"
           ' ������
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
            ' ����
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
        
        Case "�����"
    
            ' ������
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
            
            ' ���� �� �����
            cbOpt.Text = ""
            cbOpt.Enabled = False
            
'        Case "������"
'            cbOpt.List = ������
'            cbLength.List = BackKolib
'            cbAddNext.Value = True
            
        Case "��������"
            cbOpt.Text = "���"
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            
        Case "������������"
            cbOpt.List = ������������
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
       
            Dim fo As String
             If IsEmpty(FittingOption) Then fo = "�����" Else fo = FittingOption
            If Not IsEmpty(FittingLength) Then fo = fo & " " & FittingLength
            fo = Trim(fo)
                For i = 0 To cbOpt.ListCount - 1
                    If fo = cbOpt.List(i) Then
                        cbOpt.Text = cbOpt.List(i)
                        Exit For
                    End If
                Next i
        Case "���������� 3�", "���������� ����-4", "���������� ����-5", "���������� TOP-Line", _
             "���� � ���. 3�", "����. � ���. 3�", "���� �����. � ���. 3�", _
             "���� � ���. 4�", "����. � ���. 4�", "���� �����. � ���. 4�", _
             "���� � ���. ����-4", "����. � ���. ����-4", "����. ��� � ���. ����-4", "����. ���� � ���. ����-4", "���� �����. � ���. ����-4", _
             "���� � ���. ����-5", "����. � ���. ����-5", "����. ��� � ���. ����-5", "����. ���� � ���. ����-5", "���� �����. � ���. ����-5", _
             "���� � ���. TOP-Line", "����. � ���. TOP-Line", "���� ����. � ��� TOP-Line", _
             "������� � ����������", "�������-40 � ����������", "���������� 4�"
             
            
            
            cbLength.Text = ""
            cbLength.Enabled = False
            
            cbOpt.Text = ""
            
            Select Case cbFittingName.Text
            
                Case "���������� 4�"
                    cbOpt.Enabled = False
                    ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
                    cbAddNext.Value = True
                    
                    
                
                Case "���������� ����-4", "���������� ����-5"
                    cbOpt.List = OtbGorbColors
'
'                Case "���������� 4�"
'                    'bSkip = True
'                    cbOpt.List = �������
                    
                Case "������� � ����������", "�������-40 � ����������"
                    'bSkip = True
                    cbOpt.List = �������
                    
                Case "���� � ���. 4�", "����. � ���. 4�", "���� �����. � ���. 4�"
                
                    cbOpt.List = ���������4�
                    
                    If Not IsEmpty(FittingOption) Then
                        Select Case FittingOption
                            Case ���������4�(0), ���������4�(1), ���������4�(2), ���������4�(3), ���������4�(4), ���������4�(5), ���������4�(6)
                            Case Else
                            Select Case FittingOption
                            
                                Case "����������� ������", "������������ ���������", "������", "����� ������", "�������� ���������", _
                                    "����", "���������", "�����", "�������� ������", "�������� �������", "������", "������ ������", "�����", "����������� ������", _
                                    "��������", "��������", "�������", "�������", "��������� ������", "����� ������", "����������� ���������", "����� ������", "����� ����", "������"
                                
                                    FittingOption = "���"
                                    
                                Case "��� ������ ������", "��� ������ �������", "������� ��������", "��� ���������", "������ ������", "������", "�������� ����", _
                                    "���������", "�������", "�������� ���������", "�������� ������", "�������� �������", "������ �������", "������ ������", _
                                    "�����", "��������� ������", "��������� �������", "��������� ������", "���������� ������", "���������� �������", "����� �������", "����������", "���"
                                                                    
                                    FittingOption = "���"
                                    
                                Case "��������"
                                    
                                    FittingOption = "���"
                                
                                Case "���� �����"
                                                                    
                                    FittingOption = "���"
                                    
                                Case "�������", "�����", "�����", "������", "������� ����", "����� ����", "������", _
                                "������", "������ ������", "������ ����������", "������� ��������", "���� ������", "��������", "����", "������", "������� ���"
                                
                                    FittingOption = "���"
                                    
                                Case "����� ������", "������", "������ ������", "������ ������", "������ ������", "����� ������", "����", "������", "���� ����", "�������"
                                
                                    FittingOption = "����"
                                    
                                Case "������ ������", "������� �������"
                                
                                    FittingOption = "���"
                                    
                                Case "���� ����", "����� �����", "������� ������"
                                
                                    FittingOption = "���"
                                
                                Case "������ ������", "����� ������", "����� �������"
                                    
                                    FittingOption = "���"
                                    
                            End Select
                        End Select
                     End If
                    
                Case "���� � ���. ����-4", _
                     "����. � ���. ����-4", _
                     "����. ��� � ���. ����-4", _
                     "����. ���� � ���. ����-4", _
                     "���� �����. � ���. ����-4", _
                     "���� � ���. ����-5", _
                     "����. � ���. ����-5", _
                     "����. ��� � ���. ����-5", _
                     "����. ���� � ���. ����-5", _
                     "���� �����. � ���. ����-5"
                     
                     cbOpt.List = ��������������
                     ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
                     
                     If Not IsEmpty(FittingOption) Then
                        Select Case FittingOption
                            Case ��������������(0), ��������������(1), ��������������(2), ��������������(3), ��������������(4), ��������������(5)
                            Case Else
                            Select Case FittingOption
                            
                                                                   
                                Case "��� ������", "�����", "��� ���", "����� ������", "���� �����", "������", "�������� ����", "���������", "������"
                                    FittingOption = "�-���"
                                Case "��� ��", "��� �� ������", "�������", "����� �������", _
                                    "���� ������", "������", "�����", "����������", "������", "��� ��", "����� ������"
                                    FittingOption = "���"
                                Case "���"
                                    FittingOption = "�-���"
                                Case "����� ������", "����", "����� ����", "����� ����", "������� ����"
                                    FittingOption = "�-���"
                                Case "������ ������"
                                    FittingOption = "�����"
                                Case "������ ������", "��������", "�����", "������"
                                    FittingOption = "�-���"
                                Case "������"
                                    FittingOption = "�-���"
                                Case "���� ������ ��", "������ ������", "������ ������", "���� ����", "�������"
                                    FittingOption = "����"
                                Case "���� �����", "���� ���"
                                    FittingOption = "����"
                                Case "��������"
                                    FittingOption = "�����"
                                Case "���� ����"
                                    FittingOption = "���"
                            End Select
                        End Select
                     End If
                     
                Case "���������� TOP-Line", "���� � ���. TOP-Line", "����. � ���. TOP-Line", "���� ����. � ��� TOP-Line"
                    cbOpt.List = TOPLine
                  ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
                Case Else
                    cbOpt.List = OtbColors
                    ' ����� �� �����
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
            
        Case "�������� �-� Push-To-Open"
            cbOpt.Enabled = True
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            
            cbOpt.AddItem "�������������"
            cbOpt.AddItem "� ��������"
            
        Case "����������� �/�� � �����."
            cbOpt.Enabled = False
            cbLength.Enabled = True
            
            cbLength.AddItem "60��"
            cbLength.AddItem "80��"
            cbLength.AddItem "90��"
        Case "����� �/��"
            cbOpt.Enabled = False
            cbLength.Enabled = True
            cbLength.AddItem "40��"
            cbLength.AddItem "50��"
            cbLength.AddItem "60��"
            cbLength.AddItem "80��"
            cbLength.AddItem "90��"
            cbLength.AddItem "90�� � �������"
        Case "��������� �/��"
            cbOpt.Enabled = True
            cbLength.Enabled = False
            
            cbOpt.AddItem "�����"
            cbOpt.AddItem "����������"
            
        Case "��������� �/� �.�. 70��"
            cbOpt.Enabled = True
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "����"
            cbOpt.AddItem "���"
            
        Case "��������� �/� �.�.144��"
            cbOpt.Enabled = True
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "����"
            cbOpt.AddItem "���"
         Case "����������� ��� �����"
            cbOpt.Enabled = True
            cbOpt.Text = "OrgalFlex"
            cbLength.Text = ""
            cbLength.Enabled = False
             
        Case "��������� � ������"
            cbOpt.Enabled = True
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "������"
            cbOpt.AddItem "��������"
           
        Case "������� ��������������"
            cbOpt.Enabled = True
            cbOpt.AddItem "��������"
            cbOpt.AddItem "��������"
            cbOpt.Text = "��������"
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        
        Case "��������� ������� 2K", "��. ����. clip HF,HKS", "��������� � ���. ��", "������������ ���� 83 HF", "����� Blum110 �/� ��Al���", "����� Blum110 �/� ��Al���", "������������ ���� 75 HK"
            cbOpt.Enabled = False
           ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
      Case "������-� LED 30w"
            cbOpt.Enabled = False
           ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            cbAddNext.Value = True
      
        Case "���������� �������"
            cbOpt.Enabled = True
            ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
            cbOpt.AddItem "Barri (3��)"
         
        Case "���� �/�", "���� ������", "���� ����", "���� �����"
        
         cbOpt.Enabled = True
         ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
         get_st_par
         cbOpt.List = Stul_color_no
         
         Case "�-�� �������� ������"
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
            
        Case "�-�� �������� ����2"
          cbOpt.Enabled = True
            cbLength.Text = ""
            cbLength.Enabled = False
        
        Case "���� Zebra"
         cbOpt.AddItem "2273(100)�����."
         cbOpt.AddItem "2273(130)�����.�����."
         cbOpt.AddItem "2273(140)�����.�������"
         cbOpt.AddItem "2273(183)�����."
         cbOpt.AddItem "2273(310)��.�����"
         cbOpt.AddItem "2273(380)��.������"
         cbLength.Text = ""
         cbLength.Enabled = False
        
        Case "����� CLIP top"
            'cbAddNext.Value = False
        cbOpt.Enabled = True
        cbOpt.Text = ""
        cbOpt.List = zavesClipTop
        cbLength.Text = ""
        cbLength.Enabled = False
        If Not IsEmpty(FittingOption) Then
            If FittingOption = "BLUMOTION +90 ��� ��" Then cbOpt.Text = FittingOption
            If FittingOption = "BLUMOTION +45" Then cbOpt.Text = FittingOption
            If FittingOption = "+155" Then
            cbOpt.Text = FittingOption
            cbAddNext.Value = True
            End If
        End If
        Case "����� Sensys"
            
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
        ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
'        If casepropertyCurrent Is Nothing Then
'        cbAddNext.Value = True
'        ElseIf casepropertyCurrent.p_fullcn = "" Then
'        cbAddNext.Value = True
'        End If
        
            
        Case "�������� ���.Sensys"
        cbOpt.Enabled = True
        cbOpt.Text = ""
        cbOpt.List = ploschadkaSensys
        ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
         
        Case "������ CLIP TOP +155"
        cbOpt.Enabled = False
        cbLength.Text = ""
        cbLength.Enabled = False
        
        Case "������. Sensys 165"
        cbOpt.Enabled = False
        ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        'cbAddNext.Value = False
         
        Case "���. ���� Sensys 110-85 ", "���� �� ����� ���.Sensys", "���� �� ����� ���.Sensys", "�������� �� ���", "�������� �� ����"
        cbOpt.Enabled = False
        ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
      
       Case "��������� ������ BLUM"
        cbOpt.Enabled = True
        ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False
        cbOpt.AddItem "HK"
        cbOpt.AddItem "HL"
        cbOpt.AddItem "HS"
        cbOpt.AddItem "HK-S"
        
        Case "SL56 ����� �������"
        cbOpt.Enabled = False
        cbLength.Enabled = True
        cbLength.Text = "3000 ��"
        cbLength.AddItem "3000 ��"
        cbLength.AddItem "1638 ��"
        Case "SL56  ������� �������", "SL56 ������������ �������", "SL56 ��������� ���������", "SL56 ���� ���-���� �����", "SL56 ����� ������� 1638��", "SL56 ���� ���� ��� ��"
        cbOpt.Enabled = False
        ' ����� �� �����
            cbLength.Text = ""
            cbLength.Enabled = False

        Case "�����. ������� HK,HL,HS"
        cbLength.Enabled = False
        cbOpt.Enabled = False

        Case "����������� ������"
        cbOpt.List = ������
        cbLength.Enabled = False
        
        Case "������ ��������������", "������ ���������� ���"
        cbLength.Enabled = True
        cbLength.List = tbkovrLength
        cbOpt.Enabled = True
        cbOpt.List = tbkovrOpt
        
        Case "����� � ������ ����"
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
            cbOpt.AddItem "070i - ����� ���� 1/2"
            cbOpt.AddItem "361i - ����� ���� 3/4"
            cbOpt.AddItem "370L - Nuvola L"
            cbOpt.AddItem "370R - Nuvola R"
            cbOpt.AddItem "258A 20 - ����� ��� 20"
            cbOpt.AddItem "258A 15 - ����� ��� 15"
            cbOpt.AddItem "575 - ���"
            cbOpt.AddItem "����� � ������ ���� 60"
            cbOpt.AddItem "����� � ������ ���� 90"
            cbOpt.AddItem "230A 450 - ���� ���� Maxi"
            cbOpt.AddItem "230A 600 - ���� ���� Maxi"
            cbOpt.AddItem "230B 450 - ���� ���� MIDI"
            cbOpt.AddItem "230B 600 - ���� ���� MIDI"
            cbLength.Enabled = False
            If Not IsEmpty(FittingOption) Then
                cbOpt.Text = FittingOption
            End If
        
        Case "Tip-ON"
        cbOpt.Enabled = True
        cbOpt.AddItem "955 ��������"
        cbOpt.AddItem "955 �/��� ��� ����"
        cbOpt.AddItem "955� ��������"
        cbOpt.AddItem "955� ���� �/��� ��� ����"
        cbOpt.AddItem "������ �� ����"
        cbOpt.AddItem "��������� ��������������"
        cbLength.Enabled = False
        
        Case "�-�� �������� �����"
        cbOpt.Enabled = True
        cbOpt.AddItem "����1"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "����2"
        cbOpt.AddItem "����3"
        cbOpt.AddItem "�����1"
        cbOpt.AddItem "�����2"
        cbOpt.AddItem "�����3"
        cbOpt.AddItem "�����4"
        cbOpt.AddItem "�����1"
        cbOpt.AddItem "�����"
        cbOpt.AddItem "������"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "����2"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "����2"
        cbOpt.AddItem "����3"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "����2"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "�����1"
        cbOpt.AddItem "�����2"
        cbOpt.AddItem "�����1"
        cbOpt.AddItem "�����2"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "�����1"
        cbOpt.AddItem "�����2"
        cbOpt.AddItem "����1"
        cbOpt.AddItem "�����4"
        cbOpt.AddItem "�����5"
        cbOpt.AddItem "����4"
        cbOpt.AddItem "����7"
        cbOpt.AddItem "�����1"
        cbOpt.AddItem "�����2"
        cbOpt.AddItem "�����3"
        cbOpt.AddItem "�����4"
        cbOpt.AddItem "�����5"
        cbOpt.AddItem "�����6�����"
        cbOpt.AddItem "�����6�����"
        cbOpt.AddItem "�����6������"
        cbOpt.AddItem "������"
        
        cbLength.Enabled = True
        cbLength.AddItem "������ ����� 101"
        cbLength.AddItem "������ ����� 110"
        cbLength.AddItem "������ ����� 111"
        cbLength.AddItem "������ ����� 112"
        cbLength.AddItem "������ ����� 113"
        cbLength.AddItem "������ ����� 119"
        cbLength.AddItem "������ ����� 120"
        cbLength.AddItem "������ ����� 122"
        cbLength.AddItem "������ ����� 115"
        cbLength.AddItem "������ ����� 113"
        cbLength.AddItem "������ ����� 120"

        
        Case "�-�� ��������"
        cbOpt.Enabled = True
        cbOpt.AddItem "����� ��� �����"
        cbLength.Enabled = False
        
        Case "������ �������100 ��", "������ �������150 ��"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "���������� ������� 4� ��"
            cbAddNext.Value = True
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "����90 � ������ �������", "������ � ������ �������", "����90 � ��� �������150", "���+���� � ��� �������"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "������� ����"
            cbOpt.AddItem "��� 1���� G12AL07"
            cbOpt.AddItem "��� 2���� G13AL07"
            cbOpt.AddItem "���� 1���� G16AL07"
            cbOpt.AddItem "���� 2���� G15AL07"
            cbOpt.AddItem "� ��� 1���� G14AL07"
            cbOpt.AddItem "������� ���� ���� (�-�)"
            cbOpt.AddItem "���� ���"
            cbOpt.AddItem "���� ���"
            cbOpt.AddItem "���� 90����"
            cbOpt.AddItem "���� 90�����"
            cbOpt.AddItem "�-� �����"
            
            cbLength.Enabled = True
            
        Case "��������� ������� ����"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        Case "�������� ������� �������"
            cbOpt.AddItem "81/G1.1AT2"
            cbOpt.Text = "81/G1.1AT2"
            cbLength.Enabled = False

        
        Case "�����"
            cbOpt.AddItem "�����"
            cbOpt.AddItem "��������"
            cbOpt.AddItem "������"
            cbOpt.AddItem "������+"
            cbOpt.AddItem "���������"
            cbOpt.AddItem "������"
            cbOpt.AddItem "������+"
            cbOpt.AddItem "���������"
            cbOpt.AddItem "�����"
            cbOpt.AddItem "������"
            cbOpt.AddItem "��-1"
            cbOpt.AddItem "��-2"
            cbOpt.AddItem "��-3"
            cbOpt.AddItem "��-4"
            cbOpt.AddItem "��-5"
            cbOpt.AddItem "��-6"
        Case "����� �����", "����� ��������", "����� ������", "����� ������", "����� ������+", "����� ���������", "����� ��-1", "����� ��-2", "����� ��-3", "����� ��-4", "����� ��-5", _
            "����� ��-6", "����� �����", "����� ������", "����� ������+", "����� ���������"
            
            cbLength.Clear
            cbLength.AddItem "�����(����)"
            cbLength.AddItem "������(����)"
            cbOpt.Clear
            cbOpt.List = MoikaColors
        
        Case "������ ANODA �����", "������ 8*60", "������ RAFIX TAB20", "����� 5*80", "������ ���. 60*60*50"
            cbOpt.Enabled = False
            cbLength.Enabled = False
            
        
        Case "���� �/��������� ��� ������"
            cbOpt.AddItem "�� �����."
            cbOpt.AddItem "�����"
            cbOpt.Text = "�����"
            cbLength.Enabled = False
            
       Case Else
            If cbFittingName.ListIndex > -1 Then
            cbOpt.Text = ""
            cbOpt.Enabled = False
            cbLength.Text = ""
            cbLength.Enabled = False
            Else
            
            MsgBox "� ����� ��������� �� ����..."
            End If
            'Exit Function
    
    
    
    
    End Select
    

End Sub
    
Private Sub cbOpt_Change()
    If binit Then Exit Sub
    
    Select Case cbFittingName.Text
    
        Case "�����"
            If cbOpt.Text = "��������" Then
                cbLength.Enabled = True
            Else
                cbLength.Enabled = False
            End If
        Case "������ ��������������", "������ ���������� ���"
            If cbOpt.Text = "��" Then
                cbLength.Clear
                cbLength.Text = ""
                Else
                cbLength.Text = ""
                cbLength.List = tbkovrLength
                
            End If
'        Case "�������"
'            If cbOpt.Text = "CAMAR �+�" Then
'                cbAddNext.Value = True
'            End If
        Case "�����"
            If cbOpt.Text <> "" Then
                cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
            End If
        
        Case "���� �/�", "���� ������", "���� ����", "���� �����"
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
'                Case Stul(5), Stul(6) ' ����, ��
'                    cbLength.Text = "�����"
'                    cbLength.Enabled = False
'                    cbAddNext.Value = False
'                Case Stul(8), Stul(9), Stul(10), Stul(14) ' ������, ����, ������, ������
'                    cbLength.List = SitK
'                    cbLength.Enabled = True
'                    cbAddNext.Value = False
'                Case Stul(11), Stul(12) ' "������� ����" "������� ����"
'                    cbAddNext.Value = True
'                Case Stul(0), Stul(1), Stul(2), Stul(3), Stul(4), Stul(13) ' "�����" "�����" "�����" "����" "�����","������"
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
'                Case Sit(0), Sit(1), Sit(2) ' "D390" "D340" "�����"
'                    cbLength.List = SitColors
'                    cbLength.Enabled = True
'                Case Sit(3) ' "�������"
'                    cbLength.List = SitKolib
'                    cbLength.Enabled = True
'            End Select
'
'        Case "������"
'            cbLength.Enabled = True
'            cbAddNext.Value = True
        
'        Case "�����"
'            cbAddNext.Value = True
            
'        Case "��������� � �����"
'            cbAddNext.Value = True
        Case "������� ����"
            Select Case cbOpt.Text
                Case "���� ���"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.AddItem "1����(1AT2GA)"
                cbLength.AddItem "2����(3AT2GA)"
                cbLength.Text = ""
                Case "���� ���"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.AddItem "1����(1AT3GA)"
                cbLength.AddItem "2����(3AT3GA)"
                cbLength.Text = ""
                
                Case "���� 90����"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.Text = "1����(1A90B)"
                
                Case "���� 90�����"
                cbLength.Enabled = True
                cbLength.Clear
                cbLength.Text = "1����(1A90A)"
                
            End Select
       Case "����� Sensys"
'            Select Case cbOpt.Text
'            Case "165"
'                    cbAddNext.Value = True
'            End Select
        Case "�����"
            Select Case cbOpt.Text
                Case "��� ����������� BLUM", "��� ����������� FGV"
                    cbAddNext.Value = True
                Case "HF28", "HF22", "HF25", "HK-S", "��������", "��� �����������"
                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
                Case "HK", "��", "�K", "H�" ' � ������ ����������! "FGV180"
                    cbFittingName.Text = cbFittingName.Text & " " & "HK27"
                 Case "HK25", "HK27", "HK25 (TIP-ON)", "HK27 (TIP-ON)"
                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
                Case "HL23/35", "HL23/38", "HL25/35", "HL25/38", "HL27/35", "HL27/38", _
                "HL25/39", "HL27/39", "HL29/39", "HL23/39", _
                "HS A", "HS B", "HS D", "HS E", "HS G", "HS H", "HS I"
                
                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
            
            End Select
        
            
        Case "����"
'            Select Case cbOpt.Text
'                Case "80"
'                    cbFittingName.Text = cbFittingName.Text & " " & cbOpt.Text
'            End Select
            
        Case "�����"
            If cbOpt.Text = "������������� ����" Then
                cbLength.Clear
                cbLength.AddItem "60"
                cbLength.AddItem "90"
            
            Else
                Select Case cbOpt.Text
                    Case "�����"
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
            
        Case "������ � ����", _
                "������ �/� ������������", _
                "������ � ���. �����"
            Select Case cbOpt.Text
                Case "����"
                Case Else
                    cbLength.Text = "28"
            End Select
        Case "��� ������ � ��� ��"
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
        Case "����� ������� � ��� ��"
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
    On Error GoTo err_�����������������
    
    ' ������� ��������
    'Dim TasksForm As MainForm
    Dim ShipID As Long
    'Set TasksForm = New MainForm
    MainForm.Show
    ShipID = MainForm.ShipID
    
    'Set TasksForm = Nothing
    If ShipID = 0 Then Exit Sub
    Set kitchenPropertyCurrent = New kitchenProperty
    
    Set casepropertyCurrent = Nothing
    
    ' ������� ������� � �����
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
        MsgBox "������� ���������", vbInformation, "���������� ���������"
    End If
        
    Exit Sub
    
err_�����������������:
    MsgBox Error, vbCritical, "���������� ���������"
End Sub
   


