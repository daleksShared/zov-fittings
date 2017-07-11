VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorForm 
   Caption         =   "Выбор цвета бочков"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2865
   OleObjectBlob   =   "ColorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Option Compare Text
Public ColorId As Long
Public ColorName As String

Private Sub ApplyButton_Click()
    With ColorComboBox
        If .ListIndex <> -1 Then
            ColorId = .List(.ListIndex, 0)
            ColorName = .List(.ListIndex, 1)
            Hide
        Else
            MsgBox "Выберите цвет каркасов текущего заказа", vbExclamation, "Цвет бочков"
        End If
    End With
End Sub

Private Sub CancelButton_Click()
    ColorId = 0
    ColorName = ""
    Hide
End Sub


Private Sub UserForm_Initialize()
    
    ColorId = 0
    ''Dim ConnBarcode As ADODB.Connection
    InitBarcodeConnection ConnBarcode
    
    Dim cmdColor As ADODB.Command       ' Цвета ДСП
    
    Set cmdColor = New ADODB.Command       ' Цвета ДСП
    Set rsColor = New ADODB.Recordset
    Set cmdColor.ActiveConnection = ConnBarcode
    
    rsColor.CursorLocation = adUseClient
    rsColor.LockType = adLockReadOnly
    
    cmdColor.CommandType = adCmdText          ' Получаем названия цветов ДСП
    cmdColor.CommandText = "Select ColorID, ColorName,parserName,parserSecondName,isnull(bibbColor,'') as bibbColor,isnull(cambibbColor,'') as cambibbColor,parserTripleName FROM Colors ORDER BY ColorName"
    
    rsColor.Open cmdColor, , adOpenDynamic, adLockReadOnly
    
    Dim LocI As Long
    Dim TmpStr As String

    With ColorComboBox
        .Clear
        While Not (rsColor.EOF)
            .AddItem ("")
            For LocI = 0 To 1 'rsColor.Fields.Count - 1
                If (IsNull(rsColor(LocI))) Then
                    .List(.ListCount - 1, LocI) = ""
                Else
                    .List(.ListCount - 1, LocI) = rsColor(LocI)
                End If
            Next LocI
            rsColor.MoveNext
        Wend
        .ColumnCount = 2 'rsColor.Fields.Count
        
        TmpStr = ""
        For LocI = 1 To .ColumnCount
            TmpStr = TmpStr & "-1;"
        Next LocI
        
        .ColumnWidths = "0;" & TmpStr
    End With
    
    
    'ConnBarcode.Close
    
    ColorComboBox.ListIndex = -1
End Sub


