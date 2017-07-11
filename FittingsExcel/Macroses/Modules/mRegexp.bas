Attribute VB_Name = "mRegexp"
Option Explicit
Option Compare Text
Public Const patGetFirstLetters As String = "^([a-zA-Zà-ÿÀ-ß]+)" 'ØËÊ80/40/1ñïðàâà ïðàâ -> ØËÊ80/1ñïðàâà ïðàâ
Public Const patCaseIsZovModul As String = "^[ØÏ].*[ ]{1}(Ì|M)$"
Public Const patR_check As String = "^Ð[ËÍ][À-ß]*\d+.*"
Public Const patRL_full As String = "^ÐË\d+((?:[/]|[-])\d-\d+)?([(]\d+[)])?(?:[(].*[)]).*"
Public Const patRL_simple As String = "^ÐË\d+/\d+(?:[ ]|[à-ÿÀ-ßa-zA-Z]|$).*(?![(]\d+[)]).*"
Public Const patGetShtQtyfromStringBegin As String = "^.*?-\s?(\d+)\s?øò\.?"
Public Const patSHLK_check1 As String = "^Ø[ËÍ]Ê[0-9]+([/][0-9]+[/])(?:[(]á?[.]?[0-9]+[)])?(?:.*)" 'ØËÊ80/40/1ñïðàâà ïðàâ -> ØËÊ80/1ñïðàâà ïðàâ
Public Const patSHLK_check2 As String = "^Ø[ËÍ]Ê[0-9]+([/][0-9]+[/]2(?:ñïðà|ñëåâ)?)(?:[(]á?[.]?[0-9]+[)])?(?:.*)" 'ØËÊ80/40/2ñïðàâà
Public Const patSHL_check1 As String = "^ØË[0-9]+(?:[(]á?[.]?[0-9]+[)])?([\(].*[\)]){1}(?:.*)" 'ØË60(íèø,360/1) -> ØË60/1Ò
Public Const patSHL_check2 As String = "^ØË[0-9]+(?:[(]á?[.]?[0-9]+[)])?([/][1-9])(?:ÂÏ)" 'ØË80/2ÂÏ -> ØË80Ò + 2 äâåðè
Public Const patSHN_check1 As String = "^ØÍ(?:Â|Ñ)?[0-9]+(?:[(]á?[.]?[0-9]+[)])?(?:[\(].*[,]?.*[\)]){1}[ ]?(ââåðõ)(?:.*)" 'ØÍ90(355/1,355/1)ââåðõ -> ØÍ90/2Ò
Public Const patNewName As String = "^[À-ßA-Z]+[0-9]+([(]á?[.]?[0-9]+[)])?([\(].*[,]?.*[\)]){1}(?:.*)"
Public Const patGetBase As String = "^[À-ßA-Z]+[0-9]+(?:[/]\d+)?([(](?:á?[.]?)[0-9]+[)]){1}(?:.*)"
Public Const patGetBaseValue As String = "^[À-ßA-Z]+[0-9]+(?:[/]\d+)?[(]á?[.]?([0-9]+)[)]{1}(?:.*)"
Public Const patGetWidth As String = "^[À-ßA-Z]+([0-9]+)(?:.*)"
Public Const patGetWidthValue As String = "^[À-ßA-Z]+([0-9]+)(?:.*)"
Public Const patGetDepth As String = "(?:ãë|ãëóá){1}[.]?[0-9]+[ ]?(?:ñì)?[.]?"
Public Const patGetDepthValue As String = "(?:ãë|ãëóá){1}[.]?([0-9]+)[ ]?(?:ñì)?[.]?"
Public Const patZKGetDepth As String = "^.*_(\d+)[^_]*?$"
Public Const patZKGetWithoutDepth As String = "^(.*)_\d+([^_]*?)$"
Public Const patZPT As String = "[,]"
Public Const patNumber As String = "([0-9]+)"
Public Const patNumberWithSlash As String = "([0-9]+[/]{1})"
Public Const patCountInSkobochkiAfterSlash As String = "[/]([0-9]+)[^0-9/]*(?:[,]{1}|[\)]{1})?"
Public Const patGetSkobochki As String = "^[À-ßàß]+[0-9]+(?:[(]á?[.]?[0-9]+[)])?([\(].*[\)]){1}.*"
' 2015-02 Public Const patCaseFasades As String = "(?:[(]á?[.]?[0-9]+[)])?([\(]{1}.*[,]?.*[\)]{1}(?:òá\d*|ðîë\d*|øàð\d*|âï\d*|ìá\d*|êâ\d*|ââåðõ|H(?:F|KS|L)\d*)?)"
'Public Const patCaseFasades As String = "(?:[(]á?[.]?[0-9]+[)])?([\(]{1}.*[,]?.*[\)]{1}(?:[à-ÿÀ-ßa-zA-Z]+[\sà-ÿÀ-ßa-zA-Z0-9-]+\d*?.*?\d*)?)"
Public Const patCaseFasades As String = "(?:[(]á?[.]?[0-9]+[)])?([\(]{1}.*[,]?.*[\)]{1}(?:äâ|ëåâ|ïðàâ|òá|ìá|ìê|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N|H)?[\sà-ÿÀ-ßa-zA-Z0-9-]*\d*?.*?\d*)"
' 2015-02 Public Const patCaseStringAfterFasades As String = "(?:[(]á?[.]?[0-9]+[)])?[\(]{1}.*[,]?.*[\)]{1}(?:òá\d*|ìá\d*|ðîë\d*|øàð\d*|âï\d*|êâ\d*|ââåðõ|H(?:F|KS|L)\d*)?(.*)"
Public Const patCaseStringAfterFasades As String = "(?:[(]á?[.]?[0-9]+[)])?[\(]{1}.*[,]?.*[\)]{1}(?:[à-ÿÀ-ßa-zA-Z]+[à-ÿÀ-ßa-zA-Z0-9-]+\d*?.*?\d*)?(.*)"
' 2015-02 Public Const patCaseFasadesNapravlMain As String = "[\)]((?:òá|ìá|êâ|ðîë|øàð|âï|àð)\d*)"
Public Const patCaseFasadesNapravlMain As String = "[\)]((òá|ìê|ìá|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N|H)[à-ÿÀ-ßa-zA-Z0-9-]*\d*?.*?\d*)"
Public Const patCaseFasadesOnlyString As String = "(?:[(]á?[.]?[0-9]+[)])?([\(]{1}.*[,]?.*[\)]{1})"
Public Const patCaseFasadesVVERH As String = "[\)](ââåð[õ]?)"
Public Const patCaseFasadesQty As String = "[/](\d{1,})"
Public Const patCaseFasadesWidth As String = "(\d{1,}).*[/]?"
Public Const patCaseFasadesIsDver As String = "((?:[/]\d+])|(?:äâ){1}|(?:ëåâ){1}|(?:ïðàâ){1})"
Public Const patCaseFasadesIsVitr As String = "(âèòð)"
Public Const patCaseFasadesIsShufl As String = "øóô|èìèò"
Public Const patCaseFasadesIsNisha As String = "((?:íèø|íä|í.)\d*)"
Public Const patCaseFasadeGetNapravl As String = "(?:\d*[/]?\d*)?((?:(?:òá|ìá|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-]*)(?: êëèåíò)?(?:\+?\d*(?:òá|ìá|ìê|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-]*)*)"
Public Const patCaseFasadesNapravl As String = "(?:\d*[/]?\d*)?((?:\+?(?:äâ|ëåâ|ïðàâ|òá|ìá|ìê|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-]*)(?: êëèåíò)?(?:\+?\d*(?:òá|ìá|ìê|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-]*)*)"
'Public Const patCaseFasadesNapravlList As String = "(?:\d{3})?(?:[/]?)(?:\+)?(\d*(?:òá|ìá|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-]*(?:[-](?:òá|ìá|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-]*)?\d*)"
Public Const patCaseFasadesNapravlList As String = "(?:^\+())?(?:\d{3})?(?:[/]?)(\d*(?:äâ|ëåâ|ïðàâ|òá|ìá|ìê|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-]*(?:[-](?:òá|ìá|ìê|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N)[\sà-ÿÀ-ßa-zA-Z0-9-+]*)?\d*)"
' 2015-02 Public Const patCaseFasadesNapravlCustomer As String = "(?:[/]?\d+)((?:òá(?:[ì]|[á]|(?:[ -]ìîéêà))?|ìá(?:[ì]|[á])?|ðîë\d*|øàð\d*|âï\d*|êâ\d*|ñóø([à-ÿÀ-ßa-zA-Z]+)?)\d*( êëèåíò){1}(?:[+]?([à-ÿÀ-ßa-zA-Z]+)?))"
' Public Const patCaseFasadesNapravlCustomer As String = "(?:[/]?\d+)(((?:òá|ìá|êâ|ðî|øà|âï|àð|ñó)[\sà-ÿÀ-ßa-zA-Z0-9-]*\d*?.*?)\d*( êëèåíò){1}(?:[+]?(\d*(?:òá|ìá|êâ|ðî|øà|âï|àð|ñó)[\sà-ÿÀ-ßa-zA-Z0-9-]*)?))"
Public Const patCaseFasadesNapravlCustomer As String = "( êëèåíò){1}"
Public Const patCaseSHLGPFasadesNapravl As String = "[/]{1}\d?((?:òá|ìá|ìê|êâ|ðî|øà|âï|àð|ñó|C|D|in|ÂÊ|àí|M|N|H)[\sà-ÿÀ-ßa-zA-Z0-9-]*)"
Public Const patGetNumberFirst As String = "^(\d+)[à-ÿÀ-ßa-zA-Z]+.*"
Public Const patGetNumberLast As String = ".*[à-ÿÀ-ßa-zA-Z]+(\d+)$"
'Public Const patGetStringTrimNumbers As String = "^\d*([à-ÿÀ-ßa-zA-Z0-9]*[à-ÿÀ-ßa-zA-Z]+(?:[ -]{1}[à-ÿÀ-ßa-zA-Z0-9]+)?[à-ÿÀ-ßa-zA-Z]*)*\d*$"
'Public Const patGetStringTrimNumbers As String = "^\d*([à-ÿÀ-ßa-zA-Z]+[à-ÿÀ-ßa-zA-Z0-9]*[à-ÿÀ-ßa-zA-Z]+(?:[ -]{1}[à-ÿÀ-ßa-zA-Z0-9]+[à-ÿÀ-ßa-zA-Z]*)?)*\d*$"
Public Const patGetStringTrimNumbers As String = "^\d*(.*?)\d*$"


Public Const patSplitPattern As String = "(?!\+)([^\+]+)(?:(?=\+)|$)"



'Public Const patGetBase As String = "^[À-ßàß]+[0-9]+([(]á?[.]?[0-9]+[)]){1}(?:.*)"

Function TestRegExp(myPattern As String, mystring As String)

   Dim objRegExp As regexp
   Dim objMatch As Match
   Dim colMatches   As MatchCollection
   Dim objSubmatches As SubMatches
   Dim RetStr As String
   Dim i As Integer
   
   Set objRegExp = New regexp
   objRegExp.Pattern = myPattern
   objRegExp.IgnoreCase = True
   objRegExp.Global = True
   
   If (objRegExp.Test(mystring) = True) Then

    Set colMatches = objRegExp.Execute(mystring)

    For Each objMatch In colMatches
    Set objSubmatches = objMatch.SubMatches
    For i = 0 To objSubmatches.Count - 1
    MsgBox Trim(objSubmatches.Item(i))
    Next i
      RetStr = RetStr & "Match found at position "
      RetStr = RetStr & objMatch.FirstIndex & ". Match Value is '"
      RetStr = RetStr & objMatch.Value & "'." & vbCrLf
    Next
   Else
    RetStr = "String Matching Failed"
   End If
   
   
   TestRegExp = RetStr
End Function


Function regexp_check(myPattern As String, mystring As String) As Boolean
   Dim retBool As Boolean
   retBool = False
   Dim objRegExp As regexp
   Dim objMatch As Match
   Dim colMatches   As MatchCollection
   Dim objSubmatches As SubMatches

   Set objRegExp = New regexp
   objRegExp.Pattern = myPattern
   objRegExp.IgnoreCase = True
   objRegExp.Global = True
   
   retBool = objRegExp.Test(mystring)
   
   regexp_check = retBool
End Function

Function regexp_replace(myPattern As String, mystring As String, myreplace As String) As String
    Dim RetStr As String
    RetStr = mystring
    If regexp_check(myPattern, mystring) Then
           
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
                    RetStr = Replace(mystring, objSubmatches.Item(0), myreplace)
                End If
            End If
        End If
    End If
    regexp_replace = RetStr
End Function
Function regexp_count(myPattern As String, mystring As String) As Integer
    Dim Ret As Integer
    Ret = 0
    
    Dim objRegExp As regexp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim objSubmatches As SubMatches
    Set objRegExp = New regexp
    objRegExp.Pattern = myPattern
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    If (objRegExp.Test(mystring) = True) Then
        Set colMatches = objRegExp.Execute(mystring)   ' Execute search.
        If colMatches.Count > 0 Then
            Ret = colMatches.Count
        End If
    End If
    
    regexp_count = Ret
End Function
Sub regexp_ReturnSearchCollection(myPattern As String, mystring As String)
   Set splitString = New Collection
   
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
             If IsEmpty(objMatch.Value) = False Then
                Set splitStringItem = New clSplitString
                splitStringItem.str = Trim(CStr(objMatch.Value))
                splitStringItem.start = objMatch.FirstIndex + 1
                splitStringItem.length = objMatch.length
                
                splitString.Add splitStringItem
            End If
        Next objMatch
   End If
   
End Sub

Function regexp_ReturnSearchArray(myPattern As String, mystring As String) As String()
   Dim retArrayStr() As String
 Dim retArrayStrLength As Integer
   Dim objRegExp As regexp
   Dim objMatch As Match
   Dim colMatches   As MatchCollection
   Dim objSubmatches As SubMatches
   Dim RetStr As String
   Dim i As Integer
   
   Set objRegExp = New regexp
   objRegExp.Pattern = myPattern
   objRegExp.IgnoreCase = True
   objRegExp.Global = True
   
   If (objRegExp.Test(mystring) = True) Then

        Set colMatches = objRegExp.Execute(mystring)
        retArrayStrLength = -1
        For Each objMatch In colMatches
            Set objSubmatches = objMatch.SubMatches
            For i = 0 To objSubmatches.Count - 1
                If IsEmpty(objSubmatches.Item(i)) = False Then retArrayStrLength = retArrayStrLength + 1
            Next i
        Next objMatch
        ReDim retArrayStr(retArrayStrLength) As String
        retArrayStrLength = -1
        For Each objMatch In colMatches
            Set objSubmatches = objMatch.SubMatches
            For i = 0 To objSubmatches.Count - 1
             If IsEmpty(objSubmatches.Item(i)) = False Then
                retArrayStrLength = retArrayStrLength + 1
                retArrayStr(retArrayStrLength) = CStr(objSubmatches.Item(i))
                End If
            Next i
        Next objMatch
   End If
   
   
   regexp_ReturnSearchArray = retArrayStr
End Function

Function regexp_ReturnStringBySumOfMatches(myPattern As String, mystring As String) As String

Dim str As Variant

Dim resStr As String
resStr = ""
For Each str In mRegexp.regexp_ReturnSearchArray(myPattern, mystring)
    resStr = RTrim(resStr & " " & LTrim(str))
Next str

regexp_ReturnStringBySumOfMatches = LTrim(resStr)
End Function

Function regexp_ReturnSearch(myPattern As String, mystring As String) As String
    Dim RetStr As String
    RetStr = mystring
    If regexp_check(myPattern, mystring) Then
           
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
                If objSubmatches.Count > 0 Then
                    RetStr = objSubmatches.Item(0)
                End If
            End If
        End If
    Else
     RetStr = ""
    End If
    regexp_ReturnSearch = RetStr
End Function

