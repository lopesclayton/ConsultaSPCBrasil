Attribute VB_Name = "mdlGeral"
 Function PadL(cText, nLen As Integer) As String
    Dim nTam As Integer
    If Not IsNull(cText) Then
        nTam = nLen - Len(Trim(cText))
        If nTam > 0 Then
            PadL = Trim(cText) & Space(nTam)
        Else
            PadL = Left(Trim(cText), nLen)
        End If
        If Trim(cText) = "" Then
            PadL = Space(nTam)
        End If
    End If
End Function

 Function StrZero(ByVal nNumero, nTam)
    Dim nNum As String
    Dim nSubLen As Long
    Dim nLen As Long
    If IsNumeric(nNumero) Then
        nNum = Trim(CStr(nNumero))
    ElseIf IsNull(nNumero) Then
        nNum = "0"
    Else
        nNum = nNumero
    End If

    nNumero = Trim(nNumero)
    nLen = Len(nNum)
    If nLen < nTam Then
        nSubLen = nTam - nLen
        nNum = String$(nSubLen, "0") & nNum
    ElseIf nLen > nTam Then
        nNum = Right(nNum, nTam)
    End If
    StrZero = nNum
End Function

 Function RemoveNaoNumericos(Texto As String) As String
    Texto = Replace(Texto, ".", "")
    Texto = Replace(Texto, "-", "")
    Texto = Replace(Texto, "/", "")
    Texto = Replace(Texto, "_", "")
    Texto = Replace(Texto, "'", "")
    Texto = Replace(Texto, """", "")
    Texto = Replace(Texto, "(", "")
    Texto = Replace(Texto, ")", "")
    Texto = Replace(Texto, " ", "")
    RemoveNaoNumericos = Texto
End Function

Function CodigoLiberacao(CNPJ As String) As String
    Dim Codigo As String
    
    If Len(CNPJ) <> 14 Then
        Exit Function
    End If
    Codigo = Mid(CNPJ, 3, 2)
    Codigo = Codigo & Mid(CNPJ, 12, 2)
    Codigo = Codigo & Mid(CNPJ, 13, 2)
    Codigo = Codigo & Mid(CNPJ, 6, 7)
    Codigo = Codigo & Mid(CNPJ, 4, 2)

    Codigo = Codigo / 2009
    Codigo = Codigo / 9
    Codigo = Codigo / 1032
    Codigo = Replace(Codigo, ",", "")
    CodigoLiberacao = Codigo
    
End Function

Function GetStringBetween(ByVal str As String, ByVal str1 As String, ByVal str2 As String, _
                          Optional ByVal st As Long = 0, _
                          Optional ByVal N As Boolean) As String
    On Error Resume Next
    Dim S1, S2, S, L As Long
    Dim foundstr As String

    S1 = InStr(st + 1, str, str1, vbTextCompare)
    S2 = InStr(S1 + 1, str, str2, vbTextCompare)

    If S1 = 0 Or S2 = 0 Or IsNull(S1) Or IsNull(S2) Then
        foundstr = str
        If N = True Then foundstr = ""                                              'traz vazio se nao achou
    Else
        S = S1 + Len(str1)
        L = S2 - S
        foundstr = Mid(str, S, L)
    End If

    GetStringBetween = foundstr

    Set S1 = Nothing
    Set S2 = Nothing
    Set S1 = Nothing


End Function


Function Formata_Data_XML(ByVal Texto As String) As String
    Texto = Left(Texto, 10)
    Texto = Format(Texto, "dd/mm/yyyy")
    Formata_Data_XML = Texto
End Function


Public Function RemoveCaracterUTF8(ByVal Texto As String) As String

    Texto = Replace(Texto, "Ã‡", "Ç")
    Texto = Replace(Texto, "Ã•", "Õ")
    Texto = Replace(Texto, "Ã" & Chr(141), "Í")

    Texto = Replace(Texto, "Âº", "º")
    Texto = Replace(Texto, "Ã©", "é")
    Texto = Replace(Texto, "ÃŠ", "Ê") 'É
    Texto = Replace(Texto, "Ã‰", "É")
    Texto = Replace(Texto, "Ã”", "Ô")
    Texto = Replace(Texto, "Ã“", "Ó")


    Texto = Replace(Texto, "Ã" & Chr(129), "Á")
    Texto = Replace(Texto, "Ãƒ", "Ã")
    Texto = Replace(Texto, "Â§", "§")

    Texto = Replace(Texto, "Ãƒ", "Ã")

    Texto = Replace(Texto, "Ã¡", "á")

    RemoveCaracterUTF8 = Texto

End Function

