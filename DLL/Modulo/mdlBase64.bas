Attribute VB_Name = "mdlBase64"
Public Function Base64Encode(inData As String) As String
  Const Base64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim sOut As String, I As Integer
'cOut As String,

  For I = 1 To Len(inData) Step 3
    Dim nGroup As Long, pOut As String
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))

    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)

    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup

    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) & Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) & Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) & Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)

    sOut = sOut & pOut
  Next I

  If Len(sOut) > 76 Then sOut = Base64ToLong(sOut)

  Select Case Len(inData) Mod 3
  Case 1  '8 bit final
    sOut = Left(sOut, Len(sOut) - 2) & "=="
  Case 2  '16 bit final
    sOut = Left(sOut, Len(sOut) - 1) & "="
  End Select
  Base64Encode = sOut
End Function

Private Function MyASC(OneChar As String) As Integer
  If OneChar = "" Then
    MyASC = 0
  Else
    MyASC = Asc(OneChar)
  End If
End Function

Public Function Base64ToLong(ByVal base64String As String) As String
  Dim tmpBase64 As String, tmpResult As String
  tmpResult = ""
  tmpBase64 = base64String
  Do While tmpBase64 <> ""
    tmpResult = tmpResult & Mid(tmpBase64, 1, 76) & Chr(13) & Chr(10)
    tmpBase64 = Mid(tmpBase64, 77, Len(tmpBase64))
    If Len(tmpBase64) < 76 Then
      tmpResult = tmpResult & tmpBase64
      tmpBase64 = ""
    End If
  Loop
  Base64ToLong = tmpResult

End Function


