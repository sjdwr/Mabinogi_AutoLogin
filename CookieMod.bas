Attribute VB_Name = "CookieMod"
Private MyCookieName() As String
Private MyCookieValue() As String
Private MyCookieCount As Long

Private Sub setCkTTP(findString As String, requestHeader As String)
On Error GoTo Err:
    Dim stHdrPos As Long
    Dim edHdrPos As Long
    Dim tokHdr_st As Long
    Dim tokHdr_ed As Long
    Dim tokHdr_tk As Long
    Dim splitSetCk As String
    Dim name As String
    Dim value As String
    Dim index As Long
    
    stHdrPos = InStr(1, requestHeader, findString)
    
    Do While stHdrPos > 0
        edHdrPos = InStr(stHdrPos, requestHeader, vbCrLf)
        
        If edHdrPos > 0 Then
            splitSetCk = Mid(requestHeader, stHdrPos + Len(findString), edHdrPos - stHdrPos - Len(findString))
  
            tokHdr_st = 1
            tokHdr_tk = InStr(1, splitSetCk, "=")
  
            Do While tokHdr_tk > 0
                name = LTrim(RTrim(Mid(splitSetCk, tokHdr_st, tokHdr_tk - tokHdr_st)))
                
                tokHdr_tk = tokHdr_tk + 1
                tokHdr_ed = InStr(tokHdr_tk, splitSetCk, ";")
                
                If tokHdr_ed > 0 Then
                    value = LTrim(RTrim(Mid(splitSetCk, tokHdr_tk, tokHdr_ed - tokHdr_tk)))
                    tokHdr_st = tokHdr_ed + 1
                Else
                    value = LTrim(RTrim(Mid(splitSetCk, tokHdr_tk)))
                End If
                
                If MyCookieCount = 0 Then
                    ReDim MyCookieName(1)
                    ReDim MyCookieValue(1)
                ElseIf MyCookieCount > 0 Then
                    For index = 0 To MyCookieCount - 1
                        If MyCookieName(index) = name Then
                            MyCookieValue(index) = value
                            Exit For
                        End If
                    Next index
                    
                    ReDim Preserve MyCookieName(MyCookieCount + 1)
                    ReDim Preserve MyCookieValue(MyCookieCount + 1)
                End If
                
                MyCookieName(MyCookieCount) = name
                MyCookieValue(MyCookieCount) = value
                
                MyCookieCount = MyCookieCount + 1
                
                If tokHdr_ed > 0 Then
                    tokHdr_tk = InStr(tokHdr_ed + 1, splitSetCk, "=")
                Else
                    tokHdr_tk = 0
                End If
            Loop
  
        End If
        
        stHdrPos = InStr(edHdrPos + 2, requestHeader, findString)
    Loop
Err:
End Sub

Public Sub UnloadCookie()
    Erase MyCookieName
    Erase MyCookieValue
    MyCookieCount = 0
End Sub

Public Sub setCookie(requestHeader As String)
    Call setCkTTP("Set-Cookie:", requestHeader)
    Call setCkTTP("set-cookie:", requestHeader)
    Call setCkTTP("Cookie:", requestHeader)
    Call setCkTTP("cookie:", requestHeader)
End Sub

Public Function getCookie(parameter As String) As String
On Error GoTo Err:
    Dim index As Long
    For index = 0 To MyCookieCount - 1
        If MyCookieName(index) = parameter Then
            getCookie = MyCookieValue(index)
            Exit For
        End If
    Next index
Err:
End Function

Public Function getCookieAll() As String
On Error GoTo Err:

    Dim index As Long
    Dim str As String
    
    If MyCookieCount = 0 Then
        getCookieAll = vbNullString
        Exit Function
    End If
    
    str = MyCookieName(0) & "=" & MyCookieValue(0)
    
    For index = 1 To MyCookieCount - 1
        str = str & "; " & MyCookieName(index) & "=" & MyCookieValue(index)
    Next index
    
    getCookieAll = str
Err:
End Function
