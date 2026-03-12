Attribute VB_Name = "modRT_String"
Option Explicit

Public Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

Public Function LeftPadDigits(ByVal s As String, ByVal totalLen As Long) As String
    Dim t As String
    t = DigitsOnly(CStr(s))
    If Len(t) >= totalLen Then
        LeftPadDigits = Right$(t, totalLen)
    Else
        LeftPadDigits = String$(totalLen - Len(t), "0") & t
    End If
End Function
