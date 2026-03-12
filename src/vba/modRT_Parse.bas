Attribute VB_Name = "modRT_Parse"
Option Explicit

' Expected format: "(######) ####"  ->  6-digit ID inside parens, space, 4-digit SSN suffix
Public Function TryParseIdSsn4(ByVal v As Variant, ByRef empId6 As String, ByRef ssn4 As String) As Boolean
    Dim t As String
    t = Trim$(CStr(v))

    empId6 = vbNullString
    ssn4 = vbNullString

    If Len(t) = 0 Then Exit Function   ' blank cell — not a format error, skip silently

    If Len(t) < 13 Then
        Debug.Print "TryParseIdSsn4: too short to match pattern [" & t & "]"
        Exit Function
    End If

    If Left$(t, 1) <> "(" Then
        Debug.Print "TryParseIdSsn4: no leading '(' [" & t & "]"
        Exit Function
    End If

    Dim pClose As Long
    pClose = InStr(1, t, ")", vbBinaryCompare)
    If pClose <> 8 Then   ' "(######)" => ) must be at position 8
        Debug.Print "TryParseIdSsn4: closing ')' not at position 8 [" & t & "]"
        Exit Function
    End If

    Dim idPart As String
    idPart = Mid$(t, 2, 6)
    If Not IsNumeric(idPart) Then
        Debug.Print "TryParseIdSsn4: ID portion not numeric [" & idPart & "]"
        Exit Function
    End If

    Dim ssnPart As String
    ssnPart = Trim$(Mid$(t, pClose + 1))
    If Len(ssnPart) <> 4 Then
        Debug.Print "TryParseIdSsn4: SSN suffix not 4 digits [" & ssnPart & "]"
        Exit Function
    End If
    If Not IsNumeric(ssnPart) Then
        Debug.Print "TryParseIdSsn4: SSN suffix not numeric [" & ssnPart & "]"
        Exit Function
    End If

    empId6 = LeftPadDigits(idPart, 6)
    ssn4 = LeftPadDigits(ssnPart, 4)

    TryParseIdSsn4 = True
End Function

Public Function MaskSsn4(ByVal ssn4 As String) As String
    If Len(ssn4) = 4 And IsNumeric(ssn4) Then
        MaskSsn4 = "XXX-XX-" & ssn4
    Else
        MaskSsn4 = vbNullString
    End If
End Function

' Expects a trailer row like "Total for Org 123 ..." — extracts the 3-digit org code
Public Function TryParseOrg3(ByVal v As Variant, ByRef org3 As String) As Boolean
    Dim t As String: t = Trim$(CStr(v))
    org3 = vbNullString
    If Len(t) = 0 Then Exit Function

    If InStr(1, t, "Org", vbTextCompare) = 0 Then
        Debug.Print "TryParseOrg3: 'Org' not found in [" & t & "]"
        Exit Function
    End If

    ' Walk backward collecting digits until we have 3
    Dim i As Long, digits As String, ch As String
    digits = vbNullString

    For i = Len(t) To 1 Step -1
        ch = Mid$(t, i, 1)
        If ch Like "#" Then
            digits = ch & digits
            If Len(digits) = 3 Then Exit For
        End If
    Next i

    If Len(digits) <> 3 Then
        Debug.Print "TryParseOrg3: could not extract 3-digit org from [" & t & "]"
        Exit Function
    End If

    org3 = digits
    TryParseOrg3 = True
End Function

' Expected format: "(NNNNNN/NNN) VendorType"  ->  6-digit VendorID, 3-digit VendorAddrID, text label
Public Function TryParseVendor(ByVal v As Variant, ByRef vendorID As String, ByRef vendorAddrID As String, ByRef vendorType As String) As Boolean
    vendorID = vbNullString
    vendorAddrID = vbNullString
    vendorType = vbNullString

    Dim t As String
    t = Trim$(CStr(v))

    If Len(t) = 0 Then Exit Function

    ' Minimum viable length: "(NNNNNN/NNN) " = 13 chars before any label
    If Len(t) < 13 Then
        Debug.Print "TryParseVendor: too short to match pattern [" & t & "]"
        Exit Function
    End If

    If Left$(t, 1) <> "(" Then
        Debug.Print "TryParseVendor: no leading '(' [" & t & "]"
        Exit Function
    End If

    ' VendorID: positions 2-7 (6 digits)
    Dim idPart As String
    idPart = Mid$(t, 2, 6)
    If Not IsNumeric(idPart) Then
        Debug.Print "TryParseVendor: VendorID portion not numeric [" & idPart & "] in [" & t & "]"
        Exit Function
    End If

    ' Separator must be '/' at position 8
    If Mid$(t, 8, 1) <> "/" Then
        Debug.Print "TryParseVendor: expected '/' at position 8 in [" & t & "]"
        Exit Function
    End If

    ' VendorAddrID: positions 9-11 (3 digits)
    Dim addrPart As String
    addrPart = Mid$(t, 9, 3)
    If Not IsNumeric(addrPart) Then
        Debug.Print "TryParseVendor: VendorAddrID portion not numeric [" & addrPart & "] in [" & t & "]"
        Exit Function
    End If

    ' Closing paren must be at position 12
    If Mid$(t, 12, 1) <> ")" Then
        Debug.Print "TryParseVendor: expected ')' at position 12 in [" & t & "]"
        Exit Function
    End If

    vendorID = idPart                       ' preserve leading zeroes as string
    vendorAddrID = addrPart                 ' preserve leading zeroes as string
    vendorType = Trim$(Mid$(t, 13))         ' everything after "(NNNNNN/NNN)"

    TryParseVendor = True
End Function

' Handles numeric strings, currency formatting, parenthetical negatives, and non-breaking spaces
Public Function TryParseNumber(ByVal v As Variant, ByRef outNum As Double) As Boolean
    On Error GoTo Fail

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then GoTo Fail

    If IsNumeric(v) Then
        outNum = CDbl(v)
        TryParseNumber = True
        Exit Function
    End If

    If VarType(v) = vbString Then
        Dim s As String: s = Trim$(CStr(v))
        If s = "" Then GoTo Fail

        Dim neg As Boolean: neg = False
        If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
            neg = True
            s = Mid$(s, 2, Len(s) - 2)
        End If

        s = Replace(s, "$", "")
        s = Replace(s, ",", "")
        s = Replace(s, " ", "")
        s = Replace(s, ChrW(&HA0), "")  ' non-breaking space common in ERP exports

        If IsNumeric(s) Then
            outNum = CDbl(s)
            If neg Then outNum = -outNum
            TryParseNumber = True
            Exit Function
        End If
    End If

Fail:
    TryParseNumber = False
End Function

