Attribute VB_Name = "repBen02"
Option Explicit

Public Sub Ben02(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Ben02..."

    On Error GoTo CleanFail

    Ben02_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub Ben02_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)
    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub
    If ur.Rows.Count < 2 Then Exit Sub

    Dim data As Variant
    data = ur.Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    ' Output columns:
    '  1 SourceSheet
    '  2 Org
    '  3 EmployeeName
    '  4 EmployeeID
    '  5 SSN_Last4
    '  6 FTE
    '  7 ProviderEffective (D-F)
    '  8 Provider (G)
    '  9 Level (H)
    ' 10 RateEffective (I-K)
    ' 11 EmployerCost (L)
    ' 12 EmployeeCost (N)
    ' 13 Cobra (O/P)
    Const OUT_COLS As Long = 13

    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To OUT_COLS)

    Dim outRow As Long: outRow = 0

    ' Carry-forward identity fields until the next identity row resets them
    Dim curName As String
    Dim curID6 As String
    Dim curSSN4 As String
    Dim curFTE As Variant

    curName = vbNullString
    curID6 = vbNullString
    curSSN4 = vbNullString
    curFTE = Empty

    ' Org is backfilled when a "Total for Org ###" trailer row is encountered
    Dim sectionStart As Long
    sectionStart = 1

    Dim r As Long
    For r = 2 To nRows

        ' ---------- Detect org section trailer row ----------
        Dim aV As Variant
        aV = data(r, 1)

        If Not IsEmpty(aV) Then
            If VarType(aV) = vbString Then
                Dim aText As String
                aText = aV

                If InStr(1, aText, "Total", vbTextCompare) > 0 Then
                    Dim org3 As String
                    If TryParseOrg3(aText, org3) Then
                        Dim i As Long
                        For i = sectionStart To outRow
                            outArr(i, 2) = org3
                        Next i
                        sectionStart = outRow + 1
                        GoTo NextRow
                    End If
                End If
            End If
        End If

        ' ---------- Detect identity row by column B "(######) ####" ----------
        Dim tmpID As String, tmpSSN As String
        If TryParseIdSsn4(data(r, 2), tmpID, tmpSSN) Then
            If Not IsEmpty(data(r, 1)) Then
                curName = Trim$(CStr(data(r, 1)))
            Else
                curName = vbNullString
            End If

            curID6 = tmpID
            curSSN4 = tmpSSN
            curFTE = data(r, 3)

            GoTo NextRow
        End If

        ' ---------- Detail row ----------
        If LenB(curID6) = 0 Then GoTo NextRow

        ' Provider (G) must exist for the row to be meaningful
        Dim provV As Variant
        provV = data(r, 7)
        If IsEmpty(provV) Then GoTo NextRow

        Dim provider As String
        provider = Trim$(CStr(provV))
        If LenB(provider) = 0 Then GoTo NextRow

        outRow = outRow + 1

        outArr(outRow, 1) = wsSrc.Name
        outArr(outRow, 3) = curName
        outArr(outRow, 4) = curID6
        outArr(outRow, 5) = MaskSsn4(curSSN4)
        outArr(outRow, 6) = curFTE

        ' Provider Effective: D-F
        outArr(outRow, 7) = Eff3(data(r, 4), data(r, 5), data(r, 6))

        ' Provider / Level
        outArr(outRow, 8) = provider
        outArr(outRow, 9) = data(r, 8)

        ' Rate Effective: I-K
        outArr(outRow, 10) = Eff3(data(r, 9), data(r, 10), data(r, 11))

        ' Employer / Employee
        outArr(outRow, 11) = data(r, 12)
        If nCols >= 14 Then outArr(outRow, 12) = data(r, 14)

        ' Cobra: prefer col O; fall back to col P
        ' TODO: simplify once Cobra data is unified in system
        If nCols >= 16 Then
            Dim vO As Variant, vP As Variant
            vO = data(r, 15)
            vP = data(r, 16)

            If Not IsEmpty(vO) Then
                If LenB(Trim$(CStr(vO))) > 0 Then
                    outArr(outRow, 13) = vO
                ElseIf Not IsEmpty(vP) Then
                    If LenB(Trim$(CStr(vP))) > 0 Then outArr(outRow, 13) = vP
                End If
            ElseIf Not IsEmpty(vP) Then
                If LenB(Trim$(CStr(vP))) > 0 Then outArr(outRow, 13) = vP
            End If

        ElseIf nCols >= 15 Then
            outArr(outRow, 13) = data(r, 15)
        End If

NextRow:
    Next r

    ' If the source ends without a trailing "Total for Org" row, the final section
    ' never gets backfilled. Log a warning and leave Org blank for those rows.
    If sectionStart <= outRow Then
        Debug.Print "Ben02_Worker: rows " & sectionStart & " to " & outRow & _
                    " have no org trailer on sheet [" & wsSrc.Name & "] — Org left blank"
    End If

    ' Write output
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Ben02_Normalized")
    wsOut.Cells.Clear

    Dim hdr(1 To 1, 1 To OUT_COLS) As Variant
    hdr(1, 1) = "SourceSheet"
    hdr(1, 2) = "Org"
    hdr(1, 3) = "EmployeeName"
    hdr(1, 4) = "EmployeeID"
    hdr(1, 5) = "SSN_Last4"
    hdr(1, 6) = "FTE"
    hdr(1, 7) = "ProviderEffective"
    hdr(1, 8) = "Provider"
    hdr(1, 9) = "Level"
    hdr(1, 10) = "RateEffective"
    hdr(1, 11) = "EmployerCost"
    hdr(1, 12) = "EmployeeCost"
    hdr(1, 13) = "Cobra"

    wsOut.Range("A1").Resize(1, OUT_COLS).Value = hdr

    wsOut.Columns(4).NumberFormat = "@"

    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, OUT_COLS).Value = outArr
    End If
End Sub

' ---- helpers ----

' Builds an effective-date range string from three columns (start / mid / end).
' Returns whichever endpoints are non-empty, joined with "-".
Private Function Eff3(ByVal v1 As Variant, ByVal v2 As Variant, ByVal v3 As Variant) As String
    Dim p1 As String, p3 As String
    p1 = ToEffPart(v1)
    p3 = ToEffPart(v3)

    If LenB(p1) = 0 Then
        Eff3 = p3
    ElseIf LenB(p3) = 0 Then
        Eff3 = p1
    Else
        Eff3 = p1 & "-" & p3
    End If
End Function

Private Function ToEffPart(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    If IsDate(v) Then
        ToEffPart = Format$(CDate(v), "mm/dd/yy")
    Else
        ToEffPart = Trim$(CStr(v))
    End If
End Function


