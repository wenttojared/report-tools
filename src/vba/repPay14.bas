Attribute VB_Name = "repPay14"
Option Explicit

Public Sub Pay14(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Pay14..."

    On Error GoTo CleanFail

    Pay14_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub Pay14_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Pay14_Normalized")
    wsOut.Cells.Clear

    ' Headers
    wsOut.Range("A1").Value = "EmployeeID"
    wsOut.Range("B1").Value = "SSN_Last4"
    wsOut.Range("C1").Value = "EmployeeName"
    wsOut.Range("D1").Value = "PayDate"
    wsOut.Range("E1").Value = "EffectiveDate"
    wsOut.Range("F1").Value = "NetPay"
    wsOut.Range("G1").Value = "VendorID"
    wsOut.Range("H1").Value = "VendorAddrID"
    wsOut.Range("I1").Value = "VendorType"
    wsOut.Range("J1").Value = "DedContribName"
    wsOut.Range("K1").Value = "DeductionAmount"
    wsOut.Range("L1").Value = "ContributionAmount"
    wsOut.Range("M1").Value = "SubjectGross_Ded"
    wsOut.Range("N1").Value = "SubjectGross_Contrib"
    wsOut.Range("O1").Value = "SourceSheet"

    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = wsSrc.Range("A1:W" & lastRow).Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    ' Output buffer -- grown in 5000-row chunks if needed
    Dim out() As Variant
    Dim outCount As Long, outCap As Long
    outCap = 5000
    ReDim out(1 To outCap, 1 To 15)
    outCount = 0

    ' Column index for the contribution amount ("CC" header); detected per employee block
    Const COL_CC_FALLBACK As Long = 8

    Dim r As Long
    r = 1

    Do While r <= nRows
        Dim bVal As String
        bVal = LCase$(Trim$(CStr(NzV(GetA(data, r, 2)))))  ' Col B

        If bVal = "deduction/contribution" Then
            Dim headerRow As Long
            headerRow = r

            ' Locate the nearest employee identity row above this header
            Dim empRow As Long
            empRow = FindNearestEmployeeHeaderAbove_Arr(data, headerRow)
            If empRow = 0 Then
                Debug.Print "Pay14_Worker: no employee header found above row " & headerRow & " on [" & wsSrc.Name & "]"
                r = r + 1
                GoTo ContinueMain
            End If

            Dim empName As String, empID As String, last4 As String
            ParseEmployeeHeader CStr(NzV(GetA(data, empRow, 1))), empName, empID, last4

            ' Pay Date label in the employee row; value is 11 columns to the right
            Const PAY_DATE_OFFSET As Long = 11
            Dim payDate As Variant
            payDate = FindHeaderRowValue_Arr(data, empRow, nCols, "Pay Date", PAY_DATE_OFFSET)

            ' Net pay from col W (23) in the employee row
            Dim netPay As Double: netPay = 0#
            Dim tmp As Double
            If TryParseNumber(NzV(GetA(data, empRow, 23)), tmp) Then netPay = tmp

            ' Detect the contribution column ("CC") from the deduction/contribution header row.
            ' The column position can shift between report versions, so we prefer detection
            ' over hardcoding and warn if it falls back to the default.
            Dim colCC As Long
            colCC = FindColInRowExact_Arr(data, headerRow, nCols, "CC")
            If colCC = 0 Then
                Debug.Print "Pay14_Worker: 'CC' column not found in header at row " & headerRow & _
                            " on [" & wsSrc.Name & "] — falling back to col " & COL_CC_FALLBACK
                colCC = COL_CC_FALLBACK
            End If

            ' Find the "Total Deductions" sentinel row; detail rows live between header+1 and sentinel-1
            Dim totalRow As Long
            totalRow = FindRowInColAContains_Arr(data, headerRow + 1, nRows, "total deductions")
            If totalRow = 0 Then
                Debug.Print "Pay14_Worker: 'total deductions' row not found after row " & headerRow & " on [" & wsSrc.Name & "]"
                r = r + 1
                GoTo ContinueMain
            End If

            Dim startRow As Long, endRow As Long
            startRow = headerRow + 1
            endRow = totalRow - 1
            If endRow < startRow Then
                r = totalRow
                GoTo ContinueMain
            End If

            Dim rr As Long
            For rr = startRow To endRow
                Dim itemName As String
                itemName = Trim$(CStr(NzV(GetA(data, rr, 2)))) ' B
                If Len(itemName) > 0 Then
                    Dim eff As Variant
                    eff = GetA(data, rr, 1) ' A

                    Dim subjDed As Double, dedAmt As Double, subjCon As Double, conAmt As Double
                    subjDed = 0#: dedAmt = 0#: subjCon = 0#: conAmt = 0#
                    Call TryParseNumber(NzV(GetA(data, rr, 4)), subjDed)        ' D
                    Call TryParseNumber(NzV(GetA(data, rr, 5)), dedAmt)         ' E
                    Call TryParseNumber(NzV(GetA(data, rr, 7)), subjCon)        ' G
                    Call TryParseNumber(NzV(GetA(data, rr, colCC)), conAmt)     ' CC

                    outCount = outCount + 1
                    If outCount > outCap Then
                        outCap = outCap + 5000
                        ReDim Preserve out(1 To outCap, 1 To 15)
                    End If

                    Dim vendorType As String, vendorID As String, vendorAddrID As String
                    If Not TryParseVendor(NzV(GetA(data, rr, 3)), vendorID, vendorAddrID, vendorType) Then
                        Debug.Print "Pay14_Worker: vendor parse failed at row " & rr & " on [" & wsSrc.Name & "] — [" & CStr(NzV(GetA(data, rr, 3))) & "]"
                    End If

                    out(outCount, 1) = empID
                    out(outCount, 2) = last4
                    out(outCount, 3) = empName
                    out(outCount, 4) = payDate
                    out(outCount, 5) = eff
                    out(outCount, 6) = netPay
                    out(outCount, 7) = vendorID
                    out(outCount, 8) = vendorAddrID
                    out(outCount, 9) = vendorType
                    out(outCount, 10) = itemName
                    out(outCount, 11) = dedAmt
                    out(outCount, 12) = conAmt
                    out(outCount, 13) = subjDed
                    out(outCount, 14) = subjCon
                    out(outCount, 15) = wsSrc.Name
                End If
            Next rr

            r = totalRow
        End If

ContinueMain:
        r = r + 1
    Loop

    ' Force ID columns to text before writing so leading zeroes are preserved.
    ' Must be applied before the bulk write -- setting it after loses the zeroes.
    wsOut.Columns(1).NumberFormat = "@"   ' EmployeeID
    wsOut.Columns(7).NumberFormat = "@"   ' VendorID
    wsOut.Columns(8).NumberFormat = "@"   ' VendorAddrID

    ' Write output directly from the working buffer -- no need for a separate copy
    If outCount > 0 Then
        wsOut.Range("A2").Resize(outCount, 15).Value2 = out
        wsOut.Range("D2").Resize(outCount, 1).NumberFormat = "mm/dd/yyyy"  ' PayDate
        wsOut.Range("E2").Resize(outCount, 1).NumberFormat = "mm/dd/yyyy"  ' EffectiveDate
    End If

End Sub

' ----- helpers -----

Private Function FindRowInColAContains_Arr(ByRef data As Variant, ByVal startRow As Long, ByVal endRow As Long, ByVal containsLower As String) As Long
    Dim r As Long
    For r = startRow To endRow
        Dim v As String
        v = LCase$(Trim$(CStr(NzV(GetA(data, r, 1)))))
        If Len(v) > 0 Then
            If InStr(1, v, containsLower, vbTextCompare) > 0 Then
                FindRowInColAContains_Arr = r
                Exit Function
            End If
        End If
    Next r
    FindRowInColAContains_Arr = 0
End Function

Private Function FindNearestEmployeeHeaderAbove_Arr(ByRef data As Variant, ByVal startRow As Long) As Long
    Dim r As Long
    For r = startRow To 1 Step -1
        Dim aVal As String
        aVal = Trim$(CStr(NzV(GetA(data, r, 1))))
        If IsEmployeeHeader(aVal) Then
            FindNearestEmployeeHeaderAbove_Arr = r
            Exit Function
        End If
    Next r
    FindNearestEmployeeHeaderAbove_Arr = 0
End Function

Private Function FindColInRowExact_Arr(ByRef data As Variant, ByVal r As Long, ByVal nCols As Long, ByVal exactText As String) As Long
    Dim c As Long
    For c = 1 To nCols
        Dim v As String
        v = Trim$(CStr(NzV(GetA(data, r, c))))
        If Len(v) > 0 Then
            If LCase$(v) = LCase$(exactText) Then
                FindColInRowExact_Arr = c
                Exit Function
            End If
        End If
    Next c
    FindColInRowExact_Arr = 0
End Function

Private Function FindHeaderRowValue_Arr(ByRef data As Variant, ByVal r As Long, ByVal nCols As Long, ByVal labelText As String, ByVal valueOffset As Long) As Variant
    Dim c As Long
    For c = 1 To nCols
        Dim v As String
        v = Trim$(CStr(NzV(GetA(data, r, c))))
        If Len(v) > 0 Then
            If LCase$(v) = LCase$(labelText) Then
                FindHeaderRowValue_Arr = GetA(data, r, c + valueOffset)
                Exit Function
            End If
        End If
    Next c
    FindHeaderRowValue_Arr = Empty
End Function


