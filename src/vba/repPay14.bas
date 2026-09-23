Attribute VB_Name = "repPay14"
Option Explicit

Public Sub Pay14(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Pay14..."

    On Error GoTo CleanFail

    ' Single shared read -- both the deduction/contribution sheet and the
    ' retirement/earnings sheet are derived from the same source array so we
    ' only hit the worksheet once.
    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then GoTo CleanExit

    Dim data As Variant
    data = wsSrc.Range("A1:W" & lastRow).Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    Pay14_Worker wsSrc, wb, data, nRows, nCols
    Pay14_Retirement_Worker wsSrc, wb, data, nRows, nCols

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

' ============================================================
' DEDUCTION / CONTRIBUTION WORKER (existing behavior, unchanged
' other than reading data/nRows/nCols from the caller instead of
' re-reading the worksheet)
' ============================================================
Private Sub Pay14_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook, ByRef data As Variant, ByVal nRows As Long, ByVal nCols As Long)
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

    ' Output buffer -- grown in 5000-row chunks if needed
    Dim out() As Variant
    Dim outCount As Long, outCap As Long
    ' nRows is a provable upper bound: each output row corresponds to exactly
    ' one distinct source row (rr), so outCount can never exceed nRows. 
    outCap = nRows
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

' ============================================================
' RETIREMENT / EARNINGS WORKER (new)
' ============================================================
' Produces one output row per earnings detail line from the "Effective /
' Source / Earnings Description / ... " section that precedes the
' Deduction/Contribution section in each employee's block.
'
' Budget-code allocation sub-rows that can appear beneath an earnings line
' (same one-to-many split pattern as repPos04, e.g. a 70/30 split across two
' account codes) are intentionally excluded -- this sheet is one row per
' earning only, per spec.
'
' Two source columns are both labeled "Pay Rate" (col E and col L). Col E is
' dropped; col L is carried through as "RetBase" per spec.
Private Sub Pay14_Retirement_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook, ByRef data As Variant, ByVal nRows As Long, ByVal nCols As Long)

    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Pay14_Retirement")
    wsOut.Cells.Clear

    ' Headers
    wsOut.Range("A1").Value = "EmployeeID"
    wsOut.Range("B1").Value = "SSN_Last4"
    wsOut.Range("C1").Value = "EmployeeName"
    wsOut.Range("D1").Value = "Effective"
    wsOut.Range("E1").Value = "Source"
    wsOut.Range("F1").Value = "EarningsDescription"
    wsOut.Range("G1").Value = "RetirePlan"
    wsOut.Range("H1").Value = "ObjectCode"
    wsOut.Range("I1").Value = "AssnWork"
    wsOut.Range("J1").Value = "PC"
    wsOut.Range("K1").Value = "CC"
    wsOut.Range("L1").Value = "Units"
    wsOut.Range("M1").Value = "RetBase"
    wsOut.Range("N1").Value = "RetEarn"
    wsOut.Range("O1").Value = "Earnings"
    wsOut.Range("P1").Value = "Adjustment"
    wsOut.Range("Q1").Value = "SourceSheet"

    Dim out() As Variant
    Dim outCount As Long, outCap As Long
    outCap = nRows
    ReDim out(1 To outCap, 1 To 17)
    outCount = 0

    Dim r As Long
    r = 1

    Do While r <= nRows
        Dim aVal As String
        aVal = Trim$(CStr(NzV(GetA(data, r, 1))))   ' Col A
        Dim bValHdr As String
        bValHdr = Trim$(CStr(NzV(GetA(data, r, 2)))) ' Col B

        If LCase$(aVal) = "effective" And LCase$(bValHdr) = "source" Then
            Dim headerRow As Long
            headerRow = r

            Dim empRow As Long
            empRow = FindNearestEmployeeHeaderAbove_Arr(data, headerRow)
            If empRow = 0 Then
                Debug.Print "Pay14_Retirement_Worker: no employee header found above row " & headerRow & " on [" & wsSrc.Name & "]"
                r = r + 1
                GoTo ContinueMain
            End If

            Dim empName As String, empID As String, last4 As String
            ParseEmployeeHeader CStr(NzV(GetA(data, empRow, 1))), empName, empID, last4

            ' Find the "Total" sentinel row that closes this section. Must be an
            ' EXACT match on "Total" -- the Deduction/Contribution section further
            ' down ends with "Total Deductions, *Reductions , Contributions", which
            ' would false-match a "starts with" or "contains" check.
            Dim totalRow As Long
            totalRow = FindRowInColAEquals_Arr(data, headerRow + 1, nRows, "Total")
            If totalRow = 0 Then
                Debug.Print "Pay14_Retirement_Worker: 'Total' sentinel not found after row " & headerRow & " on [" & wsSrc.Name & "]"
                r = r + 1
                GoTo ContinueMain
            End If

            Dim startRow As Long, endRow As Long
            startRow = headerRow + 1
            endRow = totalRow - 1

            If endRow >= startRow Then
                Dim rr As Long
                For rr = startRow To endRow
                    Dim srcVal As String
                    srcVal = Trim$(CStr(NzV(GetA(data, rr, 2))))   ' col B (Source)

                    If Len(srcVal) = 0 Then
                        Debug.Print "Pay14_Retirement_Worker: blank Source cell at row " & rr & _
                                    " on [" & wsSrc.Name & "] -- skipped"
                        GoTo NextDetailRow
                    End If

                    If IsAccountCodeCell(srcVal) Then
                        ' Budget-code allocation sub-row -- excluded from this sheet by design
                        GoTo NextDetailRow
                    End If

                    Dim adjAmt As Double, unitsAmt As Double, earnAmt As Double
                    Dim ccAmt As Double, retEarnAmt As Double, retBaseAmt As Double, pcAmt As Double
                    adjAmt = 0#: unitsAmt = 0#: earnAmt = 0#
                    ccAmt = 0#: retEarnAmt = 0#: retBaseAmt = 0#: pcAmt = 0#

                    Call TryParseNumber(NzV(GetA(data, rr, 4)), adjAmt)     ' Adjustment
                    Call TryParseNumber(NzV(GetA(data, rr, 6)), unitsAmt)   ' Units
                    Call TryParseNumber(NzV(GetA(data, rr, 7)), earnAmt)    ' Earnings
                    Call TryParseNumber(NzV(GetA(data, rr, 10)), ccAmt)     ' CC
                    Call TryParseNumber(NzV(GetA(data, rr, 11)), retEarnAmt) ' Ret Earn
                    Call TryParseNumber(NzV(GetA(data, rr, 12)), retBaseAmt) ' Pay Rate (col L) -> RetBase
                    Call TryParseNumber(NzV(GetA(data, rr, 13)), pcAmt)     ' PC

                    ' ---- Look ahead for budget-code allocation sub-row(s) ----
                    ' Attribute this earnings row's ObjectCode from whichever
                    ' immediately-following account line carries the greatest
                    ' percentage (an assignment shouldn't legitimately split
                    ' across a Certificated and a Classified account, so we only
                    ' need the dominant one, not every split). Ties keep whichever
                    ' line was encountered first. This is a read-only peek -- the
                    ' outer loop still walks these same rows itself and skips them
                    ' via the IsAccountCodeCell check above, so no separate
                    ' consumption/pointer-advance is needed here.
                    Dim objCode As String: objCode = vbNullString
                    Dim bestPct As Double: bestPct = -1#
                    Dim foundBudgetRow As Boolean: foundBudgetRow = False

                    Dim peekRow As Long: peekRow = rr + 1
                    Do While peekRow <= endRow
                        Dim peekB As String
                        peekB = Trim$(CStr(NzV(GetA(data, peekRow, 2))))
                        If Not IsAccountCodeCell(peekB) Then Exit Do

                        foundBudgetRow = True

                        Dim peekPct As Double: peekPct = 0#
                        Call TryParseNumber(NzV(GetA(data, peekRow, 1)), peekPct)

                        If peekPct > bestPct Then
                            bestPct = peekPct
                            objCode = ExtractObjectCode(peekB)
                        End If

                        peekRow = peekRow + 1
                    Loop

                    If Not foundBudgetRow Then
                        Debug.Print "Pay14_Retirement_Worker: no budget-code line found for earnings row " & rr & _
                                    " on [" & wsSrc.Name & "] -- ObjectCode left blank"
                    End If

                    outCount = outCount + 1
                    If outCount > outCap Then
                        outCap = outCap + 5000
                        ReDim Preserve out(1 To outCap, 1 To 17)
                    End If

                    out(outCount, 1) = empID
                    out(outCount, 2) = last4
                    out(outCount, 3) = empName
                    out(outCount, 4) = GetA(data, rr, 1)                          ' Effective
                    out(outCount, 5) = srcVal                                     ' Source
                    out(outCount, 6) = Trim$(CStr(NzV(GetA(data, rr, 3))))        ' EarningsDescription
                    out(outCount, 7) = Trim$(CStr(NzV(GetA(data, rr, 8))))        ' RetirePlan
                    out(outCount, 8) = objCode                                    ' ObjectCode
                    out(outCount, 9) = Trim$(CStr(NzV(GetA(data, rr, 9))))        ' AssnWork
                    out(outCount, 10) = pcAmt
                    out(outCount, 11) = ccAmt
                    out(outCount, 12) = unitsAmt
                    out(outCount, 13) = retBaseAmt                               ' RetBase
                    out(outCount, 14) = retEarnAmt
                    out(outCount, 15) = earnAmt
                    out(outCount, 16) = adjAmt
                    out(outCount, 17) = wsSrc.Name

NextDetailRow:
                Next rr
            End If

            r = totalRow
        End If

ContinueMain:
        r = r + 1
    Loop

    ' Force EmployeeID and ObjectCode to text before the bulk write
    wsOut.Columns(1).NumberFormat = "@"    ' EmployeeID
    wsOut.Columns(8).NumberFormat = "@"    ' ObjectCode

    If outCount > 0 Then
        wsOut.Range("A2").Resize(outCount, 17).Value2 = out
        wsOut.Range("D2").Resize(outCount, 1).NumberFormat = "mm/dd/yyyy"   ' Effective
        wsOut.Range("M2").Resize(outCount, 1).NumberFormat = "#,##0.00"     ' RetBase
        wsOut.Range("N2").Resize(outCount, 1).NumberFormat = "#,##0.00"     ' RetEarn
        wsOut.Range("O2").Resize(outCount, 1).NumberFormat = "#,##0.00"     ' Earnings
        wsOut.Range("P2").Resize(outCount, 1).NumberFormat = "#,##0.00"     ' Adjustment
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

' Exact (trimmed, case-insensitive) match on column A -- used for the "Total"
' sentinel, which must NOT also match "Total Deductions, *Reductions , Contributions".
Private Function FindRowInColAEquals_Arr(ByRef data As Variant, ByVal startRow As Long, ByVal endRow As Long, ByVal exactText As String) As Long
    Dim r As Long
    For r = startRow To endRow
        Dim v As String
        v = Trim$(CStr(NzV(GetA(data, r, 1))))
        If LCase$(v) = LCase$(exactText) Then
            FindRowInColAEquals_Arr = r
            Exit Function
        End If
    Next r
    FindRowInColAEquals_Arr = 0
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

' Extracts the Object code segment from a Frontline account code string of the
' form Fd-Resc-Y-Goal-Func-Objt-SO-Sch-DD1-DD2 (Objt is index 5, 0-based, after
' splitting on "-"). Mirrors the same segment indexing already used in
' repBudget04.IsSalaryBenefitsExcluded, which reads the same Objt/SO segments
' for its own purposes.
' Returns an empty string if the input doesn't have enough "-"-delimited
' segments to contain an Objt position.
Private Function ExtractObjectCode(ByVal accountCode As String) As String
    Dim parts() As String
    parts = Split(accountCode, "-")
    If UBound(parts) < 5 Then Exit Function
    ExtractObjectCode = Trim$(parts(5))
End Function
