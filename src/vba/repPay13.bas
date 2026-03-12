Attribute VB_Name = "repPay13"
Option Explicit

Public Sub Pay13(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Pay13..."

    On Error GoTo CleanFail

    Pay13_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

' ============================================================
' WORKER
' ============================================================
Private Sub Pay13_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)

    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim data As Variant
    data = ur.Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)
    If nRows < 2 Then Exit Sub

    ' ---- Output columns ----
    '  1  District
    '  2  EmployeeName
    '  3  EmployeeID
    '  4  SSN_Last4
    '  5  PayCycle
    '  6  PayDate
    '  7  T
    '  8  Code
    '  9  Description
    ' 10  Date
    ' 11  Position_Vendor_HRA
    ' 12  DeductionAmount
    ' 13  ContributionAmount
    ' 14  PayRate
    ' 15  Units
    ' 16  EarningsAmount
    ' 17  BudgetCode
    ' 18  RetirementSystem
    ' 19  PayPeriod
    ' 20  CC
    ' 21  PC
    ' 22  Wrk_Assgn
    ' 23  Rate
    ' 24  SourceSheet
    Const OUT_COLS As Long = 24

    ' ---- Pre-scan: build district block lookup ----
    ' District total rows ("Total for X School District") appear AFTER the employees
    ' they belong to, so we pre-scan to record block boundaries, then look up by row
    ' in the main pass. Same pattern as repPos04 district lookup.
    Dim distBlockStart() As Long
    Dim distBlockEnd()   As Long
    Dim distName()       As String
    Dim distCount        As Long: distCount = 0
    ReDim distBlockStart(0 To 9)
    ReDim distBlockEnd(0 To 9)
    ReDim distName(0 To 9)

    Dim psBlockStart As Long: psBlockStart = 1
    Dim psR As Long
    For psR = 1 To nRows
        Dim psAV As Variant: psAV = GetA(data, psR, 1)
        If Not IsEmpty(psAV) Then
            If VarType(psAV) = vbString Then
                Dim psText As String: psText = Trim$(CStr(psAV))
                Const TOTAL_PREFIX As String = "Total for "
                If Left$(psText, Len(TOTAL_PREFIX)) = TOTAL_PREFIX Then
                    Dim dName As String
                    dName = Trim$(Mid$(psText, Len(TOTAL_PREFIX) + 1))
                    If Len(dName) > 0 Then
                        If distCount > UBound(distBlockStart) Then
                            ReDim Preserve distBlockStart(0 To distCount + 9)
                            ReDim Preserve distBlockEnd(0 To distCount + 9)
                            ReDim Preserve distName(0 To distCount + 9)
                        End If
                        distBlockStart(distCount) = psBlockStart
                        distBlockEnd(distCount) = psR
                        distName(distCount) = dName
                        distCount = distCount + 1
                        psBlockStart = psR + 1
                    End If
                End If
            End If
        End If
    Next psR

    ' ---- Output buffer ----
    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To OUT_COLS)
    Dim outRow As Long: outRow = 0

    ' ---- Carry-forward employee identity ----
    Dim curName   As String
    Dim curID     As String
    Dim curSSN4   As String
    Dim curCycle  As String
    Dim curPayDate As Variant

    curName = vbNullString
    curID = vbNullString
    curSSN4 = vbNullString
    curCycle = vbNullString
    curPayDate = Empty

    Dim r As Long
    r = 2   ' row 1 is the column header row

    Do While r <= nRows

        Dim aV As Variant: aV = GetA(data, r, 1)
        Dim aText As String
        aText = vbNullString
        If Not IsEmpty(aV) Then
            If VarType(aV) = vbString Then aText = Trim$(CStr(aV))
        End If

        ' ----------------------------------------------------------------
        ' 1. Skip report header row (row 1 already excluded above),
        '    district total rows, footnote rows, footer row, and blanks
        ' ----------------------------------------------------------------
        If Len(aText) = 0 And IsEmpty(aV) Then GoTo NextRow

        If Left$(aText, 10) = "Total for " Then GoTo NextRow
        If Left$(aText, 1) = "*" Then GoTo NextRow
        If Left$(aText, 9) = "Selection" Then GoTo NextRow

        ' ----------------------------------------------------------------
        ' 2. Employee identity row
        '    Format: "LAST, FIRST M (######) ####"
        ' ----------------------------------------------------------------
        If IsEmployeeHeader(aText) Then
            Dim tmpID As String, tmpSSN As String
            ParseEmployeeHeader aText, curName, tmpID, tmpSSN
            curID = tmpID
            curSSN4 = MaskSsn4(tmpSSN)
            curCycle = vbNullString
            curPayDate = Empty
            GoTo NextRow
        End If

        ' ----------------------------------------------------------------
        ' 3. Pay Cycle / Pay Date row
        '    A="Pay Cycle", B=cycle value, C="Pay Date", D=date value
        ' ----------------------------------------------------------------
        If aText = "Pay Cycle" Then
            curCycle = Trim$(CStr(NzV(GetA(data, r, 2))))
            curPayDate = GetA(data, r, 4)
            GoTo NextRow
        End If

        ' ----------------------------------------------------------------
        ' 4. Numeric summary rows after D/C entries -- skip
        '    Identified by col A being numeric (not a string)
        ' ----------------------------------------------------------------
        If Not IsEmpty(aV) Then
            If VarType(aV) = vbDouble Or VarType(aV) = vbInteger Or VarType(aV) = vbLong Then
                GoTo NextRow
            End If
        End If

        ' ----------------------------------------------------------------
        ' 5. Code Line 1 -- the primary detail row
        '    T value is a short uppercase string in col A (A, D, C, etc.)
        '    Col B = Code, Col C = Description, Col D = Date
        ' ----------------------------------------------------------------
        If IsTypeFlagCell(aText) Then

            ' Resolve district for this row
            Dim curDist As String
            curDist = LookupDistrictName(r, distBlockStart, distBlockEnd, distName, distCount)

            ' Read Line 1 fields
            Dim tVal         As String: tVal = aText
            Dim codeVal      As String: codeVal = Trim$(CStr(NzV(GetA(data, r, 2))))
            Dim descVal      As String: descVal = Trim$(CStr(NzV(GetA(data, r, 3))))
            Dim dateVal      As Variant: dateVal = GetA(data, r, 4)
            Dim posVendor    As String: posVendor = Trim$(CStr(NzV(GetA(data, r, 5))))
            Dim dedAmt       As Double: dedAmt = 0#
            Dim conAmt       As Double: conAmt = 0#
            Dim payRate      As Double: payRate = 0#
            Dim units        As Double: units = 0#
            Dim earnAmt      As Double: earnAmt = 0#

            Call TryParseNumber(NzV(GetA(data, r, 6)), dedAmt)
            Call TryParseNumber(NzV(GetA(data, r, 7)), conAmt)
            Call TryParseNumber(NzV(GetA(data, r, 8)), payRate)
            Call TryParseNumber(NzV(GetA(data, r, 9)), units)
            Call TryParseNumber(NzV(GetA(data, r, 10)), earnAmt)

            ' ---- Lookahead: optional account code and/or retirement line ----
            ' Advance a pointer to peek at following rows without consuming r.
            ' The main loop variable r stays on the current code line until the
            ' end of this block, then we advance it past however many supplemental
            ' rows we consumed.
            Dim peek As Long: peek = r + 1

            ' Optional account code line: col A matches digit-digit-dash pattern,
            ' all other columns empty
            Dim budgetCode As String: budgetCode = vbNullString
            If peek <= nRows Then
                Dim peekA As Variant: peekA = GetA(data, peek, 1)
                If IsAccountCodeCell(peekA) Then
                    budgetCode = Trim$(CStr(peekA))
                    peek = peek + 1
                End If
            End If

            ' Optional retirement line: col A starts with PERS or STRS
            Dim retireSys  As String: retireSys = vbNullString
            Dim payPeriod  As String: payPeriod = vbNullString
            Dim ccVal      As Double: ccVal = 0#
            Dim pcVal      As Double: pcVal = 0#
            Dim wrkAssgn   As String: wrkAssgn = vbNullString
            Dim rateVal    As Double: rateVal = 0#

            If peek <= nRows Then
                Dim retA As Variant: retA = GetA(data, peek, 1)
                If IsRetirementCell(retA) Then
                    retireSys = Trim$(CStr(retA))
                    payPeriod = Trim$(CStr(NzV(GetA(data, peek, 3))))
                    Call TryParseNumber(NzV(GetA(data, peek, 5)), ccVal)
                    Call TryParseNumber(NzV(GetA(data, peek, 7)), pcVal)
                    wrkAssgn = Trim$(CStr(NzV(GetA(data, peek, 9))))
                    Call TryParseNumber(NzV(GetA(data, peek, 11)), rateVal)
                    peek = peek + 1
                End If
            End If

            ' ---- Emit output row ----
            outRow = outRow + 1
            If outRow > UBound(outArr, 1) Then
                ReDim Preserve outArr(1 To UBound(outArr, 1) + 500, 1 To OUT_COLS)
            End If

            outArr(outRow, 1) = curDist
            outArr(outRow, 2) = curName
            outArr(outRow, 3) = curID
            outArr(outRow, 4) = curSSN4
            outArr(outRow, 5) = curCycle
            outArr(outRow, 6) = curPayDate
            outArr(outRow, 7) = tVal
            outArr(outRow, 8) = codeVal
            outArr(outRow, 9) = descVal
            outArr(outRow, 10) = dateVal
            outArr(outRow, 11) = posVendor
            outArr(outRow, 12) = dedAmt
            outArr(outRow, 13) = conAmt
            outArr(outRow, 14) = payRate
            outArr(outRow, 15) = units
            outArr(outRow, 16) = earnAmt
            outArr(outRow, 17) = budgetCode
            outArr(outRow, 18) = retireSys
            outArr(outRow, 19) = payPeriod
            outArr(outRow, 20) = ccVal
            outArr(outRow, 21) = pcVal
            outArr(outRow, 22) = wrkAssgn
            outArr(outRow, 23) = rateVal
            outArr(outRow, 24) = wsSrc.Name

            ' Advance main loop past any supplemental rows we consumed above.
            ' Subtract 1 because the loop footer does r = r + 1.
            r = peek - 1
            GoTo NextRow
        End If

        ' ----------------------------------------------------------------
        ' 6. Anything else we have not explicitly handled -- skip silently.
        '    Add Debug.Print here if unexpected rows need investigation.
        ' ----------------------------------------------------------------

NextRow:
        r = r + 1
    Loop

    ' ---- Write output sheet ----
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Pay13_Normalized")
    wsOut.Cells.Clear

    Dim hdr(1 To 1, 1 To OUT_COLS) As Variant
    hdr(1, 1) = "District"
    hdr(1, 2) = "EmployeeName"
    hdr(1, 3) = "EmployeeID"
    hdr(1, 4) = "SSN_Last4"
    hdr(1, 5) = "PayCycle"
    hdr(1, 6) = "PayDate"
    hdr(1, 7) = "T"
    hdr(1, 8) = "Code"
    hdr(1, 9) = "Description"
    hdr(1, 10) = "Date"
    hdr(1, 11) = "Position_Vendor_HRA"
    hdr(1, 12) = "DeductionAmount"
    hdr(1, 13) = "ContributionAmount"
    hdr(1, 14) = "PayRate"
    hdr(1, 15) = "Units"
    hdr(1, 16) = "EarningsAmount"
    hdr(1, 17) = "BudgetCode"
    hdr(1, 18) = "RetirementSystem"
    hdr(1, 19) = "PayPeriod"
    hdr(1, 20) = "CC"
    hdr(1, 21) = "PC"
    hdr(1, 22) = "Wrk_Assgn"
    hdr(1, 23) = "Rate"
    hdr(1, 24) = "SourceSheet"

    wsOut.Range("A1").Resize(1, OUT_COLS).Value = hdr
    wsOut.Range("A1").Resize(1, OUT_COLS).Font.Bold = True

    ' Force ID and budget code columns to text before bulk write
    wsOut.Columns(3).NumberFormat = "@"    ' EmployeeID
    wsOut.Columns(17).NumberFormat = "@"   ' BudgetCode

    ' Date column formats
    wsOut.Columns(6).NumberFormat = "mm/dd/yyyy"    ' PayDate
    wsOut.Columns(10).NumberFormat = "mm/dd/yyyy"   ' Date

    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, OUT_COLS).Value2 = outArr
    End If

End Sub

' ============================================================
' PRIVATE HELPERS
' ============================================================

' Returns True if the cell value is a short uppercase type flag (A, D, C, etc.).
' Guards against retirement codes like PERSO(1) which are also uppercase strings.
' A type flag is 1-2 characters, all uppercase alpha only (no digits, no parens).
Private Function IsTypeFlagCell(ByVal s As String) As Boolean
    If Len(s) = 0 Or Len(s) > 2 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If (ch < "A" Or ch > "Z") Then Exit Function
    Next i
    IsTypeFlagCell = True
End Function

' Returns True if the cell value looks like an account/budget code.
' Budget codes start with two digits followed by a dash: "01-...", "13-..."
' Mirrors the IsAccountCode function in repPos04.
Private Function IsAccountCodeCell(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) < 3 Then Exit Function
    If Not (Mid$(s, 1, 1) Like "#") Then Exit Function
    If Not (Mid$(s, 2, 1) Like "#") Then Exit Function
    IsAccountCodeCell = (Mid$(s, 3, 1) = "-")
End Function

' Returns True if the cell value is a retirement system code.
' These start with PERS or STRS (case-insensitive).
Private Function IsRetirementCell(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    Dim s As String: s = UCase$(Trim$(CStr(v)))
    IsRetirementCell = (Left$(s, 4) = "PERS" Or Left$(s, 4) = "STRS")
End Function

' Returns the district name for a given row using the pre-built block table.
' Returns empty string if no block covers the row.
Private Function LookupDistrictName(ByVal rowNum As Long, _
                                    ByRef blockStart() As Long, _
                                    ByRef blockEnd() As Long, _
                                    ByRef names() As String, _
                                    ByVal count As Long) As String
    LookupDistrictName = vbNullString
    If count = 0 Then Exit Function
    Dim i As Long
    For i = 0 To count - 1
        If rowNum >= blockStart(i) And rowNum <= blockEnd(i) Then
            LookupDistrictName = names(i)
            Exit Function
        End If
    Next i
End Function


