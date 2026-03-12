Attribute VB_Name = "repPos04"
Option Explicit

Public Sub Pos04(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Pos04..."

    On Error GoTo CleanFail

    Pos04_Worker wsSrc, wb

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
Private Sub Pos04_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)

    ' ---- Read source to array ----
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
    ' 1  OrgID
    ' 2  BU
    ' 3  AssignType
    ' 4  Employee
    ' 5  EmployeeID    (Pos# from col B of header row — text-formatted)
    ' 6  Location      (col F, text-formatted, left-padded to 4 digits)
    ' 7  JobCategory   (col G)
    ' 8  JobClass      (col H)
    ' 9  CalendarDays  (extracted from col I parens)
    ' 10 Placement     (col J, label before parens)
    ' 11 Rate          (col J, value inside parens)
    ' 12 StartDate     (parsed from col E "MM/DD - MM/DD/YY")
    ' 13 EndDate       (parsed from col E)
    ' 14 FTE_Authorized (col K)
    ' 15 FTE_Assigned  (col L)
    ' 16 BudgetCode    (col D of detail rows)
    ' 17 AccountPct    (col E of detail rows, numeric %)
    ' 18 Amount        (col N salary * AccountPct / 100)
    ' 19 SourceSheet
    Const OUT_COLS As Long = 19

    Dim outArr() As Variant
    ReDim outArr(1 To nRows * 3, 1 To OUT_COLS)   ' generous upper bound; trimmed at end
    Dim outRow As Long: outRow = 0

    ' ---- Pre-scan: build district ID lookup table ----
    ' Each district block ends with "Totals for NNN - Name" AFTER all its employee
    ' rows, so carry-forward stamping would always be one district behind.
    ' Instead, pre-scan the whole array to record every district total row and
    ' the row number where that district's block STARTS (i.e. the row after the
    ' previous district's total, or row 1 for the first district).
    ' The main loop then does a lookup: given the current row, which district
    ' block does it fall in?
    '
    ' distBlockStart(i) = first data row of district i's block
    ' distBlockEnd(i)   = the "Totals for NNN" row itself (last row of block)
    ' distID(i)         = the parsed district ID string
    Dim distBlockStart() As Long
    Dim distBlockEnd()   As Long
    Dim distIDArr()      As String
    Dim distCount        As Long: distCount = 0
    ReDim distBlockStart(0 To 9)
    ReDim distBlockEnd(0 To 9)
    ReDim distIDArr(0 To 9)

    Dim psR As Long
    Dim psBlockStart As Long: psBlockStart = 1   ' first block starts at row 1

    For psR = 2 To nRows
        Dim psAV As Variant: psAV = data(psR, 1)
        If Not IsEmpty(psAV) Then
            If VarType(psAV) = vbString Then
                Dim psText As String: psText = Trim$(CStr(psAV))
                If Left$(psText, 10) = "Totals for" Then
                    Dim psOrg As String
                    If TryParseDistrictID(psText, psOrg) Then
                        If distCount > UBound(distBlockStart) Then
                            ReDim Preserve distBlockStart(0 To distCount + 9)
                            ReDim Preserve distBlockEnd(0 To distCount + 9)
                            ReDim Preserve distIDArr(0 To distCount + 9)
                        End If
                        distBlockStart(distCount) = psBlockStart
                        distBlockEnd(distCount)   = psR
                        distIDArr(distCount)      = psOrg
                        distCount = distCount + 1
                        psBlockStart = psR + 1   ' next block starts after this total row
                    End If
                End If
            End If
        End If
    Next psR

    ' Helper: given a row number, return the district ID it belongs to.
    ' Returns empty string if no district block covers this row.
    ' (Inline lookup used in main loop via LookupDistrictID function below.)

    Dim curOrg As String:    curOrg = vbNullString
    Dim curBU As String:     curBU = vbNullString

    ' Pending employee headers: each slot is a 14-element array.
    ' Slot indices:
    '   0  AssignType (String)
    '   1  Pos# / EmployeeID (Variant — may be numeric or empty)
    '   2  Employee name (String)
    '   3  StartDate (Date or String)
    '   4  EndDate (Date or String)
    '   5  Location (String, text-padded)
    '   6  JobCategory (Variant)
    '   7  JobClass (Variant)
    '   8  CalendarDays (String)
    '   9  Placement (String)
    '  10  Rate (String)
    '  11  FTE_Authorized (Variant)
    '  12  FTE_Assigned (Variant)
    '  13  Salary from col N (Double) — used to compute Amount
    Dim pendingHeaders() As Variant
    Dim pendingCount As Long:  pendingCount = 0
    ReDim pendingHeaders(0 To 9)

    ' Track whether any budget-code rows have been emitted since the last
    ' employee header row was appended to pending.  When a new employee header
    ' arrives, if detailEmittedSinceLastHeader = True we clear pending first
    ' (new employee group); if False we append.
    Dim detailEmittedSinceLastHeader As Boolean
    detailEmittedSinceLastHeader = False

    Dim r As Long
    For r = 1 To nRows

        Dim aV As Variant: aV = data(r, 1)
        Dim aText As String
        aText = vbNullString
        If Not IsEmpty(aV) Then
            If VarType(aV) = vbString Then aText = Trim$(aV)
        End If

        ' ---- Skip true header row (row 1) ----
        If r = 1 Then GoTo NextRow

        ' ----------------------------------------------------------------
        ' 1. Org/district total row — clear state, district ID comes from lookup
        ' ----------------------------------------------------------------
        If Left$(aText, 10) = "Totals for" Then
            pendingCount = 0
            detailEmittedSinceLastHeader = False
            GoTo NextRow
        End If

        ' Resolve curOrg for this row via the pre-built district block table
        curOrg = LookupDistrictID(r, distBlockStart, distBlockEnd, distIDArr, distCount)

        ' ----------------------------------------------------------------
        ' 2. Bargaining Unit header row
        '    "Bargaining Unit CEMG - Certificated Management"
        ' ----------------------------------------------------------------
        If Left$(aText, 16) = "Bargaining Unit " Then
            curBU = ParseBargainingUnit(aText)
            pendingCount = 0
            detailEmittedSinceLastHeader = False
            GoTo NextRow
        End If

        ' ----------------------------------------------------------------
        ' 3. FTE Recap block rows — skip
        ' ----------------------------------------------------------------
        If Left$(aText, 14) = "FTE Recap As o" Then GoTo NextRow

        ' ----------------------------------------------------------------
        ' 4. Selection footer row (last row) — skip
        ' ----------------------------------------------------------------
        If Left$(aText, 9) = "Selection" Then GoTo NextRow

        ' ----------------------------------------------------------------
        ' 5. Employee header row detection
        '    Identified by: col E contains a date range "MM/DD - MM/DD/YY",
        '    AND col D does not look like a budget code (budget codes start
        '    with digit-digit-dash).
        ' ----------------------------------------------------------------
        Dim dV As Variant: dV = data(r, 4)
        Dim eV As Variant: eV = data(r, 5)

        Dim eText As String: eText = vbNullString
        If Not IsEmpty(eV) Then eText = Trim$(CStr(eV))

        Dim isEmpHeader As Boolean: isEmpHeader = False
        If InStr(1, eText, " - ") > 0 And InStr(1, eText, "/") > 0 Then
            ' Confirm col D is NOT a budget code (budget codes contain "-" early)
            Dim dText As String: dText = vbNullString
            If Not IsEmpty(dV) Then dText = Trim$(CStr(dV))
            If Not IsAccountCode(dText) Then
                isEmpHeader = True
            End If
        End If

        If isEmpHeader Then
            ' If we've already emitted budget-code rows since the last header,
            ' this is a new employee group — clear pending.
            ' If not, it's a continuation (e.g. mid-year split).
            If detailEmittedSinceLastHeader Then
                pendingCount = 0
                detailEmittedSinceLastHeader = False
            End If

            If pendingCount > UBound(pendingHeaders) Then
                ReDim Preserve pendingHeaders(0 To pendingCount + 9)
            End If

            ' Parse fields
            Dim bV As Variant: bV = data(r, 2)   ' Pos# (EmployeeID)

            Dim sDate As Date, eDate As Date
            Dim hasDate As Boolean
            hasDate = TryParseDateRange(eText, sDate, eDate)

            ' Salary is in col N (14) of the EMPLOYEE HEADER ROW
            Dim salary As Double: salary = 0#
            If nCols >= 14 Then
                Call TryParseNumber(NzV(data(r, 14)), salary)
            End If

            ' Location: col F — numeric stored as number, force to 4-char string
            Dim locStr As String
            locStr = LeftPadDigits(Trim$(CStr(NzV(data(r, 6)))), 4)

            ' Placement / Rate from col J
            Dim placement As String, rate As String
            ParsePlacementRate CStr(NzV(data(r, 10))), placement, rate

            Dim hdr(0 To 13) As Variant
            hdr(0)  = aText                                  ' AssignType
            hdr(1)  = NzV(bV)                                ' EmployeeID (Pos#)
            hdr(2)  = Trim$(CStr(NzV(dV)))                   ' Employee name
            If hasDate Then
                hdr(3) = sDate
                hdr(4) = eDate
            Else
                hdr(3) = eText
                hdr(4) = Empty
            End If
            hdr(5)  = locStr                                 ' Location
            hdr(6)  = NzV(data(r, 7))                        ' JobCategory
            hdr(7)  = NzV(data(r, 8))                        ' JobClass
            hdr(8)  = ParseCalendarDays(NzV(data(r, 9)))     ' CalendarDays
            hdr(9)  = placement                              ' Placement
            hdr(10) = rate                                   ' Rate
            hdr(11) = NzV(data(r, 11))                       ' FTE Authorized
            hdr(12) = NzV(data(r, 12))                       ' FTE Assigned
            hdr(13) = salary                                 ' Salary (col N)

            pendingHeaders(pendingCount) = hdr
            pendingCount = pendingCount + 1
            GoTo NextRow
        End If

        ' ----------------------------------------------------------------
        ' 6. Budget code detail row
        '    Col D = account code string  (contains multiple "-")
        '    Col E = " (##.##%)"
        ' ----------------------------------------------------------------
        If pendingCount > 0 Then
            Dim dText2 As String: dText2 = vbNullString
            If Not IsEmpty(dV) Then dText2 = Trim$(CStr(dV))

            Dim eText2 As String: eText2 = vbNullString
            If Not IsEmpty(eV) Then eText2 = Trim$(CStr(eV))

            If IsAccountCode(dText2) And InStr(1, eText2, "%") > 0 Then

                Dim acctPct As Double
                acctPct = ParseAccountPct(eText2)

                ' Emit one output row per pending header
                Dim ph As Long
                For ph = 0 To pendingCount - 1
                    Dim h() As Variant
                    h = pendingHeaders(ph)

                    ' Amount = salary (stored on header) * pct / 100
                    Dim amount As Double
                    amount = CDbl(h(13)) * (acctPct / 100#)

                    outRow = outRow + 1
                    If outRow > UBound(outArr, 1) Then
                        ReDim Preserve outArr(1 To UBound(outArr, 1) + 500, 1 To OUT_COLS)
                    End If

                    outArr(outRow, 1)  = curOrg           ' OrgID
                    outArr(outRow, 2)  = curBU            ' BU
                    outArr(outRow, 3)  = h(0)             ' AssignType
                    outArr(outRow, 4)  = h(2)             ' Employee
                    outArr(outRow, 5)  = h(1)             ' EmployeeID (Pos#)
                    outArr(outRow, 6)  = h(5)             ' Location
                    outArr(outRow, 7)  = h(6)             ' JobCategory
                    outArr(outRow, 8)  = h(7)             ' JobClass
                    outArr(outRow, 9)  = h(8)             ' CalendarDays
                    outArr(outRow, 10) = h(9)             ' Placement
                    outArr(outRow, 11) = h(10)            ' Rate
                    outArr(outRow, 12) = h(3)             ' StartDate
                    outArr(outRow, 13) = h(4)             ' EndDate
                    outArr(outRow, 14) = h(11)            ' FTE Authorized
                    outArr(outRow, 15) = h(12)            ' FTE Assigned
                    outArr(outRow, 16) = dText2           ' BudgetCode
                    outArr(outRow, 17) = acctPct          ' AccountPct
                    outArr(outRow, 18) = amount           ' Amount
                    outArr(outRow, 19) = wsSrc.Name       ' SourceSheet
                Next ph

                detailEmittedSinceLastHeader = True
                GoTo NextRow
            End If
        End If

        ' ----------------------------------------------------------------
        ' 7. Anything else with non-empty col A = section boundary; clear state.
        ' ----------------------------------------------------------------
        If Len(aText) > 0 Then
            pendingCount = 0
            detailEmittedSinceLastHeader = False
        End If

NextRow:
    Next r

    ' ---- Write output sheet ----
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Pos04_Normalized")
    wsOut.Cells.Clear

    Dim hdrRow(1 To 1, 1 To OUT_COLS) As Variant
    hdrRow(1, 1)  = "OrgID"
    hdrRow(1, 2)  = "BU"
    hdrRow(1, 3)  = "AssignType"
    hdrRow(1, 4)  = "Employee"
    hdrRow(1, 5)  = "EmployeeID"
    hdrRow(1, 6)  = "Location"
    hdrRow(1, 7)  = "JobCategory"
    hdrRow(1, 8)  = "JobClass"
    hdrRow(1, 9)  = "CalendarDays"
    hdrRow(1, 10) = "Placement"
    hdrRow(1, 11) = "Rate"
    hdrRow(1, 12) = "StartDate"
    hdrRow(1, 13) = "EndDate"
    hdrRow(1, 14) = "FTE_Authorized"
    hdrRow(1, 15) = "FTE_Assigned"
    hdrRow(1, 16) = "BudgetCode"
    hdrRow(1, 17) = "AccountPct"
    hdrRow(1, 18) = "Amount"
    hdrRow(1, 19) = "SourceSheet"

    wsOut.Range("A1").Resize(1, OUT_COLS).Value = hdrRow
    wsOut.Range("A1").Resize(1, OUT_COLS).Font.Bold = True

    ' Text format on columns that need leading-zero preservation
    wsOut.Columns(5).NumberFormat = "@"   ' EmployeeID
    wsOut.Columns(6).NumberFormat = "@"   ' Location
    wsOut.Columns(16).NumberFormat = "@"  ' BudgetCode

    ' Date columns
    wsOut.Columns(12).NumberFormat = "mm/dd/yyyy"
    wsOut.Columns(13).NumberFormat = "mm/dd/yyyy"

    ' Numeric display
    wsOut.Columns(17).NumberFormat = "0.00"       ' AccountPct
    wsOut.Columns(18).NumberFormat = "#,##0.00"   ' Amount

    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, OUT_COLS).Value = outArr
    End If

End Sub

' ============================================================
' PRIVATE HELPERS
' ============================================================

' Returns True if the string looks like a budget/account code.
' Account codes start with two digits followed by a dash, e.g. "01-..." or "13-..."
Private Function IsAccountCode(ByVal s As String) As Boolean
    If Len(s) < 3 Then Exit Function
    If Not (Mid$(s, 1, 1) Like "#") Then Exit Function
    If Not (Mid$(s, 2, 1) Like "#") Then Exit Function
    IsAccountCode = (Mid$(s, 3, 1) = "-")
End Function

' Extract BU code (3-4 uppercase letters) from a line like
' "Bargaining Unit CEMG - Certificated Management"
Private Function ParseBargainingUnit(ByVal s As String) As String
    Dim t As String
    t = Trim$(Mid$(s, Len("Bargaining Unit ") + 1))
    Dim p As Long
    p = InStr(1, t, " ")
    If p > 0 Then
        ParseBargainingUnit = Left$(t, p - 1)
    Else
        ParseBargainingUnit = t
    End If
End Function

' Extract numeric calendar days from text like "CE TEACH (186)" -> "186"
Private Function ParseCalendarDays(ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    Dim s As String: s = CStr(v)
    Dim p1 As Long, p2 As Long
    p1 = InStrRev(s, "(")
    p2 = InStrRev(s, ")")
    If p1 = 0 Or p2 <= p1 Then Exit Function
    Dim inner As String: inner = Trim$(Mid$(s, p1 + 1, p2 - p1 - 1))
    If IsNumeric(inner) Then ParseCalendarDays = inner
End Function

' Parse placement label and rate from col J.
' Input:  "CE Teach- 14/C   (111699.00)"
' Output: placement = "CE Teach- 14/C", rate = "111699.00"
Private Sub ParsePlacementRate(ByVal s As String, ByRef placement As String, ByRef rate As String)
    placement = vbNullString
    rate = vbNullString
    If Len(s) = 0 Then Exit Sub
    Dim p1 As Long, p2 As Long
    p1 = InStrRev(s, "(")
    p2 = InStrRev(s, ")")
    If p1 > 0 And p2 > p1 Then
        rate = Trim$(Mid$(s, p1 + 1, p2 - p1 - 1))
        placement = Trim$(Left$(s, p1 - 1))
    Else
        placement = Trim$(s)
    End If
End Sub

' Parse account percentage from col E text like " (100.00%)" -> 100.0
Private Function ParseAccountPct(ByVal s As String) As Double
    ParseAccountPct = 0#
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, s, "(")
    p2 = InStr(1, s, "%")
    If p1 = 0 Or p2 = 0 Or p2 <= p1 Then Exit Function
    Dim inner As String: inner = Trim$(Mid$(s, p1 + 1, p2 - p1 - 1))
    If IsNumeric(inner) Then ParseAccountPct = CDbl(inner)
End Function

' Parse a date range string like "07/01 - 06/30/26" into Start and End dates.
' The start date is missing its year; infer using fiscal year logic (Jul=start of FY).
' Returns False if parsing fails.
Private Function TryParseDateRange(ByVal s As String, ByRef startDt As Date, ByRef endDt As Date) As Boolean
    TryParseDateRange = False
    If Len(s) = 0 Then Exit Function

    Dim sep As Long: sep = InStr(1, s, " - ")
    If sep = 0 Then Exit Function

    Dim startPart As String: startPart = Trim$(Left$(s, sep - 1))    ' "07/01"
    Dim endPart As String:   endPart = Trim$(Mid$(s, sep + 3))       ' "06/30/26"

    On Error GoTo ParseFail
    endDt = CDate(endPart)

    Dim endYear As Integer: endYear = Year(endDt)
    Dim slashPos As Long: slashPos = InStr(1, startPart, "/")
    If slashPos = 0 Then GoTo ParseFail
    Dim startMonth As Integer: startMonth = CInt(Left$(startPart, slashPos - 1))

    ' July or later -> start year is the year before the end year
    Dim startYear As Integer
    If startMonth >= 7 Then
        startYear = endYear - 1
    Else
        startYear = endYear
    End If

    startDt = CDate(startPart & "/" & Right$(CStr(startYear), 2))
    TryParseDateRange = True
    Exit Function

ParseFail:
    On Error GoTo 0
End Function

' Parse a district/org ID from Pos04 total rows.
' Handles the pattern "Totals for 005 - District Name" where the ID is a
' numeric code immediately following "Totals for ".
' This is distinct from TryParseOrg3 (which requires the word "Org" in the string)
' and may be promoted to modRT_Parse if other reports share this pattern.
' Returns True and sets distID to the leading numeric token if found.
Private Function TryParseDistrictID(ByVal s As String, ByRef distID As String) As Boolean
    TryParseDistrictID = False
    distID = vbNullString

    ' Must start with "Totals for "
    Const PREFIX As String = "Totals for "
    If Left$(s, Len(PREFIX)) <> PREFIX Then Exit Function

    ' Token immediately after prefix
    Dim t As String
    t = Trim$(Mid$(s, Len(PREFIX) + 1))

    ' Extract the leading numeric token (stop at first non-digit)
    Dim i As Long, digits As String
    digits = vbNullString
    For i = 1 To Len(t)
        Dim ch As String: ch = Mid$(t, i, 1)
        If ch Like "#" Then
            digits = digits & ch
        Else
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then Exit Function   ' no leading digits -> BU total, not district total

    distID = digits
    TryParseDistrictID = True
End Function

' Given a row number, return the district ID whose block contains that row.
' distBlockStart(i) and distBlockEnd(i) are inclusive bounds of each district block.
' Returns empty string if no block covers the row (e.g. above first district).
Private Function LookupDistrictID(ByVal rowNum As Long, _
                                   ByRef blockStart() As Long, _
                                   ByRef blockEnd()   As Long, _
                                   ByRef ids()        As String, _
                                   ByVal count        As Long) As String
    LookupDistrictID = vbNullString
    If count = 0 Then Exit Function
    Dim i As Long
    For i = 0 To count - 1
        If rowNum >= blockStart(i) And rowNum <= blockEnd(i) Then
            LookupDistrictID = ids(i)
            Exit Function
        End If
    Next i
End Function