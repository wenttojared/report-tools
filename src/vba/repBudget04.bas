Attribute VB_Name = "repBudget04"
Option Explicit

Public Sub Budget04_Convert(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Budget04 Convert..."

    On Error GoTo CleanFail

    Budget04_Convert_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub Budget04_Import(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Budget04 Import..."

    On Error GoTo CleanFail

    Budget04_Import_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

' ============================================================
' CONVERT WORKER
' ============================================================
' Produces a normalized line-item table from a Budget04 export.
' One output row per item. No subtotals. No object code filtering.
' Output columns:
'   1  District
'   2  AccountCode
'   3  AccountDescription
'   4  ItemNumber
'   5  ItemType
'   6  Comment
'   7  ItemDescription
'   8  Amount
'   9  SourceSheet
Private Sub Budget04_Convert_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)

    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim data As Variant
    data = ur.Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)
    If nRows < 2 Then Exit Sub

    ' ---- Pre-scan: build district block lookup ----
    ' Same pattern as Budget04_Import_Worker and repPos04/repPay13.
    Dim distBlockStart() As Long
    Dim distBlockEnd()   As Long
    Dim distID()         As String
    Dim distCount        As Long: distCount = 0
    ReDim distBlockStart(0 To 9)
    ReDim distBlockEnd(0 To 9)
    ReDim distID(0 To 9)

    Dim psBlockStart As Long: psBlockStart = 1
    Dim psR As Long
    For psR = 1 To nRows
        Dim psAV As Variant: psAV = data(psR, 4)
        If Not IsEmpty(psAV) Then
            If VarType(psAV) = vbString Then
                Dim psText As String: psText = Trim$(CStr(psAV))
                Dim org3 As String
                If TryParseOrg3(psText, org3) Then
                    If distCount > UBound(distBlockStart) Then
                        ReDim Preserve distBlockStart(0 To distCount + 9)
                        ReDim Preserve distBlockEnd(0 To distCount + 9)
                        ReDim Preserve distID(0 To distCount + 9)
                    End If
                    distBlockStart(distCount) = psBlockStart
                    distBlockEnd(distCount)   = psR
                    distID(distCount)         = org3
                    distCount = distCount + 1
                    psBlockStart = psR + 1
                End If
            End If
        End If
    Next psR

    Const OUT_COLS As Long = 9

    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To OUT_COLS)
    Dim outRow As Long: outRow = 0

    ' Carry-forward fields reset each time a new account code row is encountered
    Dim curAccountCode As String
    Dim curAccountDesc As String
    curAccountCode = vbNullString
    curAccountDesc = vbNullString

    Dim r As Long
    For r = 1 To nRows

        Dim aV As Variant: aV = data(r, 1)
        If IsEmpty(aV) Then GoTo NextRow

        ' ----------------------------------------------------------------
        ' Account code row: col A is a string matching the ##- budget code
        ' pattern. Resets carry-forward state for the new chunk.
        ' ----------------------------------------------------------------
        If VarType(aV) = vbString Then
            Dim aText As String: aText = Trim$(CStr(aV))
            If Len(aText) = 0 Then GoTo NextRow

            If IsBudgetCodeRow(aText) Then
                curAccountCode = aText
                curAccountDesc = Trim$(CStr(NzV(data(r, 2))))
            End If

            ' All other string rows (fund headers, expenditure label,
            ' subtotals, totals, footer) are structural noise — skip.
            GoTo NextRow
        End If

        ' ----------------------------------------------------------------
        ' Item row: col A is numeric (the item number).
        ' Only emit if we are inside a valid account code chunk.
        ' ----------------------------------------------------------------
        If VarType(aV) = vbDouble Or VarType(aV) = vbInteger Or VarType(aV) = vbLong Then
            If LenB(curAccountCode) = 0 Then GoTo NextRow

            Dim amt As Double
            If Not TryParseNumber(NzV(data(r, 5)), amt) Then GoTo NextRow

            outRow = outRow + 1
            If outRow > UBound(outArr, 1) Then
                ReDim Preserve outArr(1 To UBound(outArr, 1) + 500, 1 To OUT_COLS)
            End If

            outArr(outRow, 1) = LookupDistrict(r, distBlockStart, distBlockEnd, distID, distCount)
            outArr(outRow, 2) = curAccountCode
            outArr(outRow, 3) = curAccountDesc
            outArr(outRow, 4) = CLng(aV)                          ' ItemNumber
            outArr(outRow, 5) = Trim$(CStr(NzV(data(r, 2))))      ' ItemType
            outArr(outRow, 6) = Trim$(CStr(NzV(data(r, 3))))      ' Comment
            outArr(outRow, 7) = Trim$(CStr(NzV(data(r, 4))))      ' ItemDescription
            outArr(outRow, 8) = amt                                ' Amount
            outArr(outRow, 9) = wsSrc.Name                        ' SourceSheet
        End If

NextRow:
    Next r

    ' ---- Write output sheet ----
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Budget04_Normalized")
    wsOut.Cells.Clear

    Dim hdr(1 To 1, 1 To OUT_COLS) As Variant
    hdr(1, 1) = "District"
    hdr(1, 2) = "AccountCode"
    hdr(1, 3) = "AccountDescription"
    hdr(1, 4) = "ItemNumber"
    hdr(1, 5) = "ItemType"
    hdr(1, 6) = "Comment"
    hdr(1, 7) = "ItemDescription"
    hdr(1, 8) = "Amount"
    hdr(1, 9) = "SourceSheet"

    wsOut.Range("A1").Resize(1, OUT_COLS).Value = hdr
    wsOut.Range("A1").Resize(1, OUT_COLS).Font.Bold = True

    wsOut.Columns(1).NumberFormat = "@"        ' District
    wsOut.Columns(2).NumberFormat = "@"        ' AccountCode
    wsOut.Columns(8).NumberFormat = "#,##0.00" ' Amount

    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, OUT_COLS).Value = outArr
    End If

End Sub

' ============================================================
' IMPORT WORKER
' ============================================================
' Produces a headerless three-column budget import table:
'   District, BudgetCode, Amount
' Excludes salary and benefits lines (Object 1000-3999) where SO=00.
' Amount is taken from col C of the account code row (rounded integer).
Private Sub Budget04_Import_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)

    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim data As Variant
    data = ur.Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)
    If nRows < 2 Then Exit Sub

    ' ---- Pre-scan: build district block lookup ----
    Dim distBlockStart() As Long
    Dim distBlockEnd()   As Long
    Dim distID()         As String
    Dim distCount        As Long: distCount = 0
    ReDim distBlockStart(0 To 9)
    ReDim distBlockEnd(0 To 9)
    ReDim distID(0 To 9)

    Dim psBlockStart As Long: psBlockStart = 1
    Dim psR As Long
    For psR = 1 To nRows
        Dim psAV As Variant: psAV = data(psR, 4)
        If Not IsEmpty(psAV) Then
            If VarType(psAV) = vbString Then
                Dim psText As String: psText = Trim$(CStr(psAV))
                Dim org3 As String
                If TryParseOrg3(psText, org3) Then
                    If distCount > UBound(distBlockStart) Then
                        ReDim Preserve distBlockStart(0 To distCount + 9)
                        ReDim Preserve distBlockEnd(0 To distCount + 9)
                        ReDim Preserve distID(0 To distCount + 9)
                    End If
                    distBlockStart(distCount) = psBlockStart
                    distBlockEnd(distCount)   = psR
                    distID(distCount)         = org3
                    distCount = distCount + 1
                    psBlockStart = psR + 1
                End If
            End If
        End If
    Next psR

    Const OUT_COLS As Long = 3

    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To OUT_COLS)
    Dim outRow As Long: outRow = 0

    Dim r As Long
    For r = 1 To nRows

        Dim aV As Variant: aV = data(r, 1)
        If IsEmpty(aV) Then GoTo NextRow
        If VarType(aV) <> vbString Then GoTo NextRow

        Dim aText As String: aText = Trim$(CStr(aV))
        If Len(aText) = 0 Then GoTo NextRow
        If Not IsBudgetCodeRow(aText) Then GoTo NextRow

        If nCols < 3 Then GoTo NextRow
        Dim amtV As Variant: amtV = data(r, 3)
        If IsEmpty(amtV) Then GoTo NextRow

        Dim amt As Double
        If Not TryParseNumber(amtV, amt) Then GoTo NextRow
        If IsSalaryBenefitsExcluded(aText) Then GoTo NextRow

        outRow = outRow + 1
        outArr(outRow, 1) = LookupDistrict(r, distBlockStart, distBlockEnd, distID, distCount)
        outArr(outRow, 2) = aText
        outArr(outRow, 3) = amt

NextRow:
    Next r

    ' ---- Write output sheet ----
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Budget04_Import")
    wsOut.Cells.Clear

    ' No headers — import format requires raw data only
    wsOut.Columns(1).NumberFormat = "@"   ' District
    wsOut.Columns(2).NumberFormat = "@"   ' BudgetCode

    If outRow > 0 Then
        wsOut.Range("A1").Resize(outRow, OUT_COLS).Value = outArr
    Else
        Debug.Print "Budget04_Import_Worker: no rows passed the filter on sheet [" & wsSrc.Name & "]." & vbCrLf & _
                    "If all codes have Object 1000-3999 and SO=00 this is expected — " & _
                    "check the source report for non-salary/benefits lines."
    End If

End Sub

' ============================================================
' SHARED PRIVATE HELPERS
' ============================================================

' Returns True if the string is a Frontline budget code row.
' Budget codes start with two digits followed by a dash: "01-...", "12-..."
Private Function IsBudgetCodeRow(ByVal s As String) As Boolean
    If Len(s) < 3 Then Exit Function
    If Not (Mid$(s, 1, 1) Like "#") Then Exit Function
    If Not (Mid$(s, 2, 1) Like "#") Then Exit Function
    IsBudgetCodeRow = (Mid$(s, 3, 1) = "-")
End Function

' Given a row number, returns the district ID whose pre-scanned block
' contains that row. Returns empty string if no block covers the row.
Private Function LookupDistrict(ByVal rowNum As Long, _
                                ByRef blockStart() As Long, _
                                ByRef blockEnd()   As Long, _
                                ByRef ids()        As String, _
                                ByVal count        As Long) As String
    LookupDistrict = vbNullString
    If count = 0 Then Exit Function
    Dim i As Long
    For i = 0 To count - 1
        If rowNum >= blockStart(i) And rowNum <= blockEnd(i) Then
            LookupDistrict = ids(i)
            Exit Function
        End If
    Next i
End Function

' Returns True if the row should be EXCLUDED from the import output.
' Exclusion rule: Object code is in range 1000-3999 (salary & benefits)
' AND Sub-Object is "00".
'
' Account code format: Fd-Resc-Y-Goal-Func-Objt-SO-Sch-DD1-DD2
'   0-based segment indices after Split by "-":
'     5 = Objt
'     6 = SO
Private Function IsSalaryBenefitsExcluded(ByVal s As String) As Boolean
    IsSalaryBenefitsExcluded = False

    Dim parts() As String
    parts = Split(s, "-")

    If UBound(parts) < 6 Then Exit Function

    Dim objStr As String: objStr = Trim$(parts(5))
    Dim soStr  As String: soStr  = Trim$(parts(6))

    If Not IsNumeric(objStr) Then Exit Function
    Dim obj As Long: obj = CLng(objStr)
    If obj < 1000 Or obj > 3999 Then Exit Function

    IsSalaryBenefitsExcluded = (soStr = "00")
End Function
