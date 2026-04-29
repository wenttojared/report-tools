Attribute VB_Name = "repBudget04"
Option Explicit

Public Sub Budget04(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Budget04..."

    On Error GoTo CleanFail

    Budget04_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


' ----------- WORKER -----------
Private Sub Budget04_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)

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
    ' "Total for Org 000 - District Name" trailer rows appear AFTER the account
    ' codes they belong to, so we pre-scan to record block boundaries and look
    ' up by row in the main pass. Same pattern as repPos04/repPay13.
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
    ' TROUBLESHOOTING NOTE:
    ' The original report I built this for had "Total for Org 000 - District Name"
    ' in col A but the one I received from Business that they're running this macro
    ' for has it in col C. If OrgId suddently stops working check this part and adjust 
    ' the column index as needed. 
    ' e.g. if it's in col A then change data(psR, 3) to data(psR, 1) and test.
        Dim psAV As Variant: psAV = data(psR, 3) 
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
                    distBlockEnd(distCount) = psR
                    distID(distCount) = org3
                    distCount = distCount + 1
                    psBlockStart = psR + 1
                End If
            End If
        End If
    Next psR

    ' Outputs three columns: OrgID, BudgetCode, Amount. 
    ' No headers (import format)
    Const OUT_COLS As Long = 3

    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To OUT_COLS)
    Dim outRow As Long: outRow = 0

    Dim r As Long
    For r = 1 To nRows

        Dim aV As Variant: aV = data(r, 1)

        ' Only string values in col A can be account code rows
        If IsEmpty(aV) Then GoTo NextRow
        If VarType(aV) <> vbString Then GoTo NextRow

        Dim aText As String: aText = Trim$(CStr(aV))
        If Len(aText) = 0 Then GoTo NextRow

        ' Account code rows start with two digits followed by a dash
        ' e.g. "01-6054-0-0000-2140-1317-00-800-405-000"
        ' All other structural rows (fund headers, totals, footer) do not match this.
        If Not IsBudgetCodeRow(aText) Then GoTo NextRow

        ' Amount is in col C (index 3)
        Dim amtV As Variant
        If nCols >= 3 Then
            amtV = data(r, 3)
        Else
            GoTo NextRow
        End If

        ' Skip rows where col C is empty or non-numeric
        If IsEmpty(amtV) Then GoTo NextRow
        Dim amt As Double
        If Not TryParseNumber(amtV, amt) Then GoTo NextRow

        ' Apply salary/benefits + SO=00 exclusion filter
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

    ' No headers... import format requires raw data only
    ' Force text on columns that must preserve leading zeroes
    wsOut.Columns(1).NumberFormat = "@"   ' District
    wsOut.Columns(2).NumberFormat = "@"   ' BudgetCode

    If outRow > 0 Then
        wsOut.Range("A1").Resize(outRow, OUT_COLS).Value = outArr
    Else
        Debug.Print "Budget04_Worker: no rows passed the filter on sheet [" & wsSrc.Name & "]." & vbCrLf & _
                    "If all codes have Object 1000-3999 and SO=00 this is expected � " & _
                    "check the source report for non-salary/benefits lines."
    End If

End Sub

' ----------- PRIVATE HELPERS -----------

' Returns True if the string is a Frontline budget code row.
' Budget codes start with two digits followed by a dash: "01-...", "12-..."
' All structural rows (fund headers, totals, selection footer, item rows)
' either start with non-digit characters or are numeric types (not strings).
Private Function IsBudgetCodeRow(ByVal s As String) As Boolean
    If Len(s) < 3 Then Exit Function
    If Not (Mid$(s, 1, 1) Like "#") Then Exit Function
    If Not (Mid$(s, 2, 1) Like "#") Then Exit Function
    IsBudgetCodeRow = (Mid$(s, 3, 1) = "-")
End Function

' Given a row number, returns the district ID (org3) whose pre-scanned block
' contains that row. Returns empty string if no block covers the row.
Private Function LookupDistrict(ByVal rowNum As Long, _
                                ByRef blockStart() As Long, _
                                ByRef blockEnd() As Long, _
                                ByRef ids() As String, _
                                ByVal count As Long) As String
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
' Exclusion rule: Object code is in range 1000-3999 (salary & benefits) AND
' Sub-Object is "00".
'
' Account code format: Fd-Resc-Y-Goal-Func-Objt-SO-Sch-DD1-DD2
'   Segment index (1-based after Split by "-"):
'     1 = Fd
'     2 = Resc
'     3 = Y
'     4 = Goal
'     5 = Func
'     6 = Objt   <-- Object code
'     7 = SO     <-- Sub-Object
'     8 = Sch
'     9 = DD1
'    10 = DD2
Private Function IsSalaryBenefitsExcluded(ByVal s As String) As Boolean
    IsSalaryBenefitsExcluded = False

    Dim parts() As String
    parts = Split(s, "-")

    ' Need at least 7 segments to read Object and SO
    If UBound(parts) < 6 Then Exit Function

    Dim objStr As String: objStr = Trim$(parts(5))   ' 0-based index 5 = 6th segment
    Dim soStr As String:  soStr = Trim$(parts(6))    ' 0-based index 6 = 7th segment

    ' Object must be numeric and in range 1000-3999
    If Not IsNumeric(objStr) Then Exit Function
    Dim obj As Long: obj = CLng(objStr)
    If obj < 1000 Or obj > 3999 Then Exit Function

    ' Sub-Object must be "00"
    IsSalaryBenefitsExcluded = (soStr = "00")
End Function


