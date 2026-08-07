Attribute VB_Name = "repPay07"
Option Explicit

Public Sub Pay07(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Pay07..."

    On Error GoTo CleanFail

    Pay07_Worker wsSrc, wb

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
' Source report has NO header row -- data begins on row 1.
'
' Row types (detected via VarType, since Value2 returns date-serial
' rows as Double, not vbDate):
'   - Data row:        col A is numeric (a date serial)
'   - Pay-date subtotal:  col A = "Total for MM/DD/YY", col C = "Count"        -> skip
'   - Org total (boundary): col A = "Total for [District]", col C = "Total Count" -> triggers OrgID backfill
'   - Footer "Fund NN" rows: skip (unrecognized, fall through)
'   - Footer "Org" rows:  col A = "Org", col B = 3-digit Org ID -- pre-scanned
'                          separately and matched to org totals ORDINALLY
'                          (Nth "Total for [District]" <-> Nth footer "Org" row),
'                          NOT by row-range containment like the Pos04/Pay13/
'                          Budget04 district lookups. The footer appears in the
'                          same top-to-bottom order as the org blocks above it,
'                          so a simple ordinal pointer is sufficient and much
'                          cheaper than a block-range table.
'
' Output routing (col B = warrant/ACH identifier, col D = payee format):
'   - col B starts "ACH-"            -> ACH sheet. Number = text after "ACH-".
'                                        Col D expected in "(NNNNNN) NNNN" format.
'   - col B all-digit, starts with "1":
'       - col D matches "(NNNNNN) NNNN"  -> standard Physical Warrant
'       - col D matches "NNNNNN/N"       -> Trailing Warrant (vendor)
'   - anything else                  -> logged and skipped
'
' All three output types are accumulated into one combined array (with a
' discriminator column) so the OrgID backfill logic only has to be written
' once; the combined array is filtered into three sheets at write time.
Private Sub Pay07_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)

    Dim ur As Range
    Set ur = wsSrc.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim data As Variant
    data = ur.Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)
    If nRows < 1 Then Exit Sub

    ' ---- Sheet-type discriminator values ----
    Const TYPE_PHYSICAL As Long = 1
    Const TYPE_TRAILING As Long = 2
    Const TYPE_ACH As Long = 3

    ' ---- Combined output columns ----
    Const COL_TYPE As Long = 1
    Const COL_ORG As Long = 2
    Const COL_DATE As Long = 3
    Const COL_ID As Long = 4
    Const COL_SUFFIX As Long = 5
    Const COL_NAME As Long = 6
    Const COL_AMT As Long = 7
    Const COL_NUM As Long = 8
    Const COL_SRC As Long = 9
    Const OUT_COLS As Long = 9

    ' ---- Pre-scan: footer "Org" rows, in order ----
    Dim orgIDList() As String
    Dim orgIDCount As Long: orgIDCount = 0
    ReDim orgIDList(0 To 9)

    Dim psR As Long
    For psR = 1 To nRows
        Dim psA As Variant: psA = data(psR, 1)
        If Not IsEmpty(psA) Then
            If VarType(psA) = vbString Then
                If Trim$(CStr(psA)) = "Org" Then
                    If orgIDCount > UBound(orgIDList) Then
                        ReDim Preserve orgIDList(0 To orgIDCount + 9)
                    End If
                    orgIDList(orgIDCount) = Trim$(CStr(NzV(GetA(data, psR, 2))))
                    orgIDCount = orgIDCount + 1
                End If
            End If
        End If
    Next psR

    ' ---- Main pass ----
    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To OUT_COLS)   ' upper bound: one output row per source row, never exceeded
    Dim outRow As Long: outRow = 0

    Dim blockStart As Long: blockStart = 1    ' first un-backfilled row in outArr
    Dim orgBlockIdx As Long: orgBlockIdx = 0  ' pointer into orgIDList

    Dim r As Long
    For r = 1 To nRows

        Dim aV As Variant: aV = data(r, 1)

        ' ------------------------------------------------------------
        ' Data row: col A is a date serial (numeric)
        ' ------------------------------------------------------------
        If VarType(aV) = vbDouble Or VarType(aV) = vbInteger Or VarType(aV) = vbLong Then

            Dim bText As String: bText = Trim$(CStr(NzV(GetA(data, r, 2))))
            Dim dText As String: dText = Trim$(CStr(NzV(GetA(data, r, 4))))
            Dim nameVal As String: nameVal = Trim$(CStr(NzV(GetA(data, r, 5))))

            Dim amt As Double: amt = 0#
            Call TryParseNumber(NzV(GetA(data, r, 3)), amt)

            Dim sheetType As Long: sheetType = 0
            Dim idOut As String, suffixOut As String, numOut As String
            idOut = vbNullString: suffixOut = vbNullString: numOut = vbNullString

            If Left$(bText, 4) = "ACH-" Then
                numOut = Mid$(bText, 5)

                Dim achID As String, achSSN As String
                If TryParseIdSsn4(dText, achID, achSSN) Then
                    sheetType = TYPE_ACH
                    idOut = achID
                    suffixOut = MaskSsn4(achSSN)
                Else
                    Debug.Print "Pay07_Worker: ACH row col D did not match employee ID format at row " & r & _
                                " on [" & wsSrc.Name & "] — [" & dText & "]"
                End If

            ElseIf IsWarrantNumber(bText) Then
                numOut = bText

                Dim wID As String, wSSN As String
                Dim vID As String, vAddr As String
                If TryParseIdSsn4(dText, wID, wSSN) Then
                    sheetType = TYPE_PHYSICAL
                    idOut = wID
                    suffixOut = MaskSsn4(wSSN)
                ElseIf TryParseVendorSlash(dText, vID, vAddr) Then
                    sheetType = TYPE_TRAILING
                    idOut = vID
                    suffixOut = vAddr
                Else
                    Debug.Print "Pay07_Worker: warrant row col D matched neither employee nor vendor format at row " & r & _
                                " on [" & wsSrc.Name & "] — [" & dText & "]"
                End If

            Else
                Debug.Print "Pay07_Worker: unrecognized column B format at row " & r & _
                            " on [" & wsSrc.Name & "] — [" & bText & "]"
            End If

            If sheetType > 0 Then
                outRow = outRow + 1

                outArr(outRow, COL_TYPE) = sheetType
                outArr(outRow, COL_ORG) = vbNullString          ' backfilled at the next org total row
                outArr(outRow, COL_DATE) = CDate(aV)
                outArr(outRow, COL_ID) = idOut
                outArr(outRow, COL_SUFFIX) = suffixOut
                outArr(outRow, COL_NAME) = nameVal
                outArr(outRow, COL_AMT) = amt
                outArr(outRow, COL_NUM) = numOut
                outArr(outRow, COL_SRC) = wsSrc.Name
            End If

            GoTo NextRow
        End If

        ' ------------------------------------------------------------
        ' Not a data row -- check for the org-total boundary via col C
        ' ------------------------------------------------------------
        Dim cText As String: cText = vbNullString
        Dim cV As Variant: cV = GetA(data, r, 3)
        If Not IsEmpty(cV) Then
            If VarType(cV) = vbString Then cText = Trim$(CStr(cV))
        End If

        If cText = "Total Count" Then
            Dim thisOrgID As String
            If orgBlockIdx < orgIDCount Then
                thisOrgID = orgIDList(orgBlockIdx)
            Else
                thisOrgID = vbNullString
                Debug.Print "Pay07_Worker: no footer Org ID available for block ending at row " & r & _
                            " on [" & wsSrc.Name & "] — OrgID left blank"
            End If

            Dim k As Long
            For k = blockStart To outRow
                outArr(k, COL_ORG) = thisOrgID
            Next k

            orgBlockIdx = orgBlockIdx + 1
            blockStart = outRow + 1
            GoTo NextRow
        End If

        ' Pay-date subtotal rows (col C = "Count") and footer "Fund NN" / "Org"
        ' rows fall through here and are silently skipped -- expected, not an error.

NextRow:
    Next r

    ' If the source ends without a final org-total row, warn (mirrors repBen02)
    If blockStart <= outRow Then
        Debug.Print "Pay07_Worker: rows " & blockStart & " to " & outRow & _
                    " have no org total trailer on sheet [" & wsSrc.Name & "] — OrgID left blank"
    End If

    ' ---- Split combined array into three output sheets ----
    WritePay07Sheet wb, "Pay07_Physical", outArr, outRow, TYPE_PHYSICAL, _
        Array("OrgID", "CheckDate", "EmployeeID", "SSN_Last4", "EmployeeName", "Amount", "WarrantNumber", "SourceSheet")

    WritePay07Sheet wb, "Pay07_Trailing", outArr, outRow, TYPE_TRAILING, _
        Array("OrgID", "CheckDate", "VendorID", "VendorAddrNum", "VendorName", "Amount", "WarrantNumber", "SourceSheet")

    WritePay07Sheet wb, "Pay07_ACH", outArr, outRow, TYPE_ACH, _
        Array("OrgID", "CheckDate", "EmployeeID", "SSN_Last4", "EmployeeName", "Amount", "ACHNumber", "SourceSheet")

End Sub

' ============================================================
' PRIVATE HELPERS
' ============================================================

' Filters the combined array down to rows matching targetType and writes
' them to the named output sheet with the supplied headers.
Private Sub WritePay07Sheet(ByVal wb As Workbook, ByVal sheetName As String, _
                             ByRef combinedArr As Variant, ByVal combinedRowCount As Long, _
                             ByVal targetType As Long, ByVal headers As Variant)

    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, sheetName)
    wsOut.Cells.Clear

    Dim hCount As Long: hCount = UBound(headers) - LBound(headers) + 1

    Dim c As Long
    For c = 0 To hCount - 1
        wsOut.Cells(1, c + 1).Value = headers(c)
    Next c
    wsOut.Range("A1").Resize(1, hCount).Font.Bold = True

    Dim filtered() As Variant
    If combinedRowCount > 0 Then
        ReDim filtered(1 To combinedRowCount, 1 To hCount)
    End If

    Dim fRow As Long: fRow = 0
    Dim i As Long
    For i = 1 To combinedRowCount
        If combinedArr(i, 1) = targetType Then
            fRow = fRow + 1
            filtered(fRow, 1) = combinedArr(i, 2)   ' OrgID
            filtered(fRow, 2) = combinedArr(i, 3)   ' CheckDate
            filtered(fRow, 3) = combinedArr(i, 4)   ' ID
            filtered(fRow, 4) = combinedArr(i, 5)   ' Suffix
            filtered(fRow, 5) = combinedArr(i, 6)   ' Name
            filtered(fRow, 6) = combinedArr(i, 7)   ' Amount
            filtered(fRow, 7) = combinedArr(i, 8)   ' WarrantNumber / ACHNumber
            filtered(fRow, 8) = combinedArr(i, 9)   ' SourceSheet
        End If
    Next i

    ' Text-format ID/leading-zero columns before the bulk write
    wsOut.Columns(1).NumberFormat = "@"             ' OrgID
    wsOut.Columns(2).NumberFormat = "mm/dd/yyyy"    ' CheckDate
    wsOut.Columns(3).NumberFormat = "@"             ' ID
    wsOut.Columns(4).NumberFormat = "@"             ' Suffix
    wsOut.Columns(7).NumberFormat = "@"             ' WarrantNumber / ACHNumber

    If fRow > 0 Then
        wsOut.Range("A2").Resize(fRow, hCount).Value = filtered
    End If

End Sub

' Returns True if s is a pure-digit warrant number starting with "1".
' Per Frontline export convention, physical warrant numbers always begin with 1;
' ACH rows are excluded earlier via the "ACH-" prefix check, so no overlap.
Private Function IsWarrantNumber(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) <> "1" Then Exit Function

    Dim i As Long
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Function
    Next i

    IsWarrantNumber = True
End Function
