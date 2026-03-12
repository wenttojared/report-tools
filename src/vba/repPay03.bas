Attribute VB_Name = "repPay03"
Option Explicit

Public Sub Pay03(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Pay03..."

    On Error GoTo CleanFail

    Pay03_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub Pay03_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)

    Dim wsOut As Worksheet
    Dim trueLast As Long, lastRow As Long
    Dim r As Long

    Dim src As Variant          ' source data A:M
    Dim outArr() As Variant     ' output A:O (15 cols)

    Dim v1 As String, v2 As String
    Dim emp As String
    Dim id6 As String, ssn4 As String, ssnMasked As String
    Dim curPayDate As Variant

    ' Create or clear output sheet
    Set wsOut = GetOrCreateSheet(wb, "Clean")
    wsOut.Cells.Clear

    ' Headers
    wsOut.Range("A1:O1").Value = Array( _
        "Pay Date", "Employee", "ID", "SSN4", "Net Pay", "Gross", "RetireGross", "Retire", _
        "OASDIGross", "OASDI", "MediGross", "Medi", "Taxes", "MiscDed/Red", "Summer Pay" _
    )

    ' Formats (force ID + SSN to text)
    wsOut.Columns(1).NumberFormat = "mm/dd/yyyy"
    wsOut.Columns(3).NumberFormat = "@"
    wsOut.Columns(4).NumberFormat = "@"
    wsOut.Range("A1:O1").Font.Bold = True

    ' Determine last row (exclude footer section)
    trueLast = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastRow = Application.Max(2, trueLast - 3)

    ' Read source data once: A:M (13 columns)
    src = wsSrc.Range("A1:M" & lastRow).Value2

    ' ---------- Pass 1: count output rows ----------
    Dim outCount As Long
    outCount = 0
    curPayDate = Empty

    For r = 2 To UBound(src, 1)
        v1 = Trim$(CStr(src(r, 1))) ' col A
        v2 = Trim$(CStr(src(r, 2))) ' col B

        If Len(v1) = 0 And Len(v2) = 0 Then GoTo NextRow_Count

        ' Pay Date row
        If LCase$(Left$(v1, 8)) = "pay date" Then
            Dim dtTxt1 As String
            dtTxt1 = Trim$(Mid$(v1, 9))
            If Len(dtTxt1) > 0 Then
                On Error Resume Next
                curPayDate = CDate(dtTxt1)
                On Error GoTo 0
            End If
            GoTo NextRow_Count
        End If

        ' Per-pay-date totals
        If LCase$(Left$(v1, 18)) = "total for pay date" Then GoTo NextRow_Count

        ' Data row test
        If Len(v1) > 0 And Len(v2) > 0 Then
            If TryParseIdSsn4(src(r, 2), id6, ssn4) Then
                outCount = outCount + 1
            End If
        End If

NextRow_Count:
    Next r

    If outCount = 0 Then Exit Sub

    ' ---------- Pass 2: allocate exact array + fill ----------
    ReDim outArr(1 To outCount, 1 To 15)

    Dim outR As Long
    outR = 0
    curPayDate = Empty

    For r = 2 To UBound(src, 1)
        v1 = Trim$(CStr(src(r, 1))) ' col A
        v2 = Trim$(CStr(src(r, 2))) ' col B

        If Len(v1) = 0 And Len(v2) = 0 Then GoTo NextRow_Fill

        ' Pay Date row
        If LCase$(Left$(v1, 8)) = "pay date" Then
            Dim dtTxt2 As String
            dtTxt2 = Trim$(Mid$(v1, 9))
            If Len(dtTxt2) > 0 Then
                On Error Resume Next
                curPayDate = CDate(dtTxt2)
                On Error GoTo 0
            End If
            GoTo NextRow_Fill
        End If

        ' Per-pay-date totals
        If LCase$(Left$(v1, 18)) = "total for pay date" Then GoTo NextRow_Fill

        ' Data row test
        If Len(v1) > 0 And Len(v2) > 0 Then
            If Not TryParseIdSsn4(src(r, 2), id6, ssn4) Then GoTo NextRow_Fill

            emp = v1
            ssnMasked = MaskSsn4(ssn4)

            outR = outR + 1

            outArr(outR, 1) = curPayDate
            outArr(outR, 2) = emp
            outArr(outR, 3) = id6
            outArr(outR, 4) = ssnMasked

            ' Map C..M to output cols 5..15
            outArr(outR, 5) = src(r, 3)    ' Net Pay
            outArr(outR, 6) = src(r, 4)    ' Gross
            outArr(outR, 7) = src(r, 5)    ' RetireGross
            outArr(outR, 8) = src(r, 6)    ' Retire
            outArr(outR, 9) = src(r, 7)    ' OASDIGross
            outArr(outR, 10) = src(r, 8)   ' OASDI
            outArr(outR, 11) = src(r, 9)   ' MediGross
            outArr(outR, 12) = src(r, 10)  ' Medi
            outArr(outR, 13) = src(r, 11)  ' Taxes
            outArr(outR, 14) = src(r, 12)  ' MiscDed/Red
            outArr(outR, 15) = src(r, 13)  ' Summer Pay
        End If

NextRow_Fill:
    Next r

    ' Write output once
    wsOut.Range("A2").Resize(outCount, 15).Value2 = outArr

End Sub

