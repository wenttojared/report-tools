Attribute VB_Name = "repFiscal05"
Option Explicit

Public Sub Fiscal05(ByVal wsSrc As Worksheet)
    Dim wb As Workbook: Set wb = wsSrc.Parent

    Dim g As New cAppPerfGuard
    g.Start "ReportTools: Fiscal05..."

    On Error GoTo CleanFail

    Fiscal05_Worker wsSrc, wb

CleanExit:
    g.Finish
    Exit Sub

CleanFail:
    g.Finish
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub Fiscal05_Worker(ByVal wsSrc As Worksheet, ByVal wb As Workbook)
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wb, "Fiscal05_Normalized")
    wsOut.Cells.Clear

    ' Output headers
    Dim headers As Variant
    headers = Array( _
        "District", _
        "AccountType", _
        "Fd-Resc-Y-Goal-Func-Objt-SO-Sch-DD1-DD2", _
        "Description", _
        "AdoptedBudget", _
        "Revised", _
        "Encumbered_or_Debit", _
        "Expenditure_or_Credit", _
        "AccountBalance", _
        "SourceSheet" _
    )

    Dim c As Long
    For c = LBound(headers) To UBound(headers)
        wsOut.Cells(1, c + 1).Value = headers(c)
    Next c

    Dim rng As Range
    Set rng = wsSrc.UsedRange
    If rng Is Nothing Then Exit Sub

    Dim arr As Variant
    arr = rng.Value2

    Dim rMax As Long, cMax As Long
    rMax = UBound(arr, 1)
    cMax = UBound(arr, 2)

    If cMax < 7 Then Exit Sub

    Dim outArr() As Variant
    ReDim outArr(1 To rMax, 1 To 10)

    Dim outRow As Long: outRow = 0

    ' Track where each org/account block starts so we can backfill district and account type
    ' when the corresponding trailer row is encountered
    Dim districtBlockStart As Long: districtBlockStart = 1
    Dim acctBlockStart As Long: acctBlockStart = 1

    Dim v1 As String
    Dim acctType As String
    Dim i As Long

    For i = 1 To rMax
        v1 = Trim$(CStr(arr(i, 1)))
        If Len(v1) = 0 Then GoTo NextI

        ' Check account total BEFORE org total: an "Ending Balance Accounts" row can
        ' appear immediately before a "Total for Org" row and must not be misclassified
        If IsAccountTotalRow(v1) Then
            acctType = ParseAccountTypeFromTotal(v1)

            If acctBlockStart <= outRow Then
                Dim k As Long
                For k = acctBlockStart To outRow
                    outArr(k, 2) = acctType
                Next k
            End If

            acctBlockStart = outRow + 1
            GoTo NextI
        End If

        If IsOrgTotalRow(v1) Then
            Dim dist As String
            dist = ParseDistrictFromOrgTotal(v1)

            If districtBlockStart <= outRow Then
                Dim j As Long
                For j = districtBlockStart To outRow
                    outArr(j, 1) = dist
                Next j
            End If

            districtBlockStart = outRow + 1
            acctBlockStart = outRow + 1   ' reset account block at org boundary

            GoTo NextI
        End If

        If IsDetailDataRow(v1) Then
            outRow = outRow + 1
            outArr(outRow, 3) = arr(i, 1)
            outArr(outRow, 4) = arr(i, 2)
            outArr(outRow, 5) = arr(i, 3)
            outArr(outRow, 6) = arr(i, 4)
            outArr(outRow, 7) = arr(i, 5)
            outArr(outRow, 8) = arr(i, 6)
            outArr(outRow, 9) = arr(i, 7)
            outArr(outRow, 10) = wsSrc.Name
        End If

NextI:
    Next i

    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, 10).Value = outArr
    End If

End Sub

' ----- helpers -----

Private Function IsOrgTotalRow(ByVal s As String) As Boolean
    IsOrgTotalRow = (Left$(s, 13) = "Total for Org")
End Function

Private Function ParseDistrictFromOrgTotal(ByVal s As String) As String
    ' e.g. "Total for Org 123 - District Name" -> "123"
    Dim parts() As String
    parts = Split(s, " ")
    If UBound(parts) >= 3 Then
        ParseDistrictFromOrgTotal = parts(3)
    Else
        ParseDistrictFromOrgTotal = vbNullString
    End If
End Function

Private Function IsAccountTotalRow(ByVal s As String) As Boolean
    ' Must not be an org total even though it also starts with "Total for"
    If IsOrgTotalRow(s) Then
        IsAccountTotalRow = False
    Else
        IsAccountTotalRow = (Left$(s, 9) = "Total for" And InStr(1, s, " Accounts", vbTextCompare) > 0)
    End If
End Function

Private Function ParseAccountTypeFromTotal(ByVal s As String) As String
    Dim x As String
    x = s

    x = Replace$(x, "Total for ", vbNullString, 1, 1, vbTextCompare)
    x = Replace$(x, " Accounts", vbNullString, 1, 1, vbTextCompare)

    ParseAccountTypeFromTotal = Trim$(x)
End Function

' Detects account code rows by format: starts with a digit and contains at least one hyphen.
' Frontline account codes follow the pattern: Fd-Resc-Y-Goal-Func-Objt-SO-Sch-DD1-DD2
' e.g. "0-0000-0-0000-0000-0000-00-0000-00-00"
Private Function IsDetailDataRow(ByVal s As String) As Boolean
    If Len(s) < 6 Then
        IsDetailDataRow = False
        Exit Function
    End If

    If Left$(s, 1) < "0" Or Left$(s, 1) > "9" Then
        IsDetailDataRow = False
        Exit Function
    End If

    IsDetailDataRow = (InStr(1, s, "-", vbBinaryCompare) > 0)
End Function


