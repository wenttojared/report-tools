Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

Public Sub Ribbon_Run(ByVal control As Object)
    On Error GoTo Fail

    Dim id As String
    id = CStr(control.id)

    ' Convention: take text after last underscore
    ' Examples:
    '   btnHRPayroll_Pay03    -> Pay03
    '   btnFinance_Fiscal05   -> Fiscal05
    '   btnHRPayroll_Ben02    -> Ben02
    Dim p As Long
    p = InStrRev(id, "_")
    If p = 0 Or p = Len(id) Then Err.Raise 5, "Ribbon_Run", "Unexpected control ID: " & id

    Dim reportCode As String
    reportCode = Mid$(id, p + 1)

    Dim procName As String
    procName = "Run_" & reportCode & "_WithPicker"

    Application.Run procName
    Exit Sub

Fail:
    MsgBox "ReportTools: Unable to run this command." & vbCrLf & _
           "Control ID: " & CStr(control.id) & vbCrLf & _
           "Error: " & Err.Description, vbExclamation, "ReportTools"
End Sub


