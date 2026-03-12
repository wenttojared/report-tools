Attribute VB_Name = "modRT_Errors"
Option Explicit

Public Sub RT_HandleError(ByVal reportName As String, _
                          Optional ByVal wb As Workbook = Nothing, _
                          Optional ByVal ws As Worksheet = Nothing)
    Dim msg As String
    msg = "ReportTools error" & vbCrLf & vbCrLf & _
          "Report: " & reportName & vbCrLf

    If Not wb Is Nothing Then msg = msg & "Workbook: " & wb.Name & vbCrLf
    If Not ws Is Nothing Then msg = msg & "Sheet: " & ws.Name & vbCrLf

    msg = msg & vbCrLf & _
          "Error " & Err.Number & ": " & Err.Description

    Debug.Print Now & " | " & reportName & " | " & Err.Number & " | " & Err.Description

    MsgBox msg, vbExclamation, "ReportTools"
End Sub


