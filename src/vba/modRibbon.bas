Attribute VB_Name = "modRibbon"
Option Explicit
Private gRibbon As Object

Public Sub Ribbon_OnLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub Ribbon_About(control As Object)

    Dim msg As String

    msg = "Report Tools" & vbCrLf & vbCrLf & _
          "Version: " & ReportTools_Version() & vbCrLf & _
          "Build Date: " & ReportTools_Build() & vbCrLf & _
          "Author: " & ReportTools_Author() & vbCrLf & vbCrLf & _
          "Licensed under MIT License" & vbCrLf & _
          "Source: https://github.com/wenttojared/report-tools"

    MsgBox msg, vbInformation, "About Report Tools"

End Sub

