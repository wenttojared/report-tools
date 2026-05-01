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
          "Copyright: " & ReportTools_Copyright() & vbCrLf & vbCrLf & _
          "Licensed under MIT License" & vbCrLf & _
          "Source: https://github.com/wenttojared/report-tools"

    MsgBox msg, vbInformation, "About Report Tools"

End Sub

Public Sub Ribbon_ExportSettings(ByVal control As Object)

    Dim msg As String

    msg = "Export Settings" & vbCrLf & vbCrLf & _
          "These macros are designed to work with reports exported using the " & _
          """Excel Data"" option in Frontline CA ERP. Other export formats " & _
          "(PDF, CSV, etc.) are not supported and will produce unexpected results." & vbCrLf & vbCrLf & _
          "Some reports require specific settings to be configured before running. " & _
          "Where this applies, an Export Guide will appear in the report's submenu." & vbCrLf & vbCrLf & _
          "If no Export Guide is listed for a report, it is designed to work with " & _
          "default report settings and no additional configuration is needed."

    MsgBox msg, vbInformation, "Export Settings"

End Sub

Public Sub Ribbon_Guide(ByVal control As Object)
    On Error GoTo Fail

    ' Convention: control ID is "btnGuide_ReportCode"
    ' e.g. "btnGuide_Budget04" -> "Budget04"
    Dim id As String
    id = CStr(control.id)

    Const PREFIX As String = "btnGuide_"
    If Left$(id, Len(PREFIX)) <> PREFIX Then
        Err.Raise 5, "Ribbon_Guide", "Unexpected guide control ID: " & id
    End If

    Dim reportCode As String
    reportCode = Mid$(id, Len(PREFIX) + 1)

    Dim msg As String
    msg = GetGuideText(reportCode)

    If LenB(msg) = 0 Then
        msg = "No export guide is available for " & reportCode & "."
    End If

    MsgBox msg, vbInformation, reportCode & " Report Settings"
    Exit Sub

Fail:
    MsgBox "ReportTools: Unable to display export guide." & vbCrLf & _
           "Control ID: " & CStr(control.id) & vbCrLf & _
           "Error: " & Err.Description, vbExclamation, "ReportTools"
End Sub

' Returns the export guide text for the given report code.
' Add a Case entry here when a new report requires specific export settings.
Private Function GetGuideText(ByVal reportCode As String) As String
    Select Case reportCode

        Case "Budget04Import"
            GetGuideText = "Before exporting, open the report settings and configure " & _
                        "the following option:" & vbCrLf & vbCrLf & _
                        "  4 - Account Sort/Group Options" & vbCrLf & _
                        "       Sort/Group 1: Resc" & vbCrLf & vbCrLf & _
                        "The macro will not produce correct results if this setting " & _
                        "is not configured before exporting."
        Case Else
            GetGuideText = ""

    End Select
End Function
