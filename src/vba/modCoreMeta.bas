Attribute VB_Name = "modCoreMeta"
Option Explicit
Private Const ADDIN_VERSION As String = "0.6.0"
Private Const ADDIN_BUILD_DATE As String = "2026-02-25"
Private Const ADDIN_AUTHOR As String = "Jared Rippey"
Private Const ADDIN_COPYRIGHT As String = "Copyright (c) 2026 Jared Rippey"

Public Function ReportTools_Version() As String
    ReportTools_Version = ADDIN_VERSION
End Function
Public Function ReportTools_Author() As String
    ReportTools_Author = ADDIN_AUTHOR
End Function
Public Function ReportTools_Build() As String
    ReportTools_Build = ADDIN_BUILD_DATE
End Function
Public Function ReportTools_Copyright() As String
    ReportTools_Copyright = ADDIN_COPYRIGHT
End Function

