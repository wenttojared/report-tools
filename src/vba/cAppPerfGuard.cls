VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAppPerfGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mCalculation As XlCalculation
Private mDisplayStatusBar As Boolean
Private mStatusBar As Variant
Private mActive As Boolean

Public Sub Start(Optional ByVal statusText As String = vbNullString)
    If mActive Then Exit Sub ' prevent double-begin misuse

    With Application
        mScreenUpdating = .ScreenUpdating
        mEnableEvents = .EnableEvents
        mCalculation = .Calculation
        mDisplayStatusBar = .DisplayStatusBar
        mStatusBar = .StatusBar

        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual

        If Len(statusText) > 0 Then
            .DisplayStatusBar = True
            .StatusBar = statusText
        End If
    End With

    mActive = True
End Sub

Public Sub Finish()
    If Not mActive Then Exit Sub
    Restore
End Sub

Private Sub Class_Terminate()
    ' Safety net: restore even if caller forgets Finish or errors out
    If mActive Then Restore
End Sub

Private Sub Restore()
    On Error Resume Next
    With Application
        .ScreenUpdating = mScreenUpdating
        .EnableEvents = mEnableEvents
        .Calculation = mCalculation
        .DisplayStatusBar = mDisplayStatusBar
        .StatusBar = mStatusBar
    End With
    On Error GoTo 0

    mActive = False
End Sub

