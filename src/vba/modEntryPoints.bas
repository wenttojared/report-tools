Attribute VB_Name = "modEntryPoints"
Option Explicit

' HR/PAYROLL - Reports - Payroll

Public Sub Run_Pay03_WithPicker()
    Dim wb As Workbook, ws As Worksheet
    Set wb = PickWorkbook("Pick the workbook that contains the Pay03 export:")
    If wb Is Nothing Then Exit Sub

    Set ws = PickWorksheet(wb, "Pick the sheet that contains the Pay03 export:")
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Pay03 ws
    Exit Sub
    
Fail:
    RT_HandleError "Pay03", wb, ws
End Sub
Public Sub Run_Pay13_WithPicker()
    Dim wb As Workbook, ws As Worksheet
    Set wb = PickWorkbook("Pick the workbook that contains the Pay13 export:")
    If wb Is Nothing Then Exit Sub

    Set ws = PickWorksheet(wb, "Pick the sheet that contains the Pay13 export:")
    If ws Is Nothing Then Exit Sub
    
    On Error GoTo Fail
    Pay13 ws
    Exit Sub
    
Fail:
    RT_HandleError "Pay13", wb, ws
End Sub
Public Sub Run_Pay14_WithPicker()
    Dim wb As Workbook, ws As Worksheet
    Set wb = PickWorkbook("Pick the workbook that contains the Pay14 export:")
    If wb Is Nothing Then Exit Sub

    Set ws = PickWorksheet(wb, "Pick the sheet that contains the Pay14 export:")
    If ws Is Nothing Then Exit Sub
    
    On Error GoTo Fail
    Pay14 ws
    Exit Sub
    
Fail:
    RT_HandleError "Pay14", wb, ws
End Sub

' HR/PAYROLL - Reports - Position Control
Public Sub Run_Pos04_WithPicker()
    Dim wb As Workbook, ws As Worksheet
    Set wb = PickWorkbook("Pick the workbook that contains the Pos04 export:")
    If wb Is Nothing Then Exit Sub

    Set ws = PickWorksheet(wb, "Pick the sheet that contains the Pos04 export:")
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Pos04 ws
    Exit Sub

Fail:
    RT_HandleError "Pos04", wb, ws
End Sub

' HR/PAYROLL - Reports - Benefits
Public Sub Run_Ben02_WithPicker()
    Dim wb As Workbook, ws As Worksheet
    Set wb = PickWorkbook("Pick the workbook that contains the Ben02 export:")
    If wb Is Nothing Then Exit Sub

    Set ws = PickWorksheet(wb, "Pick the sheet that contains the Ben02 export:")
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fail
    Ben02 ws
    Exit Sub
    
Fail:
    RT_HandleError "Ben02", wb, ws
End Sub

' FINANCE - Reports - Fiscal
Public Sub Run_Fiscal05_WithPicker()
    Dim wb As Workbook, ws As Worksheet
    Set wb = PickWorkbook("Pick the workbook that contains the Fiscal05 export:")
    If wb Is Nothing Then Exit Sub

    Set ws = PickWorksheet(wb, "Pick the sheet that contains the Fiscal05 export:")
    If ws Is Nothing Then Exit Sub
    
    On Error GoTo Fail
    Fiscal05 ws
    Exit Sub

Fail:
    RT_HandleError "Fiscal05", wb, ws
End Sub


