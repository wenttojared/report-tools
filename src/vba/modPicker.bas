Attribute VB_Name = "modPicker"
Option Explicit

Public Function PickWorkbook(Optional ByVal prompt As String = "Pick source workbook:") As Workbook

    Dim wb As Workbook
    Dim validWbs As Collection
    Dim i As Long
    Dim list As String
    Dim choice As Variant

    Set validWbs = New Collection

    ' Build filtered list
    For Each wb In Application.Workbooks

        ' Skip PERSONAL.XLSB
        If LCase$(wb.Name) = "personal.xlsb" Then GoTo NextWb

        ' Skip add-ins (.xlam, .xla)
        If wb.IsAddin Then GoTo NextWb

        ' Skip hidden workbooks
        If wb.Windows(1).Visible = False Then GoTo NextWb

        validWbs.Add wb

NextWb:
    Next wb

    ' No valid workbooks
    If validWbs.Count = 0 Then Exit Function

    ' Only one valid workbook? Auto-select...
    If validWbs.Count = 1 Then
        Set PickWorkbook = validWbs(1)
        Exit Function
    End If

    ' More than one? Prompt user...
    list = prompt & vbCrLf & vbCrLf

    For i = 1 To validWbs.Count
        list = list & i & ") " & validWbs(i).Name & vbCrLf
    Next i

    choice = Application.InputBox(list & vbCrLf & "Enter number:", "Report Tools", 1, Type:=1)

    If choice = False Then Exit Function
    If choice < 1 Or choice > validWbs.Count Then Exit Function

    Set PickWorkbook = validWbs(choice)

End Function


Public Function PickWorksheet(ByVal wb As Workbook, Optional ByVal prompt As String = "Pick source sheet:") As Worksheet
    Dim ws As Worksheet
    Dim validWs As Collection
    Dim i As Long, list As String
    Dim choice As Variant

    If wb Is Nothing Then Exit Function

    ' If the active sheet is a visible worksheet in this workbook, just use it
    On Error Resume Next
    If Not ActiveSheet Is Nothing Then
        If ActiveSheet.Parent Is wb Then
            If TypeName(ActiveSheet) = "Worksheet" Then
                If ActiveSheet.Visible = xlSheetVisible Then
                    Set PickWorksheet = ActiveSheet
                    Exit Function
                End If
            End If
        End If
    End If
    On Error GoTo 0

    Set validWs = New Collection

    ' Build filtered list of visible worksheets
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            validWs.Add ws
        End If
    Next ws

    If validWs.Count = 0 Then Exit Function

    ' Only one visible sheet? Auto-select
    If validWs.Count = 1 Then
        Set PickWorksheet = validWs(1)
        Exit Function
    End If

    ' Prompt
    list = prompt & " (" & wb.Name & ")" & vbCrLf & vbCrLf
    For i = 1 To validWs.Count
        list = list & i & ") " & validWs(i).Name & vbCrLf
    Next i

    choice = Application.InputBox(list & vbCrLf & "Enter number:", "Report Tools", 1, Type:=1)
    If choice = False Then Exit Function
    If choice < 1 Or choice > validWs.Count Then Exit Function

    Set PickWorksheet = validWs(choice)
End Function


