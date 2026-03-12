Attribute VB_Name = "modRT_Array"
Option Explicit

Public Function NzV(ByVal v As Variant) As Variant
    If IsError(v) Then
        NzV = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzV = ""
    Else
        NzV = v
    End If
End Function

Public Function GetA(ByRef data As Variant, ByVal r As Long, ByVal c As Long) As Variant
    ' Safe-ish access: if c beyond bounds, return Empty
    If c < 1 Then
        GetA = Empty
    ElseIf c > UBound(data, 2) Then
        GetA = Empty
    Else
        GetA = data(r, c)
    End If
End Function


Public Function Slice2D(ByRef arr As Variant, ByVal startRow As Long, ByVal endRow As Long, ByVal cols As Long) As Variant
    Dim r As Long, c As Long, out As Variant
    ReDim out(1 To (endRow - startRow + 1), 1 To cols)
    For r = startRow To endRow
        For c = 1 To cols
            out(r - startRow + 1, c) = arr(r, c)
        Next c
    Next r
    Slice2D = out
End Function


