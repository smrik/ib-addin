Attribute VB_Name = "modShapeTools"
Option Explicit

Public Function FAS_SetSelectedObjectsDoNotMoveOrSize() As Boolean
    Dim selectedShapeRange As Object
    Dim targetShape As Shape
    Dim changedCount As Long

    On Error Resume Next
    Set selectedShapeRange = Selection.ShapeRange
    On Error GoTo 0

    If selectedShapeRange Is Nothing Then
        FAS_SetStatus "Select one or more shapes/images/charts first."
        Exit Function
    End If

    For Each targetShape In selectedShapeRange
        If targetShape.Placement <> xlFreeFloating Then
            targetShape.Placement = xlFreeFloating
            changedCount = changedCount + 1
        End If
    Next targetShape

    FAS_SetStatus changedCount & " selected object(s) set to don't move or size."
    FAS_SetSelectedObjectsDoNotMoveOrSize = True
End Function

Public Sub FAS_SetShapesDoNotMoveOrSize()
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim targetShape As Shape
    Dim changedCount As Long
    Dim response As VbMsgBoxResult

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        FAS_SetStatus "No active user workbook available."
        Exit Sub
    End If

    response = MsgBox("Set all shapes in the active workbook to 'Don't move or size with cells'?", vbQuestion + vbOKCancel, "FAS Shapes")
    If response <> vbOK Then
        Exit Sub
    End If

    For Each targetSheet In targetWorkbook.Worksheets
        For Each targetShape In targetSheet.Shapes
            If targetShape.Placement <> xlFreeFloating Then
                targetShape.Placement = xlFreeFloating
                changedCount = changedCount + 1
            End If
        Next targetShape
    Next targetSheet

    FAS_SetStatus changedCount & " shape(s) updated."
End Sub
