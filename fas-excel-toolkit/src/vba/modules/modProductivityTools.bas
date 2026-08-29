Attribute VB_Name = "modProductivityTools"
Option Explicit

Private mHasCurrentFontColor As Boolean
Private mCurrentFontColor As Long

' Capture the active cell font color for the next Apply action.
Public Sub FAS_CaptureCurrentColor()
    Dim selectedRange As Range

    Set selectedRange = FAS_GetSelectedRange()
    If selectedRange Is Nothing Then Exit Sub

    On Error GoTo CaptureFailed
    mCurrentFontColor = selectedRange.Cells(1, 1).Font.Color
    mHasCurrentFontColor = True
    FAS_SetStatus "Current font color captured."
    Exit Sub

CaptureFailed:
    mHasCurrentFontColor = False
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not capture the current font color."
End Sub

' Apply the captured color to the selected cells without taking over Ctrl+Y.
Public Sub FAS_ApplyCurrentColor()
    Dim selectedRange As Range

    If Not mHasCurrentFontColor Then
        FAS_SetStatus "Capture a font color first."
        Exit Sub
    End If

    Set selectedRange = FAS_GetSelectedRange()
    If selectedRange Is Nothing Then Exit Sub

    On Error GoTo ApplyFailed
    selectedRange.Font.Color = mCurrentFontColor
    FAS_SetStatus "Captured font color applied."
    Exit Sub

ApplyFailed:
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not apply the captured font color."
End Sub

' Put a delta sign in the active cell. Existing text is preserved and gets a separator.
Public Sub FAS_InsertDelta()
    Dim activeCell As Range
    Dim currentText As String

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a worksheet cell first."
        Exit Sub
    End If

    Set activeCell = ActiveCell
    On Error GoTo InsertFailed

    currentText = CStr(activeCell.Value2)
    If Len(currentText) = 0 Then
        activeCell.Value2 = ChrW$(&H394)
    ElseIf Application.Visible Then
        If MsgBox("Append a delta sign to the active cell?", vbQuestion + vbOKCancel, "FAS Insert Delta") <> vbOK Then
            Exit Sub
        End If
        activeCell.Value2 = currentText & " " & ChrW$(&H394)
    Else
        activeCell.Value2 = currentText & " " & ChrW$(&H394)
    End If

    FAS_SetStatus "Delta sign inserted."
    Exit Sub

InsertFailed:
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not insert a delta sign in the active cell."
End Sub

' Match Alt+H, O, I and leave a little breathing room after AutoFit.
Public Sub FAS_AutoFitWidthWithMargin()
    Dim selectedRange As Range
    Dim area As Range
    Dim columnItem As Range
    Dim widthValue As Double

    Set selectedRange = FAS_GetSelectedRange()
    If selectedRange Is Nothing Then Exit Sub

    On Error GoTo AutoFitFailed
    For Each area In selectedRange.Areas
        For Each columnItem In area.Columns
            If Not columnItem.EntireColumn.Hidden Then
                columnItem.EntireColumn.AutoFit
                widthValue = CDbl(columnItem.EntireColumn.ColumnWidth)
                columnItem.EntireColumn.ColumnWidth = WorksheetFunction.Min(255, Round(widthValue * 1.2, 2))
            End If
        Next columnItem
    Next area

    FAS_SetStatus "Column width autofit with 20% margin applied."
    Exit Sub

AutoFitFailed:
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not autofit the selected columns."
End Sub
