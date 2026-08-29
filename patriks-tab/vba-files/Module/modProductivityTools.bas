Attribute VB_Name = "modProductivityTools"
Option Explicit

Private mHasCurrentFontColor As Boolean
Private mCurrentFontColor As Long

' Capture the active cell font color for the next Apply action.
Public Sub FAS_CaptureCurrentColor()
    Dim selectedRange As Range

    Set selectedRange = FAS_GetSelectedRange()
    If selectedRange Is Nothing Then
        Exit Sub
    End If

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

' Apply the captured color without taking over Excel's native Ctrl+Y repeat command.
Public Sub FAS_ApplyCurrentColor()
    Dim selectedRange As Range
    Dim undoStarted As Boolean

    If Not mHasCurrentFontColor Then
        FAS_SetStatus "Capture a font color first."
        Exit Sub
    End If

    Set selectedRange = FAS_GetSelectedRange()
    If selectedRange Is Nothing Then
        Exit Sub
    End If

    On Error GoTo ApplyFailed
    If Not FAS_UndoBeginRange(selectedRange, "Apply font color") Then Exit Sub
    undoStarted = True
    selectedRange.Font.Color = mCurrentFontColor
    FAS_SetStatus "Captured font color applied."
    FAS_UndoCommit
    Exit Sub

ApplyFailed:
    If undoStarted Then FAS_UndoCancel
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not apply the captured font color."
End Sub

' Insert a delta sign in the active cell. Existing text is preserved.
Public Sub FAS_InsertDelta()
    ' Do not name this local activeCell. VBA is case-insensitive, so such a local
    ' shadows the global ActiveCell property and "Set activeCell = ActiveCell"
    ' assigns the uninitialised local to itself; every later use raises error 91.
    Dim targetCell As Range
    Dim currentText As String
    Dim undoStarted As Boolean

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a worksheet cell first."
        Exit Sub
    End If

    Set targetCell = Application.ActiveCell
    On Error GoTo InsertFailed

    If targetCell.HasFormula Then
        FAS_SetStatus "Delta was not inserted into a formula cell. Select a blank or text cell."
        Exit Sub
    End If

    If IsError(targetCell.Value2) Then
        FAS_SetStatus "Delta was not inserted into an error cell. Select a blank or text cell."
        Exit Sub
    End If

    If Not FAS_UndoBeginRange(targetCell, "Insert delta") Then Exit Sub
    undoStarted = True
    currentText = CStr(targetCell.Value2)
    If Len(currentText) = 0 Then
        targetCell.Value2 = ChrW$(916)
    Else
        targetCell.Value2 = currentText & " " & ChrW$(916)
    End If

    FAS_SetStatus "Delta sign inserted."
    FAS_UndoCommit
    Exit Sub

InsertFailed:
    If undoStarted Then FAS_UndoCancel
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
    Dim undoStarted As Boolean

    Set selectedRange = FAS_GetSelectedRange()
    If selectedRange Is Nothing Then
        Exit Sub
    End If

    On Error GoTo AutoFitFailed
    If Not FAS_UndoBeginRange(selectedRange, "AutoFit columns") Then Exit Sub
    undoStarted = True
    For Each area In selectedRange.Areas
        For Each columnItem In area.Columns
            If Not columnItem.EntireColumn.Hidden Then
                columnItem.EntireColumn.AutoFit
                widthValue = CDbl(columnItem.EntireColumn.ColumnWidth)
                columnItem.EntireColumn.ColumnWidth = Application.Min(255, Round(widthValue * 1.2, 2))
            End If
        Next columnItem
    Next area

    FAS_SetStatus "Column width autofit with 20% margin applied."
    FAS_UndoCommit
    Exit Sub

AutoFitFailed:
    If undoStarted Then FAS_UndoCancel
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not autofit the selected columns."
End Sub
