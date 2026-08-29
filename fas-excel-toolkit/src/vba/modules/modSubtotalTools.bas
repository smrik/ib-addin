Attribute VB_Name = "modSubtotalTools"
Option Explicit

Public Sub FAS_InsertSubtotalAboveSelection()
    Dim selectedRange As Range
    Dim destinationRange As Range
    Dim columnIndex As Long
    Dim response As VbMsgBoxResult
    Dim formulaAddress As String

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select rows or columns to subtotal first."
        Exit Sub
    End If

    Set selectedRange = Selection
    If selectedRange.Areas.Count > 1 Then
        FAS_SetStatus "Subtotal needs one contiguous selection."
        Exit Sub
    End If

    If selectedRange.Row = 1 Then
        FAS_SetStatus "Selection starts in row 1; no row above for subtotal formulas."
        Exit Sub
    End If

    Set destinationRange = selectedRange.Offset(-1, 0).Resize(1, selectedRange.Columns.Count)
    If Application.WorksheetFunction.CountA(destinationRange) > 0 Then
        response = MsgBox("Overwrite the row above the selection with subtotal formulas?", vbQuestion + vbOKCancel, "FAS Above Subtotal")
        If response <> vbOK Then
            Exit Sub
        End If
    End If

    For columnIndex = 1 To selectedRange.Columns.Count
        formulaAddress = selectedRange.Columns(columnIndex).Address(False, False)
        destinationRange.Cells(1, columnIndex).Formula = "=SUBTOTAL(9," & formulaAddress & ")"
        destinationRange.Cells(1, columnIndex).NumberFormat = selectedRange.Cells(1, columnIndex).NumberFormat
    Next columnIndex

    FAS_SetStatus "Placed subtotal formulas in the row above " & selectedRange.Address(False, False) & "."
End Sub

Public Sub FAS_InsertSubtotalBelowSelection()
    Dim selectedRange As Range
    Dim destinationRange As Range
    Dim columnIndex As Long
    Dim response As VbMsgBoxResult
    Dim formulaAddress As String

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select rows or columns to subtotal first."
        Exit Sub
    End If

    Set selectedRange = Selection
    If selectedRange.Areas.Count > 1 Then
        FAS_SetStatus "Subtotal needs one contiguous selection."
        Exit Sub
    End If

    If selectedRange.Rows.Count = selectedRange.Worksheet.Rows.Count Or _
        selectedRange.Row + selectedRange.Rows.Count > selectedRange.Worksheet.Rows.Count Then
        FAS_SetStatus "Selection has no row below for subtotal formulas."
        Exit Sub
    End If

    Set destinationRange = selectedRange.Offset(selectedRange.Rows.Count, 0).Resize(1, selectedRange.Columns.Count)
    If Application.WorksheetFunction.CountA(destinationRange) > 0 Then
        response = MsgBox("Overwrite the row below the selection with subtotal formulas?", vbQuestion + vbOKCancel, "FAS Subtotal")
        If response <> vbOK Then
            Exit Sub
        End If
    End If

    For columnIndex = 1 To selectedRange.Columns.Count
        formulaAddress = selectedRange.Columns(columnIndex).Address(False, False)
        destinationRange.Cells(1, columnIndex).Formula = "=SUBTOTAL(109," & formulaAddress & ")"
        destinationRange.Cells(1, columnIndex).NumberFormat = selectedRange.Cells(1, columnIndex).NumberFormat
        destinationRange.Cells(1, columnIndex).Font.Bold = True
        destinationRange.Cells(1, columnIndex).Borders(xlEdgeTop).LineStyle = xlContinuous
        destinationRange.Cells(1, columnIndex).Borders(xlEdgeTop).Weight = xlThin
    Next columnIndex

    FAS_SetStatus "Inserted subtotal row below " & selectedRange.Address(False, False) & "."
End Sub
