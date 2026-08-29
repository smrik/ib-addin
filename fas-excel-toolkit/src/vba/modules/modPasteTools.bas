Attribute VB_Name = "modPasteTools"
Option Explicit

Public Sub FAS_PasteValuesOnly()
    Dim targetRange As Range

    Set targetRange = FAS_GetSelectedRange()
    If targetRange Is Nothing Then
        Exit Sub
    End If

    On Error GoTo PasteFailed
    targetRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    FAS_SetStatus "Pasted values."
    Exit Sub

PasteFailed:
    Application.CutCopyMode = False
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

    FAS_SetStatus "Could not paste values. Copy a valid range first."
End Sub

Public Sub FAS_PasteAsTextOnly()
    Dim targetRange As Range

    Set targetRange = FAS_GetSelectedRange()
    If targetRange Is Nothing Then
        Exit Sub
    End If

    On Error GoTo PasteFailed
    targetRange.NumberFormat = "@"
    targetRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    FAS_SetStatus "Pasted as text."
    Exit Sub

PasteFailed:
    Application.CutCopyMode = False
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

    FAS_SetStatus "Could not paste text. Copy a valid range first."
End Sub

' Copy only cells that remain visible after filters or manual hiding.
Public Sub FAS_CopyVisibleOnly()
    Dim selectedRange As Range
    Dim visibleRange As Range

    Set selectedRange = FAS_GetSelectedRange()
    If selectedRange Is Nothing Then Exit Sub

    On Error GoTo CopyFailed
    Set visibleRange = selectedRange.SpecialCells(xlCellTypeVisible)
    visibleRange.Copy
    FAS_SetStatus "Visible cells copied."
    Exit Sub

CopyFailed:
    Application.CutCopyMode = False
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "No visible cells could be copied from the selection."
End Sub

Public Function FAS_GetSelectedRange() As Range
    If TypeName(Selection) <> "Range" Then
        If Not Application.Visible Then
            Err.Raise vbObjectError + 5100, "FAS_GetSelectedRange", "Select a worksheet range before using this tool."
        End If

        FAS_SetStatus "Select a worksheet range first."
        Exit Function
    End If

    On Error Resume Next
    Set FAS_GetSelectedRange = Selection
    On Error GoTo 0

    If FAS_GetSelectedRange Is Nothing Then
        If Not Application.Visible Then
            Err.Raise vbObjectError + 5101, "FAS_GetSelectedRange", "The current selection is not a valid worksheet range."
        End If

        FAS_SetStatus "Current selection is not a valid range."
    End If
End Function
