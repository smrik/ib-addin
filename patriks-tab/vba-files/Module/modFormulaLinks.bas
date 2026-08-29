Attribute VB_Name = "modFormulaLinks"
Option Explicit

Private mLinkSource As Range

' Store a source range in memory. The next paste links to this range.
Public Sub FAS_CopyLinks()
    Dim sourceRange As Range

    Set sourceRange = FAS_GetSelectedRange()
    If sourceRange Is Nothing Then
        Exit Sub
    End If

    If sourceRange.Areas.Count <> 1 Then
        FAS_SetStatus "Select one contiguous source range for linked paste."
        Exit Sub
    End If

    Set mLinkSource = sourceRange
    FAS_SetStatus "Link source copied: " & sourceRange.Worksheet.Name & "!" & sourceRange.Address(False, False) & ". Select a target cell and use Paste Links."
End Sub

' Paste source formatting and then rebuild the contents as same-workbook links/values.
Public Sub FAS_PasteLinks()
    Dim targetRange As Range
    Dim targetBlock As Range
    Dim sourceCell As Range
    Dim targetCell As Range
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim response As VbMsgBoxResult
    Dim undoStarted As Boolean

    If mLinkSource Is Nothing Then
        FAS_SetStatus "No link source. Select the source range and run Copy Links first; normal Copy is not enough."
        Exit Sub
    End If

    Set targetRange = FAS_GetSelectedRange()
    If targetRange Is Nothing Then
        Exit Sub
    End If

    If targetRange.Areas.Count <> 1 Then
        FAS_SetStatus "Select one target cell or range for linked paste."
        Exit Sub
    End If

    On Error GoTo PasteFailed
    Set sourceSheet = mLinkSource.Worksheet
    Set targetSheet = targetRange.Worksheet
    Set sourceWorkbook = sourceSheet.Parent
    Set targetWorkbook = targetSheet.Parent
    If Not (sourceWorkbook Is targetWorkbook) Then
        FAS_SetStatus "Linked paste only supports source and target sheets in the same workbook."
        Exit Sub
    End If

    Set targetBlock = targetRange.Cells(1, 1).Resize(mLinkSource.Rows.Count, mLinkSource.Columns.Count)
    ' Intersect only works for ranges on the same worksheet. Cross-sheet linked paste is valid.
    If sourceSheet Is targetSheet Then
        If Not (Intersect(mLinkSource, targetBlock) Is Nothing) Then
            FAS_SetStatus "Choose a target range outside the copied source range."
            Exit Sub
        End If
    End If

    If Application.Visible Then
        If Application.WorksheetFunction.CountA(targetBlock) > 0 Then
            response = MsgBox("Overwrite the target range with linked values and formulas?", vbQuestion + vbOKCancel, "FAS Paste Links")
            If response <> vbOK Then
                Exit Sub
            End If
        End If
    End If

    ' Copy formats first. Contents are assigned below so constants and blanks keep their meaning.
    If Not FAS_UndoBeginRange(targetBlock, "Paste links") Then Exit Sub
    undoStarted = True
    mLinkSource.Copy
    targetBlock.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    For columnIndex = 1 To mLinkSource.Columns.Count
        On Error Resume Next
        targetBlock.Columns(columnIndex).EntireColumn.ColumnWidth = mLinkSource.Columns(columnIndex).EntireColumn.ColumnWidth
        On Error GoTo PasteFailed
    Next columnIndex

    For rowIndex = 1 To mLinkSource.Rows.Count
        On Error Resume Next
        targetBlock.Rows(rowIndex).RowHeight = mLinkSource.Rows(rowIndex).RowHeight
        On Error GoTo PasteFailed
    Next rowIndex

    For rowIndex = 1 To mLinkSource.Rows.Count
        For columnIndex = 1 To mLinkSource.Columns.Count
            Set sourceCell = mLinkSource.Cells(rowIndex, columnIndex)
            Set targetCell = targetBlock.Cells(rowIndex, columnIndex)
            FAS_CopyLinkedCell sourceCell, targetCell
        Next columnIndex
    Next rowIndex

    FAS_SetStatus "Linked formulas, values, blanks, and source formatting pasted."
    FAS_UndoCommit
    Exit Sub

PasteFailed:
    If undoStarted Then FAS_UndoCancel
    Application.CutCopyMode = False
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not paste linked formulas. Copy a valid same-workbook source range first."
End Sub

Private Sub FAS_CopyLinkedCell(ByVal sourceCell As Range, ByVal targetCell As Range)
    Dim sheetPart As String

    If IsEmpty(sourceCell.Value2) Then
        targetCell.ClearContents
    ElseIf sourceCell.HasFormula Then
        sheetPart = Replace(sourceCell.Worksheet.Name, "'", "''")
        targetCell.Formula = "='" & sheetPart & "'!" & sourceCell.Address(True, True, xlA1)
    Else
        targetCell.Value2 = sourceCell.Value2
    End If
End Sub
