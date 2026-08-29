Attribute VB_Name = "modFormulaLinks"
Option Explicit

Private mLinkSource As Range

' Store a source range in memory. The next paste links to this range and copies its formats.
Public Sub FAS_CopyLinks()
    Dim sourceRange As Range

    Set sourceRange = FAS_GetSelectedRange()
    If sourceRange Is Nothing Then Exit Sub

    If sourceRange.Areas.Count <> 1 Then
        FAS_SetStatus "Select one contiguous source range for linked paste."
        Exit Sub
    End If

    Set mLinkSource = sourceRange
    FAS_SetStatus "Link source copied. Select a target cell and use Paste Links."
End Sub

' Paste direct links plus source formatting starting at the selected target cell.
Public Sub FAS_PasteLinks()
    Dim targetRange As Range
    Dim targetBlock As Range
    Dim sourceCell As Range
    Dim targetCell As Range
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim sourceWorkbook As Workbook
    Dim response As VbMsgBoxResult

    If mLinkSource Is Nothing Then
        FAS_SetStatus "Copy a source range first."
        Exit Sub
    End If

    Set targetRange = FAS_GetSelectedRange()
    If targetRange Is Nothing Then Exit Sub

    If targetRange.Areas.Count <> 1 Then
        FAS_SetStatus "Select one target cell or range for linked paste."
        Exit Sub
    End If

    On Error GoTo PasteFailed
    Set sourceWorkbook = mLinkSource.Parent.Parent
    Set targetBlock = targetRange.Cells(1, 1).Resize(mLinkSource.Rows.Count, mLinkSource.Columns.Count)

    If Application.Visible Then
        If Application.WorksheetFunction.CountA(targetBlock) > 0 Then
            response = MsgBox("Overwrite the target range with linked formulas?", vbQuestion + vbOKCancel, "FAS Paste Links")
            If response <> vbOK Then Exit Sub
        End If
    End If

    ' Formats are copied once; formulas are assigned cell by cell so they always point back to source.
    mLinkSource.Copy
    targetBlock.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    For rowIndex = 1 To mLinkSource.Rows.Count
        For columnIndex = 1 To mLinkSource.Columns.Count
            Set sourceCell = mLinkSource.Cells(rowIndex, columnIndex)
            Set targetCell = targetBlock.Cells(rowIndex, columnIndex)
            targetCell.Formula = FAS_BuildLinkFormula(sourceWorkbook, sourceCell)
        Next columnIndex
    Next rowIndex

    FAS_SetStatus "Linked formulas and source formatting pasted."
    Exit Sub

PasteFailed:
    Application.CutCopyMode = False
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    FAS_SetStatus "Could not paste linked formulas. Copy a valid source range first."
End Sub

Private Function FAS_BuildLinkFormula(ByVal sourceWorkbook As Workbook, ByVal sourceCell As Range) As String
    Dim workbookPart As String
    Dim sheetPart As String

    workbookPart = vbNullString
    If Not sourceWorkbook Is ActiveWorkbook Then
        workbookPart = "[" & sourceWorkbook.Name & "]"
    End If

    sheetPart = Replace(workbookPart & sourceCell.Worksheet.Name, "'", "''")
    FAS_BuildLinkFormula = "='" & sheetPart & "'!" & sourceCell.Address(True, True, xlA1)
End Function
