Attribute VB_Name = "modUndoManager"
Option Explicit

Private Const FAS_UNDO_MAX_CELLS As Double = 250000

Private mFASUndoSourceWorkbook As Object
Private mFASUndoSourceSheet As Object
Private mFASUndoBackupWorkbook As Object
Private mFASUndoBackupSheet As Object
Private mFASUndoSourceAddress As String
Private mFASUndoDescription As String
Private mFASUndoKind As String
Private mFASUndoCreatedSheet As Object
Private mFASUndoRows As Long
Private mFASUndoColumns As Long
Private mFASUndoColumnWidths() As Variant
Private mFASUndoRowHeights() As Variant

Public Function FAS_UndoMaxCellCountForTests() As Double
    FAS_UndoMaxCellCountForTests = FAS_UNDO_MAX_CELLS
End Function

Public Function FAS_UndoBeginRange(ByVal targetRange As Object, ByVal actionDescription As String) As Boolean
    Dim previousWorkbook As Object
    Dim previousSheet As Object
    Dim previousSelection As Object
    Dim backupRange As Object
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean
    Dim columnIndex As Long
    Dim rowIndex As Long

    If targetRange Is Nothing Then Exit Function
    If targetRange.Areas.Count <> 1 Then
        FAS_SetStatus "FAS undo needs one contiguous range; action cancelled."
        Exit Function
    End If

    If CDbl(targetRange.CountLarge) > FAS_UNDO_MAX_CELLS Then
        FAS_SetStatus "FAS undo snapshot is limited to 250,000 cells; action cancelled."
        Exit Function
    End If

    On Error GoTo CaptureFailed
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    Set previousWorkbook = Application.ActiveWorkbook
    Set previousSheet = Application.ActiveSheet
    If TypeName(Selection) = "Range" Then Set previousSelection = Selection

    FAS_UndoClearState

    Set mFASUndoSourceWorkbook = targetRange.Worksheet.Parent
    Set mFASUndoSourceSheet = targetRange.Worksheet
    mFASUndoSourceAddress = targetRange.Address
    mFASUndoDescription = actionDescription
    mFASUndoKind = "range"
    mFASUndoRows = targetRange.Rows.Count
    mFASUndoColumns = targetRange.Columns.Count

    ReDim mFASUndoColumnWidths(1 To mFASUndoColumns)
    For columnIndex = 1 To mFASUndoColumns
        mFASUndoColumnWidths(columnIndex) = targetRange.Columns(columnIndex).EntireColumn.ColumnWidth
    Next columnIndex

    ReDim mFASUndoRowHeights(1 To mFASUndoRows)
    For rowIndex = 1 To mFASUndoRows
        mFASUndoRowHeights(rowIndex) = targetRange.Rows(rowIndex).EntireRow.RowHeight
    Next rowIndex

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set mFASUndoBackupWorkbook = Application.Workbooks.Add
    Set mFASUndoBackupSheet = mFASUndoBackupWorkbook.Worksheets(1)
    mFASUndoBackupSheet.Name = "FAS Undo"
    Set backupRange = mFASUndoBackupSheet.Range("A1").Resize(mFASUndoRows, mFASUndoColumns)
    ' Destination copy keeps the user's pending clipboard source available for paste commands.
    targetRange.Copy Destination:=backupRange
    mFASUndoBackupWorkbook.Windows(1).Visible = False
    mFASUndoBackupWorkbook.Saved = True

    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    FAS_UndoRestoreFocus previousWorkbook, previousSheet, previousSelection

    FAS_UndoBeginRange = True
    Exit Function

CaptureFailed:
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    FAS_UndoClearState
    FAS_UndoRestoreFocus previousWorkbook, previousSheet, previousSelection
    FAS_SetStatus "FAS undo snapshot could not be created; action cancelled."
End Function

Public Sub FAS_UndoCommit()
    If mFASUndoBackupWorkbook Is Nothing Or mFASUndoKind <> "range" Then Exit Sub

    On Error GoTo RegisterFailed
    Application.OnUndo "Undo FAS: " & mFASUndoDescription, FAS_GetQualifiedMacroName("FAS_UndoLastAction")
    Exit Sub

RegisterFailed:
    FAS_UndoClearState
    FAS_SetStatus "FAS edit completed, but its custom undo could not be registered."
End Sub

Public Sub FAS_UndoRegisterCreatedSheet(ByVal createdSheet As Object, ByVal actionDescription As String)
    If createdSheet Is Nothing Then Exit Sub

    FAS_UndoClearState
    Set mFASUndoSourceWorkbook = createdSheet.Parent
    Set mFASUndoCreatedSheet = createdSheet
    mFASUndoDescription = actionDescription
    mFASUndoKind = "sheet"

    On Error GoTo RegisterFailed
    Application.OnUndo "Undo FAS: " & mFASUndoDescription, FAS_GetQualifiedMacroName("FAS_UndoLastAction")
    Exit Sub

RegisterFailed:
    FAS_UndoClearState
    FAS_SetStatus "Divider sheet created, but its custom undo could not be registered."
End Sub

Public Sub FAS_UndoCancel()
    FAS_UndoClearState
End Sub

Public Sub FAS_UndoInvalidate()
    FAS_UndoClearState
End Sub

Public Sub FAS_UndoLastAction()
    Dim targetRange As Object
    Dim backupRange As Object
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean

    If mFASUndoKind = "sheet" Then
        On Error GoTo SheetRestoreFailed
        oldDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        mFASUndoCreatedSheet.Delete
        Application.DisplayAlerts = oldDisplayAlerts
        FAS_SetStatus "Undid FAS: " & mFASUndoDescription & "."
        FAS_UndoClearState
        Exit Sub
    End If

    If mFASUndoBackupWorkbook Is Nothing Or mFASUndoSourceSheet Is Nothing Then
        FAS_SetStatus "There is no FAS edit available to undo."
        Exit Sub
    End If

    On Error GoTo RestoreFailed
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    Set targetRange = mFASUndoSourceSheet.Range(mFASUndoSourceAddress)
    Set backupRange = mFASUndoBackupSheet.Range("A1").Resize(mFASUndoRows, mFASUndoColumns)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    backupRange.Copy Destination:=targetRange

    For columnIndex = 1 To mFASUndoColumns
        targetRange.Columns(columnIndex).EntireColumn.ColumnWidth = mFASUndoColumnWidths(columnIndex)
    Next columnIndex

    For rowIndex = 1 To mFASUndoRows
        targetRange.Rows(rowIndex).EntireRow.RowHeight = mFASUndoRowHeights(rowIndex)
    Next rowIndex

    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error Resume Next
    mFASUndoSourceWorkbook.Activate
    mFASUndoSourceSheet.Activate
    targetRange.Select
    On Error GoTo 0
    FAS_SetStatus "Undid FAS: " & mFASUndoDescription & "."
    FAS_UndoClearState
    Exit Sub

RestoreFailed:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    FAS_SetStatus "Could not restore the last FAS edit. The source workbook or range may have changed."
    Exit Sub

SheetRestoreFailed:
    Application.DisplayAlerts = oldDisplayAlerts
    FAS_SetStatus "Could not remove the FAS-created sheet. It may have been renamed or deleted."
End Sub

Public Sub FAS_UndoShutdown()
    FAS_UndoClearState
End Sub

Private Sub FAS_UndoClearState()
    Dim oldDisplayAlerts As Boolean

    On Error Resume Next
    ' Clear the callback with the snapshot so discarded state cannot remain callable.
    Application.OnUndo vbNullString, vbNullString
    If Not mFASUndoBackupWorkbook Is Nothing Then
        oldDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        mFASUndoBackupWorkbook.Close SaveChanges:=False
        Application.DisplayAlerts = oldDisplayAlerts
    End If
    On Error GoTo 0

    Set mFASUndoBackupSheet = Nothing
    Set mFASUndoBackupWorkbook = Nothing
    Set mFASUndoSourceSheet = Nothing
    Set mFASUndoSourceWorkbook = Nothing
    Set mFASUndoCreatedSheet = Nothing
    mFASUndoSourceAddress = vbNullString
    mFASUndoDescription = vbNullString
    mFASUndoKind = vbNullString
    mFASUndoRows = 0
    mFASUndoColumns = 0
    Erase mFASUndoColumnWidths
    Erase mFASUndoRowHeights
End Sub

Private Sub FAS_UndoRestoreFocus(ByVal previousWorkbook As Object, ByVal previousSheet As Object, ByVal previousSelection As Object)
    On Error Resume Next
    If Not previousWorkbook Is Nothing Then previousWorkbook.Activate
    If Not previousSheet Is Nothing Then previousSheet.Activate
    If Not previousSelection Is Nothing Then previousSelection.Select
    On Error GoTo 0
End Sub
