Option Explicit

Private Sub Workbook_Open()
    FAS_RegisterShortcuts
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    FAS_UnregisterShortcuts
End Sub
