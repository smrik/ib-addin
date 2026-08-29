Attribute VB_Name = "modAddinLifecycle"
Option Explicit

Public Sub FAS_RegisterShortcuts()
    FAS_RegisterConfiguredShortcuts
End Sub

Public Sub FAS_UnregisterShortcuts()
    FAS_UnregisterConfiguredShortcuts
End Sub

Public Function FAS_GetQualifiedMacroName(ByVal macroName As String) As String
    FAS_GetQualifiedMacroName = "'" & ThisWorkbook.Name & "'!" & macroName
End Function
