Attribute VB_Name = "modAddinLifecycle"
Option Explicit

Public Sub FAS_RegisterShortcuts()
    Application.OnKey "^+N", FAS_GetQualifiedMacroName("FAS_ToggleNavigator")
    Application.OnKey "^+J", FAS_GetQualifiedMacroName("FAS_ToggleCommandPalette")
    Application.OnKey "^+P", FAS_GetQualifiedMacroName("FAS_InsertSubtotalAboveSelection")
    Application.OnKey "^+T", FAS_GetQualifiedMacroName("FAS_GoToTB")
    Application.OnKey "^+G", FAS_GetQualifiedMacroName("FAS_GoToGL")
    Application.OnKey "^%{UP}", FAS_GetQualifiedMacroName("FAS_CycleDirectPrecedent")
    Application.OnKey "^%{DOWN}", FAS_GetQualifiedMacroName("FAS_BackToTraceOrigin")
End Sub

Public Sub FAS_UnregisterShortcuts()
    Application.OnKey "^+N"
    Application.OnKey "^+J"
    Application.OnKey "^+P"
    Application.OnKey "^+T"
    Application.OnKey "^+G"
    Application.OnKey "^%{UP}"
    Application.OnKey "^%{DOWN}"
End Sub

Private Function FAS_GetQualifiedMacroName(ByVal macroName As String) As String
    FAS_GetQualifiedMacroName = "'" & ThisWorkbook.Name & "'!" & macroName
End Function
