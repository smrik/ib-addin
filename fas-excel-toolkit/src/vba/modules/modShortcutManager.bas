Attribute VB_Name = "modShortcutManager"
Option Explicit

Private Const FAS_SHORTCUTS_ENABLED_KEY As String = "Shortcuts.Enabled"
Private Const FAS_SHORTCUT_PREFIX As String = "Shortcut."

Public Sub FAS_RegisterConfiguredShortcuts()
    Dim a As Variant
    Dim i As Long
    Dim keyText As String

    FAS_UnregisterConfiguredShortcuts

    If Not FAS_GetBooleanSetting(FAS_SHORTCUTS_ENABLED_KEY, True) Then
        FAS_SetStatus "Shortcuts are off."
        Exit Sub
    End If

    a = FAS_GetShortcutActions()
    For i = LBound(a) To UBound(a)
        keyText = FAS_GetShortcutKey(CStr(a(i)(0)), CStr(a(i)(2)))
        If Len(keyText) > 0 Then
            Application.OnKey keyText, FAS_QualifiedMacro(CStr(a(i)(3)))
        End If
    Next i
End Sub

Public Sub FAS_UnregisterConfiguredShortcuts()
    Dim a As Variant
    Dim i As Long
    Dim keyText As String
    Dim defaultKey As String

    On Error Resume Next
    a = FAS_GetShortcutActions()
    For i = LBound(a) To UBound(a)
        defaultKey = CStr(a(i)(2))
        keyText = FAS_GetShortcutKey(CStr(a(i)(0)), defaultKey)
        If Len(defaultKey) > 0 Then Application.OnKey defaultKey
        If Len(keyText) > 0 And StrComp(keyText, defaultKey, vbTextCompare) <> 0 Then Application.OnKey keyText
    Next i
    On Error GoTo 0
End Sub

Public Function FAS_TryRunShortcutCommand(ByVal commandText As String) As Boolean
    Dim rawText As String
    Dim c As String
    Dim rest As String
    Dim actionId As String
    Dim keyText As String
    Dim p As Long

    rawText = Trim$(commandText)
    c = LCase$(Trim$(rawText))

    Select Case c
        Case "shortcut", "shortcuts", "keys"
            FAS_ShowShortcutStatus
            FAS_TryRunShortcutCommand = True
            Exit Function

        Case "shortcut on", "shortcuts on", "keys on"
            FAS_SetBooleanSetting FAS_SHORTCUTS_ENABLED_KEY, True
            FAS_RegisterConfiguredShortcuts
            FAS_SetStatus "Shortcuts on."
            FAS_TryRunShortcutCommand = True
            Exit Function

        Case "shortcut off", "shortcuts off", "keys off"
            FAS_SetBooleanSetting FAS_SHORTCUTS_ENABLED_KEY, False
            FAS_UnregisterConfiguredShortcuts
            FAS_SetStatus "Shortcuts off."
            FAS_TryRunShortcutCommand = True
            Exit Function

        Case "shortcut reset", "shortcuts reset", "keys reset", "shortcut default", "shortcuts default"
            FAS_ResetShortcuts
            FAS_RegisterConfiguredShortcuts
            FAS_SetStatus "Shortcuts reset."
            FAS_TryRunShortcutCommand = True
            Exit Function
    End Select

    If Left$(LCase$(rawText), 13) = "shortcut set " Then
        rest = Trim$(Mid$(rawText, 14))
    ElseIf Left$(LCase$(rawText), 9) = "shortcut " Then
        rest = Trim$(Mid$(rawText, 10))
    Else
        Exit Function
    End If

    p = InStr(1, rest, " ", vbBinaryCompare)
    If p = 0 Then
        If FAS_IsKnownShortcutAction(rest) Then
            FAS_ShowShortcutStatus rest
            FAS_TryRunShortcutCommand = True
        End If
        Exit Function
    End If

    actionId = LCase$(Trim$(Left$(rest, p - 1)))
    keyText = Trim$(Mid$(rest, p + 1))
    If Not FAS_IsKnownShortcutAction(actionId) Then
        FAS_SetStatus "Unknown shortcut action: " & actionId
        FAS_TryRunShortcutCommand = True
        Exit Function
    End If

    If Len(keyText) = 0 Then
        FAS_SetStatus "Shortcut key cannot be blank."
        FAS_TryRunShortcutCommand = True
        Exit Function
    End If

    FAS_SetSetting FAS_SHORTCUT_PREFIX & actionId, keyText
    FAS_RegisterConfiguredShortcuts
    FAS_SetStatus "Shortcut saved for " & actionId & "."
    FAS_TryRunShortcutCommand = True
End Function

Public Function FAS_IsKnownShortcutAction(ByVal actionId As String) As Boolean
    Dim actions As Variant
    Dim i As Long

    actions = FAS_GetShortcutActions()
    For i = LBound(actions) To UBound(actions)
        If StrComp(LCase$(Trim$(actionId)), CStr(actions(i)(0)), vbTextCompare) = 0 Then
            FAS_IsKnownShortcutAction = True
            Exit Function
        End If
    Next i
End Function

Public Sub FAS_ResetShortcuts()
    Dim actions As Variant
    Dim i As Long

    actions = FAS_GetShortcutActions()
    For i = LBound(actions) To UBound(actions)
        FAS_DeleteSetting FAS_SHORTCUT_PREFIX & CStr(actions(i)(0))
    Next i
End Sub

Public Sub FAS_ShowShortcutStatus(Optional ByVal actionId As String = vbNullString)
    Dim actions As Variant
    Dim i As Long
    Dim statusText As String

    If Len(Trim$(actionId)) > 0 Then
        If Not FAS_IsKnownShortcutAction(actionId) Then
            FAS_SetStatus "Unknown shortcut action: " & actionId
            Exit Sub
        End If
    End If

    actions = FAS_GetShortcutActions()
    For i = LBound(actions) To UBound(actions)
        If Len(Trim$(actionId)) = 0 Or StrComp(actionId, CStr(actions(i)(0)), vbTextCompare) = 0 Then
            statusText = statusText & CStr(actions(i)(1)) & ": " & FAS_GetShortcutKey(CStr(actions(i)(0)), CStr(actions(i)(2))) & vbCrLf
        End If
    Next i

    If Application.Visible Then
        MsgBox statusText, vbInformation, "FAS Shortcuts"
    Else
        FAS_SetStatus Replace(Replace(statusText, vbCrLf, "; "), vbCr, vbNullString)
    End If
End Sub

Private Function FAS_GetShortcutActions() As Variant
    Dim actions(0 To 6) As Variant

    actions(0) = Array("navigator", "Navigator", "^+N", "FAS_ToggleNavigator")
    actions(1) = Array("palette", "Command Palette", "^+J", "FAS_ToggleCommandPalette")
    actions(2) = Array("subtotalabove", "Subtotal Above", "^+P", "FAS_InsertSubtotalAboveSelection")
    actions(3) = Array("tb", "TB", "^+T", "FAS_GoToTB")
    actions(4) = Array("gl", "GL", "^+G", "FAS_GoToGL")
    actions(5) = Array("precedent", "Cycle Precedent", "^%{UP}", "FAS_CycleDirectPrecedent")
    actions(6) = Array("back", "Back to Trace Origin", "^%{DOWN}", "FAS_BackToTraceOrigin")

    FAS_GetShortcutActions = actions
End Function

Private Function FAS_GetShortcutKey(ByVal actionId As String, ByVal defaultKey As String) As String
    FAS_GetShortcutKey = Trim$(FAS_GetSetting(FAS_SHORTCUT_PREFIX & actionId, defaultKey))
End Function

Private Function FAS_QualifiedMacro(ByVal macroName As String) As String
    FAS_QualifiedMacro = "'" & ThisWorkbook.Name & "'!" & macroName
End Function
