Attribute VB_Name = "modShortcutManager"
Option Explicit

Private Const FAS_SHORTCUTS_ENABLED_KEY As String = "Shortcuts.Enabled"
Private Const FAS_SHORTCUT_PREFIX As String = "Shortcut."

Private mRegisteredKeys As Object
Private mFASShortcutDefinitions As Collection

Public Sub FAS_RegisterConfiguredShortcuts()
    Dim definitions As Collection
    Dim item As Variant
    Dim displayText As String
    Dim keySpec As String
    Dim errorText As String
    Dim keyName As String

    FAS_UnregisterConfiguredShortcuts

    If Not FAS_GetBooleanSetting(FAS_SHORTCUTS_ENABLED_KEY, True) Then
        FAS_SetStatus "FAS shortcuts are off."
        Exit Sub
    End If

    Set mRegisteredKeys = CreateObject("Scripting.Dictionary")
    mRegisteredKeys.CompareMode = vbTextCompare
    Set definitions = FAS_GetShortcutDefinitions()

    For Each item In definitions
        displayText = FAS_GetConfiguredShortcut(CStr(item(0)))
        If Len(displayText) > 0 Then
            errorText = vbNullString
            keySpec = FAS_ShortcutDisplayToOnKey(displayText, errorText)
            If Len(keySpec) = 0 Then
                FAS_SetStatus "Skipped invalid shortcut for " & CStr(item(1)) & "."
            Else
                keyName = LCase$(keySpec)
                If mRegisteredKeys.Exists(keyName) Then
                    FAS_SetStatus "Skipped duplicate shortcut " & displayText & "."
                Else
                    Application.OnKey keySpec, FAS_GetQualifiedMacroName(CStr(item(4)))
                    mRegisteredKeys.Add keyName, keySpec
                End If
            End If
        End If
    Next item
End Sub

Public Sub FAS_UnregisterConfiguredShortcuts()
    Dim definitions As Collection
    Dim item As Variant
    Dim registeredKey As Variant
    Dim defaultKey As String
    Dim configuredKey As String
    Dim errorText As String

    On Error Resume Next

    If Not mRegisteredKeys Is Nothing Then
        For Each registeredKey In mRegisteredKeys.Items
            Application.OnKey CStr(registeredKey)
        Next registeredKey
    End If

    Set definitions = FAS_GetShortcutDefinitions()
    For Each item In definitions
        errorText = vbNullString
        defaultKey = FAS_ShortcutDisplayToOnKey(CStr(item(2)), errorText)
        If Len(defaultKey) > 0 Then Application.OnKey defaultKey

        errorText = vbNullString
        configuredKey = FAS_ShortcutDisplayToOnKey(FAS_GetConfiguredShortcut(CStr(item(0))), errorText)
        If Len(configuredKey) > 0 And StrComp(configuredKey, defaultKey, vbTextCompare) <> 0 Then
            Application.OnKey configuredKey
        End If
    Next item

    Set mRegisteredKeys = Nothing
    On Error GoTo 0
End Sub

Public Function FAS_GetShortcutDefinitions() As Collection
    Dim definitions As Collection

    If Not mFASShortcutDefinitions Is Nothing Then
        Set FAS_GetShortcutDefinitions = mFASShortcutDefinitions
        Exit Function
    End If

    Set definitions = New Collection
    definitions.Add Array("navigator", "Navigator", "Ctrl+Shift+N", "FAS_ToggleNavigator", "navigator nav")
    definitions.Add Array("palette", "Command Palette", "Ctrl+Shift+J", "FAS_ToggleCommandPalette", "palette commandpalette")
    definitions.Add Array("subtotalabove", "Subtotal Above", "Ctrl+Shift+P", "FAS_InsertSubtotalAboveSelection", "subtotalabove subtotal")
    definitions.Add Array("tb", "Go to TB", "Ctrl+Shift+T", "FAS_GoToTB", "tb t")
    definitions.Add Array("gl", "Go to GL", "Ctrl+Shift+G", "FAS_GoToGL", "gl g")
    definitions.Add Array("precedent", "Cycle Precedent", "Ctrl+Alt+Up", "FAS_CycleDirectPrecedent", "precedent pre")
    definitions.Add Array("back", "Back to Trace Origin", "Ctrl+Alt+Down", "FAS_BackToTraceOrigin", "back")
    definitions.Add Array("numberformat", "Cycle Number Format", "Ctrl+Shift+1", "FAS_CycleNumberFormat", "number format fmt1")
    definitions.Add Array("percentformat", "Cycle Percent Format", "Ctrl+Shift+5", "FAS_CyclePercentFormat", "percent percentage format fmt5")
    definitions.Add Array("currencyformat", "Cycle Currency Format", "Ctrl+Shift+4", "FAS_CycleCurrencyFormat", "currency format fmt4")

    Set mFASShortcutDefinitions = definitions
    Set FAS_GetShortcutDefinitions = mFASShortcutDefinitions
End Function

Public Function FAS_GetConfiguredShortcut(ByVal actionId As String) As String
    Dim definitions As Collection
    Dim item As Variant

    Set definitions = FAS_GetShortcutDefinitions()
    For Each item In definitions
        If StrComp(actionId, CStr(item(0)), vbTextCompare) = 0 Then
            FAS_GetConfiguredShortcut = FAS_GetSetting(FAS_SHORTCUT_PREFIX & CStr(item(0)), CStr(item(2)))
            Exit Function
        End If
    Next item
End Function

Public Function FAS_GetDefaultShortcut(ByVal actionId As String) As String
    Dim definitions As Collection
    Dim item As Variant

    Set definitions = FAS_GetShortcutDefinitions()
    For Each item In definitions
        If StrComp(actionId, CStr(item(0)), vbTextCompare) = 0 Then
            FAS_GetDefaultShortcut = CStr(item(2))
            Exit Function
        End If
    Next item
End Function

Public Function FAS_IsKnownShortcutAction(ByVal actionId As String) As Boolean
    Dim definitions As Collection
    Dim item As Variant

    Set definitions = FAS_GetShortcutDefinitions()
    For Each item In definitions
        If StrComp(LCase$(Trim$(actionId)), LCase$(CStr(item(0))), vbTextCompare) = 0 Then
            FAS_IsKnownShortcutAction = True
            Exit Function
        End If
    Next item
End Function

Public Function FAS_SaveShortcut(ByVal actionId As String, ByVal shortcutText As String, Optional ByRef errorText As String) As Boolean
    Dim definitions As Collection
    Dim item As Variant
    Dim normalizedText As String
    Dim keySpec As String
    Dim otherKeySpec As String

    errorText = vbNullString
    If Not FAS_IsKnownShortcutAction(actionId) Then
        errorText = "Unknown shortcut action."
        Exit Function
    End If

    normalizedText = FAS_NormalizeShortcutDisplay(shortcutText, errorText)
    If Len(errorText) > 0 Then Exit Function

    If Len(normalizedText) > 0 Then
        keySpec = FAS_ShortcutDisplayToOnKey(normalizedText, errorText)
        If Len(errorText) > 0 Or Len(keySpec) = 0 Then
            If Len(errorText) = 0 Then errorText = "The shortcut is not valid."
            Exit Function
        End If

        Set definitions = FAS_GetShortcutDefinitions()
        For Each item In definitions
            If StrComp(actionId, CStr(item(0)), vbTextCompare) <> 0 Then
                otherKeySpec = FAS_ShortcutDisplayToOnKey(FAS_GetConfiguredShortcut(CStr(item(0))), errorText)
                errorText = vbNullString
                If Len(otherKeySpec) > 0 And StrComp(keySpec, otherKeySpec, vbTextCompare) = 0 Then
                    errorText = "That shortcut is already assigned to " & CStr(item(1)) & "."
                    Exit Function
                End If
            End If
        Next item
    End If

    If Len(normalizedText) = 0 Or StrComp(normalizedText, FAS_GetDefaultShortcut(actionId), vbTextCompare) = 0 Then
        FAS_DeleteSetting FAS_SHORTCUT_PREFIX & actionId
    Else
        FAS_SetSetting FAS_SHORTCUT_PREFIX & actionId, normalizedText
    End If

    FAS_RegisterConfiguredShortcuts
    FAS_SaveShortcut = True
End Function

Public Sub FAS_ResetShortcuts()
    Dim definitions As Collection
    Dim item As Variant

    Set definitions = FAS_GetShortcutDefinitions()
    For Each item In definitions
        FAS_DeleteSetting FAS_SHORTCUT_PREFIX & CStr(item(0))
    Next item
    FAS_DeleteSetting FAS_SHORTCUTS_ENABLED_KEY
End Sub

Public Function FAS_GetShortcutLabelForCommand(ByVal commandText As String) As String
    Dim definitions As Collection
    Dim item As Variant
    Dim aliasText As Variant
    Dim normalizedCommand As String

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    Set definitions = FAS_GetShortcutDefinitions()

    For Each item In definitions
        For Each aliasText In Split(CStr(item(4)), " ")
            If Len(CStr(aliasText)) > 0 And normalizedCommand = FAS_NormalizeCommandText(CStr(aliasText)) Then
                FAS_GetShortcutLabelForCommand = FAS_GetConfiguredShortcut(CStr(item(0)))
                Exit Function
            End If
        Next aliasText
    Next item
End Function

Public Function FAS_TryRunShortcutCommand(ByVal commandText As String) As Boolean
    Dim normalizedText As String

    normalizedText = LCase$(Trim$(commandText))
    Select Case normalizedText
        Case "shortcut", "shortcuts", "keys", "shortcut settings", "shortcutsettings"
            FAS_ShowShortcutSettings
            FAS_TryRunShortcutCommand = True
        Case "shortcut on", "shortcuts on", "keys on"
            FAS_SetBooleanSetting FAS_SHORTCUTS_ENABLED_KEY, True
            FAS_RegisterConfiguredShortcuts
            FAS_SetStatus "FAS shortcuts on."
            FAS_TryRunShortcutCommand = True
        Case "shortcut off", "shortcuts off", "keys off"
            FAS_SetBooleanSetting FAS_SHORTCUTS_ENABLED_KEY, False
            FAS_UnregisterConfiguredShortcuts
            FAS_SetStatus "FAS shortcuts off."
            FAS_TryRunShortcutCommand = True
        Case "shortcut reset", "shortcuts reset", "keys reset", "shortcut default", "shortcuts default"
            If Not Application.Visible Or MsgBox("Restore the default FAS keyboard shortcuts?", vbQuestion + vbYesNo, "FAS Shortcut Settings") = vbYes Then
                FAS_ResetShortcuts
                FAS_RegisterConfiguredShortcuts
                FAS_SetStatus "FAS shortcuts reset to defaults."
            End If
            FAS_TryRunShortcutCommand = True
    End Select
End Function

Public Function FAS_IsKnownShortcutCommand(ByVal commandText As String) As Boolean
    Dim normalizedText As String

    normalizedText = LCase$(Trim$(commandText))
    Select Case normalizedText
        Case "shortcut", "shortcuts", "keys", "shortcut settings", "shortcutsettings", _
             "shortcut on", "shortcuts on", "keys on", "shortcut off", "shortcuts off", "keys off", _
             "shortcut reset", "shortcuts reset", "keys reset", "shortcut default", "shortcuts default"
            FAS_IsKnownShortcutCommand = True
    End Select
End Function

Public Function FAS_NormalizeShortcutDisplay(ByVal shortcutText As String, Optional ByRef errorText As String) As String
    Dim rawText As String
    Dim parts As Variant
    Dim keyText As String
    Dim displayKey As String
    Dim onKeyText As String
    Dim hasCtrl As Boolean
    Dim hasAlt As Boolean
    Dim hasShift As Boolean
    Dim partIndex As Long
    Dim partText As String

    errorText = vbNullString
    rawText = Replace(Trim$(shortcutText), " ", vbNullString)
    If Len(rawText) = 0 Then Exit Function

    parts = Split(rawText, "+")
    If UBound(parts) < 1 Then
        errorText = "Use a modifier and key, for example Ctrl+Shift+J."
        Exit Function
    End If

    For partIndex = 0 To UBound(parts) - 1
        partText = LCase$(CStr(parts(partIndex)))
        Select Case partText
            Case "ctrl", "control"
                If hasCtrl Then errorText = "Ctrl appears more than once."
                hasCtrl = True
            Case "alt", "option"
                If hasAlt Then errorText = "Alt appears more than once."
                hasAlt = True
            Case "shift"
                If hasShift Then errorText = "Shift appears more than once."
                hasShift = True
            Case Else
                errorText = "Only Ctrl, Alt, and Shift can be modifiers."
        End Select
        If Len(errorText) > 0 Then Exit Function
    Next partIndex

    keyText = CStr(parts(UBound(parts)))
    If Not FAS_NormalizeShortcutKey(keyText, displayKey, onKeyText, errorText) Then Exit Function
    If Not hasCtrl And Not hasAlt And Not hasShift Then
        errorText = "Use at least one modifier: Ctrl, Alt, or Shift."
        Exit Function
    End If

    If hasCtrl Then FAS_NormalizeShortcutDisplay = "Ctrl+"
    If hasAlt Then FAS_NormalizeShortcutDisplay = FAS_NormalizeShortcutDisplay & "Alt+"
    If hasShift Then FAS_NormalizeShortcutDisplay = FAS_NormalizeShortcutDisplay & "Shift+"
    FAS_NormalizeShortcutDisplay = FAS_NormalizeShortcutDisplay & displayKey
End Function

Public Function FAS_ShortcutDisplayToOnKey(ByVal shortcutText As String, Optional ByRef errorText As String) As String
    Dim normalizedText As String
    Dim parts As Variant
    Dim prefixText As String

    errorText = vbNullString
    normalizedText = FAS_NormalizeShortcutDisplay(shortcutText, errorText)
    If Len(errorText) > 0 Or Len(normalizedText) = 0 Then Exit Function

    parts = Split(normalizedText, "+")
    If UBound(parts) > 0 Then
        If StrComp(CStr(parts(0)), "Ctrl", vbTextCompare) = 0 Then prefixText = prefixText & "^"
        If StrComp(CStr(parts(0)), "Alt", vbTextCompare) = 0 Then prefixText = prefixText & "%"
        If StrComp(CStr(parts(0)), "Shift", vbTextCompare) = 0 Then prefixText = prefixText & "+"
    End If
    If UBound(parts) > 1 Then
        If StrComp(CStr(parts(1)), "Alt", vbTextCompare) = 0 Then prefixText = prefixText & "%"
        If StrComp(CStr(parts(1)), "Shift", vbTextCompare) = 0 Then prefixText = prefixText & "+"
    End If
    If UBound(parts) > 2 Then
        If StrComp(CStr(parts(2)), "Shift", vbTextCompare) = 0 Then prefixText = prefixText & "+"
    End If

    FAS_ShortcutDisplayToOnKey = prefixText & FAS_OnKeyTokenForDisplayKey(CStr(parts(UBound(parts))))
End Function

Private Function FAS_NormalizeShortcutKey(ByVal keyText As String, ByRef displayKey As String, ByRef onKeyText As String, ByRef errorText As String) As Boolean
    Dim numericPart As Long
    Dim normalizedKey As String

    normalizedKey = LCase$(Trim$(keyText))
    Select Case normalizedKey
        Case "up", "uparrow"
            displayKey = "Up"
            onKeyText = "{UP}"
        Case "down", "downarrow"
            displayKey = "Down"
            onKeyText = "{DOWN}"
        Case "left", "leftarrow"
            displayKey = "Left"
            onKeyText = "{LEFT}"
        Case "right", "rightarrow"
            displayKey = "Right"
            onKeyText = "{RIGHT}"
        Case "enter", "return"
            displayKey = "Enter"
            onKeyText = "~"
        Case "escape", "esc"
            displayKey = "Esc"
            onKeyText = "{ESC}"
        Case "tab"
            displayKey = "Tab"
            onKeyText = "{TAB}"
        Case "home"
            displayKey = "Home"
            onKeyText = "{HOME}"
        Case "end"
            displayKey = "End"
            onKeyText = "{END}"
        Case "pageup", "pgup"
            displayKey = "PageUp"
            onKeyText = "{PGUP}"
        Case "pagedown", "pgdn"
            displayKey = "PageDown"
            onKeyText = "{PGDN}"
        Case "insert", "ins"
            displayKey = "Insert"
            onKeyText = "{INSERT}"
        Case "delete", "del"
            displayKey = "Delete"
            onKeyText = "{DELETE}"
        Case "backspace", "back"
            displayKey = "Backspace"
            onKeyText = "{BACKSPACE}"
        Case Else
            If Len(normalizedKey) = 1 And normalizedKey Like "[a-z0-9]" Then
                displayKey = UCase$(normalizedKey)
                onKeyText = displayKey
            ElseIf Left$(normalizedKey, 1) = "f" And IsNumeric(Mid$(normalizedKey, 2)) Then
                numericPart = CLng(Mid$(normalizedKey, 2))
                If numericPart < 1 Or numericPart > 12 Then
                    errorText = "Function keys F1 through F12 are supported."
                    Exit Function
                End If
                displayKey = "F" & numericPart
                onKeyText = "{F" & numericPart & "}"
            Else
                errorText = "Unsupported key. Use a letter, number, arrow, Enter, or F1-F12."
                Exit Function
            End If
    End Select

    FAS_NormalizeShortcutKey = True
End Function

Private Function FAS_OnKeyTokenForDisplayKey(ByVal displayKey As String) As String
    Select Case LCase$(displayKey)
        Case "up": FAS_OnKeyTokenForDisplayKey = "{UP}"
        Case "down": FAS_OnKeyTokenForDisplayKey = "{DOWN}"
        Case "left": FAS_OnKeyTokenForDisplayKey = "{LEFT}"
        Case "right": FAS_OnKeyTokenForDisplayKey = "{RIGHT}"
        Case "enter": FAS_OnKeyTokenForDisplayKey = "~"
        Case "esc": FAS_OnKeyTokenForDisplayKey = "{ESC}"
        Case "tab": FAS_OnKeyTokenForDisplayKey = "{TAB}"
        Case "home": FAS_OnKeyTokenForDisplayKey = "{HOME}"
        Case "end": FAS_OnKeyTokenForDisplayKey = "{END}"
        Case "pageup": FAS_OnKeyTokenForDisplayKey = "{PGUP}"
        Case "pagedown": FAS_OnKeyTokenForDisplayKey = "{PGDN}"
        Case "insert": FAS_OnKeyTokenForDisplayKey = "{INSERT}"
        Case "delete": FAS_OnKeyTokenForDisplayKey = "{DELETE}"
        Case "backspace": FAS_OnKeyTokenForDisplayKey = "{BACKSPACE}"
        Case Else
            If Left$(LCase$(displayKey), 1) = "f" Then
                FAS_OnKeyTokenForDisplayKey = "{" & UCase$(displayKey) & "}"
            Else
                FAS_OnKeyTokenForDisplayKey = UCase$(displayKey)
            End If
    End Select
End Function
