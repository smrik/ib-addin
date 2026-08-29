Attribute VB_Name = "modSettings"
Option Explicit

Private Const FAS_SETTINGS_DIR As String = "FAS_Toolkit\Settings"
Private Const FAS_SETTINGS_FILE As String = "settings.ini"

Public Function FAS_GetSettingsPath() As String
    FAS_GetSettingsPath = Environ$("APPDATA") & "\" & FAS_SETTINGS_DIR & "\" & FAS_SETTINGS_FILE
End Function

Public Function FAS_GetSetting(ByVal keyName As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim d As Object
    Set d = FAS_LoadSettings()
    If d.Exists(keyName) Then
        FAS_GetSetting = CStr(d(keyName))
    Else
        FAS_GetSetting = defaultValue
    End If
End Function

Public Sub FAS_SetSetting(ByVal keyName As String, ByVal valueText As String)
    Dim d As Object
    Set d = FAS_LoadSettings()
    d(keyName) = valueText
    FAS_SaveSettings d
End Sub

Public Sub FAS_DeleteSetting(ByVal keyName As String)
    Dim d As Object
    Set d = FAS_LoadSettings()
    If d.Exists(keyName) Then d.Remove keyName
    FAS_SaveSettings d
End Sub

Public Function FAS_GetBooleanSetting(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    Dim v As String
    v = LCase$(Trim$(FAS_GetSetting(keyName, IIf(defaultValue, "true", "false"))))
    FAS_GetBooleanSetting = (v = "true" Or v = "yes" Or v = "1" Or v = "on")
End Function

Public Sub FAS_SetBooleanSetting(ByVal keyName As String, ByVal valueFlag As Boolean)
    FAS_SetSetting keyName, IIf(valueFlag, "true", "false")
End Sub

Public Function FAS_TryRunSettingsCommand(ByVal commandText As String) As Boolean
    Dim c As String
    c = FAS_NormalizeCommandText(commandText)

    Select Case c
        Case "settings", "setting", "config", "status"
            FAS_ShowSettingsStatus
            FAS_TryRunSettingsCommand = True

        Case "settings reset", "setting reset", "config reset"
            If MsgBox("Reset FAS Toolkit settings?", vbQuestion + vbYesNo, "FAS Toolkit") = vbYes Then
                FAS_ResetSettings
                FAS_RegisterShortcuts
                FAS_SetStatus "Settings reset."
            Else
                FAS_SetStatus "Settings reset cancelled."
            End If
            FAS_TryRunSettingsCommand = True
    End Select
End Function

Public Function FAS_IsSettingsCommand(ByVal commandText As String) As Boolean
    Dim c As String
    c = FAS_NormalizeCommandText(commandText)
    Select Case c
        Case "settings", "setting", "config", "status", "settings reset", "setting reset", "config reset"
            FAS_IsSettingsCommand = True
    End Select
End Function

Public Sub FAS_ShowSettingsStatus()
    MsgBox "FAS Toolkit settings" & vbCrLf & vbCrLf & _
        "Settings file:" & vbCrLf & FAS_GetSettingsPath() & vbCrLf & vbCrLf & _
        "Shortcuts enabled: " & IIf(FAS_GetBooleanSetting("Shortcuts.Enabled", True), "Yes", "No"), _
        vbInformation, "FAS Toolkit"
End Sub

Public Sub FAS_ResetSettings()
    On Error Resume Next
    If Len(Dir$(FAS_GetSettingsPath())) > 0 Then Kill FAS_GetSettingsPath()
    On Error GoTo 0
End Sub

Private Function FAS_LoadSettings() As Object
    Dim d As Object
    Dim p As String
    Dim f As Integer
    Dim lineText As String
    Dim pos As Long
    Dim k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    p = FAS_GetSettingsPath()
    If Len(Dir$(p)) = 0 Then
        Set FAS_LoadSettings = d
        Exit Function
    End If

    f = FreeFile
    Open p For Input As #f
    Do Until EOF(f)
        Line Input #f, lineText
        lineText = Trim$(lineText)
        pos = InStr(1, lineText, "=", vbBinaryCompare)
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" And pos > 1 Then
            k = Trim$(Left$(lineText, pos - 1))
            If Len(k) > 0 Then d(k) = Mid$(lineText, pos + 1)
        End If
    Loop
    Close #f

    Set FAS_LoadSettings = d
End Function

Private Sub FAS_SaveSettings(ByVal d As Object)
    Dim p As String
    Dim folderPath As String
    Dim f As Integer
    Dim k As Variant

    folderPath = Environ$("APPDATA") & "\" & FAS_SETTINGS_DIR
    FAS_EnsureFolder folderPath

    p = FAS_GetSettingsPath()
    f = FreeFile
    Open p For Output As #f
    Print #f, "# FAS Toolkit settings"
    For Each k In d.Keys
        Print #f, CStr(k) & "=" & CStr(d(k))
    Next k
    Close #f
End Sub

Public Sub FAS_EnsureFolder(ByVal folderPath As String)
    Dim parts As Variant
    Dim cur As String
    Dim i As Long

    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    parts = Split(folderPath, "\")
    cur = CStr(parts(0))
    For i = 1 To UBound(parts)
        cur = cur & "\" & CStr(parts(i))
        If Len(Dir$(cur, vbDirectory)) = 0 Then
            On Error Resume Next
            MkDir cur
            On Error GoTo 0
        End If
    Next i
End Sub
