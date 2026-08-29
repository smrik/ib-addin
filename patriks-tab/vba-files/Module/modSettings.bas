Attribute VB_Name = "modSettings"
Option Explicit

Private Const FAS_SETTINGS_DIR As String = "FAS_Toolkit\Settings"
Private Const FAS_SETTINGS_FILE As String = "settings.ini"

Private Const FAS_SETTINGS_FORM_NAME As String = "frmFASSettings"

' Keep this late-bound because XVBA can import a changed module before it
' imports a newly added UserForm during an incremental sync.
Public gFASSettings As Object

Public Sub FAS_ShowSettings()
    On Error GoTo SettingsFormMissing
    Set gFASSettings = FAS_GetOrCreateSettings()
    If Not gFASSettings.Visible Then gFASSettings.Show vbModeless
    Exit Sub

SettingsFormMissing:
    Set gFASSettings = Nothing
    FAS_SetStatus "FAS Settings is not loaded. Run one full XVBA Production export."
End Sub

Public Sub FAS_ClearSettingsReference()
    Set gFASSettings = Nothing
End Sub

Private Function FAS_GetOrCreateSettings() As Object
    If gFASSettings Is Nothing Then
        Set gFASSettings = UserForms.Add(FAS_SETTINGS_FORM_NAME)
    End If
    Set FAS_GetOrCreateSettings = gFASSettings
End Function

Public Function FAS_GetSettingsPath() As String
    FAS_GetSettingsPath = Environ$("APPDATA") & "\" & FAS_SETTINGS_DIR & "\" & FAS_SETTINGS_FILE
End Function

Public Function FAS_GetSetting(ByVal keyName As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim settings As Object

    Set settings = FAS_LoadSettings()
    If settings.Exists(keyName) Then
        FAS_GetSetting = CStr(settings(keyName))
    Else
        FAS_GetSetting = defaultValue
    End If
End Function

Public Sub FAS_SetSetting(ByVal keyName As String, ByVal valueText As String)
    Dim settings As Object

    Set settings = FAS_LoadSettings()
    settings(keyName) = valueText
    FAS_SaveSettings settings
End Sub

Public Function FAS_GetBooleanSetting(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    Dim valueText As String

    valueText = LCase$(Trim$(FAS_GetSetting(keyName, IIf(defaultValue, "true", "false"))))
    FAS_GetBooleanSetting = (valueText = "true" Or valueText = "yes" Or valueText = "1" Or valueText = "on")
End Function

Public Sub FAS_SetBooleanSetting(ByVal keyName As String, ByVal valueFlag As Boolean)
    FAS_SetSetting keyName, IIf(valueFlag, "true", "false")
End Sub

Public Sub FAS_DeleteSetting(ByVal keyName As String)
    Dim settings As Object

    Set settings = FAS_LoadSettings()
    If settings.Exists(keyName) Then settings.Remove keyName
    FAS_SaveSettings settings
End Sub

Public Sub FAS_ResetSettings()
    On Error Resume Next
    If Len(Dir$(FAS_GetSettingsPath())) > 0 Then Kill FAS_GetSettingsPath()
    On Error GoTo 0
End Sub

Private Function FAS_LoadSettings() As Object
    Dim settings As Object
    Dim fileNumber As Integer
    Dim lineText As String
    Dim separatorPosition As Long
    Dim keyName As String

    Set settings = CreateObject("Scripting.Dictionary")
    settings.CompareMode = vbTextCompare

    On Error GoTo LoadFailed
    If Len(Dir$(FAS_GetSettingsPath())) = 0 Then
        Set FAS_LoadSettings = settings
        Exit Function
    End If

    fileNumber = FreeFile
    Open FAS_GetSettingsPath() For Input As #fileNumber
    Do Until EOF(fileNumber)
        Line Input #fileNumber, lineText
        lineText = Trim$(lineText)
        separatorPosition = InStr(1, lineText, "=", vbBinaryCompare)
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" And separatorPosition > 1 Then
            keyName = Trim$(Left$(lineText, separatorPosition - 1))
            If Len(keyName) > 0 Then
                settings(keyName) = Mid$(lineText, separatorPosition + 1)
            End If
        End If
    Loop
    Close #fileNumber
    Set FAS_LoadSettings = settings
    Exit Function

LoadFailed:
    If fileNumber > 0 Then Close #fileNumber
    Set FAS_LoadSettings = settings
End Function

Private Sub FAS_SaveSettings(ByVal settings As Object)
    Dim fileNumber As Integer
    Dim keyName As Variant

    On Error GoTo SaveFailed
    FAS_EnsureSettingsFolder

    fileNumber = FreeFile
    Open FAS_GetSettingsPath() For Output As #fileNumber
    Print #fileNumber, "# FAS Toolkit settings"
    For Each keyName In settings.Keys
        Print #fileNumber, CStr(keyName) & "=" & CStr(settings(keyName))
    Next keyName
    Close #fileNumber
    Exit Sub

SaveFailed:
    If fileNumber > 0 Then Close #fileNumber
End Sub

Private Sub FAS_EnsureSettingsFolder()
    Dim fileSystem As Object
    Dim folderPath As String

    folderPath = Environ$("APPDATA") & "\" & FAS_SETTINGS_DIR
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    FAS_EnsureFolderRecursive fileSystem, folderPath
End Sub

Private Sub FAS_EnsureFolderRecursive(ByVal fileSystem As Object, ByVal folderPath As String)
    Dim parentPath As String

    If Len(folderPath) = 0 Or fileSystem.FolderExists(folderPath) Then
        Exit Sub
    End If

    parentPath = fileSystem.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And Not fileSystem.FolderExists(parentPath) Then
        FAS_EnsureFolderRecursive fileSystem, parentPath
    End If

    If Not fileSystem.FolderExists(folderPath) Then
        fileSystem.CreateFolder folderPath
    End If
End Sub
