Attribute VB_Name = "modStatusBar"
Option Explicit

Public Sub FAS_SetStatus(ByVal messageText As String)
    On Error Resume Next
    If Len(Trim$(messageText)) = 0 Then
        Application.StatusBar = False
    Else
        Application.StatusBar = "FAS: " & messageText
    End If
    On Error GoTo 0
End Sub

Public Sub FAS_ClearStatus()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub
