Attribute VB_Name = "modFASNavigator"
Option Explicit

Public gFASSidebar As frmFASSidebar

Public Sub FAS_ShowNavigator()
    Set gFASSidebar = FAS_GetOrCreateNavigator()
    gFASSidebar.Show vbModeless
End Sub

Public Sub FAS_ToggleNavigator()
    Set gFASSidebar = FAS_GetOrCreateNavigator()

    If FAS_IsNavigatorVisible(gFASSidebar) Then
        Unload gFASSidebar
        Set gFASSidebar = Nothing
    Else
        gFASSidebar.Show vbModeless
    End If
End Sub

Public Sub FAS_ClearNavigatorReference()
    Set gFASSidebar = Nothing
End Sub

Private Function FAS_GetOrCreateNavigator() As frmFASSidebar
    If gFASSidebar Is Nothing Then
        Set gFASSidebar = New frmFASSidebar
    End If

    Set FAS_GetOrCreateNavigator = gFASSidebar
End Function

Private Function FAS_IsNavigatorVisible(ByVal navigatorForm As frmFASSidebar) As Boolean
    On Error Resume Next
    FAS_IsNavigatorVisible = navigatorForm.Visible
    On Error GoTo 0
End Function
