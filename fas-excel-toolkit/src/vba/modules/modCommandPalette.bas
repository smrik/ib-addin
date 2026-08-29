Attribute VB_Name = "modCommandPalette"
Option Explicit

Public gFASCommandPalette As frmFASCommandPalette

Public Sub FAS_ShowCommandPalette(Optional ByVal initialQuery As String = vbNullString)
    Set gFASCommandPalette = FAS_GetOrCreateCommandPalette()
    gFASCommandPalette.SetInitialQuery initialQuery

    If Not gFASCommandPalette.Visible Then
        gFASCommandPalette.Show vbModeless
    End If

    gFASCommandPalette.FocusSearch
End Sub

Public Sub FAS_ToggleCommandPalette()
    Set gFASCommandPalette = FAS_GetOrCreateCommandPalette()

    If gFASCommandPalette.Visible Then
        Unload gFASCommandPalette
        Set gFASCommandPalette = Nothing
    Else
        FAS_ShowCommandPalette vbNullString
    End If
End Sub

Public Sub FAS_ClearCommandPaletteReference()
    Set gFASCommandPalette = Nothing
End Sub

Private Function FAS_GetOrCreateCommandPalette() As frmFASCommandPalette
    If gFASCommandPalette Is Nothing Then
        Set gFASCommandPalette = New frmFASCommandPalette
    End If

    Set FAS_GetOrCreateCommandPalette = gFASCommandPalette
End Function
