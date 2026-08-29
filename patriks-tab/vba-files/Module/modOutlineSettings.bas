Attribute VB_Name = "modOutlineSettings"
Option Explicit

Public Const FAS_DEFAULT_OUTLINE_ABOVE_LEFT As Boolean = True

Private Const FAS_OUTLINE_ABOVE_LEFT_SETTING As String = "View.OutlineAboveLeft"

Public Function FAS_GetOutlineAboveLeft() As Boolean
    FAS_GetOutlineAboveLeft = FAS_GetBooleanSetting(FAS_OUTLINE_ABOVE_LEFT_SETTING, FAS_DEFAULT_OUTLINE_ABOVE_LEFT)
End Function

Public Sub FAS_SetOutlineAboveLeft(ByVal outlineAboveLeft As Boolean)
    FAS_SetBooleanSetting FAS_OUTLINE_ABOVE_LEFT_SETTING, outlineAboveLeft
    FAS_ApplyOutlinePreference

    If outlineAboveLeft Then
        FAS_SetStatus "Outline summaries set above and left of detail."
    Else
        FAS_SetStatus "Outline summaries set below and right of detail."
    End If
End Sub

Public Sub FAS_ApplyOutlinePreference(Optional ByVal targetWorksheet As Object)
    Dim outlineSheet As Object

    If Not targetWorksheet Is Nothing Then Set outlineSheet = targetWorksheet
    If outlineSheet Is Nothing Then
        If TypeName(Application.ActiveSheet) = "Worksheet" Then Set outlineSheet = Application.ActiveSheet
    End If
    If outlineSheet Is Nothing Or TypeName(outlineSheet) <> "Worksheet" Then Exit Sub

    On Error GoTo ApplyFailed
    If FAS_GetOutlineAboveLeft() Then
        outlineSheet.Outline.SummaryRow = xlAbove
        outlineSheet.Outline.SummaryColumn = xlLeft
    Else
        outlineSheet.Outline.SummaryRow = xlBelow
        outlineSheet.Outline.SummaryColumn = xlRight
    End If
    Exit Sub

ApplyFailed:
    On Error Resume Next
    If Application.Visible Then FAS_SetStatus "Could not update outline direction for this worksheet."
    On Error GoTo 0
End Sub
