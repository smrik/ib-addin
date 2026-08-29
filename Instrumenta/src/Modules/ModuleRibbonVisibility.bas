Attribute VB_Name = "ModuleRibbonVisibility"
'MIT License
'
'Copyright (c) 2021 - 2026 iappyx

Option Explicit

Private Const RIBBON_VISIBILITY_APP As String = "Instrumenta"
Private Const RIBBON_VISIBILITY_SECTION As String = "RibbonVisibility"

Private Const RIBBON_KEY_STYLES As String = "Styles"
Private Const RIBBON_KEY_TABLES As String = "Tables"
Private Const RIBBON_KEY_MAIL_MERGE As String = "MailMerge"
Private Const RIBBON_KEY_ADVANCED_EXTRAS As String = "AdvancedExtras"

Public Sub InstrumentaGetVisibleWithUserPreference(control As IRibbonControl, ByRef visible)
    visible = InstrumentaBaseVisibleForControl(control) And RibbonVisibilityIsEnabledForControl(control.id)
End Sub

Public Sub RibbonVisibilityShowSettings()
    MsgBox RibbonVisibilitySummaryText(), vbInformation, "Instrumenta Ribbon"
End Sub

Public Sub RibbonVisibilityToggleStyles()
    RibbonVisibilityToggleKey RIBBON_KEY_STYLES, "Styles"
End Sub

Public Sub RibbonVisibilityToggleTables()
    RibbonVisibilityToggleKey RIBBON_KEY_TABLES, "Tables"
End Sub

Public Sub RibbonVisibilityToggleMailMerge()
    RibbonVisibilityToggleKey RIBBON_KEY_MAIL_MERGE, "Mail Merge"
End Sub

Public Sub RibbonVisibilityToggleAdvancedExtras()
    RibbonVisibilityToggleKey RIBBON_KEY_ADVANCED_EXTRAS, "Advanced extras"
End Sub

Public Sub RibbonVisibilityShowAll()
    SaveSetting RIBBON_VISIBILITY_APP, RIBBON_VISIBILITY_SECTION, RIBBON_KEY_STYLES, "True"
    SaveSetting RIBBON_VISIBILITY_APP, RIBBON_VISIBILITY_SECTION, RIBBON_KEY_TABLES, "True"
    SaveSetting RIBBON_VISIBILITY_APP, RIBBON_VISIBILITY_SECTION, RIBBON_KEY_MAIL_MERGE, "True"
    SaveSetting RIBBON_VISIBILITY_APP, RIBBON_VISIBILITY_SECTION, RIBBON_KEY_ADVANCED_EXTRAS, "True"
    InstrumentaResetRibbonVisibility
    MsgBox "All major ribbon groups are visible again.", vbInformation, "Instrumenta Ribbon"
End Sub

Private Sub RibbonVisibilityToggleKey(ByVal keyName As String, ByVal labelText As String)
    Dim currentValue As Boolean
    currentValue = RibbonVisibilityGet(keyName)
    SaveSetting RIBBON_VISIBILITY_APP, RIBBON_VISIBILITY_SECTION, keyName, CStr(Not currentValue)
    InstrumentaResetRibbonVisibility

    If currentValue Then
        MsgBox labelText & " hidden.", vbInformation, "Instrumenta Ribbon"
    Else
        MsgBox labelText & " shown.", vbInformation, "Instrumenta Ribbon"
    End If
End Sub

Private Function RibbonVisibilitySummaryText() As String
    RibbonVisibilitySummaryText = _
        "Ribbon visibility:" & vbCrLf & vbCrLf & _
        "Styles: " & RibbonVisibilityStateText(RIBBON_KEY_STYLES) & vbCrLf & _
        "Tables: " & RibbonVisibilityStateText(RIBBON_KEY_TABLES) & vbCrLf & _
        "Mail Merge: " & RibbonVisibilityStateText(RIBBON_KEY_MAIL_MERGE) & vbCrLf & _
        "Advanced extras: " & RibbonVisibilityStateText(RIBBON_KEY_ADVANCED_EXTRAS)
End Function

Private Function RibbonVisibilityStateText(ByVal keyName As String) As String
    If RibbonVisibilityGet(keyName) Then
        RibbonVisibilityStateText = "Shown"
    Else
        RibbonVisibilityStateText = "Hidden"
    End If
End Function

Private Function RibbonVisibilityGet(ByVal keyName As String) As Boolean
    RibbonVisibilityGet = (GetSetting(RIBBON_VISIBILITY_APP, RIBBON_VISIBILITY_SECTION, keyName, "True") = "True")
End Function

Private Function RibbonVisibilityIsEnabledForControl(ByVal controlId As String) As Boolean
    Select Case controlId
        Case "StylesGroup", "TabViewStylesGroup"
            RibbonVisibilityIsEnabledForControl = RibbonVisibilityGet(RIBBON_KEY_STYLES)
        Case "TableGroup", "TabViewTableGroup"
            RibbonVisibilityIsEnabledForControl = RibbonVisibilityGet(RIBBON_KEY_TABLES)
        Case "MailMergeMenu", "TabViewMailMergeMenu"
            RibbonVisibilityIsEnabledForControl = RibbonVisibilityGet(RIBBON_KEY_MAIL_MERGE)
        Case "ReplaceColorSplitButton", "PyramidBuilder", "SelectionPane", "TagsSplitButton", "ShowSlideGraderButton", "ShowFeatureSearchButton", "ShowAboutDialogButton", "TabViewISCRGroup", "TabViewAllGroup"
            RibbonVisibilityIsEnabledForControl = RibbonVisibilityGet(RIBBON_KEY_ADVANCED_EXTRAS)
        Case Else
            RibbonVisibilityIsEnabledForControl = True
    End Select
End Function

Private Function InstrumentaBaseVisibleForControl(control As IRibbonControl) As Boolean
    If InstrumentaVisible = "InstrumentaVisible" Then
        InstrumentaBaseVisibleForControl = True
    ElseIf control.Tag Like InstrumentaVisible Then
        InstrumentaBaseVisibleForControl = True
    End If
End Function
