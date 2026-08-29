Attribute VB_Name = "ModuleFASAlignment"
'MIT License
'
'Copyright (c) 2021 - 2026 iappyx
'
'FAS alignment helpers extended for configurable named zones, separate place
'and size actions, and preset-driven layouts.

Option Explicit

Private Const FAS_APP As String = "InstrumentaFASAlignment"
Private Const FAS_SECTION_GENERAL As String = "General"
Private Const FAS_SECTION_GUTTERS As String = "Gutters"
Private Const FAS_SECTION_ZONES As String = "Zones"
Private Const FAS_SECTION_PRESETS As String = "Presets"

Private Const FAS_KEY_ACTIVE_ZONE As String = "ActiveZone"
Private Const FAS_KEY_PROFILE_VERSION As String = "ProfileVersion"
Private Const FAS_KEY_HORIZONTAL_GUTTER As String = "HorizontalCm"
Private Const FAS_KEY_VERTICAL_GUTTER As String = "VerticalCm"
Private Const FAS_KEY_PRESET_NAMES As String = "PresetNames"

Private Const FAS_ZONE_CONTENT As String = "Content"
Private Const FAS_ZONE_HEADING As String = "Title"
Private Const FAS_ZONE_SUBHEADING As String = "Subtitle"

Private Const FAS_EDGE_LEFT As String = "LeftCm"
Private Const FAS_EDGE_RIGHT As String = "RightCm"
Private Const FAS_EDGE_TOP As String = "TopCm"
Private Const FAS_EDGE_BOTTOM As String = "BottomCm"

Private Const FAS_DEFAULT_PROFILE_VERSION As String = "FAS-FM-wide-2026-07-09"
Private Const FAS_DEFAULT_LEFT_CM As Double = 0.83
Private Const FAS_DEFAULT_WIDTH_CM As Double = 32.18
Private Const FAS_DEFAULT_RIGHT_CM As Double = 33.01
Private Const FAS_DEFAULT_TITLE_TOP_CM As Double = 1.3
Private Const FAS_DEFAULT_TITLE_HEIGHT_CM As Double = 1#
Private Const FAS_DEFAULT_SUBTITLE_TOP_CM As Double = 2.3
Private Const FAS_DEFAULT_SUBTITLE_HEIGHT_CM As Double = 1#
Private Const FAS_DEFAULT_CONTENT_TOP_CM As Double = 4.3
Private Const FAS_DEFAULT_CONTENT_HEIGHT_CM As Double = 12.2
Private Const FAS_DEFAULT_HORIZONTAL_GUTTER_CM As Double = 0.83
Private Const FAS_DEFAULT_VERTICAL_GUTTER_CM As Double = 0.83

Private Type FASRect
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type FASPreset
    ZoneName As String
    HorizontalPlacement As String
    VerticalPlacement As String
    WidthRule As String
    HeightRule As String
End Type

Public Sub FASPlaceLeft()
    FAS_ApplyPlacementToActiveZone "left", vbNullString
End Sub

Public Sub FASPlaceCenterX()
    FAS_ApplyPlacementToActiveZone "center", vbNullString
End Sub

Public Sub FASPlaceRight()
    FAS_ApplyPlacementToActiveZone "right", vbNullString
End Sub

Public Sub FASPlaceTop()
    FAS_ApplyPlacementToActiveZone vbNullString, "top"
End Sub

Public Sub FASPlaceCenterY()
    FAS_ApplyPlacementToActiveZone vbNullString, "middle"
End Sub

Public Sub FASPlaceBottom()
    FAS_ApplyPlacementToActiveZone vbNullString, "bottom"
End Sub

Public Sub FASPlaceTopLeft()
    FAS_ApplyPlacementToActiveZone "left", "top"
End Sub

Public Sub FASPlaceTopRight()
    FAS_ApplyPlacementToActiveZone "right", "top"
End Sub

Public Sub FASPlaceBottomLeft()
    FAS_ApplyPlacementToActiveZone "left", "bottom"
End Sub

Public Sub FASPlaceBottomRight()
    FAS_ApplyPlacementToActiveZone "right", "bottom"
End Sub

Public Sub FASPlaceColumnHalf1()
    FAS_ApplyPlacementToActiveZone "col-2-1", vbNullString
End Sub

Public Sub FASPlaceColumnHalf2()
    FAS_ApplyPlacementToActiveZone "col-2-2", vbNullString
End Sub

Public Sub FASPlaceColumnThird1()
    FAS_ApplyPlacementToActiveZone "col-3-1", vbNullString
End Sub

Public Sub FASPlaceColumnThird2()
    FAS_ApplyPlacementToActiveZone "col-3-2", vbNullString
End Sub

Public Sub FASPlaceColumnThird3()
    FAS_ApplyPlacementToActiveZone "col-3-3", vbNullString
End Sub

Public Sub FASPlaceColumnQuarter1()
    FAS_ApplyPlacementToActiveZone "col-4-1", vbNullString
End Sub

Public Sub FASPlaceColumnQuarter2()
    FAS_ApplyPlacementToActiveZone "col-4-2", vbNullString
End Sub

Public Sub FASPlaceColumnQuarter3()
    FAS_ApplyPlacementToActiveZone "col-4-3", vbNullString
End Sub

Public Sub FASPlaceColumnQuarter4()
    FAS_ApplyPlacementToActiveZone "col-4-4", vbNullString
End Sub

Public Sub FASPlaceRowHalf1()
    FAS_ApplyPlacementToActiveZone vbNullString, "row-2-1"
End Sub

Public Sub FASPlaceRowHalf2()
    FAS_ApplyPlacementToActiveZone vbNullString, "row-2-2"
End Sub

Public Sub FASPlaceRowThird1()
    FAS_ApplyPlacementToActiveZone vbNullString, "row-3-1"
End Sub

Public Sub FASPlaceRowThird2()
    FAS_ApplyPlacementToActiveZone vbNullString, "row-3-2"
End Sub

Public Sub FASPlaceRowThird3()
    FAS_ApplyPlacementToActiveZone vbNullString, "row-3-3"
End Sub

Public Sub FASSizeFitFullZone()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "left", "top", "full", "full"
End Sub

Public Sub FASSizeColumnHalf1()
    FAS_ApplySizeToActiveZone "col-2", vbNullString
End Sub

Public Sub FASSizeColumnHalf2()
    FAS_ApplySizeToActiveZone "col-2", vbNullString
End Sub

Public Sub FASSizeColumnThird1()
    FAS_ApplySizeToActiveZone "col-3", vbNullString
End Sub

Public Sub FASSizeColumnThird2()
    FAS_ApplySizeToActiveZone "col-3", vbNullString
End Sub

Public Sub FASSizeColumnThird3()
    FAS_ApplySizeToActiveZone "col-3", vbNullString
End Sub

Public Sub FASSizeColumnQuarter1()
    FAS_ApplySizeToActiveZone "col-4", vbNullString
End Sub

Public Sub FASSizeColumnQuarter2()
    FAS_ApplySizeToActiveZone "col-4", vbNullString
End Sub

Public Sub FASSizeColumnQuarter3()
    FAS_ApplySizeToActiveZone "col-4", vbNullString
End Sub

Public Sub FASSizeColumnQuarter4()
    FAS_ApplySizeToActiveZone "col-4", vbNullString
End Sub

Public Sub FASSizeRowHalf1()
    FAS_ApplySizeToActiveZone vbNullString, "row-2"
End Sub

Public Sub FASSizeRowHalf2()
    FAS_ApplySizeToActiveZone vbNullString, "row-2"
End Sub

Public Sub FASSizeRowThird1()
    FAS_ApplySizeToActiveZone vbNullString, "row-3"
End Sub

Public Sub FASSizeRowThird2()
    FAS_ApplySizeToActiveZone vbNullString, "row-3"
End Sub

Public Sub FASSizeRowThird3()
    FAS_ApplySizeToActiveZone vbNullString, "row-3"
End Sub

Public Sub FASPresetFullZone()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "left", "top", "full", "full"
End Sub

Public Sub FASPresetLeftHalf()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "col-2-1", "top", "col-2", "full"
End Sub

Public Sub FASPresetRightHalf()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "col-2-2", "top", "col-2", "full"
End Sub

Public Sub FASPresetLeftThird()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "col-3-1", "top", "col-3", "full"
End Sub

Public Sub FASPresetCenterThird()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "col-3-2", "top", "col-3", "full"
End Sub

Public Sub FASPresetRightThird()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "col-3-3", "top", "col-3", "full"
End Sub

Public Sub FASPresetBottomHalf()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "left", "row-2-2", "full", "row-2"
End Sub

Public Sub FASPresetBottomRightThird()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "col-3-3", "row-3-3", "col-3", "row-3"
End Sub

Public Sub FASPresetBottomRightSmallBlock()
    FAS_ApplyPresetToZone FAS_GetActiveZoneName(), "col-3-3", "row-2-2", "col-3", "row-2"
End Sub

Public Sub FASPresetRunSaved()
    Dim presetName As String
    Dim preset As FASPreset

    presetName = FAS_PromptForPresetName("Run saved preset", True)
    If Len(presetName) = 0 Then Exit Sub
    If Not FAS_TryLoadPreset(presetName, preset) Then
        MsgBox "Preset '" & presetName & "' was not found.", vbExclamation, "FAS Preset"
        Exit Sub
    End If

    FAS_ApplyPresetToZone preset.ZoneName, preset.HorizontalPlacement, preset.VerticalPlacement, preset.WidthRule, preset.HeightRule
End Sub

Public Sub FASPresetSaveCustom()
    Dim presetName As String
    Dim preset As FASPreset
    Dim zoneName As String

    presetName = Trim$(InputBox("Preset name", "Save FAS Preset"))
    If Len(presetName) = 0 Then Exit Sub

    zoneName = FAS_PromptForZoneName("Target zone", FAS_GetActiveZoneName())
    If Len(zoneName) = 0 Then Exit Sub

    preset.ZoneName = zoneName
    preset.HorizontalPlacement = FAS_PromptForHorizontalPlacement("Horizontal placement", "left")
    If Len(preset.HorizontalPlacement) = 0 Then Exit Sub

    preset.VerticalPlacement = FAS_PromptForVerticalPlacement("Vertical placement", "top")
    If Len(preset.VerticalPlacement) = 0 Then Exit Sub

    preset.WidthRule = FAS_PromptForWidthRule("Width rule", "keep")
    If Len(preset.WidthRule) = 0 Then Exit Sub

    preset.HeightRule = FAS_PromptForHeightRule("Height rule", "keep")
    If Len(preset.HeightRule) = 0 Then Exit Sub

    FAS_SavePreset presetName, preset
    MsgBox "Preset '" & presetName & "' saved.", vbInformation, "FAS Preset"
End Sub

Public Sub FASPresetRenameSaved()
    Dim presetName As String
    Dim newName As String
    Dim preset As FASPreset

    presetName = FAS_PromptForPresetName("Rename saved preset", True)
    If Len(presetName) = 0 Then Exit Sub
    If Not FAS_TryLoadPreset(presetName, preset) Then
        MsgBox "Preset '" & presetName & "' was not found.", vbExclamation, "FAS Preset"
        Exit Sub
    End If

    newName = Trim$(InputBox("New name for preset '" & presetName & "'", "Rename FAS Preset", presetName))
    If Len(newName) = 0 Or StrComp(newName, presetName, vbTextCompare) = 0 Then Exit Sub

    FAS_SavePreset newName, preset
    FAS_DeletePreset presetName
    MsgBox "Preset renamed to '" & newName & "'.", vbInformation, "FAS Preset"
End Sub

Public Sub FASPresetDeleteSaved()
    Dim presetName As String

    presetName = FAS_PromptForPresetName("Delete saved preset", True)
    If Len(presetName) = 0 Then Exit Sub

    If MsgBox("Delete preset '" & presetName & "'?", vbQuestion + vbYesNo, "Delete FAS Preset") <> vbYes Then Exit Sub

    FAS_DeletePreset presetName
    MsgBox "Preset '" & presetName & "' deleted.", vbInformation, "FAS Preset"
End Sub

Public Sub FASPresetShowSaved()
    Dim namesText As String
    namesText = FAS_GetPresetNamesText()

    If Len(namesText) = 0 Then
        MsgBox "No saved presets yet.", vbInformation, "FAS Presets"
        Exit Sub
    End If

    MsgBox "Saved presets:" & vbCrLf & vbCrLf & Replace(namesText, Chr(30), vbCrLf), vbInformation, "FAS Presets"
End Sub

Public Sub FASSetupShowSummary()
    MsgBox FAS_BuildSummaryText(), vbInformation, "FAS Setup"
End Sub

Public Sub FASAlignmentShowProfile()
    FASSetupShowSummary
End Sub

Public Sub FASAlignmentSetup()
    Dim choice As String

    choice = Trim$(InputBox( _
        "Choose setup action:" & vbCrLf & _
        "1 = Show summary" & vbCrLf & _
        "2 = Set active zone" & vbCrLf & _
        "3 = Edit Content zone" & vbCrLf & _
        "4 = Edit Title zone" & vbCrLf & _
        "5 = Edit Subtitle zone" & vbCrLf & _
        "6 = Edit gutters" & vbCrLf & _
        "7 = Save preset" & vbCrLf & _
        "8 = Rename preset" & vbCrLf & _
        "9 = Delete preset" & vbCrLf & _
        "0 = Reset defaults", _
        "FAS Setup", "1"))

    Select Case choice
        Case "1"
            FASSetupShowSummary
        Case "2"
            FASSetupSelectActiveZone
        Case "3"
            FASSetupEditContentZone
        Case "4"
            FASSetupEditHeadingZone
        Case "5"
            FASSetupEditSubheadingZone
        Case "6"
            FASSetupEditGutters
        Case "7"
            FASPresetSaveCustom
        Case "8"
            FASPresetRenameSaved
        Case "9"
            FASPresetDeleteSaved
        Case "0"
            FASAlignmentResetDefaults
    End Select
End Sub

Public Sub FASAlignmentResetDefaults()
    FAS_SaveZoneDefaults FAS_ZONE_CONTENT
    FAS_SaveZoneDefaults FAS_ZONE_HEADING
    FAS_SaveZoneDefaults FAS_ZONE_SUBHEADING
    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, FAS_ZONE_CONTENT
    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_PROFILE_VERSION, FAS_DEFAULT_PROFILE_VERSION
    SaveSetting FAS_APP, FAS_SECTION_GUTTERS, FAS_KEY_HORIZONTAL_GUTTER, FAS_FormatInvariant(FAS_DEFAULT_HORIZONTAL_GUTTER_CM)
    SaveSetting FAS_APP, FAS_SECTION_GUTTERS, FAS_KEY_VERTICAL_GUTTER, FAS_FormatInvariant(FAS_DEFAULT_VERTICAL_GUTTER_CM)

    MsgBox "FAS defaults restored for the Forvis Mazars wide layout.", vbInformation, "FAS Setup"
End Sub

Public Sub FASSetupSelectActiveZone()
    Dim zoneName As String
    zoneName = FAS_PromptForZoneName("Default active zone", FAS_GetActiveZoneName())
    If Len(zoneName) = 0 Then Exit Sub

    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, zoneName
    MsgBox "Active zone set to " & zoneName & ".", vbInformation, "FAS Setup"
End Sub

Public Sub FASSetActiveZoneContent()
    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, FAS_ZONE_CONTENT
End Sub

Public Sub FASSetActiveZoneHeading()
    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, FAS_ZONE_HEADING
End Sub

Public Sub FASSetActiveZoneSubheading()
    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, FAS_ZONE_SUBHEADING
End Sub

Public Sub FASSetupEditContentZone()
    FAS_EditZone FAS_ZONE_CONTENT
End Sub

Public Sub FASSetupEditHeadingZone()
    FAS_EditZone FAS_ZONE_HEADING
End Sub

Public Sub FASSetupEditSubheadingZone()
    FAS_EditZone FAS_ZONE_SUBHEADING
End Sub

Public Sub FASSetupEditGutters()
    Dim horizontalCm As Double
    Dim verticalCm As Double

    If Not FAS_PromptCmValue("Horizontal gutter (cm)", FAS_GetGutterCm(True), horizontalCm) Then Exit Sub
    If Not FAS_PromptCmValue("Vertical gutter (cm)", FAS_GetGutterCm(False), verticalCm) Then Exit Sub

    If horizontalCm < 0 Or verticalCm < 0 Then
        MsgBox "Gutters cannot be negative.", vbExclamation, "FAS Setup"
        Exit Sub
    End If

    SaveSetting FAS_APP, FAS_SECTION_GUTTERS, FAS_KEY_HORIZONTAL_GUTTER, FAS_FormatInvariant(horizontalCm)
    SaveSetting FAS_APP, FAS_SECTION_GUTTERS, FAS_KEY_VERTICAL_GUTTER, FAS_FormatInvariant(verticalCm)
    MsgBox "FAS gutters saved.", vbInformation, "FAS Setup"
End Sub

Public Sub FASCanvasSnapLeft()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "left", vbNullString
End Sub

Public Sub FASCanvasSnapRight()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "right", vbNullString
End Sub

Public Sub FASCanvasSnapTop()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, vbNullString, "top"
End Sub

Public Sub FASCanvasSnapBottom()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, vbNullString, "bottom"
End Sub

Public Sub FASCanvasSnapTopLeft()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "left", "top"
End Sub

Public Sub FASCanvasSnapTopRight()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "right", "top"
End Sub

Public Sub FASCanvasSnapBottomLeft()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "left", "bottom"
End Sub

Public Sub FASCanvasSnapBottomRight()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "right", "bottom"
End Sub

Public Sub FASCanvasCenter()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "center", "middle"
End Sub

Public Sub FASFitContentBox()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "left", "top", "full", "full"
End Sub

Public Sub FASPlaceHalf1()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-2-1", vbNullString
End Sub

Public Sub FASPlaceHalf2()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-2-2", vbNullString
End Sub

Public Sub FASPlaceThird1()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-3-1", vbNullString
End Sub

Public Sub FASPlaceThird2()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-3-2", vbNullString
End Sub

Public Sub FASPlaceThird3()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-3-3", vbNullString
End Sub

Public Sub FASPlaceQuarter1()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-4-1", vbNullString
End Sub

Public Sub FASPlaceQuarter2()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-4-2", vbNullString
End Sub

Public Sub FASPlaceQuarter3()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-4-3", vbNullString
End Sub

Public Sub FASPlaceQuarter4()
    FAS_ApplyPlacementToNamedZone FAS_ZONE_CONTENT, "col-4-4", vbNullString
End Sub

Public Sub FASFitHalf1()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-2-1", "top", "col-2", "full"
End Sub

Public Sub FASFitHalf2()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-2-2", "top", "col-2", "full"
End Sub

Public Sub FASFitThird1()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-3-1", "top", "col-3", "full"
End Sub

Public Sub FASFitThird2()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-3-2", "top", "col-3", "full"
End Sub

Public Sub FASFitThird3()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-3-3", "top", "col-3", "full"
End Sub

Public Sub FASFitQuarter1()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-4-1", "top", "col-4", "full"
End Sub

Public Sub FASFitQuarter2()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-4-2", "top", "col-4", "full"
End Sub

Public Sub FASFitQuarter3()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-4-3", "top", "col-4", "full"
End Sub

Public Sub FASFitQuarter4()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "col-4-4", "top", "col-4", "full"
End Sub

Public Sub FASFitRowHalf1()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "left", "row-2-1", "full", "row-2"
End Sub

Public Sub FASFitRowHalf2()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "left", "row-2-2", "full", "row-2"
End Sub

Public Sub FASFitRowThird1()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "left", "row-3-1", "full", "row-3"
End Sub

Public Sub FASFitRowThird2()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "left", "row-3-2", "full", "row-3"
End Sub

Public Sub FASFitRowThird3()
    FAS_ApplyPresetToZone FAS_ZONE_CONTENT, "left", "row-3-3", "full", "row-3"
End Sub

Public Function FASGetActiveZoneName() As String
    FASGetActiveZoneName = FAS_GetActiveZoneName()
End Function

Public Function FASGetActiveZoneBounds(ByRef zoneLeft As Single, ByRef zoneTop As Single, ByRef zoneWidth As Single, ByRef zoneHeight As Single) As Boolean
    Dim zoneRect As FASRect

    If Not FAS_GetZoneRect(FAS_GetActiveZoneName(), zoneRect) Then Exit Function
    zoneLeft = zoneRect.Left
    zoneTop = zoneRect.Top
    zoneWidth = zoneRect.Width
    zoneHeight = zoneRect.Height
    FASGetActiveZoneBounds = True
End Function

Public Sub FASApplyEditCell(ByVal horizontalAnchor As String, ByVal verticalAnchor As String, ByVal fitShape As Boolean, ByVal gridColumns As Long, ByVal gridRows As Long, ByVal exactColumn As Long, ByVal exactRow As Long)
    Dim selectedShapes As ShapeRange
    Dim zoneRect As FASRect
    Dim targetRect As FASRect
    Dim columnRect As FASRect
    Dim rowRect As FASRect
    Dim slideShape As shape
    Dim rangeLeft As Single
    Dim rangeTop As Single
    Dim rangeRight As Single
    Dim rangeBottom As Single
    Dim targetLeft As Single
    Dim targetTop As Single
    Dim deltaX As Single
    Dim deltaY As Single
    Dim scaleFactor As Double
    Dim isFirst As Boolean

    Set selectedShapes = FAS_GetSelectedShapeRange(False)
    If selectedShapes Is Nothing Then
        FASEditModeSetStatus "Select one or more shapes"
        Exit Sub
    End If
    If fitShape And selectedShapes.count <> 1 Then
        FASEditModeSetStatus "Fit requires exactly one shape"
        Exit Sub
    End If
    If Not FAS_GetZoneRect(FAS_GetActiveZoneName(), zoneRect) Then Exit Sub

    targetRect = zoneRect
    If gridColumns >= 2 And gridColumns <= 4 And exactColumn >= 1 And exactColumn <= gridColumns Then
        If Not FAS_GetColumnSlotRect(zoneRect, gridColumns, exactColumn, columnRect) Then Exit Sub
        targetRect.Left = columnRect.Left
        targetRect.Width = columnRect.Width
    End If
    If gridRows >= 2 And gridRows <= 4 And exactRow >= 1 And exactRow <= gridRows Then
        If Not FAS_GetRowSlotRect(zoneRect, gridRows, exactRow, rowRect) Then Exit Sub
        targetRect.Top = rowRect.Top
        targetRect.Height = rowRect.Height
    End If

    If fitShape Then
        Set slideShape = selectedShapes(1)
        If slideShape.LockAspectRatio = msoTrue Then
            scaleFactor = targetRect.Width / GetRealWidth(slideShape)
            If GetRealHeight(slideShape) * scaleFactor > targetRect.Height Then
                scaleFactor = targetRect.Height / GetRealHeight(slideShape)
            End If
            SetRealWidth slideShape, CSng(GetRealWidth(slideShape) * scaleFactor)
        Else
            SetRealWidth slideShape, targetRect.Width
            SetRealHeight slideShape, targetRect.Height
        End If

        targetLeft = FAS_ResolveHorizontalLeft(targetRect, horizontalAnchor, GetRealWidth(slideShape))
        targetTop = FAS_ResolveVerticalTop(targetRect, verticalAnchor, GetRealHeight(slideShape))
        SetRealLeft slideShape, targetLeft
        SetRealTop slideShape, targetTop
        FASEditModeSetStatus "Fit to " & gridColumns & " x " & gridRows & " cell"
        Exit Sub
    End If

    isFirst = True
    For Each slideShape In selectedShapes
        If isFirst Then
            rangeLeft = GetRealLeft(slideShape)
            rangeTop = GetRealTop(slideShape)
            rangeRight = rangeLeft + GetRealWidth(slideShape)
            rangeBottom = rangeTop + GetRealHeight(slideShape)
            isFirst = False
        Else
            If GetRealLeft(slideShape) < rangeLeft Then rangeLeft = GetRealLeft(slideShape)
            If GetRealTop(slideShape) < rangeTop Then rangeTop = GetRealTop(slideShape)
            If GetRealLeft(slideShape) + GetRealWidth(slideShape) > rangeRight Then rangeRight = GetRealLeft(slideShape) + GetRealWidth(slideShape)
            If GetRealTop(slideShape) + GetRealHeight(slideShape) > rangeBottom Then rangeBottom = GetRealTop(slideShape) + GetRealHeight(slideShape)
        End If
    Next slideShape

    If rangeRight - rangeLeft > zoneRect.Width Or rangeBottom - rangeTop > zoneRect.Height Then
        FASEditModeSetStatus "Selection is larger than the active zone"
        Exit Sub
    End If

    targetLeft = FAS_ResolveHorizontalLeft(zoneRect, horizontalAnchor, rangeRight - rangeLeft)
    targetTop = FAS_ResolveVerticalTop(zoneRect, verticalAnchor, rangeBottom - rangeTop)
    deltaX = targetLeft - rangeLeft
    deltaY = targetTop - rangeTop

    For Each slideShape In selectedShapes
        SetRealLeft slideShape, GetRealLeft(slideShape) + deltaX
        SetRealTop slideShape, GetRealTop(slideShape) + deltaY
    Next slideShape
    FASEditModeSetStatus "Placed " & horizontalAnchor & " / " & verticalAnchor
End Sub

Private Sub FAS_EditZone(ByVal zoneName As String)
    Dim leftCm As Double
    Dim rightCm As Double
    Dim topCm As Double
    Dim bottomCm As Double

    If Not FAS_PromptCmValue(zoneName & " left edge (cm)", FAS_GetZoneValueCm(zoneName, FAS_EDGE_LEFT), leftCm) Then Exit Sub
    If Not FAS_PromptCmValue(zoneName & " right edge (cm)", FAS_GetZoneValueCm(zoneName, FAS_EDGE_RIGHT), rightCm) Then Exit Sub
    If Not FAS_PromptCmValue(zoneName & " top edge (cm)", FAS_GetZoneValueCm(zoneName, FAS_EDGE_TOP), topCm) Then Exit Sub
    If Not FAS_PromptCmValue(zoneName & " bottom edge (cm)", FAS_GetZoneValueCm(zoneName, FAS_EDGE_BOTTOM), bottomCm) Then Exit Sub

    If rightCm <= leftCm Then
        MsgBox zoneName & " right edge must be greater than left edge.", vbExclamation, "FAS Setup"
        Exit Sub
    End If

    If bottomCm <= topCm Then
        MsgBox zoneName & " bottom edge must be greater than top edge.", vbExclamation, "FAS Setup"
        Exit Sub
    End If

    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_LEFT), FAS_FormatInvariant(leftCm)
    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_RIGHT), FAS_FormatInvariant(rightCm)
    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_TOP), FAS_FormatInvariant(topCm)
    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_BOTTOM), FAS_FormatInvariant(bottomCm)
    MsgBox zoneName & " zone saved.", vbInformation, "FAS Setup"
End Sub

Private Sub FAS_ApplyPlacementToActiveZone(ByVal horizontalRule As String, ByVal verticalRule As String)
    FAS_ApplyPlacementToNamedZone FAS_GetActiveZoneName(), horizontalRule, verticalRule
End Sub

Private Sub FAS_ApplyPlacementToNamedZone(ByVal zoneName As String, ByVal horizontalRule As String, ByVal verticalRule As String)
    Dim selectedShapes As ShapeRange
    Dim zoneRect As FASRect
    Dim slideShape As shape

    Set selectedShapes = FAS_GetSelectedShapeRange()
    If selectedShapes Is Nothing Then Exit Sub
    If Not FAS_GetZoneRect(zoneName, zoneRect) Then Exit Sub

    For Each slideShape In selectedShapes
        If Len(horizontalRule) > 0 Then
            SetRealLeft slideShape, FAS_ResolveHorizontalLeft(zoneRect, horizontalRule, GetRealWidth(slideShape))
        End If

        If Len(verticalRule) > 0 Then
            SetRealTop slideShape, FAS_ResolveVerticalTop(zoneRect, verticalRule, GetRealHeight(slideShape))
        End If
    Next slideShape
End Sub

Private Sub FAS_ApplySizeToActiveZone(ByVal widthRule As String, ByVal heightRule As String)
    Dim selectedShapes As ShapeRange
    Dim zoneRect As FASRect
    Dim slideShape As shape

    Set selectedShapes = FAS_GetSelectedShapeRange()
    If selectedShapes Is Nothing Then Exit Sub
    If Not FAS_GetZoneRect(FAS_GetActiveZoneName(), zoneRect) Then Exit Sub

    For Each slideShape In selectedShapes
        FAS_ApplySizeRules slideShape, zoneRect, widthRule, heightRule
    Next slideShape
End Sub

Private Sub FAS_ApplyPresetToZone(ByVal zoneName As String, ByVal horizontalRule As String, ByVal verticalRule As String, ByVal widthRule As String, ByVal heightRule As String)
    Dim selectedShapes As ShapeRange
    Dim zoneRect As FASRect
    Dim slideShape As shape

    Set selectedShapes = FAS_GetSelectedShapeRange()
    If selectedShapes Is Nothing Then Exit Sub
    If Not FAS_GetZoneRect(zoneName, zoneRect) Then Exit Sub

    For Each slideShape In selectedShapes
        FAS_ApplySizeRules slideShape, zoneRect, widthRule, heightRule

        If Len(horizontalRule) > 0 Then
            SetRealLeft slideShape, FAS_ResolveHorizontalLeft(zoneRect, horizontalRule, GetRealWidth(slideShape))
        End If

        If Len(verticalRule) > 0 Then
            SetRealTop slideShape, FAS_ResolveVerticalTop(zoneRect, verticalRule, GetRealHeight(slideShape))
        End If
    Next slideShape
End Sub

Private Sub FAS_ApplySizeRules(ByVal slideShape As shape, ByRef zoneRect As FASRect, ByVal widthRule As String, ByVal heightRule As String)
    Dim targetWidth As Single
    Dim targetHeight As Single

    If Len(widthRule) > 0 And LCase$(widthRule) <> "keep" Then
        targetWidth = FAS_ResolveWidth(zoneRect, widthRule)
        If targetWidth > 0 Then SetRealWidth slideShape, targetWidth
    End If

    If Len(heightRule) > 0 And LCase$(heightRule) <> "keep" Then
        targetHeight = FAS_ResolveHeight(zoneRect, heightRule)
        If targetHeight > 0 Then SetRealHeight slideShape, targetHeight
    End If
End Sub

Private Function FAS_GetSelectedShapeRange(Optional ByVal showMessage As Boolean = True) As ShapeRange
    On Error GoTo NoSelection

    If Application.ActiveWindow.Selection.Type = ppSelectionShapes Then
        If Application.ActiveWindow.Selection.HasChildShapeRange Then
            Set FAS_GetSelectedShapeRange = Application.ActiveWindow.Selection.ChildShapeRange
        Else
            Set FAS_GetSelectedShapeRange = Application.ActiveWindow.Selection.ShapeRange
        End If
    ElseIf Application.ActiveWindow.Selection.Type = ppSelectionText Then
        Set FAS_GetSelectedShapeRange = Application.ActiveWindow.Selection.ShapeRange
    End If

    If showMessage And FAS_GetSelectedShapeRange Is Nothing Then
        MsgBox "Select one or more shapes first.", vbInformation, "FAS Alignment"
    End If
    Exit Function

NoSelection:
    If showMessage Then MsgBox "Select one or more shapes first.", vbInformation, "FAS Alignment"
End Function

Private Function FAS_GetZoneRect(ByVal zoneName As String, ByRef zoneRect As FASRect) As Boolean
    Dim leftCm As Double
    Dim rightCm As Double
    Dim topCm As Double
    Dim bottomCm As Double

    FAS_EnsureDefaultProfile

    leftCm = FAS_GetZoneValueCm(zoneName, FAS_EDGE_LEFT)
    rightCm = FAS_GetZoneValueCm(zoneName, FAS_EDGE_RIGHT)
    topCm = FAS_GetZoneValueCm(zoneName, FAS_EDGE_TOP)
    bottomCm = FAS_GetZoneValueCm(zoneName, FAS_EDGE_BOTTOM)

    If rightCm <= leftCm Then
        MsgBox zoneName & " right edge must be greater than left edge.", vbExclamation, "FAS Alignment"
        Exit Function
    End If

    If bottomCm <= topCm Then
        MsgBox zoneName & " bottom edge must be greater than top edge.", vbExclamation, "FAS Alignment"
        Exit Function
    End If

    zoneRect.Left = FAS_CmToPoints(leftCm)
    zoneRect.Top = FAS_CmToPoints(topCm)
    zoneRect.Width = FAS_CmToPoints(rightCm - leftCm)
    zoneRect.Height = FAS_CmToPoints(bottomCm - topCm)
    FAS_GetZoneRect = True
End Function

Private Function FAS_ResolveHorizontalLeft(ByRef zoneRect As FASRect, ByVal rule As String, ByVal shapeWidth As Single) As Single
    Dim slotRect As FASRect
    Dim slotCount As Long
    Dim slotIndex As Long
    Dim normalizedRule As String

    normalizedRule = LCase$(Trim$(rule))

    Select Case normalizedRule
        Case "left"
            FAS_ResolveHorizontalLeft = zoneRect.Left
        Case "center"
            FAS_ResolveHorizontalLeft = zoneRect.Left + (zoneRect.Width - shapeWidth) / 2
        Case "right"
            FAS_ResolveHorizontalLeft = zoneRect.Left + zoneRect.Width - shapeWidth
        Case Else
            If FAS_TryParseSlotPlacement(normalizedRule, "col", slotCount, slotIndex) Then
                If FAS_GetColumnSlotRect(zoneRect, slotCount, slotIndex, slotRect) Then
                    FAS_ResolveHorizontalLeft = slotRect.Left
                Else
                    FAS_ResolveHorizontalLeft = zoneRect.Left
                End If
            Else
                FAS_ResolveHorizontalLeft = zoneRect.Left
            End If
    End Select
End Function

Private Function FAS_ResolveVerticalTop(ByRef zoneRect As FASRect, ByVal rule As String, ByVal shapeHeight As Single) As Single
    Dim slotRect As FASRect
    Dim slotCount As Long
    Dim slotIndex As Long
    Dim normalizedRule As String

    normalizedRule = LCase$(Trim$(rule))

    Select Case normalizedRule
        Case "top"
            FAS_ResolveVerticalTop = zoneRect.Top
        Case "middle"
            FAS_ResolveVerticalTop = zoneRect.Top + (zoneRect.Height - shapeHeight) / 2
        Case "bottom"
            FAS_ResolveVerticalTop = zoneRect.Top + zoneRect.Height - shapeHeight
        Case Else
            If FAS_TryParseSlotPlacement(normalizedRule, "row", slotCount, slotIndex) Then
                If FAS_GetRowSlotRect(zoneRect, slotCount, slotIndex, slotRect) Then
                    FAS_ResolveVerticalTop = slotRect.Top
                Else
                    FAS_ResolveVerticalTop = zoneRect.Top
                End If
            Else
                FAS_ResolveVerticalTop = zoneRect.Top
            End If
    End Select
End Function

Private Function FAS_ResolveWidth(ByRef zoneRect As FASRect, ByVal rule As String) As Single
    Dim normalizedRule As String
    normalizedRule = LCase$(Trim$(rule))

    Select Case normalizedRule
        Case "full"
            FAS_ResolveWidth = zoneRect.Width
        Case "col-2"
            FAS_ResolveWidth = FAS_ColumnWidth(zoneRect, 2)
        Case "col-3"
            FAS_ResolveWidth = FAS_ColumnWidth(zoneRect, 3)
        Case "col-4"
            FAS_ResolveWidth = FAS_ColumnWidth(zoneRect, 4)
    End Select
End Function

Private Function FAS_ResolveHeight(ByRef zoneRect As FASRect, ByVal rule As String) As Single
    Dim normalizedRule As String
    normalizedRule = LCase$(Trim$(rule))

    Select Case normalizedRule
        Case "full"
            FAS_ResolveHeight = zoneRect.Height
        Case "row-2"
            FAS_ResolveHeight = FAS_RowHeight(zoneRect, 2)
        Case "row-3"
            FAS_ResolveHeight = FAS_RowHeight(zoneRect, 3)
    End Select
End Function

Private Function FAS_GetColumnSlotRect(ByRef zoneRect As FASRect, ByVal slotCount As Long, ByVal slotIndex As Long, ByRef slotRect As FASRect) As Boolean
    Dim gutter As Single

    If slotCount < 1 Or slotIndex < 1 Or slotIndex > slotCount Then Exit Function

    gutter = FAS_CmToPoints(FAS_GetGutterCm(True))
    slotRect.Width = (zoneRect.Width - gutter * (slotCount - 1)) / slotCount
    If slotRect.Width <= 0 Then
        MsgBox "Horizontal gutter is too large for this zone width.", vbExclamation, "FAS Alignment"
        Exit Function
    End If

    slotRect.Left = zoneRect.Left + (slotIndex - 1) * (slotRect.Width + gutter)
    slotRect.Top = zoneRect.Top
    slotRect.Height = zoneRect.Height
    FAS_GetColumnSlotRect = True
End Function

Private Function FAS_GetRowSlotRect(ByRef zoneRect As FASRect, ByVal slotCount As Long, ByVal slotIndex As Long, ByRef slotRect As FASRect) As Boolean
    Dim gutter As Single

    If slotCount < 1 Or slotIndex < 1 Or slotIndex > slotCount Then Exit Function

    gutter = FAS_CmToPoints(FAS_GetGutterCm(False))
    slotRect.Height = (zoneRect.Height - gutter * (slotCount - 1)) / slotCount
    If slotRect.Height <= 0 Then
        MsgBox "Vertical gutter is too large for this zone height.", vbExclamation, "FAS Alignment"
        Exit Function
    End If

    slotRect.Left = zoneRect.Left
    slotRect.Top = zoneRect.Top + (slotIndex - 1) * (slotRect.Height + gutter)
    slotRect.Width = zoneRect.Width
    FAS_GetRowSlotRect = True
End Function

Private Function FAS_ColumnWidth(ByRef zoneRect As FASRect, ByVal slotCount As Long) As Single
    Dim slotRect As FASRect
    If FAS_GetColumnSlotRect(zoneRect, slotCount, 1, slotRect) Then FAS_ColumnWidth = slotRect.Width
End Function

Private Function FAS_RowHeight(ByRef zoneRect As FASRect, ByVal slotCount As Long) As Single
    Dim slotRect As FASRect
    If FAS_GetRowSlotRect(zoneRect, slotCount, 1, slotRect) Then FAS_RowHeight = slotRect.Height
End Function

Private Function FAS_TryParseSlotPlacement(ByVal rule As String, ByVal expectedPrefix As String, ByRef slotCount As Long, ByRef slotIndex As Long) As Boolean
    Dim parts() As String

    parts = Split(rule, "-")
    If UBound(parts) <> 2 Then Exit Function
    If parts(0) <> expectedPrefix Then Exit Function
    If Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then Exit Function

    slotCount = CLng(parts(1))
    slotIndex = CLng(parts(2))
    FAS_TryParseSlotPlacement = True
End Function

Private Function FAS_GetActiveZoneName() As String
    Dim rawValue As String
    FAS_EnsureDefaultProfile
    rawValue = GetSetting(FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, FAS_ZONE_CONTENT)
    FAS_GetActiveZoneName = FAS_NormalizeZoneName(rawValue)
End Function

Private Function FAS_NormalizeZoneName(ByVal rawValue As String) As String
    Select Case LCase$(Trim$(rawValue))
        Case "heading", "title"
            FAS_NormalizeZoneName = FAS_ZONE_HEADING
        Case "subheading", "subtitle"
            FAS_NormalizeZoneName = FAS_ZONE_SUBHEADING
        Case Else
            FAS_NormalizeZoneName = FAS_ZONE_CONTENT
    End Select
End Function

Private Function FAS_PromptForZoneName(ByVal promptTitle As String, ByVal defaultZone As String) As String
    Dim userInput As String
    userInput = Trim$(InputBox(promptTitle & vbCrLf & vbCrLf & "Options: Content, Title, Subtitle", "FAS Setup", defaultZone))
    If Len(userInput) = 0 Then Exit Function
    FAS_PromptForZoneName = FAS_NormalizeZoneName(userInput)
End Function

Private Function FAS_GetZoneValueCm(ByVal zoneName As String, ByVal edgeKey As String) As Double
    Dim storageKey As String
    Dim storedValue As String
    Dim parsedValue As Double

    FAS_EnsureDefaultProfile

    storageKey = FAS_ZoneSettingKey(zoneName, edgeKey)
    storedValue = GetSetting(FAS_APP, FAS_SECTION_ZONES, storageKey, vbNullString)

    If Len(storedValue) > 0 Then
        If FAS_TryParseCm(storedValue, parsedValue) Then
            FAS_GetZoneValueCm = parsedValue
            Exit Function
        End If
    End If

    FAS_GetZoneValueCm = FAS_DefaultZoneValueCm(zoneName, edgeKey)
End Function

Private Sub FAS_EnsureDefaultProfile()
    If GetSetting(FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_PROFILE_VERSION, vbNullString) = FAS_DEFAULT_PROFILE_VERSION Then
        Exit Sub
    End If

    FAS_SaveZoneDefaults FAS_ZONE_CONTENT
    FAS_SaveZoneDefaults FAS_ZONE_HEADING
    FAS_SaveZoneDefaults FAS_ZONE_SUBHEADING
    SaveSetting FAS_APP, FAS_SECTION_GUTTERS, FAS_KEY_HORIZONTAL_GUTTER, FAS_FormatInvariant(FAS_DEFAULT_HORIZONTAL_GUTTER_CM)
    SaveSetting FAS_APP, FAS_SECTION_GUTTERS, FAS_KEY_VERTICAL_GUTTER, FAS_FormatInvariant(FAS_DEFAULT_VERTICAL_GUTTER_CM)
    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, FAS_NormalizeZoneName(GetSetting(FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_ACTIVE_ZONE, FAS_ZONE_CONTENT))
    SaveSetting FAS_APP, FAS_SECTION_GENERAL, FAS_KEY_PROFILE_VERSION, FAS_DEFAULT_PROFILE_VERSION
End Sub

Private Function FAS_DefaultZoneValueCm(ByVal zoneName As String, ByVal edgeKey As String) As Double
    Dim normalizedZoneName As String
    normalizedZoneName = FAS_NormalizeZoneName(zoneName)

    Select Case normalizedZoneName
        Case FAS_ZONE_HEADING
            Select Case edgeKey
                Case FAS_EDGE_LEFT
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_LEFT_CM
                Case FAS_EDGE_RIGHT
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_RIGHT_CM
                Case FAS_EDGE_TOP
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_TITLE_TOP_CM
                Case FAS_EDGE_BOTTOM
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_TITLE_TOP_CM + FAS_DEFAULT_TITLE_HEIGHT_CM
            End Select
        Case FAS_ZONE_SUBHEADING
            Select Case edgeKey
                Case FAS_EDGE_LEFT
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_LEFT_CM
                Case FAS_EDGE_RIGHT
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_RIGHT_CM
                Case FAS_EDGE_TOP
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_SUBTITLE_TOP_CM
                Case FAS_EDGE_BOTTOM
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_SUBTITLE_TOP_CM + FAS_DEFAULT_SUBTITLE_HEIGHT_CM
            End Select
        Case Else
            Select Case edgeKey
                Case FAS_EDGE_LEFT
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_LEFT_CM
                Case FAS_EDGE_RIGHT
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_RIGHT_CM
                Case FAS_EDGE_TOP
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_CONTENT_TOP_CM
                Case FAS_EDGE_BOTTOM
                    FAS_DefaultZoneValueCm = FAS_DEFAULT_CONTENT_TOP_CM + FAS_DEFAULT_CONTENT_HEIGHT_CM
            End Select
    End Select
End Function

Private Function FAS_GetGutterCm(ByVal horizontalGutter As Boolean) As Double
    Dim keyName As String
    Dim defaultValue As Double
    Dim storedValue As String
    Dim parsedValue As Double

    FAS_EnsureDefaultProfile

    If horizontalGutter Then
        keyName = FAS_KEY_HORIZONTAL_GUTTER
        defaultValue = FAS_DEFAULT_HORIZONTAL_GUTTER_CM
    Else
        keyName = FAS_KEY_VERTICAL_GUTTER
        defaultValue = FAS_DEFAULT_VERTICAL_GUTTER_CM
    End If

    storedValue = GetSetting(FAS_APP, FAS_SECTION_GUTTERS, keyName, vbNullString)
    If Len(storedValue) > 0 Then
        If FAS_TryParseCm(storedValue, parsedValue) Then
            FAS_GetGutterCm = parsedValue
            Exit Function
        End If
    End If

    FAS_GetGutterCm = defaultValue
End Function

Private Sub FAS_SaveZoneDefaults(ByVal zoneName As String)
    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_LEFT), FAS_FormatInvariant(FAS_DefaultZoneValueCm(zoneName, FAS_EDGE_LEFT))
    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_RIGHT), FAS_FormatInvariant(FAS_DefaultZoneValueCm(zoneName, FAS_EDGE_RIGHT))
    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_TOP), FAS_FormatInvariant(FAS_DefaultZoneValueCm(zoneName, FAS_EDGE_TOP))
    SaveSetting FAS_APP, FAS_SECTION_ZONES, FAS_ZoneSettingKey(zoneName, FAS_EDGE_BOTTOM), FAS_FormatInvariant(FAS_DefaultZoneValueCm(zoneName, FAS_EDGE_BOTTOM))
End Sub

Private Function FAS_ZoneSettingKey(ByVal zoneName As String, ByVal edgeKey As String) As String
    FAS_ZoneSettingKey = zoneName & edgeKey
End Function

Private Function FAS_BuildSummaryText() As String
    FAS_BuildSummaryText = _
        "Active zone: " & FAS_GetActiveZoneName() & vbCrLf & vbCrLf & _
        FAS_BuildZoneSummary(FAS_ZONE_CONTENT) & vbCrLf & vbCrLf & _
        FAS_BuildZoneSummary(FAS_ZONE_HEADING) & vbCrLf & vbCrLf & _
        FAS_BuildZoneSummary(FAS_ZONE_SUBHEADING) & vbCrLf & vbCrLf & _
        FAS_BuildSlotSummary(FAS_ZONE_CONTENT) & vbCrLf & vbCrLf & _
        "Horizontal gutter: " & FAS_FormatCm(FAS_GetGutterCm(True)) & " cm" & vbCrLf & _
        "Vertical gutter: " & FAS_FormatCm(FAS_GetGutterCm(False)) & " cm" & vbCrLf & vbCrLf & _
        FAS_BuildPresetSummary()
End Function

Private Function FAS_BuildZoneSummary(ByVal zoneName As String) As String
    FAS_BuildZoneSummary = zoneName & ":" & vbCrLf & _
        "  Left " & FAS_FormatCm(FAS_GetZoneValueCm(zoneName, FAS_EDGE_LEFT)) & " cm" & vbCrLf & _
        "  Right " & FAS_FormatCm(FAS_GetZoneValueCm(zoneName, FAS_EDGE_RIGHT)) & " cm" & vbCrLf & _
        "  Top " & FAS_FormatCm(FAS_GetZoneValueCm(zoneName, FAS_EDGE_TOP)) & " cm" & vbCrLf & _
        "  Bottom " & FAS_FormatCm(FAS_GetZoneValueCm(zoneName, FAS_EDGE_BOTTOM)) & " cm"
End Function

Private Function FAS_BuildSlotSummary(ByVal zoneName As String) As String
    Dim zoneWidth As Double
    Dim zoneHeight As Double
    Dim horizontalGutter As Double
    Dim verticalGutter As Double
    Dim halfWidth As Double
    Dim thirdWidth As Double
    Dim quarterWidth As Double
    Dim halfHeight As Double

    zoneWidth = FAS_GetZoneValueCm(zoneName, FAS_EDGE_RIGHT) - FAS_GetZoneValueCm(zoneName, FAS_EDGE_LEFT)
    zoneHeight = FAS_GetZoneValueCm(zoneName, FAS_EDGE_BOTTOM) - FAS_GetZoneValueCm(zoneName, FAS_EDGE_TOP)
    horizontalGutter = FAS_GetGutterCm(True)
    verticalGutter = FAS_GetGutterCm(False)
    halfWidth = (zoneWidth - horizontalGutter) / 2
    thirdWidth = (zoneWidth - horizontalGutter * 2) / 3
    quarterWidth = (zoneWidth - horizontalGutter * 3) / 4
    halfHeight = (zoneHeight - verticalGutter) / 2

    FAS_BuildSlotSummary = zoneName & " calculated slots:" & vbCrLf & _
        "  Full width " & FAS_FormatCm(zoneWidth) & " cm" & vbCrLf & _
        "  1/2 width " & FAS_FormatCm(halfWidth) & " cm" & vbCrLf & _
        "  1/3 width " & FAS_FormatCm(thirdWidth) & " cm" & vbCrLf & _
        "  2/3 span " & FAS_FormatCm(thirdWidth * 2 + horizontalGutter) & " cm" & vbCrLf & _
        "  1/4 width " & FAS_FormatCm(quarterWidth) & " cm" & vbCrLf & _
        "  1/2 height " & FAS_FormatCm(halfHeight) & " cm"
End Function

Private Function FAS_BuildPresetSummary() As String
    Dim namesText As String
    namesText = FAS_GetPresetNamesText()

    If Len(namesText) = 0 Then
        FAS_BuildPresetSummary = "Saved presets: none"
    Else
        FAS_BuildPresetSummary = "Saved presets:" & vbCrLf & Replace(namesText, Chr(30), vbCrLf)
    End If
End Function

Private Function FAS_PromptCmValue(ByVal promptText As String, ByVal currentValue As Double, ByRef parsedValue As Double) As Boolean
    Dim inputValue As String

    inputValue = InputBox(promptText, "FAS Setup", FAS_FormatCm(currentValue))
    If Len(inputValue) = 0 Then Exit Function
    If Not FAS_TryParseCm(inputValue, parsedValue) Then
        MsgBox "Please enter a number in centimeters.", vbExclamation, "FAS Setup"
        Exit Function
    End If

    FAS_PromptCmValue = True
End Function

Private Function FAS_PromptForHorizontalPlacement(ByVal promptTitle As String, ByVal defaultValue As String) As String
    FAS_PromptForHorizontalPlacement = FAS_PromptForChoice(promptTitle, defaultValue, _
        "Options:" & vbCrLf & _
        "left, center, right" & vbCrLf & _
        "col-2-1, col-2-2" & vbCrLf & _
        "col-3-1, col-3-2, col-3-3" & vbCrLf & _
        "col-4-1, col-4-2, col-4-3, col-4-4")
End Function

Private Function FAS_PromptForVerticalPlacement(ByVal promptTitle As String, ByVal defaultValue As String) As String
    FAS_PromptForVerticalPlacement = FAS_PromptForChoice(promptTitle, defaultValue, _
        "Options:" & vbCrLf & _
        "top, middle, bottom" & vbCrLf & _
        "row-2-1, row-2-2" & vbCrLf & _
        "row-3-1, row-3-2, row-3-3")
End Function

Private Function FAS_PromptForWidthRule(ByVal promptTitle As String, ByVal defaultValue As String) As String
    FAS_PromptForWidthRule = FAS_PromptForChoice(promptTitle, defaultValue, _
        "Options:" & vbCrLf & _
        "keep, full, col-2, col-3, col-4")
End Function

Private Function FAS_PromptForHeightRule(ByVal promptTitle As String, ByVal defaultValue As String) As String
    FAS_PromptForHeightRule = FAS_PromptForChoice(promptTitle, defaultValue, _
        "Options:" & vbCrLf & _
        "keep, full, row-2, row-3")
End Function

Private Function FAS_PromptForChoice(ByVal promptTitle As String, ByVal defaultValue As String, ByVal optionsText As String) As String
    FAS_PromptForChoice = LCase$(Trim$(InputBox(promptTitle & vbCrLf & vbCrLf & optionsText, "FAS Setup", defaultValue)))
End Function

Private Function FAS_GetPresetNamesText() As String
    FAS_GetPresetNamesText = GetSetting(FAS_APP, FAS_SECTION_PRESETS, FAS_KEY_PRESET_NAMES, vbNullString)
End Function

Private Function FAS_PromptForPresetName(ByVal promptTitle As String, ByVal showPresetList As Boolean) As String
    Dim promptText As String
    promptText = promptTitle

    If showPresetList Then
        promptText = promptText & vbCrLf & vbCrLf & "Saved presets:" & vbCrLf
        If Len(FAS_GetPresetNamesText()) = 0 Then
            promptText = promptText & "(none)"
        Else
            promptText = promptText & Replace(FAS_GetPresetNamesText(), Chr(30), vbCrLf)
        End If
    End If

    FAS_PromptForPresetName = Trim$(InputBox(promptText, "FAS Presets"))
End Function

Private Sub FAS_SavePreset(ByVal presetName As String, ByRef preset As FASPreset)
    Dim existingName As String
    existingName = FAS_FindPresetName(presetName)

    If Len(existingName) = 0 Then
        FAS_AddPresetName presetName
    Else
        presetName = existingName
    End If

    SaveSetting FAS_APP, FAS_SECTION_PRESETS, presetName, FAS_SerializePreset(preset)
End Sub

Private Sub FAS_DeletePreset(ByVal presetName As String)
    Dim actualName As String
    actualName = FAS_FindPresetName(presetName)
    If Len(actualName) = 0 Then Exit Sub

    DeleteSetting FAS_APP, FAS_SECTION_PRESETS, actualName
    FAS_RemovePresetName actualName
End Sub

Private Function FAS_TryLoadPreset(ByVal presetName As String, ByRef preset As FASPreset) As Boolean
    Dim actualName As String
    Dim rawValue As String

    actualName = FAS_FindPresetName(presetName)
    If Len(actualName) = 0 Then Exit Function

    rawValue = GetSetting(FAS_APP, FAS_SECTION_PRESETS, actualName, vbNullString)
    If Len(rawValue) = 0 Then Exit Function

    FAS_TryLoadPreset = FAS_TryDeserializePreset(rawValue, preset)
End Function

Private Function FAS_SerializePreset(ByRef preset As FASPreset) As String
    FAS_SerializePreset = preset.ZoneName & Chr(31) & preset.HorizontalPlacement & Chr(31) & preset.VerticalPlacement & Chr(31) & preset.WidthRule & Chr(31) & preset.HeightRule
End Function

Private Function FAS_TryDeserializePreset(ByVal rawValue As String, ByRef preset As FASPreset) As Boolean
    Dim parts() As String

    parts = Split(rawValue, Chr(31))
    If UBound(parts) <> 4 Then Exit Function

    preset.ZoneName = FAS_NormalizeZoneName(parts(0))
    preset.HorizontalPlacement = parts(1)
    preset.VerticalPlacement = parts(2)
    preset.WidthRule = parts(3)
    preset.HeightRule = parts(4)
    FAS_TryDeserializePreset = True
End Function

Private Function FAS_FindPresetName(ByVal presetName As String) As String
    Dim namesText As String
    Dim names() As String
    Dim i As Long

    namesText = FAS_GetPresetNamesText()
    If Len(namesText) = 0 Then Exit Function

    names = Split(namesText, Chr(30))
    For i = LBound(names) To UBound(names)
        If StrComp(names(i), presetName, vbTextCompare) = 0 Then
            FAS_FindPresetName = names(i)
            Exit Function
        End If
    Next i
End Function

Private Sub FAS_AddPresetName(ByVal presetName As String)
    Dim namesText As String

    namesText = FAS_GetPresetNamesText()
    If Len(namesText) = 0 Then
        SaveSetting FAS_APP, FAS_SECTION_PRESETS, FAS_KEY_PRESET_NAMES, presetName
    Else
        SaveSetting FAS_APP, FAS_SECTION_PRESETS, FAS_KEY_PRESET_NAMES, namesText & Chr(30) & presetName
    End If
End Sub

Private Sub FAS_RemovePresetName(ByVal presetName As String)
    Dim namesText As String
    Dim names() As String
    Dim rebuiltText As String
    Dim i As Long

    namesText = FAS_GetPresetNamesText()
    If Len(namesText) = 0 Then Exit Sub

    names = Split(namesText, Chr(30))
    For i = LBound(names) To UBound(names)
        If StrComp(names(i), presetName, vbTextCompare) <> 0 Then
            If Len(rebuiltText) = 0 Then
                rebuiltText = names(i)
            Else
                rebuiltText = rebuiltText & Chr(30) & names(i)
            End If
        End If
    Next i

    SaveSetting FAS_APP, FAS_SECTION_PRESETS, FAS_KEY_PRESET_NAMES, rebuiltText
End Sub

Private Function FAS_CmToPoints(ByVal valueCm As Double) As Single
    FAS_CmToPoints = CSng(valueCm * 72 / 2.54)
End Function

Private Function FAS_PointsToCm(ByVal valuePoints As Double) As Double
    FAS_PointsToCm = valuePoints * 2.54 / 72
End Function

Private Function FAS_TryParseCm(ByVal rawValue As String, ByRef parsedValue As Double) As Boolean
    Dim normalizedValue As String

    normalizedValue = Trim$(rawValue)
    normalizedValue = Replace(normalizedValue, "cm", "", 1, -1, vbTextCompare)
    normalizedValue = Replace(normalizedValue, " ", "")
    normalizedValue = Replace(normalizedValue, ".", GetDecimalSeperator())
    normalizedValue = Replace(normalizedValue, ",", GetDecimalSeperator())

    If Len(normalizedValue) = 0 Then Exit Function
    If Not IsNumeric(normalizedValue) Then Exit Function

    parsedValue = CDbl(normalizedValue)
    FAS_TryParseCm = True
End Function

Private Function FAS_FormatCm(ByVal valueCm As Double) As String
    FAS_FormatCm = Replace(Format$(valueCm, "0.00"), ".", GetDecimalSeperator())
End Function

Private Function FAS_FormatInvariant(ByVal valueCm As Double) As String
    FAS_FormatInvariant = Replace(Format$(valueCm, "0.00"), GetDecimalSeperator(), ".")
End Function
