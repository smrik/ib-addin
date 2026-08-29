Attribute VB_Name = "ModuleFASProductivity"
Option Explicit

Private Type FASLinePreset
    Key As String
    DisplayName As String
    ColorValue As Long
    DashStyleValue As Long
    WeightValue As Single
End Type

Private mHasCurrentColor As Boolean
Private mCurrentColor As Long
Private mCurrentColorTarget As String

Private Const FAS_DEFAULT_LIGHTEN_PERCENT As Single = 20
Private Const FAS_DEFAULT_LINE_PRESET As String = "blue-dashed-1pt"

' Capture the first selected text, fill, or outline color for the next Apply action.
Public Sub FASCaptureCurrentColor()
    Dim selectionType As Long
    Dim selectedShapes As Object
    Dim firstShape As Object

    mHasCurrentColor = False
    mCurrentColorTarget = vbNullString

    On Error GoTo CaptureFailed
    selectionType = ActiveWindow.Selection.Type

    If selectionType = ppSelectionText Then
        mCurrentColor = ActiveWindow.Selection.TextRange.Font.Color.RGB
        mCurrentColorTarget = "font"
        mHasCurrentColor = True
    ElseIf selectionType = ppSelectionShapes Then
        Set selectedShapes = ActiveWindow.Selection.ShapeRange
        If selectedShapes.Count = 0 Then Exit Sub
        Set firstShape = selectedShapes.Item(1)
        mCurrentColorTarget = FAS_GetShapeColorTarget(firstShape)
        If mCurrentColorTarget = "fill" Then
            mCurrentColor = firstShape.Fill.ForeColor.RGB
        Else
            mCurrentColor = firstShape.Line.ForeColor.RGB
        End If
        mHasCurrentColor = True
    End If

    FAS_LogProductivity "Current " & mCurrentColorTarget & " color captured."
    Exit Sub

CaptureFailed:
    mHasCurrentColor = False
    FAS_LogProductivity "Could not capture the current color."
End Sub

' Apply the captured color to the current text or selected shapes.
Public Sub FASApplyCurrentColor()
    Dim selectionType As Long
    Dim selectedShapes As Object
    Dim shp As Object

    If Not mHasCurrentColor Then
        FAS_LogProductivity "Capture a color first."
        Exit Sub
    End If

    On Error GoTo ApplyFailed
    selectionType = ActiveWindow.Selection.Type
    If mCurrentColorTarget = "font" And selectionType = ppSelectionText Then
        ActiveWindow.Selection.TextRange.Font.Color.RGB = mCurrentColor
    ElseIf selectionType = ppSelectionShapes Then
        Set selectedShapes = ActiveWindow.Selection.ShapeRange
        For Each shp In selectedShapes
            FAS_SetShapeColor shp, mCurrentColorTarget, mCurrentColor
        Next shp
    Else
        FAS_LogProductivity "Select text or shapes before applying a color."
        Exit Sub
    End If

    FAS_LogProductivity "Captured color applied."
    Exit Sub

ApplyFailed:
    FAS_LogProductivity "Could not apply the captured color."
End Sub

' Lighten the selected color toward white. Transparency is left unchanged.
Public Sub FASLightenSelectedColor(Optional ByVal lightenPercent As Single = -1)
    Dim selectionType As Long
    Dim selectedShapes As Object
    Dim shp As Object
    Dim targetName As String
    Dim colorValue As Long

    If lightenPercent < 0 Then
        lightenPercent = CSng(GetSetting("Instrumenta", "Productivity", "LightenPercent", CStr(FAS_DEFAULT_LIGHTEN_PERCENT)))
    Else
        If lightenPercent > 100 Then lightenPercent = 100
        SaveSetting "Instrumenta", "Productivity", "LightenPercent", CStr(lightenPercent)
    End If

    On Error GoTo LightenFailed
    selectionType = ActiveWindow.Selection.Type
    If selectionType = ppSelectionText Then
        colorValue = ActiveWindow.Selection.TextRange.Font.Color.RGB
        ActiveWindow.Selection.TextRange.Font.Color.RGB = FAS_BlendColorToWhite(colorValue, lightenPercent)
    ElseIf selectionType = ppSelectionShapes Then
        Set selectedShapes = ActiveWindow.Selection.ShapeRange
        For Each shp In selectedShapes
            targetName = FAS_GetShapeColorTarget(shp)
            If targetName = "fill" Then
                colorValue = shp.Fill.ForeColor.RGB
            Else
                colorValue = shp.Line.ForeColor.RGB
            End If
            FAS_SetShapeColor shp, targetName, FAS_BlendColorToWhite(colorValue, lightenPercent)
        Next shp
    Else
        FAS_LogProductivity "Select text or shapes before lightening a color."
        Exit Sub
    End If

    FAS_LogProductivity "Color lightened toward white."
    Exit Sub

LightenFailed:
    FAS_LogProductivity "Could not lighten the selected color."
End Sub

' Convert selected circles to PowerPoint arc AutoShapes while preserving placement and style.
Public Sub FASConvertSelectedCirclesToArcs()
    Dim selectedShapes As Object
    Dim shp As Object
    Dim originalLeft As Single
    Dim originalTop As Single
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim originalRotation As Single
    Dim convertedCount As Long

    Set selectedShapes = FAS_GetSelectedShapeRange()
    If selectedShapes Is Nothing Then Exit Sub

    On Error GoTo ConvertFailed
    For Each shp In selectedShapes
        If shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeOval Then
            originalLeft = shp.Left
            originalTop = shp.Top
            originalWidth = shp.Width
            originalHeight = shp.Height
            originalRotation = shp.Rotation

            shp.AutoShapeType = msoShapeArc
            shp.Left = originalLeft
            shp.Top = originalTop
            shp.Width = originalWidth
            shp.Height = originalHeight
            shp.Rotation = originalRotation
            convertedCount = convertedCount + 1
        End If
    Next shp

    FAS_LogProductivity CStr(convertedCount) & " circle(s) converted to arcs."
    Exit Sub

ConvertFailed:
    FAS_LogProductivity "Circle-to-arc conversion stopped after " & convertedCount & " shape(s)."
End Sub

' Convenience path for the full two-circle workflow.
Public Sub FASConvertCirclesToArcsAndConnectCenters()
    FASConvertSelectedCirclesToArcs
    FASConnectSelectedArcCenters
End Sub

' Add one attached straight connector between the two selected shapes.
Public Sub FASConnectSelectedArcCenters()
    Dim selectedShapes As Object
    Dim slideObject As Object
    Dim fromShape As Object
    Dim toShape As Object
    Dim connector As Object
    Dim bestFromSite As Long
    Dim bestToSite As Long
    Dim fromSite As Long
    Dim toSite As Long
    Dim bestSurface As Double
    Dim currentSurface As Double

    Set selectedShapes = FAS_GetSelectedShapeRange()
    If selectedShapes Is Nothing Then Exit Sub
    If selectedShapes.Count <> 2 Then
        FAS_LogProductivity "Select exactly two circles/arcs for a center connector."
        Exit Sub
    End If

    Set fromShape = selectedShapes.Item(1)
    Set toShape = selectedShapes.Item(2)
    If fromShape.ConnectionSiteCount = 0 Or toShape.ConnectionSiteCount = 0 Then
        FAS_LogProductivity "The selected shapes have no connector sites."
        Exit Sub
    End If

    On Error GoTo ConnectorFailed
    Set slideObject = ActiveWindow.View.Slide
    Set connector = slideObject.Shapes.AddConnector( _
        msoConnectorStraight, _
        fromShape.Left + fromShape.Width / 2, _
        fromShape.Top + fromShape.Height / 2, _
        toShape.Left + toShape.Width / 2, _
        toShape.Top + toShape.Height / 2)

    bestSurface = 1E+30
    For fromSite = 1 To fromShape.ConnectionSiteCount
        For toSite = 1 To toShape.ConnectionSiteCount
            With connector.ConnectorFormat
                .BeginConnect fromShape, fromSite
                .EndConnect toShape, toSite
            End With
            connector.RerouteConnections
            currentSurface = CDbl(connector.Width) * CDbl(connector.Height)
            If currentSurface < bestSurface Then
                bestSurface = currentSurface
                bestFromSite = fromSite
                bestToSite = toSite
            End If
        Next toSite
    Next fromSite

    With connector.ConnectorFormat
        .BeginConnect fromShape, bestFromSite
        .EndConnect toShape, bestToSite
    End With
    connector.RerouteConnections
    FAS_ApplyLinePresetToShape connector, FAS_GetLastLinePresetKey()
    FAS_LogProductivity "Center connector added using the shortest attached route."
    Exit Sub

ConnectorFailed:
    On Error Resume Next
    If Not connector Is Nothing Then connector.Delete
    On Error GoTo 0
    FAS_LogProductivity "Could not connect the selected shapes."
End Sub

' Fast path: apply the last selected preset without reopening the preset menu.
Public Sub FASApplyLastLinePreset()
    FAS_ApplyLinePresetKey FAS_GetLastLinePresetKey(), False
End Sub

Public Sub FASApplyBlueDashedLinePreset()
    FAS_ApplyLinePresetKey "blue-dashed-1pt", True
End Sub

Public Sub FASApplyBlueSolidLinePreset()
    FAS_ApplyLinePresetKey "blue-solid-1pt", True
End Sub

Public Sub FASApplyGrayDashedLinePreset()
    FAS_ApplyLinePresetKey "gray-dashed-1pt", True
End Sub

Public Sub FASApplyGraySolidLinePreset()
    FAS_ApplyLinePresetKey "gray-solid-1pt", True
End Sub

Private Sub FAS_ApplyLinePresetKey(ByVal presetKey As String, ByVal rememberChoice As Boolean)
    Dim selectedShapes As Object
    Dim shp As Object
    Dim preset As FASLinePreset

    Set selectedShapes = FAS_GetSelectedShapeRange()
    If selectedShapes Is Nothing Then Exit Sub

    On Error GoTo PresetFailed
    For Each shp In selectedShapes
        FAS_ApplyLinePresetToShape shp, presetKey
    Next shp

    If rememberChoice Then SaveSetting "Instrumenta", "Productivity", "LastLinePreset", FAS_NormalizePresetKey(presetKey)
    preset = FAS_GetLinePreset(presetKey)
    FAS_LogProductivity "Line preset applied: " & preset.DisplayName
    Exit Sub

PresetFailed:
    FAS_LogProductivity "Could not apply line preset."
End Sub

Private Sub FAS_ApplyLinePresetToShape(ByVal shp As Object, ByVal presetKey As String)
    Dim preset As FASLinePreset

    preset = FAS_GetLinePreset(presetKey)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = preset.ColorValue
    shp.Line.DashStyle = preset.DashStyleValue
    shp.Line.Weight = preset.WeightValue
End Sub

Private Function FAS_GetLastLinePresetKey() As String
    Dim savedKey As String

    savedKey = FAS_NormalizePresetKey(GetSetting("Instrumenta", "Productivity", "LastLinePreset", FAS_DEFAULT_LINE_PRESET))
    If Len(savedKey) = 0 Then savedKey = FAS_DEFAULT_LINE_PRESET
    FAS_GetLastLinePresetKey = savedKey
End Function

Private Function FAS_NormalizePresetKey(ByVal presetKey As String) As String
    Select Case LCase$(Trim$(presetKey))
        Case "blue-dashed-1pt", "blue-solid-1pt", "gray-dashed-1pt", "gray-solid-1pt"
            FAS_NormalizePresetKey = LCase$(Trim$(presetKey))
        Case Else
            FAS_NormalizePresetKey = FAS_DEFAULT_LINE_PRESET
    End Select
End Function

Private Function FAS_GetLinePreset(ByVal presetKey As String) As FASLinePreset
    Dim preset As FASLinePreset

    preset.Key = FAS_NormalizePresetKey(presetKey)
    preset.WeightValue = 1
    preset.DashStyleValue = msoLineDash

    Select Case preset.Key
        Case "blue-solid-1pt"
            preset.DisplayName = "Blue solid 1 pt"
            preset.ColorValue = RGB(31, 78, 121)
            preset.DashStyleValue = msoLineSolid
        Case "gray-dashed-1pt"
            preset.DisplayName = "Gray dashed 1 pt"
            preset.ColorValue = RGB(128, 128, 128)
        Case "gray-solid-1pt"
            preset.DisplayName = "Gray solid 1 pt"
            preset.ColorValue = RGB(128, 128, 128)
            preset.DashStyleValue = msoLineSolid
        Case Else
            preset.Key = "blue-dashed-1pt"
            preset.DisplayName = "Blue dashed 1 pt"
            preset.ColorValue = RGB(31, 78, 121)
    End Select

    FAS_GetLinePreset = preset
End Function

Private Function FAS_GetSelectedShapeRange() As Object
    On Error Resume Next
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Function
    Set FAS_GetSelectedShapeRange = ActiveWindow.Selection.ShapeRange
    On Error GoTo 0
End Function

Private Function FAS_GetShapeColorTarget(ByVal shp As Object) As String
    On Error GoTo UseLine
    If shp.Fill.Visible <> msoFalse And shp.Fill.Transparency < 1 Then
        FAS_GetShapeColorTarget = "fill"
        Exit Function
    End If

UseLine:
    FAS_GetShapeColorTarget = "line"
End Function

Private Sub FAS_SetShapeColor(ByVal shp As Object, ByVal targetName As String, ByVal colorValue As Long)
    If targetName = "fill" Then
        shp.Fill.ForeColor.RGB = colorValue
    ElseIf targetName = "line" Then
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = colorValue
    End If
End Sub

Private Function FAS_BlendColorToWhite(ByVal colorValue As Long, ByVal lightenPercent As Single) As Long
    Dim fraction As Double
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    fraction = CDbl(lightenPercent) / 100
    If fraction < 0 Then fraction = 0
    If fraction > 1 Then fraction = 1

    redValue = colorValue And &HFF&
    greenValue = (colorValue \ &H100&) And &HFF&
    blueValue = (colorValue \ &H10000) And &HFF&
    FAS_BlendColorToWhite = RGB( _
        redValue + CLng((255 - redValue) * fraction), _
        greenValue + CLng((255 - greenValue) * fraction), _
        blueValue + CLng((255 - blueValue) * fraction))
End Function

Private Sub FAS_LogProductivity(ByVal messageText As String)
    Debug.Print "Instrumenta: " & messageText
End Sub
