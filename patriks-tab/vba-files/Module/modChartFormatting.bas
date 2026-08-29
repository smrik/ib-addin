Attribute VB_Name = "modChartFormatting"
Option Explicit

Private Const FAS_CHART_TEXT_COLOR As Long = 4934470
Private Const FAS_CHART_GRID_COLOR As Long = 15132390
Private Const FAS_CHART_FONT_NAME As String = "Arial"
Private Const FAS_CHART_FONT_SIZE As Long = 8
Private Const FAS_CHART_VALUE_AXIS_TITLE As String = "EUR 000"
Private Const FAS_CHART_HEADER_TOP As Double = 4
Private Const FAS_CHART_HEADER_LEFT As Double = 4
Private Const FAS_CHART_HEADER_GAP As Double = 2
Private Const FAS_CHART_PLOT_TOP_FALLBACK As Double = 42
Private Const FAS_CHART_PLOT_BOTTOM_PAD As Double = 10

Public Sub FAS_FormatSelectedChartFAS()
    Dim targetCharts As Collection
    Dim chartItem As Variant
    Dim formattedCount As Long

    Set targetCharts = FAS_GetSelectedCharts()
    If targetCharts.Count = 0 Then
        FAS_SetStatus "Select a chart first."
        Exit Sub
    End If

    For Each chartItem In targetCharts
        FAS_FormatChartLayout chartItem
        formattedCount = formattedCount + 1
    Next chartItem

    FAS_SetStatus "Formatted " & formattedCount & " chart(s)."
End Sub

Public Sub FAS_FormatSelectedSeriesFAS()
    Dim targetSeries As Collection
    Dim seriesItem As Variant
    Dim activeChartObject As Chart
    Dim formattedCount As Long

    Set targetSeries = FAS_GetTargetSeries()
    If targetSeries.Count = 0 Then
        FAS_SetStatus "Select a data series or chart first."
        Exit Sub
    End If

    For Each seriesItem In targetSeries
        FAS_FormatSeriesLabelsAndLine seriesItem
        formattedCount = formattedCount + 1
    Next seriesItem

    Set activeChartObject = FAS_GetActiveChart()
    If Not activeChartObject Is Nothing Then
        FAS_FormatChartAxisTitles activeChartObject
    End If

    FAS_SetStatus "Formatted " & formattedCount & " series."
End Sub

Public Sub FAS_ChartLegendBottom()
    FAS_SetSelectedChartLegend xlLegendPositionBottom
End Sub

Public Sub FAS_ChartLegendRight()
    FAS_SetSelectedChartLegend xlLegendPositionRight
End Sub

Public Sub FAS_ChartLegendOff()
    Dim targetCharts As Collection
    Dim chartItem As Variant

    Set targetCharts = FAS_GetSelectedCharts()
    If targetCharts.Count = 0 Then
        FAS_SetStatus "Select a chart first."
        Exit Sub
    End If

    For Each chartItem In targetCharts
        chartItem.HasLegend = False
    Next chartItem

    FAS_SetStatus "Legend removed."
End Sub

Public Sub FAS_ChartGridlinesLight()
    Dim targetCharts As Collection
    Dim chartItem As Variant

    Set targetCharts = FAS_GetSelectedCharts()
    If targetCharts.Count = 0 Then
        FAS_SetStatus "Select a chart first."
        Exit Sub
    End If

    For Each chartItem In targetCharts
        FAS_FormatGridlines chartItem, True
    Next chartItem

    FAS_SetStatus "Gridlines set to light."
End Sub

Public Sub FAS_ChartGridlinesOff()
    Dim targetCharts As Collection
    Dim chartItem As Variant

    Set targetCharts = FAS_GetSelectedCharts()
    If targetCharts.Count = 0 Then
        FAS_SetStatus "Select a chart first."
        Exit Sub
    End If

    For Each chartItem In targetCharts
        FAS_FormatGridlines chartItem, False
    Next chartItem

    FAS_SetStatus "Gridlines removed."
End Sub

Public Function FAS_RunChartFormatCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    If Len(normalizedCommand) = 0 Then Exit Function

    Select Case normalizedCommand
        Case "chartfas", "chartformat", "chartstyle", "chartclean", "chartdeck"
            FAS_FormatSelectedChartFAS
        Case "seriesfas", "seriesformat", "datalabelsfas", "labelsfas", "linefas"
            FAS_FormatSelectedSeriesFAS
        Case "legendbottom", "chartlegendbottom"
            FAS_ChartLegendBottom
        Case "legendright", "chartlegendright"
            FAS_ChartLegendRight
        Case "legendoff", "chartlegendoff"
            FAS_ChartLegendOff
        Case "gridlight", "gridlineslight", "chartgridlight"
            FAS_ChartGridlinesLight
        Case "gridoff", "gridlinesoff", "chartgridoff"
            FAS_ChartGridlinesOff
        Case Else
            Exit Function
    End Select

    FAS_RunChartFormatCommand = True
End Function

Public Function FAS_IsChartFormatCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    Select Case normalizedCommand
        Case "chartfas", "chartformat", "chartstyle", "chartclean", "chartdeck", _
            "seriesfas", "seriesformat", "datalabelsfas", "labelsfas", "linefas", _
            "legendbottom", "chartlegendbottom", "legendright", "chartlegendright", _
            "legendoff", "chartlegendoff", "gridlight", "gridlineslight", "chartgridlight", _
            "gridoff", "gridlinesoff", "chartgridoff"
            FAS_IsChartFormatCommand = True
    End Select
End Function

Private Sub FAS_FormatChartLayout(ByVal targetChart As Chart)
    If targetChart Is Nothing Then Exit Sub

    FAS_FormatChartText targetChart
    FAS_FormatChartTitle targetChart
    FAS_FormatChartLegend targetChart
    FAS_FormatAxes targetChart
    FAS_FormatGridlines targetChart, False
    FAS_FormatChartFrame targetChart
    FAS_MaximizePlotArea targetChart
End Sub

Private Sub FAS_FormatChartText(ByVal targetChart As Chart)
    On Error Resume Next
    With targetChart.ChartArea.Format.TextFrame2.TextRange.Font
        .Name = FAS_CHART_FONT_NAME
        .Size = FAS_CHART_FONT_SIZE
        .Fill.ForeColor.RGB = FAS_CHART_TEXT_COLOR
        .Bold = msoFalse
    End With
    On Error GoTo 0
End Sub

Private Sub FAS_FormatChartTitle(ByVal targetChart As Chart)
    If Not targetChart.HasTitle Then Exit Sub

    On Error Resume Next
    targetChart.ChartTitle.IncludeInLayout = False
    targetChart.ChartTitle.Left = FAS_CHART_HEADER_LEFT
    targetChart.ChartTitle.Top = FAS_CHART_HEADER_TOP
    With targetChart.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = FAS_CHART_FONT_NAME
        .Size = FAS_CHART_FONT_SIZE
        .Fill.ForeColor.RGB = FAS_CHART_TEXT_COLOR
        .Bold = msoTrue
    End With
    On Error GoTo 0
End Sub

Private Sub FAS_FormatChartLegend(ByVal targetChart As Chart)
    If Not targetChart.HasLegend Then
        targetChart.HasLegend = True
    End If

    On Error Resume Next
    targetChart.Legend.Position = xlLegendPositionTop
    targetChart.Legend.IncludeInLayout = False
    targetChart.Legend.Left = FAS_CHART_HEADER_LEFT
    targetChart.Legend.Top = FAS_GetLegendTop(targetChart)
    targetChart.Legend.Width = targetChart.ChartArea.Width - (FAS_CHART_HEADER_LEFT * 2)
    targetChart.Legend.Height = 16
    With targetChart.Legend.Format.TextFrame2.TextRange.Font
        .Name = FAS_CHART_FONT_NAME
        .Size = FAS_CHART_FONT_SIZE
        .Fill.ForeColor.RGB = FAS_CHART_TEXT_COLOR
        .Bold = msoFalse
    End With
    On Error GoTo 0
End Sub

Private Function FAS_GetLegendTop(ByVal targetChart As Chart) As Double
    FAS_GetLegendTop = FAS_CHART_HEADER_TOP + 12
    On Error Resume Next
    If targetChart.HasTitle Then
        FAS_GetLegendTop = targetChart.ChartTitle.Top + targetChart.ChartTitle.Height + FAS_CHART_HEADER_GAP
    End If
    On Error GoTo 0
End Function

Private Sub FAS_FormatAxes(ByVal targetChart As Chart)
    Dim valueAxis As Axis
    Dim categoryAxis As Axis

    On Error Resume Next
    Set valueAxis = targetChart.Axes(xlValue)
    Set categoryAxis = targetChart.Axes(xlCategory)
    On Error GoTo 0

    If Not valueAxis Is Nothing Then
        FAS_FormatAxisText valueAxis
        On Error Resume Next
        valueAxis.MajorTickMark = xlTickMarkNone
        valueAxis.MinorTickMark = xlTickMarkNone
        valueAxis.Format.Line.Visible = msoFalse
        On Error GoTo 0
    End If

    If Not categoryAxis Is Nothing Then
        FAS_FormatAxisText categoryAxis
        On Error Resume Next
        categoryAxis.MajorTickMark = xlTickMarkNone
        categoryAxis.MinorTickMark = xlTickMarkNone
        categoryAxis.TickLabelPosition = xlTickLabelPositionLow
        categoryAxis.Format.Line.Visible = msoTrue
        categoryAxis.Format.Line.ForeColor.RGB = FAS_CHART_TEXT_COLOR
        categoryAxis.Format.Line.Weight = 0.75
        On Error GoTo 0
    End If

    FAS_FormatChartAxisTitles targetChart
End Sub

Private Sub FAS_FormatAxisText(ByVal targetAxis As Axis)
    On Error Resume Next
    With targetAxis.TickLabels.Font
        .Name = FAS_CHART_FONT_NAME
        .Size = FAS_CHART_FONT_SIZE
        .Color = FAS_CHART_TEXT_COLOR
        .Bold = False
    End With
    If targetAxis.HasTitle Then
        With targetAxis.AxisTitle.Format.TextFrame2.TextRange.Font
            .Name = FAS_CHART_FONT_NAME
            .Size = FAS_CHART_FONT_SIZE
            .Fill.ForeColor.RGB = FAS_CHART_TEXT_COLOR
            .Bold = msoFalse
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub FAS_FormatChartAxisTitles(ByVal targetChart As Chart)
    Dim valueAxis As Axis
    Dim categoryAxis As Axis

    On Error Resume Next
    Set categoryAxis = targetChart.Axes(xlCategory)
    If Not categoryAxis Is Nothing Then categoryAxis.HasTitle = False

    Set valueAxis = targetChart.Axes(xlValue)
    If Not valueAxis Is Nothing Then
        valueAxis.HasTitle = True
        If valueAxis.HasTitle Then
            valueAxis.AxisTitle.Text = FAS_CHART_VALUE_AXIS_TITLE
            valueAxis.AxisTitle.Orientation = xlHorizontal
            valueAxis.AxisTitle.Left = FAS_CHART_HEADER_LEFT
            valueAxis.AxisTitle.Top = FAS_GetLegendTop(targetChart) + 14
            With valueAxis.AxisTitle.Format.TextFrame2.TextRange.Font
                .Name = FAS_CHART_FONT_NAME
                .Size = FAS_CHART_FONT_SIZE
                .Fill.ForeColor.RGB = FAS_CHART_TEXT_COLOR
                .Bold = msoFalse
            End With
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub FAS_FormatGridlines(ByVal targetChart As Chart, ByVal showGridlines As Boolean)
    Dim valueAxis As Axis
    Dim categoryAxis As Axis

    On Error Resume Next
    Set valueAxis = targetChart.Axes(xlValue)
    Set categoryAxis = targetChart.Axes(xlCategory)
    FAS_SetAxisGridlines valueAxis, showGridlines
    FAS_SetAxisGridlines categoryAxis, showGridlines
    On Error GoTo 0
End Sub

Private Sub FAS_SetAxisGridlines(ByVal targetAxis As Axis, ByVal showGridlines As Boolean)
    If targetAxis Is Nothing Then Exit Sub

    On Error Resume Next
    targetAxis.HasMajorGridlines = showGridlines
    targetAxis.HasMinorGridlines = False
    If showGridlines Then
        targetAxis.MajorGridlines.Format.Line.Visible = msoTrue
        targetAxis.MajorGridlines.Format.Line.ForeColor.RGB = FAS_CHART_GRID_COLOR
        targetAxis.MajorGridlines.Format.Line.Weight = 0.5
    End If
    On Error GoTo 0
End Sub

Private Sub FAS_FormatChartFrame(ByVal targetChart As Chart)
    On Error Resume Next
    targetChart.ChartArea.Format.Fill.Visible = msoFalse
    targetChart.ChartArea.Format.Line.Visible = msoFalse
    targetChart.PlotArea.Format.Fill.Visible = msoFalse
    targetChart.PlotArea.Format.Line.Visible = msoFalse
    On Error GoTo 0
End Sub

Private Sub FAS_MaximizePlotArea(ByVal targetChart As Chart)
    Dim plotTop As Double
    Dim plotLeft As Double
    Dim plotWidth As Double
    Dim plotHeight As Double

    If targetChart Is Nothing Then Exit Sub

    plotTop = FAS_GetLegendTop(targetChart) + 20
    If plotTop < FAS_CHART_PLOT_TOP_FALLBACK Then plotTop = FAS_CHART_PLOT_TOP_FALLBACK
    plotLeft = 2
    plotWidth = targetChart.ChartArea.Width - 6
    plotHeight = targetChart.ChartArea.Height - plotTop - FAS_CHART_PLOT_BOTTOM_PAD

    If plotWidth < 50 Or plotHeight < 50 Then Exit Sub

    On Error Resume Next
    targetChart.PlotArea.Left = plotLeft
    targetChart.PlotArea.Top = plotTop
    targetChart.PlotArea.Width = plotWidth
    targetChart.PlotArea.Height = plotHeight
    On Error GoTo 0
End Sub

Private Sub FAS_FormatSeriesLabelsAndLine(ByVal targetSeries As Series)
    Dim seriesColor As Long

    If targetSeries Is Nothing Then Exit Sub
    seriesColor = FAS_GetSeriesColor(targetSeries)

    On Error Resume Next
    targetSeries.Format.Line.Visible = msoTrue
    targetSeries.Format.Line.ForeColor.RGB = seriesColor
    targetSeries.Format.Line.Weight = 1.5
    targetSeries.MarkerForegroundColor = seriesColor
    targetSeries.MarkerBackgroundColor = seriesColor
    targetSeries.MarkerStyle = xlMarkerStyleCircle
    targetSeries.MarkerSize = 5

    targetSeries.HasDataLabels = True
    targetSeries.DataLabels.Position = xlLabelPositionAbove
    targetSeries.DataLabels.NumberFormatLinked = True
    targetSeries.DataLabels.Font.Name = FAS_CHART_FONT_NAME
    targetSeries.DataLabels.Font.Size = FAS_CHART_FONT_SIZE
    targetSeries.DataLabels.Font.Color = seriesColor
    targetSeries.DataLabels.Format.Fill.Visible = msoFalse
    targetSeries.DataLabels.Format.Line.Visible = msoFalse
    On Error GoTo 0
End Sub

Private Function FAS_GetSeriesColor(ByVal targetSeries As Series) As Long
    On Error Resume Next
    FAS_GetSeriesColor = targetSeries.Format.Line.ForeColor.RGB
    If FAS_GetSeriesColor = 0 Then FAS_GetSeriesColor = targetSeries.MarkerForegroundColor
    If FAS_GetSeriesColor = 0 Then FAS_GetSeriesColor = FAS_CHART_TEXT_COLOR
    On Error GoTo 0
End Function

Private Sub FAS_SetSelectedChartLegend(ByVal legendPosition As XlLegendPosition)
    Dim targetCharts As Collection
    Dim chartItem As Variant

    Set targetCharts = FAS_GetSelectedCharts()
    If targetCharts.Count = 0 Then
        FAS_SetStatus "Select a chart first."
        Exit Sub
    End If

    For Each chartItem In targetCharts
        chartItem.HasLegend = True
        chartItem.Legend.Position = legendPosition
        FAS_FormatChartText chartItem
    Next chartItem

    FAS_SetStatus "Legend position updated."
End Sub

Private Function FAS_GetSelectedCharts() As Collection
    Dim targetCharts As Collection
    Dim selectedShapeRange As Object
    Dim selectedShape As Variant
    Dim activeChartObject As Chart

    Set targetCharts = New Collection

    Set activeChartObject = FAS_GetActiveChart()
    If Not activeChartObject Is Nothing Then targetCharts.Add activeChartObject

    If targetCharts.Count = 0 Then
        On Error Resume Next
        Set selectedShapeRange = Selection.ShapeRange
        On Error GoTo 0

        If Not selectedShapeRange Is Nothing Then
            For Each selectedShape In selectedShapeRange
                On Error Resume Next
                If selectedShape.HasChart Then targetCharts.Add selectedShape.Chart
                On Error GoTo 0
            Next selectedShape
        End If
    End If

    Set FAS_GetSelectedCharts = targetCharts
End Function

Private Function FAS_GetTargetSeries() As Collection
    Dim targetSeries As Collection
    Dim selectedSeries As Series
    Dim activeChartObject As Chart
    Dim seriesItem As Variant

    Set targetSeries = New Collection

    On Error Resume Next
    If TypeName(Selection) = "Series" Then
        Set selectedSeries = Selection
    End If
    On Error GoTo 0

    If Not selectedSeries Is Nothing Then
        targetSeries.Add selectedSeries
    End If

    Set activeChartObject = FAS_GetActiveChart()
    If targetSeries.Count = 0 Then
        If Not activeChartObject Is Nothing Then
            On Error Resume Next
            For Each seriesItem In activeChartObject.SeriesCollection
                targetSeries.Add seriesItem
            Next seriesItem
            On Error GoTo 0
        End If
    End If

    Set FAS_GetTargetSeries = targetSeries
End Function

Private Function FAS_GetActiveChart() As Chart
    On Error Resume Next
    Set FAS_GetActiveChart = ActiveChart
    On Error GoTo 0
End Function
