Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

Public Sub Ribbon_ShowCommandPalette(ByVal control As IRibbonControl)
    FAS_ShowCommandPalette
End Sub

Public Sub Ribbon_ShowNavigator(ByVal control As IRibbonControl)
    FAS_ShowNavigator
End Sub

Public Sub Ribbon_ShowCommandHelp(ByVal control As IRibbonControl)
    FAS_ShowCommandHelp
End Sub

Public Sub Ribbon_JumpToSheet(ByVal control As IRibbonControl)
    FAS_JumpToSheet
End Sub

Public Sub Ribbon_GoToPL(ByVal control As IRibbonControl)
    FAS_GoToPL
End Sub

Public Sub Ribbon_GoToBS(ByVal control As IRibbonControl)
    FAS_GoToBS
End Sub

Public Sub Ribbon_GoToTB(ByVal control As IRibbonControl)
    FAS_GoToTB
End Sub

Public Sub Ribbon_GoToGL(ByVal control As IRibbonControl)
    FAS_GoToGL
End Sub

Public Sub Ribbon_GoToMapping(ByVal control As IRibbonControl)
    FAS_GoToMapping
End Sub

Public Sub Ribbon_ShowWorkbookTabsMenu(ByVal control As IRibbonControl)
    FAS_ShowWorkbookTabsMenu
End Sub

Public Sub Ribbon_FindTBAccount(ByVal control As IRibbonControl)
    FAS_FindTBAccount
End Sub

Public Sub Ribbon_SearchLedgerIndex(ByVal control As IRibbonControl)
    FAS_SearchLedgerIndex
End Sub

Public Sub Ribbon_CopyFormulasSmart(ByVal control As IRibbonControl)
    FAS_CopyFormulasSmart
End Sub

Public Sub Ribbon_ReplaceInSelectedFormulas(ByVal control As IRibbonControl)
    FAS_ReplaceInSelectedFormulas
End Sub

Public Sub Ribbon_ShowPrecedents(ByVal control As IRibbonControl)
    FAS_ShowPrecedents
End Sub

Public Sub Ribbon_CycleDirectPrecedent(ByVal control As IRibbonControl)
    FAS_CycleDirectPrecedent
End Sub

Public Sub Ribbon_BackToTraceOrigin(ByVal control As IRibbonControl)
    FAS_BackToTraceOrigin
End Sub

Public Sub Ribbon_PasteValuesOnly(ByVal control As IRibbonControl)
    FAS_PasteValuesOnly
End Sub

Public Sub Ribbon_PasteAsTextOnly(ByVal control As IRibbonControl)
    FAS_PasteAsTextOnly
End Sub

Public Sub Ribbon_CopyVisibleOnly(ByVal control As IRibbonControl)
    FAS_CopyVisibleOnly
End Sub

Public Sub Ribbon_CopyLinks(ByVal control As IRibbonControl)
    FAS_CopyLinks
End Sub

Public Sub Ribbon_PasteLinks(ByVal control As IRibbonControl)
    FAS_PasteLinks
End Sub

Public Sub Ribbon_WrapSelectedFormulasRound(ByVal control As IRibbonControl)
    FAS_WrapSelectedFormulasRound
End Sub

Public Sub Ribbon_WrapSelectedFormulasIfError(ByVal control As IRibbonControl)
    FAS_WrapSelectedFormulasIfError
End Sub

Public Sub Ribbon_CaptureCurrentColor(ByVal control As IRibbonControl)
    FAS_CaptureCurrentColor
End Sub

Public Sub Ribbon_ApplyCurrentColor(ByVal control As IRibbonControl)
    FAS_ApplyCurrentColor
End Sub

Public Sub Ribbon_InsertDelta(ByVal control As IRibbonControl)
    FAS_InsertDelta
End Sub

Public Sub Ribbon_AutoFitWidthWithMargin(ByVal control As IRibbonControl)
    FAS_AutoFitWidthWithMargin
End Sub

Public Sub Ribbon_InsertSubtotalAboveSelection(ByVal control As IRibbonControl)
    FAS_InsertSubtotalAboveSelection
End Sub

Public Sub Ribbon_InsertSubtotalBelowSelection(ByVal control As IRibbonControl)
    FAS_InsertSubtotalBelowSelection
End Sub

Public Sub Ribbon_FormatSelectedChartFAS(ByVal control As IRibbonControl)
    FAS_FormatSelectedChartFAS
End Sub

Public Sub Ribbon_FormatSelectedSeriesFAS(ByVal control As IRibbonControl)
    FAS_FormatSelectedSeriesFAS
End Sub

Public Sub Ribbon_ChartLegendBottom(ByVal control As IRibbonControl)
    FAS_ChartLegendBottom
End Sub

Public Sub Ribbon_ChartLegendRight(ByVal control As IRibbonControl)
    FAS_ChartLegendRight
End Sub

Public Sub Ribbon_ChartLegendOff(ByVal control As IRibbonControl)
    FAS_ChartLegendOff
End Sub

Public Sub Ribbon_ChartGridlinesLight(ByVal control As IRibbonControl)
    FAS_ChartGridlinesLight
End Sub

Public Sub Ribbon_ChartGridlinesOff(ByVal control As IRibbonControl)
    FAS_ChartGridlinesOff
End Sub

Public Sub Ribbon_InsertDividerSheet(ByVal control As IRibbonControl)
    FAS_InsertDividerSheet
End Sub

Public Sub Ribbon_ShowActiveWorkbookBookmarks(ByVal control As IRibbonControl)
    FAS_ShowActiveWorkbookBookmarks
End Sub

Public Sub Ribbon_ShowActiveWorkbookMappings(ByVal control As IRibbonControl)
    FAS_ShowActiveWorkbookMappings
End Sub

Public Sub Ribbon_ShowDrilldown(ByVal control As IRibbonControl)
    FAS_ShowDrilldown
End Sub

Public Sub Ribbon_SetShapesDoNotMoveOrSize(ByVal control As IRibbonControl)
    FAS_SetShapesDoNotMoveOrSize
End Sub
