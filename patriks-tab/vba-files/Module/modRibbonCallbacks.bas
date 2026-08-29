Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

' Written from Ribbon state callbacks to prove which packaged code Excel loads.
Private Const FAS_RIBBON_DIAGNOSTIC_FILE As String = "fas-ribbon-diagnostics.log"

Public Sub Ribbon_ShowCommandPalette(ByVal control As IRibbonControl)
    FAS_ShowCommandPalette
End Sub

Public Sub Ribbon_ShowNavigator(ByVal control As IRibbonControl)
    FAS_ShowNavigator
End Sub

Public Sub Ribbon_ShowCommandHelp(ByVal control As IRibbonControl)
    FAS_ShowCommandHelp
End Sub

Public Sub Ribbon_ShowShortcutSettings(ByVal control As IRibbonControl)
    FAS_ShowShortcutSettings
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

Public Sub Ribbon_WrapSelectedFormulasRound(ByVal control As IRibbonControl)
    FAS_WrapSelectedFormulasRound
End Sub

Public Sub Ribbon_WrapSelectedFormulasIfError(ByVal control As IRibbonControl)
    FAS_WrapSelectedFormulasIfError
End Sub

Public Sub Ribbon_CycleNumberFormat(ByVal control As IRibbonControl)
    FAS_CycleNumberFormat
End Sub

Public Sub Ribbon_CyclePercentFormat(ByVal control As IRibbonControl)
    FAS_CyclePercentFormat
End Sub

Public Sub Ribbon_CycleCurrencyFormat(ByVal control As IRibbonControl)
    FAS_CycleCurrencyFormat
End Sub

Public Sub Ribbon_BuildSUMIFS(ByVal control As IRibbonControl)
    FAS_BuildSUMIFS
End Sub

Public Sub Ribbon_BuildCOUNTIFS(ByVal control As IRibbonControl)
    FAS_BuildCOUNTIFS
End Sub

Public Sub Ribbon_CaptureIFSPattern(ByVal control As IRibbonControl)
    FAS_CaptureIFSPattern
End Sub

Public Sub Ribbon_FillIFSPattern(ByVal control As IRibbonControl)
    FAS_FillIFSPattern
End Sub

Public Sub Ribbon_SwitchToSUMIFS(ByVal control As IRibbonControl)
    FAS_SwitchToSUMIFS
End Sub

Public Sub Ribbon_SwitchToCOUNTIFS(ByVal control As IRibbonControl)
    FAS_SwitchToCOUNTIFS
End Sub

Public Sub Ribbon_SwitchToAVERAGEIFS(ByVal control As IRibbonControl)
    FAS_SwitchToAVERAGEIFS
End Sub

Public Sub Ribbon_SwitchToMINIFS(ByVal control As IRibbonControl)
    FAS_SwitchToMINIFS
End Sub

Public Sub Ribbon_SwitchToMAXIFS(ByVal control As IRibbonControl)
    FAS_SwitchToMAXIFS
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

Public Sub Ribbon_InsertSubtotalAboveSelection(ByVal control As IRibbonControl)
    FAS_InsertSubtotalAboveSelection
End Sub

Public Sub Ribbon_InsertSubtotalBelowSelection(ByVal control As IRibbonControl)
    FAS_InsertSubtotalBelowSelection
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

Public Sub Ribbon_CreateCapexSummary(ByVal control As IRibbonControl)
    FAS_CreateCapexSummary
End Sub

Public Sub Ribbon_CreateGroupByReport(ByVal control As IRibbonControl)
    FAS_CreateGroupByReport
End Sub

Public Sub Ribbon_CreateGroupByReportWithCount(ByVal control As IRibbonControl)
    FAS_CreateGroupByReportWithCount
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

Public Sub Ribbon_ToggleHideGridlines(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    FAS_SetHideGridlines pressed
    FAS_InvalidateRibbonControl "tglFASHideGridlines"
End Sub

Public Sub Ribbon_GetHideGridlinesPressed(ByVal control As IRibbonControl, ByRef returnedVal)
    On Error GoTo ReadFailed
    FAS_LogRibbonDiagnostic "Ribbon_GetHideGridlinesPressed", "start"
    returnedVal = FAS_GetHideGridlines()
    FAS_LogRibbonDiagnostic "Ribbon_GetHideGridlinesPressed", "success=" & CStr(returnedVal)
    Exit Sub

ReadFailed:
    returnedVal = False
    FAS_LogRibbonDiagnostic "Ribbon_GetHideGridlinesPressed", _
        "error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Public Sub Ribbon_ToggleOutlineAboveLeft(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    FAS_SetOutlineAboveLeft pressed
    FAS_InvalidateRibbonControl "tglFASOutlineAboveLeft"
End Sub

Public Sub Ribbon_GetOutlineAboveLeftPressed(ByVal control As IRibbonControl, ByRef returnedVal)
    On Error GoTo ReadFailed
    FAS_LogRibbonDiagnostic "Ribbon_GetOutlineAboveLeftPressed", "start"
    returnedVal = FAS_GetOutlineAboveLeft()
    FAS_LogRibbonDiagnostic "Ribbon_GetOutlineAboveLeftPressed", "success=" & CStr(returnedVal)
    Exit Sub

ReadFailed:
    returnedVal = False
    FAS_LogRibbonDiagnostic "Ribbon_GetOutlineAboveLeftPressed", _
        "error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Private Sub FAS_LogRibbonDiagnostic(ByVal callbackName As String, ByVal messageText As String)
    Dim fileNumber As Integer
    Dim logPath As String

    ' Diagnostics must never create another Ribbon callback failure.
    On Error Resume Next
    logPath = Environ$("TEMP") & "\" & FAS_RIBBON_DIAGNOSTIC_FILE
    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        " | " & callbackName & " | " & messageText
    Close #fileNumber
    On Error GoTo 0
End Sub
