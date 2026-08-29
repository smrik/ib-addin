Attribute VB_Name = "modCommandRegistry"
Option Explicit

Public Function FAS_RunCommand(ByVal commandText As String) As Boolean
    Dim targetWorkbook As Workbook
    Dim normalizedCommand As String
    Dim groupKey As String
    Dim slotIndex As Long
    Dim targetSheet As Worksheet
    Dim searchText As String
    Dim projectSearchKind As String
    Dim destinationAddress As String
    Dim precedentIndex As Long

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    If Len(normalizedCommand) = 0 Then
        Exit Function
    End If

    On Error GoTo CommandFailed

    If FAS_TryRunCommandHelp(commandText) Then
        FAS_RunCommand = True
        Exit Function
    End If

    If FAS_TryRunIFSCommand(commandText) Then
        FAS_RunCommand = True
        Exit Function
    End If

    If FAS_TryRunFormatCycleCommand(commandText) Then
        FAS_RunCommand = True
        Exit Function
    End If

    If FAS_TryRunShortcutCommand(commandText) Then
        FAS_RunCommand = True
        Exit Function
    End If

    If FAS_TryRunFormulaWrapperCommand(commandText) Then
        FAS_RunCommand = True
        Exit Function
    End If

    If FAS_TryRunGroupByCommand(commandText) Then
        FAS_RunCommand = True
        Exit Function
    End If

    If FAS_TryRunCapexCommand(commandText) Then
        FAS_RunCommand = True
        Exit Function
    End If

    If FAS_IsBookmarkCommand(commandText) Then
        FAS_RunCommand = FAS_RunBookmarkCommand(commandText)
        Exit Function
    End If

    If FAS_IsChartFormatCommand(commandText) Then
        FAS_RunCommand = FAS_RunChartFormatCommand(commandText)
        Exit Function
    End If

    If FAS_IsSlideLayoutCommand(commandText) Then
        FAS_RunCommand = FAS_RunSlideLayoutCommand(commandText)
        Exit Function
    End If

    If FAS_TryParseFormulaCopyCommand(commandText, destinationAddress) Then
        FAS_RunCommand = FAS_CopySelectionFormulasSmartToAddress(destinationAddress)
        Exit Function
    End If

    If FAS_TryParsePrecedentCommand(commandText, precedentIndex) Then
        FAS_RunCommand = FAS_JumpToPrecedentByIndex(precedentIndex)
        If Not FAS_RunCommand Then
            FAS_SetStatus "No supported precedent " & precedentIndex & " for active cell."
        End If
        Exit Function
    End If

    If FAS_TryParseProjectSearchCommand(commandText, projectSearchKind, searchText) Then
        FAS_RunCommand = FAS_RunProjectSearch(projectSearchKind, searchText)
        Exit Function
    End If

    If FAS_TryParseSheetLookupCommand(commandText, groupKey, searchText) Then
        FAS_RunCommand = FAS_GoToCommonSheetMatch(groupKey, searchText)
        If Not FAS_RunCommand And Application.Visible Then
            FAS_SetStatus "No " & FAS_GetGroupLabel(groupKey) & " match found for '" & searchText & "'."
        End If
        Exit Function
    End If

    Select Case normalizedCommand
        Case "visible", "vis", "copyvisible", "visiblecopy"
            FAS_CopyVisibleOnly
            FAS_RunCommand = True
        Case "linkcopy", "copylinks", "copylink"
            FAS_CopyLinks
            FAS_RunCommand = True
        Case "linkpaste", "pastelinks", "pastelink"
            FAS_PasteLinks
            FAS_RunCommand = True
        Case "colorcapture", "capturecolor", "getcolor"
            FAS_CaptureCurrentColor
            FAS_RunCommand = True
        Case "colorapply", "applycolor", "usecolor"
            FAS_ApplyCurrentColor
            FAS_RunCommand = True
        Case "delta", "deltasign"
            FAS_InsertDelta
            FAS_RunCommand = True
        Case "autofit", "fitwidth", "width20", "autofitwidth"
            FAS_AutoFitWidthWithMargin
            FAS_RunCommand = True
        Case "settings", "options", "config"
            FAS_ShowSettings
            FAS_RunCommand = True
        Case "val", "values"
            FAS_PasteValuesOnly
            FAS_RunCommand = True
            Exit Function
        Case "txt", "text"
            FAS_PasteAsTextOnly
            FAS_RunCommand = True
            Exit Function
        Case "rep", "replace"
            FAS_ReplaceInSelectedFormulas
            FAS_RunCommand = True
            Exit Function
        Case "back"
            FAS_RunCommand = FAS_GoBackToTraceOrigin()
            Exit Function
        Case "tabs"
            FAS_ShowWorkbookTabsMenu
            FAS_RunCommand = True
            Exit Function
        Case "acct", "account"
            FAS_FindTBAccount
            FAS_RunCommand = True
            Exit Function
        Case "dkdebug", "idxdebug"
            FAS_RunCommand = FAS_ShowProjectSearchDebug()
            Exit Function
        Case "fcopy", "formulacopy", "smartcopy"
            FAS_CopyFormulasSmart
            FAS_RunCommand = True
            Exit Function
        Case "sub", "subtotal"
            FAS_InsertSubtotalBelowSelection
            FAS_RunCommand = True
            Exit Function
        Case "div", "divider"
            FAS_InsertDividerSheet
            FAS_RunCommand = True
            Exit Function
        Case "drill", "dd"
            FAS_RunCommand = FAS_ShowDrilldownForSelection()
            Exit Function
        Case "drillclear", "ddclear"
            FAS_RunCommand = FAS_ClearDrilldownShape()
            Exit Function
        Case "fixobj", "fixshape", "float"
            FAS_RunCommand = FAS_SetSelectedObjectsDoNotMoveOrSize()
            Exit Function
        Case "shapefix", "nofloat", "fixshapes"
            FAS_SetShapesDoNotMoveOrSize
            FAS_RunCommand = True
            Exit Function
        Case "maps", "mapinfo", "mappings"
            FAS_ShowActiveWorkbookMappings
            FAS_RunCommand = True
            Exit Function
        Case "mapreset", "resetmaps"
            FAS_RunCommand = FAS_ResetActiveWorkbookMappings()
            Exit Function
    End Select

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    If FAS_ParseJumpCommand(commandText, groupKey, slotIndex) Then
        Set targetSheet = FAS_GetResolvedGroupSheet(targetWorkbook, groupKey, slotIndex)
    Else
        Set targetSheet = FAS_FindBestSheetByQuery(targetWorkbook, commandText)
    End If

    If targetSheet Is Nothing Then
        FAS_SetUnknownCommandStatus commandText
        Exit Function
    End If

    targetSheet.Activate
    FAS_RunCommand = True
    Exit Function

CommandFailed:
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.description
    End If

    FAS_SetStatus "Command failed: " & Err.description
End Function

Public Function FAS_IsKnownCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String
    Dim groupKey As String
    Dim slotIndex As Long
    Dim searchText As String
    Dim projectSearchKind As String
    Dim destinationAddress As String
    Dim precedentIndex As Long

    If FAS_IsBookmarkCommand(commandText) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_IsCommandHelpCommand(commandText) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_IsIFSCommand(commandText) Or FAS_IsFormatCycleCommand(commandText) Or _
        FAS_IsKnownShortcutCommand(commandText) Or _
        FAS_IsFormulaWrapperCommand(commandText) Or _
        FAS_IsGroupByCommand(commandText) Or _
        FAS_IsCapexCommand(commandText) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_IsChartFormatCommand(commandText) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_IsSlideLayoutCommand(commandText) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_TryParseFormulaCopyCommand(commandText, destinationAddress) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_TryParsePrecedentCommand(commandText, precedentIndex) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_TryParseProjectSearchCommand(commandText, projectSearchKind, searchText) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    If FAS_TryParseSheetLookupCommand(commandText, groupKey, searchText) Then
        FAS_IsKnownCommand = True
        Exit Function
    End If

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    Select Case normalizedCommand
        Case "visible", "vis", "copyvisible", "visiblecopy", _
            "linkcopy", "copylinks", "copylink", "linkpaste", "pastelinks", "pastelink", _
            "colorcapture", "capturecolor", "getcolor", "colorapply", "applycolor", "usecolor", _
            "delta", "deltasign", "autofit", "fitwidth", "width20", "autofitwidth", _
            "settings", "options", "config", _
            "val", "values", "txt", "text", "rep", "replace", "back", "tabs", _
            "acct", "account", "fcopy", "formulacopy", "smartcopy", "sub", "subtotal", "div", "divider", _
            "dk", "ledger", "ix", "idx", "dktb", "dkmap", "dkdebug", "idxdebug", _
            "drill", "dd", "drillclear", "ddclear", "fixobj", "fixshape", "float", _
            "shapefix", "nofloat", "fixshapes", "maps", "mapinfo", "mappings", "mapreset", "resetmaps"
            FAS_IsKnownCommand = True
            Exit Function
    End Select

    FAS_IsKnownCommand = FAS_ParseJumpCommand(commandText, groupKey, slotIndex)
End Function

Public Function FAS_GetShortcutBuilderCommandList() As String
    FAS_GetShortcutBuilderCommandList = FAS_GetCommandCatalogShortcutList()
End Function

Private Sub FAS_SetUnknownCommandStatus(ByVal commandText As String)
    Dim suggestionText As String

    suggestionText = FAS_GetCommandHelpSuggestionText(commandText)
    If Len(suggestionText) > 0 Then
        FAS_SetStatus "No command or sheet matched '" & commandText & "'. " & suggestionText
    Else
        FAS_SetStatus "No command or sheet matched '" & commandText & "'. Try: help"
    End If
End Sub
