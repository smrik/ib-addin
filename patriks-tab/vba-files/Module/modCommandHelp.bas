Attribute VB_Name = "modCommandHelp"
Option Explicit

Public Sub FAS_ShowCommandHelp()
    FAS_ShowCommandPalette vbNullString
End Sub

Public Function FAS_TryRunCommandHelp(ByVal commandText As String) As Boolean
    Dim queryText As Variant

    If Not FAS_IsCommandHelpCommand(commandText, queryText) Then Exit Function

    FAS_ShowCommandPalette CStr(queryText)
    FAS_TryRunCommandHelp = True
End Function

Public Function FAS_IsCommandHelpCommand(ByVal commandText As String, Optional ByRef queryText As Variant) As Boolean
    Dim rawText As String
    Dim normalizedCommand As String

    rawText = Trim$(commandText)
    normalizedCommand = FAS_NormalizeCommandText(rawText)
    If Len(rawText) = 0 Then Exit Function

    If rawText = "?" Or normalizedCommand = "help" Or normalizedCommand = "commands" Or normalizedCommand = "cmd" Then
        queryText = vbNullString
        FAS_IsCommandHelpCommand = True
        Exit Function
    End If

    If Left$(rawText, 1) = "?" Then
        queryText = Trim$(Mid$(rawText, 2))
        FAS_IsCommandHelpCommand = True
        Exit Function
    End If

    queryText = FAS_CommandHelpPayload(rawText, "help")
    If Len(CStr(queryText)) > 0 Then
        FAS_IsCommandHelpCommand = True
        Exit Function
    End If

    queryText = FAS_CommandHelpPayload(rawText, "commands")
    If Len(CStr(queryText)) > 0 Then
        FAS_IsCommandHelpCommand = True
        Exit Function
    End If

    queryText = FAS_CommandHelpPayload(rawText, "cmd")
    If Len(CStr(queryText)) > 0 Then
        FAS_IsCommandHelpCommand = True
    End If
End Function

Public Function FAS_GetCommandHelpSuggestionText(ByVal queryText As String) As String
    Dim matches As Collection
    Dim item As Variant
    Dim resultText As String
    Dim index As Long

    Set matches = FAS_FilterCommandCatalog(queryText)
    For Each item In matches
        index = index + 1
        If Len(resultText) > 0 Then resultText = resultText & ", "
        resultText = resultText & CStr(item(0))
        If index >= 5 Then Exit For
    Next item

    If Len(resultText) > 0 Then
        FAS_GetCommandHelpSuggestionText = "Try: " & resultText
    End If
End Function

Public Function FAS_GetCommandCatalogCount() As Long
    Dim catalog As Collection

    Set catalog = FAS_GetCommandCatalog()
    FAS_GetCommandCatalogCount = catalog.Count
End Function

Public Function FAS_CountCommandCatalogMatches(ByVal queryText As String) As Long
    Dim matches As Collection

    Set matches = FAS_FilterCommandCatalog(queryText)
    FAS_CountCommandCatalogMatches = matches.Count
End Function

Public Function FAS_FilterCommandCatalog(ByVal queryText As String) As Collection
    Dim catalog As Collection
    Dim matches As Collection
    Dim normalizedQuery As String
    Dim item As Variant
    Dim matchScore As Long

    Set catalog = FAS_GetCommandCatalog()
    Set matches = New Collection
    normalizedQuery = FAS_NormalizeSheetText(queryText)

    If Len(normalizedQuery) = 0 Then
        For Each item In catalog
            matches.Add item
        Next item
    Else
        For matchScore = 3 To 1 Step -1
            For Each item In catalog
                If FAS_CommandCatalogItemMatchScore(item, normalizedQuery) = matchScore Then
                    matches.Add item
                End If
            Next item
        Next matchScore
    End If

    Set FAS_FilterCommandCatalog = matches
End Function

Public Function FAS_GetCommandCatalogShortcutList() As String
    Dim catalog As Collection
    Dim item As Variant
    Dim commandText As String
    Dim resultText As String

    Set catalog = FAS_GetCommandCatalog()
    For Each item In catalog
        commandText = CStr(item(0))
        If InStr(1, commandText, "<", vbBinaryCompare) = 0 Then
            commandText = FAS_NormalizeCommandText(commandText)
            If Len(commandText) > 0 Then
                If Len(resultText) > 0 Then resultText = resultText & ","
                resultText = resultText & commandText
            End If
        End If
    Next item

    FAS_GetCommandCatalogShortcutList = resultText
End Function

Public Function FAS_GetCommandCatalog() As Collection
    Dim catalog As Collection
    Set catalog = New Collection

    FAS_AddCommandHelp catalog, "help", "commands cmd ?", "command palette find search", "Show every available FAS command", "Help"

    FAS_AddCommandHelp catalog, "p", "p1 pl", "profit loss pnl sheet jump", "Jump to the primary P&L sheet", "Navigation"
    FAS_AddCommandHelp catalog, "p2", "pl2", "profit loss consolidated sheet jump", "Jump to P&L slot 2", "Navigation"
    FAS_AddCommandHelp catalog, "bs", "b bs1", "balance sheet jump", "Jump to the primary balance sheet", "Navigation"
    FAS_AddCommandHelp catalog, "bs2", "b2", "balance sheet jump", "Jump to balance-sheet slot 2", "Navigation"
    FAS_AddCommandHelp catalog, "t", "t1 tb bb", "trial balance opening balance sheet jump", "Jump to the primary TB or BB sheet", "Navigation"
    FAS_AddCommandHelp catalog, "t2", "tb2 bb2", "trial balance opening balance sheet jump", "Jump to TB or BB slot 2", "Navigation"
    FAS_AddCommandHelp catalog, "g", "g1 gl", "general ledger sheet jump", "Jump to the primary GL sheet", "Navigation"
    FAS_AddCommandHelp catalog, "g2", "gl2", "general ledger sheet jump", "Jump to GL slot 2", "Navigation"
    FAS_AddCommandHelp catalog, "m", "m1 map mapping", "classification mapping sheet jump", "Jump to the primary Mapping sheet", "Navigation"
    FAS_AddCommandHelp catalog, "m2", "map2 mapping2", "classification mapping sheet jump", "Jump to Mapping slot 2", "Navigation"
    FAS_AddCommandHelp catalog, "tabs", "", "workbook sheets list", "Choose a worksheet from a numbered list", "Navigation"

    FAS_AddCommandHelp catalog, "tb <search>", "bb acct account", "trial balance account find", "Find text or an account in the mapped TB or BB", "Search"
    FAS_AddCommandHelp catalog, "gl <search>", "", "general ledger find", "Find text in the mapped GL", "Search"
    FAS_AddCommandHelp catalog, "map <search>", "mapping class klas", "classification mapping find", "Find text in the mapped classification sheet", "Search"
    FAS_AddCommandHelp catalog, "dk <search>", "ledger ix idx", "project ledger index find", "Search the active project ledger index", "Search"
    FAS_AddCommandHelp catalog, "dktb <search>", "", "project trial balance index find", "Search the active project TB index", "Search"
    FAS_AddCommandHelp catalog, "dkmap <search>", "", "project mapping index find", "Search the active project mapping index", "Search"
    FAS_AddCommandHelp catalog, "dkdebug", "idxdebug", "project search diagnostics", "Show active project search diagnostics", "Search"

    FAS_AddCommandHelp catalog, "val", "values", "paste clipboard", "Paste values only", "Paste"
    FAS_AddCommandHelp catalog, "txt", "text", "paste clipboard", "Paste as plain text", "Paste"
    FAS_AddCommandHelp catalog, "visible", "vis copyvisible visiblecopy", "paste clipboard filtered hidden rows", "Copy visible cells only", "Paste"
    FAS_AddCommandHelp catalog, "linkcopy", "copylinks copylink", "formula link reference copy", "Copy the selection for pasting as links", "Paste"
    FAS_AddCommandHelp catalog, "linkpaste", "pastelinks pastelink", "formula link reference paste", "Paste the copied range as linked formulas", "Paste"
    FAS_AddCommandHelp catalog, "round", "roundformula wrapround", "formula wrap rounding", "Wrap selected formulas in ROUND using the saved digits", "Formulas"
    FAS_AddCommandHelp catalog, "round <digits>", "roundformula wrapround", "formula wrap rounding decimals", "Wrap selected formulas in ROUND with the given decimals", "Formulas"
    FAS_AddCommandHelp catalog, "iferror", "errorblank wrapiferror", "formula wrap error blank", "Wrap selected formulas in IFERROR returning blank", "Formulas"
    FAS_AddCommandHelp catalog, "delta", "deltasign", "insert symbol change variance", "Insert a delta sign", "Formatting"
    FAS_AddCommandHelp catalog, "autofit", "fitwidth width20 autofitwidth", "column width fit margin", "Auto-fit column width with a margin", "Formatting"
    FAS_AddCommandHelp catalog, "colorcapture", "capturecolor getcolor", "eyedropper color pick format", "Capture the active cell colour", "Formatting"
    FAS_AddCommandHelp catalog, "colorapply", "applycolor usecolor", "eyedropper color apply format", "Apply the captured colour to the selection", "Formatting"
    FAS_AddCommandHelp catalog, "settings", "options config", "settings preferences shortcuts configure", "Open the FAS settings dialog", "Settings"
    FAS_AddCommandHelp catalog, "rep", "replace", "formula replace selected", "Replace text in selected formulas", "Formulas"
    FAS_AddCommandHelp catalog, "fcopy", "formulacopy smartcopy", "formula copy destination", "Smart-copy selected formulas using a destination picker", "Formulas"
    FAS_AddCommandHelp catalog, "fcopy <cell>", "formulacopy smartcopy", "formula copy address destination", "Smart-copy selected formulas to an address", "Formulas"
    FAS_AddCommandHelp catalog, "sub", "subtotal", "formula total selected", "Insert SUBTOTAL formulas below the selection", "Formulas"
    FAS_AddCommandHelp catalog, "pre", "precedent precedents pre1", "formula upstream dependency trace", "Jump to direct precedent 1", "Precedents"
    FAS_AddCommandHelp catalog, "pre2", "precedent2", "formula upstream dependency trace", "Jump to direct precedent 2", "Precedents"
    FAS_AddCommandHelp catalog, "back", "", "formula precedent origin return", "Return to the last precedent origin", "Precedents"

    FAS_AddCommandHelp catalog, "sumifs", "buildsumifs ifssum", "formula criteria sum measure", "Build a SUMIFS formula from picked ranges", "IFS Tools"
    FAS_AddCommandHelp catalog, "countifs", "buildcountifs ifscount", "formula criteria count", "Build a COUNTIFS formula from picked ranges", "IFS Tools"
    FAS_AddCommandHelp catalog, "captureifs", "captureifspattern ifscapture", "formula pattern copy", "Capture a direct IFS formula for safe filling", "IFS Tools"
    FAS_AddCommandHelp catalog, "fillifs", "fillifspattern ifsfill", "formula pattern fill", "Fill blanks and matching direct IFS formulas", "IFS Tools"
    FAS_AddCommandHelp catalog, "switch sumifs", "switchsumifs switchifssum", "formula convert", "Switch selected formulas to SUMIFS", "IFS Tools"
    FAS_AddCommandHelp catalog, "switch countifs", "switchcountifs switchifscount", "formula convert", "Switch selected formulas to COUNTIFS", "IFS Tools"
    FAS_AddCommandHelp catalog, "switch averageifs", "switchaverageifs switchifsaverage", "formula convert", "Switch selected formulas to AVERAGEIFS", "IFS Tools"
    FAS_AddCommandHelp catalog, "switch minifs", "switchminifs switchifsmin", "formula convert", "Switch selected formulas to MINIFS", "IFS Tools"
    FAS_AddCommandHelp catalog, "switch maxifs", "switchmaxifs switchifsmax", "formula convert", "Switch selected formulas to MAXIFS", "IFS Tools"

    FAS_AddCommandHelp catalog, "format number", "number numberformat fmt1", "format number custom cycle", "Cycle through your saved number formats", "Formatting"
    FAS_AddCommandHelp catalog, "format percent", "percent percentage percentformat fmt5", "format percent custom cycle", "Cycle through your saved percent formats", "Formatting"
    FAS_AddCommandHelp catalog, "format currency", "currency currencyformat fmt4", "format currency custom cycle", "Cycle through your saved currency formats", "Formatting"
    FAS_AddCommandHelp catalog, "format settings", "formats formatsettings", "format settings custom", "Open FAS shortcut and format settings", "Setup"
    FAS_AddCommandHelp catalog, "shortcuts", "keys shortcut settings", "keyboard shortcuts settings", "Open FAS shortcut and format settings", "Setup"

    FAS_AddCommandHelp catalog, "group by", "groupby group groupbyreport", "report grouping sumifs editable", "Create an editable non-spill Group By report", "Group By"
    FAS_AddCommandHelp catalog, "group by count", "groupbycount groupcount", "report grouping countifs editable", "Create a Group By report with COUNTIFS", "Group By"
    FAS_AddCommandHelp catalog, "capex", "capexsummary", "report grouping capital expenditure", "Create the CAPEX Group By preset", "Group By"

    FAS_AddCommandHelp catalog, "bm", "bms bookmarks", "saved locations list", "List local bookmarks", "Bookmarks"
    FAS_AddCommandHelp catalog, "bm set <name>", "mark", "save bookmark active cell", "Save the active cell as a local bookmark", "Bookmarks"
    FAS_AddCommandHelp catalog, "bm <name>", "", "open bookmark jump", "Jump to a local bookmark", "Bookmarks"
    FAS_AddCommandHelp catalog, "bmreset", "resetbookmarks", "clear delete bookmarks", "Clear local bookmarks for this workbook", "Bookmarks"

    FAS_AddCommandHelp catalog, "chart fas", "chart format chart style chart clean chart deck", "graph house style", "Apply the complete FAS chart format", "Charts"
    FAS_AddCommandHelp catalog, "series fas", "series format labels fas line fas", "data labels markers line", "Format selected chart series and labels", "Charts"
    FAS_AddCommandHelp catalog, "legend bottom", "chart legend bottom", "chart legend position", "Move the chart legend to the bottom", "Charts"
    FAS_AddCommandHelp catalog, "legend right", "chart legend right", "chart legend position", "Move the chart legend to the right", "Charts"
    FAS_AddCommandHelp catalog, "legend off", "chart legend off", "chart legend remove", "Remove the chart legend", "Charts"
    FAS_AddCommandHelp catalog, "grid light", "gridlines light chart grid", "chart gridlines", "Apply light chart gridlines", "Charts"
    FAS_AddCommandHelp catalog, "grid off", "gridlines off chart grid", "chart gridlines remove", "Remove chart gridlines", "Charts"

    FAS_AddCommandHelp catalog, "box full", "box content", "slide layout outline", "Insert a full-content outline box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box half", "box 1/2", "slide layout outline", "Insert a half-width outline box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box third", "box 1/3", "slide layout outline", "Insert a one-third-width outline box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box 2/3", "box two thirds", "slide layout outline", "Insert a two-thirds-width outline box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box quarter", "box 1/4", "slide layout outline", "Insert a quarter-width outline box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box row", "box half row", "slide layout outline height", "Insert a half-height content box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box small", "box block", "slide layout outline", "Insert a third-width by half-height box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box title", "", "slide heading outline", "Insert a title-area outline box", "Slide boxes"
    FAS_AddCommandHelp catalog, "box subtitle", "", "slide subheading outline", "Insert a subtitle-area outline box", "Slide boxes"

    FAS_AddCommandHelp catalog, "size full", "size content chart full", "shape chart slide layout", "Resize the selected object to full content", "Slide sizing"
    FAS_AddCommandHelp catalog, "size half", "size 1/2 chart half", "shape chart slide layout", "Resize the selected object to half width", "Slide sizing"
    FAS_AddCommandHelp catalog, "size third", "size 1/3 chart third", "shape chart slide layout", "Resize the selected object to one-third width", "Slide sizing"
    FAS_AddCommandHelp catalog, "size 2/3", "size two thirds chart 2/3", "shape chart slide layout", "Resize the selected object to two-thirds width", "Slide sizing"
    FAS_AddCommandHelp catalog, "size quarter", "size 1/4 chart quarter", "shape chart slide layout", "Resize the selected object to quarter width", "Slide sizing"
    FAS_AddCommandHelp catalog, "size row", "size half row chart row", "shape chart slide layout height", "Resize the selected object to half content height", "Slide sizing"
    FAS_AddCommandHelp catalog, "size small", "size block chart small", "shape chart slide layout", "Resize to a third-width by half-height block", "Slide sizing"
    FAS_AddCommandHelp catalog, "layoutdebug", "boxdebug slidedebug faslayout", "slide sizing diagnostics", "Show calculated FAS slide dimensions", "Slide sizing"

    FAS_AddCommandHelp catalog, "div", "divider", "worksheet separator", "Insert a colored divider worksheet", "Tools"
    FAS_AddCommandHelp catalog, "drill", "dd", "mapping tb gl drilldown floating", "Create a floating drilldown panel", "Tools"
    FAS_AddCommandHelp catalog, "drillclear", "ddclear", "remove drilldown floating", "Remove the active sheet drilldown panel", "Tools"
    FAS_AddCommandHelp catalog, "fixobj", "fixshape float", "selected chart image shape placement", "Keep selected objects from moving with cells", "Objects"
    FAS_AddCommandHelp catalog, "shapefix", "nofloat fixshapes", "workbook chart image shape placement", "Keep all workbook objects from moving with cells", "Objects"
    FAS_AddCommandHelp catalog, "maps", "mapinfo mappings", "sheet assignments setup", "Show local sheet mappings", "Setup"
    FAS_AddCommandHelp catalog, "mapreset", "resetmaps", "sheet assignments clear", "Clear local sheet mappings", "Setup"

    Set FAS_GetCommandCatalog = catalog
End Function

Private Function FAS_CommandCatalogItemMatchScore(ByVal item As Variant, ByVal normalizedQuery As String) As Long
    Dim normalizedCommand As String
    Dim normalizedAliases As String
    Dim aliasToken As Variant
    Dim searchText As String

    normalizedCommand = FAS_NormalizeSheetText(CStr(item(0)))
    If StrComp(normalizedCommand, normalizedQuery, vbTextCompare) = 0 Then
        FAS_CommandCatalogItemMatchScore = 3
        Exit Function
    End If

    normalizedAliases = FAS_NormalizeSheetText(CStr(item(1)))
    For Each aliasToken In Split(normalizedAliases, " ")
        If Len(CStr(aliasToken)) > 0 And StrComp(CStr(aliasToken), normalizedQuery, vbTextCompare) = 0 Then
            FAS_CommandCatalogItemMatchScore = 2
            Exit Function
        End If
    Next aliasToken

    searchText = FAS_NormalizeSheetText(CStr(item(0)) & " " & CStr(item(1)) & " " & CStr(item(2)) & " " & CStr(item(3)) & " " & CStr(item(4)))
    If InStr(1, searchText, normalizedQuery, vbTextCompare) > 0 Then
        FAS_CommandCatalogItemMatchScore = 1
    End If
End Function

Private Function FAS_CommandHelpPayload(ByVal rawText As String, ByVal prefixText As String) As String
    Dim normalizedRaw As String
    Dim normalizedPrefix As String

    normalizedRaw = LCase$(Trim$(rawText))
    normalizedPrefix = LCase$(Trim$(prefixText)) & " "
    If Left$(normalizedRaw, Len(normalizedPrefix)) = normalizedPrefix Then
        FAS_CommandHelpPayload = Trim$(Mid$(rawText, Len(normalizedPrefix) + 1))
    End If
End Function

Private Sub FAS_AddCommandHelp(ByVal catalog As Collection, ByVal commandName As String, ByVal aliases As String, ByVal keywords As String, ByVal description As String, ByVal groupName As String)
    catalog.Add Array(commandName, aliases, keywords, description, groupName)
End Sub
