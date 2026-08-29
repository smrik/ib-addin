Attribute VB_Name = "modProjectSearch"
Option Explicit

Private Const FAS_PROJECT_SEARCH_SHAPE_NAME As String = "FAS_ProjectSearchHUD"
Private Const FAS_PROJECT_SEARCH_TIMEOUT_SECONDS As Double = 20#

Public Sub FAS_SearchLedgerIndex()
    Dim searchText As String

    searchText = InputBox("Search the active FAS project ledger index:", "FAS Ledger Search")
    If Len(Trim$(searchText)) = 0 Then
        Exit Sub
    End If

    FAS_ShowProjectSearchResults "ledger", searchText
End Sub

Public Function FAS_TryParseProjectSearchCommand(ByVal commandText As String, ByRef kind As String, ByRef searchText As String) As Boolean
    Dim rawText As String
    Dim normalizedText As String

    rawText = Trim$(commandText)
    normalizedText = FAS_NormalizeSheetText(rawText)
    If Len(normalizedText) = 0 Then
        Exit Function
    End If

    If FAS_ProjectSearchHasPrefix(normalizedText, "dk") Then
        kind = "ledger"
        searchText = FAS_ProjectSearchTextAfterPrefix(rawText, "dk")
    ElseIf FAS_ProjectSearchHasPrefix(normalizedText, "ledger") Then
        kind = "ledger"
        searchText = FAS_ProjectSearchTextAfterPrefix(rawText, "ledger")
    ElseIf FAS_ProjectSearchHasPrefix(normalizedText, "idx") Then
        kind = "ledger"
        searchText = FAS_ProjectSearchTextAfterPrefix(rawText, "idx")
    ElseIf FAS_ProjectSearchHasPrefix(normalizedText, "ix") Then
        kind = "ledger"
        searchText = FAS_ProjectSearchTextAfterPrefix(rawText, "ix")
    ElseIf FAS_ProjectSearchHasPrefix(normalizedText, "dktb") Then
        kind = "tb"
        searchText = FAS_ProjectSearchTextAfterPrefix(rawText, "dktb")
    ElseIf FAS_ProjectSearchHasPrefix(normalizedText, "dkmap") Then
        kind = "mapping"
        searchText = FAS_ProjectSearchTextAfterPrefix(rawText, "dkmap")
    End If

    FAS_TryParseProjectSearchCommand = (Len(kind) > 0 And Len(Trim$(searchText)) > 0)
End Function

Private Function FAS_ProjectSearchHasPrefix(ByVal normalizedText As String, ByVal prefixText As String) As Boolean
    Dim expectedPrefix As String

    expectedPrefix = FAS_NormalizeSheetText(prefixText) & " "
    FAS_ProjectSearchHasPrefix = (Left$(normalizedText, Len(expectedPrefix)) = expectedPrefix)
End Function

Public Function FAS_RunProjectSearch(ByVal kind As String, ByVal searchText As String) As Boolean
    FAS_RunProjectSearch = FAS_ShowProjectSearchResults(kind, searchText)
End Function

Public Function FAS_ShowProjectSearchDebug() As Boolean
    Dim projectFolder As String
    Dim databasePath As String
    Dim duckDbPath As String
    Dim fileSystem As Object
    Dim debugText As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    projectFolder = FAS_GetActiveProjectFolder()
    If Len(projectFolder) > 0 Then
        databasePath = fileSystem.BuildPath(fileSystem.BuildPath(projectFolder, "index"), "fas_index.duckdb")
        duckDbPath = FAS_GetDuckDbExecutablePath(projectFolder)
    End If

    debugText = "Active project folder: " & IIf(Len(projectFolder) = 0, "(not found)", projectFolder) & vbCrLf & _
        "Index DB: " & IIf(Len(databasePath) = 0, "(not found)", databasePath) & vbCrLf & _
        "Index DB exists: " & CStr(Len(databasePath) > 0 And fileSystem.FileExists(databasePath)) & vbCrLf & _
        "DuckDB path: " & IIf(Len(duckDbPath) = 0, "(not found)", duckDbPath) & vbCrLf & _
        "DuckDB exists: " & CStr(Len(duckDbPath) > 0 And fileSystem.FileExists(duckDbPath)) & vbCrLf

    FAS_RenderProjectSearchShape "debug", "project search", debugText
    FAS_SetStatus "Project search debug shown."
    FAS_ShowProjectSearchDebug = True
End Function

Private Function FAS_ShowProjectSearchResults(ByVal kind As String, ByVal searchText As String) As Boolean
    Dim scriptPath As String
    Dim commandText As String
    Dim resultText As String

    searchText = Trim$(searchText)
    If Len(searchText) = 0 Then
        Exit Function
    End If

    FAS_SetStatus "Searching " & kind & " index for '" & searchText & "'..."
    commandText = FAS_BuildDirectDuckDbSearchCommand(kind, searchText)
    If Len(commandText) = 0 Then
        FAS_SetStatus "Searching " & kind & " index via PowerShell fallback..."
        scriptPath = FAS_GetProjectQueryScriptPath()
        If Len(scriptPath) = 0 Then
            FAS_SetStatus "Project index or search script not found."
            Exit Function
        End If

        commandText = "powershell -NoProfile -ExecutionPolicy Bypass -File " & FAS_QuoteShellArg(scriptPath) & _
            " -Kind " & FAS_QuoteShellArg(kind) & _
            " -Query " & FAS_QuoteShellArg(searchText) & _
            " -Limit 15"
    Else
        FAS_SetStatus "Searching " & kind & " index via DuckDB..."
    End If

    resultText = FAS_RunShellAndCapture(commandText, FAS_PROJECT_SEARCH_TIMEOUT_SECONDS)
    If Len(Trim$(resultText)) = 0 Then
        FAS_SetStatus "No output from project search."
        Exit Function
    End If

    FAS_RenderProjectSearchShape kind, searchText, resultText
    FAS_SetStatus "Project search complete."
    FAS_ShowProjectSearchResults = True
End Function

Private Function FAS_BuildDirectDuckDbSearchCommand(ByVal kind As String, ByVal searchText As String) As String
    Dim projectFolder As String
    Dim databasePath As String
    Dim duckDbPath As String
    Dim sqlText As String
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    projectFolder = FAS_GetActiveProjectFolder()
    If Len(projectFolder) = 0 Then
        Exit Function
    End If

    databasePath = fileSystem.BuildPath(fileSystem.BuildPath(projectFolder, "index"), "fas_index.duckdb")
    If Not fileSystem.FileExists(databasePath) Then
        Exit Function
    End If

    duckDbPath = FAS_GetDuckDbExecutablePath(projectFolder)
    If Len(duckDbPath) = 0 Then
        Exit Function
    End If

    sqlText = FAS_BuildProjectSearchSql(kind, searchText, 15)
    If Len(sqlText) = 0 Then
        Exit Function
    End If

    FAS_BuildDirectDuckDbSearchCommand = FAS_QuoteShellArg(duckDbPath) & " -csv " & _
        FAS_QuoteShellArg(databasePath) & " -c " & FAS_QuoteShellArg(sqlText)
End Function

Private Function FAS_GetActiveProjectFolder() As String
    Dim fileSystem As Object
    Dim projectsPath As String
    Dim jsonText As String
    Dim activeProjectName As String
    Dim projectFolder As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    projectsPath = Environ$("APPDATA") & "\FAS_Toolkit\Projects\projects.json"
    If Not fileSystem.FileExists(projectsPath) Then
        Exit Function
    End If

    jsonText = FAS_ReadTextFile(projectsPath)
    activeProjectName = FAS_GetJsonStringProperty(jsonText, "activeProject")
    If Len(activeProjectName) = 0 Then
        Exit Function
    End If

    projectFolder = Environ$("APPDATA") & "\FAS_Toolkit\Projects\" & FAS_SanitizeProjectFolderName(activeProjectName)
    If fileSystem.FolderExists(projectFolder) Then
        FAS_GetActiveProjectFolder = projectFolder
    End If
End Function

Private Function FAS_GetDuckDbExecutablePath(ByVal projectFolder As String) As String
    Dim fileSystem As Object
    Dim rememberedPath As String
    Dim rememberedPathFile As String
    Dim appDataPath As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    rememberedPathFile = fileSystem.BuildPath(fileSystem.BuildPath(projectFolder, "index"), "duckdb_path.txt")
    If fileSystem.FileExists(rememberedPathFile) Then
        rememberedPath = Trim$(FAS_ReadTextFile(rememberedPathFile))
        If Len(rememberedPath) > 0 And fileSystem.FileExists(rememberedPath) Then
            FAS_GetDuckDbExecutablePath = rememberedPath
            Exit Function
        End If
    End If

    appDataPath = Environ$("APPDATA") & "\FAS_Toolkit\Tools\duckdb.exe"
    If fileSystem.FileExists(appDataPath) Then
        FAS_GetDuckDbExecutablePath = appDataPath
        Exit Function
    End If
End Function

Private Function FAS_BuildProjectSearchSql(ByVal kind As String, ByVal searchText As String, ByVal limitRows As Long) As String
    Dim tokens As Variant
    Dim token As Variant
    Dim tableName As String
    Dim whereSql As String
    Dim escapedToken As String

    Select Case LCase$(Trim$(kind))
        Case "ledger", "tb", "mapping"
            tableName = LCase$(Trim$(kind))
        Case Else
            Exit Function
    End Select

    tokens = Split(Trim$(searchText), " ")
    For Each token In tokens
        If Len(Trim$(CStr(token))) > 0 Then
            escapedToken = FAS_EscapeSqlLiteral(LCase$(Trim$(CStr(token))))
            If Len(whereSql) > 0 Then
                whereSql = whereSql & " AND "
            End If
            whereSql = whereSql & "lower(coalesce(raw_text, '') || ' ' || coalesce(account, '') || ' ' || coalesce(description, '') || ' ' || coalesce(classification, '')) LIKE '%" & escapedToken & "%'"
        End If
    Next token

    If Len(whereSql) = 0 Then
        Exit Function
    End If

    FAS_BuildProjectSearchSql = _
        "SELECT company, year, source_file, source_sheet, source_row, posting_date, account, account_name, description, debit, credit, amount, classification, substr(raw_text, 1, 220) AS snippet " & _
        "FROM " & tableName & " WHERE " & whereSql & " LIMIT " & CStr(limitRows) & ";"
End Function

Private Function FAS_ProjectSearchTextAfterPrefix(ByVal rawText As String, ByVal prefixText As String) As String
    Dim result As String
    Dim firstChar As String

    result = Trim$(Mid$(rawText, Len(prefixText) + 1))
    Do While Len(result) > 0
        firstChar = Left$(result, 1)
        If firstChar <> " " And firstChar <> ":" And firstChar <> ";" And firstChar <> "=" Then
            Exit Do
        End If
        result = Trim$(Mid$(result, 2))
    Loop

    FAS_ProjectSearchTextAfterPrefix = result
End Function

Private Function FAS_GetProjectQueryScriptPath() As String
    Dim fileSystem As Object
    Dim candidates As Collection
    Dim workbookPath As String
    Dim item As Variant
    Dim envRepo As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set candidates = New Collection

    envRepo = Environ$("FAS_TOOLKIT_REPO")
    If Len(envRepo) > 0 Then
        candidates.Add fileSystem.BuildPath(envRepo, "tools\query_search_index.ps1")
    End If

    workbookPath = ThisWorkbook.Path
    If Len(workbookPath) > 0 Then
        candidates.Add fileSystem.BuildPath(workbookPath, "tools\query_search_index.ps1")
        candidates.Add fileSystem.BuildPath(fileSystem.GetParentFolderName(workbookPath), "tools\query_search_index.ps1")
    End If

    candidates.Add Environ$("APPDATA") & "\FAS_Toolkit\Tools\query_search_index.ps1"

    For Each item In candidates
        If fileSystem.FileExists(CStr(item)) Then
            FAS_GetProjectQueryScriptPath = CStr(item)
            Exit Function
        End If
    Next item
End Function

Private Function FAS_RunShellAndCapture(ByVal commandText As String, ByVal timeoutSeconds As Double) As String
    Dim shell As Object
    Dim execObject As Object
    Dim startTime As Double
    Dim outputText As String
    Dim errorText As String

    Set shell = CreateObject("WScript.Shell")
    Set execObject = shell.Exec(commandText)
    startTime = Timer

    Do While execObject.Status = 0
        DoEvents
        If FAS_ElapsedSeconds(startTime) > timeoutSeconds Then
            FAS_RunShellAndCapture = "Project search timed out after " & CStr(CLng(timeoutSeconds)) & " seconds."
            Exit Function
        End If
    Loop

    outputText = execObject.StdOut.ReadAll
    errorText = execObject.StdErr.ReadAll

    If Len(Trim$(errorText)) > 0 Then
        If Len(Trim$(outputText)) > 0 Then
            FAS_RunShellAndCapture = outputText & vbCrLf & "Errors:" & vbCrLf & errorText
        Else
            FAS_RunShellAndCapture = errorText
        End If
    Else
        FAS_RunShellAndCapture = outputText
    End If
End Function

Private Function FAS_ElapsedSeconds(ByVal startTime As Double) As Double
    If Timer >= startTime Then
        FAS_ElapsedSeconds = Timer - startTime
    Else
        FAS_ElapsedSeconds = (86400# - startTime) + Timer
    End If
End Function

Private Function FAS_QuoteShellArg(ByVal valueText As String) As String
    FAS_QuoteShellArg = Chr$(34) & Replace(valueText, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function FAS_EscapeSqlLiteral(ByVal valueText As String) As String
    FAS_EscapeSqlLiteral = Replace(valueText, "'", "''")
End Function

Private Function FAS_ReadTextFile(ByVal filePath As String) As String
    Dim fileSystem As Object
    Dim textStream As Object
    Dim stream As Object

    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    If Not stream Is Nothing Then
        stream.Type = 2
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile filePath
        FAS_ReadTextFile = stream.ReadText
        stream.Close
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textStream = fileSystem.OpenTextFile(filePath, 1, False, -2)
    FAS_ReadTextFile = textStream.ReadAll
    textStream.Close
End Function

Private Function FAS_GetJsonStringProperty(ByVal jsonText As String, ByVal propertyName As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim patternText As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    patternText = Chr$(34) & propertyName & Chr$(34) & "\s*:\s*" & Chr$(34) & "([^" & Chr$(34) & "]*)" & Chr$(34)
    regex.Pattern = patternText

    Set matches = regex.Execute(jsonText)
    If matches.Count > 0 Then
        FAS_GetJsonStringProperty = Replace(CStr(matches(0).SubMatches(0)), "\" & Chr$(34), Chr$(34))
    End If
End Function

Private Function FAS_SanitizeProjectFolderName(ByVal valueText As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each item In invalidChars
        valueText = Replace(valueText, CStr(item), "_")
    Next item

    If Len(valueText) = 0 Then
        valueText = "Project"
    End If

    FAS_SanitizeProjectFolderName = valueText
End Function

Private Sub FAS_RenderProjectSearchShape(ByVal kind As String, ByVal searchText As String, ByVal resultText As String)
    Dim sourceCell As Range
    Dim targetSheet As Worksheet
    Dim targetShape As Shape
    Dim shapeLeft As Double
    Dim shapeTop As Double
    Dim shapeWidth As Double
    Dim shapeHeight As Double
    Dim displayText As String
    Dim lineCount As Long

    If TypeName(Selection) = "Range" Then
        Set sourceCell = ActiveCell
    End If
    If sourceCell Is Nothing Then
        Set sourceCell = ActiveSheet.Range("B2")
    End If

    Set targetSheet = sourceCell.Worksheet
    FAS_DeleteProjectSearchShape targetSheet

    displayText = "FAS " & UCase$(kind) & " Search - " & searchText & vbCrLf & String$(64, "-") & vbCrLf & resultText
    If Len(displayText) > 7000 Then
        displayText = Left$(displayText, 7000) & vbCrLf & "... truncated"
    End If

    shapeWidth = 680
    lineCount = UBound(Split(displayText, vbCrLf)) + 1
    shapeHeight = 80 + (lineCount * 12)
    If shapeHeight > 430 Then
        shapeHeight = 430
    End If

    shapeLeft = sourceCell.Left
    shapeTop = sourceCell.Top - shapeHeight - 8
    If shapeTop < 5 Then
        shapeTop = sourceCell.Top + sourceCell.Height + 8
    End If

    Set targetShape = targetSheet.Shapes.AddShape(msoShapeRoundedRectangle, shapeLeft, shapeTop, shapeWidth, shapeHeight)
    targetShape.Name = FAS_PROJECT_SEARCH_SHAPE_NAME
    targetShape.Placement = xlFreeFloating
    targetShape.Fill.ForeColor.RGB = RGB(235, 246, 255)
    targetShape.Line.ForeColor.RGB = RGB(90, 120, 145)
    targetShape.Line.Weight = 1
    targetShape.TextFrame.Characters.Text = displayText
    targetShape.TextFrame.MarginLeft = 8
    targetShape.TextFrame.MarginRight = 8
    targetShape.TextFrame.MarginTop = 6
    targetShape.TextFrame.MarginBottom = 6
    targetShape.TextFrame.AutoSize = False
    targetShape.TextFrame.Characters.Font.Name = "Arial"
    targetShape.TextFrame.Characters.Font.Size = 8
    targetShape.TextFrame.Characters.Font.Color = RGB(35, 45, 55)
    targetShape.ZOrder msoBringToFront
End Sub

Private Sub FAS_DeleteProjectSearchShape(ByVal targetSheet As Worksheet)
    On Error Resume Next
    targetSheet.Shapes(FAS_PROJECT_SEARCH_SHAPE_NAME).Delete
    On Error GoTo 0
End Sub
