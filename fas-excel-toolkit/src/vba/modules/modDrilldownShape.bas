Attribute VB_Name = "modDrilldownShape"
Option Explicit

Private Const FAS_DRILLDOWN_SHAPE_NAME As String = "FAS_DrilldownHUD"
Private Const FAS_DRILLDOWN_MAX_TERMS As Long = 16
Private Const FAS_DRILLDOWN_MAX_MAPPING As Long = 10
Private Const FAS_DRILLDOWN_MAX_TB As Long = 15
Private Const FAS_DRILLDOWN_MAX_GL As Long = 15
Private Const FAS_DRILLDOWN_MAX_TRACE_LINES As Long = 8
Private Const FAS_DRILLDOWN_MAX_TRACE_DEPTH As Long = 5
Private Const FAS_DRILLDOWN_MAX_TRACE_NODES As Long = 220
Private Const FAS_DRILLDOWN_MAX_RANGE_SAMPLE As Long = 48
Private Const FAS_DRILLDOWN_MAX_SUMIF_SCAN As Long = 600

Public Sub FAS_ShowDrilldown()
    FAS_ShowDrilldownForSelection
End Sub

Public Function FAS_ShowDrilldownForSelection() As Boolean
    Dim targetWorkbook As Workbook
    Dim sourceCell As Range
    Dim terms As Collection
    Dim traceLines As Collection
    Dim visitedCells As Object
    Dim resultText As String
    Dim traceNodes As Long
    Dim traceTruncated As Long
    Dim unsupportedCount As Long
    Dim truncatedHits As Long

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a P&L cell before drilldown."
        Exit Function
    End If

    Set sourceCell = ActiveCell
    If sourceCell Is Nothing Then
        Set sourceCell = Selection.Cells(1, 1)
    End If

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        FAS_SetStatus "No active user workbook available."
        Exit Function
    End If

    Set terms = New Collection
    Set traceLines = New Collection
    Set visitedCells = CreateObject("Scripting.Dictionary")

    FAS_SeedDrilldownTermsFromRow sourceCell, terms
    FAS_WalkDrilldownPrecedents sourceCell, terms, traceLines, visitedCells, 0, traceNodes, traceTruncated, unsupportedCount

    If terms.Count = 0 Then
        FAS_SetStatus "No useful drilldown terms found from precedents."
        Exit Function
    End If

    resultText = FAS_BuildDrilldownText(targetWorkbook, sourceCell, terms, traceLines, traceNodes, traceTruncated, unsupportedCount, truncatedHits)
    FAS_RenderDrilldownShape sourceCell, resultText

    If truncatedHits > 0 Then
        FAS_SetStatus "Drilldown created. +" & truncatedHits & " more match(es) truncated."
    Else
        FAS_SetStatus "Drilldown created."
    End If

    FAS_ShowDrilldownForSelection = True
End Function

Public Function FAS_ClearDrilldownShape() As Boolean
    Dim targetSheet As Worksheet

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Set targetSheet = ActiveSheet
    FAS_DeleteDrilldownShape targetSheet
    FAS_ClearDrilldownShape = True
    FAS_SetStatus "Drilldown shape cleared."
End Function

Public Function FAS_CountDrilldownShapesOnActiveSheet() As Long
    Dim targetShape As Shape

    For Each targetShape In ActiveSheet.Shapes
        If StrComp(targetShape.Name, FAS_DRILLDOWN_SHAPE_NAME, vbTextCompare) = 0 Then
            FAS_CountDrilldownShapesOnActiveSheet = FAS_CountDrilldownShapesOnActiveSheet + 1
        End If
    Next targetShape
End Function

Public Function FAS_GetActiveDrilldownPlacement() As Long
    Dim targetShape As Shape

    On Error Resume Next
    Set targetShape = ActiveSheet.Shapes(FAS_DRILLDOWN_SHAPE_NAME)
    On Error GoTo 0

    If targetShape Is Nothing Then
        Exit Function
    End If

    FAS_GetActiveDrilldownPlacement = targetShape.Placement
End Function

Public Function FAS_GetActiveDrilldownText() As String
    Dim targetShape As Shape

    On Error Resume Next
    Set targetShape = ActiveSheet.Shapes(FAS_DRILLDOWN_SHAPE_NAME)
    On Error GoTo 0

    If targetShape Is Nothing Then
        Exit Function
    End If

    FAS_GetActiveDrilldownText = targetShape.TextFrame.Characters.Text
End Function

Private Sub FAS_SeedDrilldownTermsFromRow(ByVal sourceCell As Range, ByVal terms As Collection)
    Dim currentCell As Range
    Dim startColumn As Long
    Dim columnIndex As Long
    Dim cellText As String

    If sourceCell Is Nothing Then
        Exit Sub
    End If

    startColumn = sourceCell.Column - 8
    If startColumn < 1 Then
        startColumn = 1
    End If

    For columnIndex = sourceCell.Column - 1 To startColumn Step -1
        Set currentCell = sourceCell.Worksheet.Cells(sourceCell.Row, columnIndex)
        cellText = Trim$(CStr(currentCell.Text))
        If FAS_IsUsefulDrilldownText(cellText) Then
            FAS_AddDrilldownTerm terms, cellText
            FAS_AddAccountLikeTerms terms, cellText
        End If
    Next columnIndex

    cellText = Trim$(CStr(sourceCell.Text))
    If FAS_IsUsefulDrilldownText(cellText) Then
        FAS_AddDrilldownTerm terms, cellText
        FAS_AddAccountLikeTerms terms, cellText
    End If
End Sub

Private Sub FAS_WalkDrilldownPrecedents(ByVal sourceCell As Range, ByVal terms As Collection, ByVal traceLines As Collection, ByVal visitedCells As Object, ByVal depth As Long, ByRef traceNodes As Long, ByRef traceTruncated As Long, ByRef unsupportedCount As Long)
    Dim cellKey As String
    Dim directPrecedents As Collection
    Dim walkedCells As Collection
    Dim sampledCells As Collection
    Dim precedentItem As Variant
    Dim targetCell As Range
    Dim precedentRange As Range

    If sourceCell Is Nothing Then
        Exit Sub
    End If

    If depth > FAS_DRILLDOWN_MAX_TRACE_DEPTH Then
        traceTruncated = traceTruncated + 1
        Exit Sub
    End If

    If traceNodes >= FAS_DRILLDOWN_MAX_TRACE_NODES Then
        traceTruncated = traceTruncated + 1
        Exit Sub
    End If

    cellKey = FAS_BuildCellKey(sourceCell)
    If visitedCells.Exists(cellKey) Then
        Exit Sub
    End If
    visitedCells.Add cellKey, True

    traceNodes = traceNodes + 1
    FAS_AddCellTraceLine traceLines, sourceCell, depth
    FAS_AddCellValueTerms terms, sourceCell

    If Not sourceCell.HasFormula Then
        Exit Sub
    End If

    Set walkedCells = New Collection
    If FAS_TryResolveSumifMatchedCells(sourceCell, walkedCells) Then
        For Each precedentItem In walkedCells
            Set targetCell = precedentItem
            FAS_WalkDrilldownPrecedents targetCell, terms, traceLines, visitedCells, depth + 1, traceNodes, traceTruncated, unsupportedCount
        Next precedentItem
        Exit Sub
    End If

    Set directPrecedents = FAS_DrillParsePrecedentRanges(sourceCell, unsupportedCount)
    For Each precedentItem In directPrecedents
        Set precedentRange = precedentItem
        Set sampledCells = FAS_GetSampleCellsFromRange(precedentRange)
        For Each targetCell In sampledCells
            FAS_WalkDrilldownPrecedents targetCell, terms, traceLines, visitedCells, depth + 1, traceNodes, traceTruncated, unsupportedCount
        Next targetCell
    Next precedentItem
End Sub

Private Function FAS_DrillParsePrecedentRanges(ByVal sourceCell As Range, ByRef unsupportedCount As Long) As Collection
    Dim precedents As Collection
    Dim formulaText As String
    Dim crossSheetRegex As Object
    Dim sameSheetRegex As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim sheetName As String
    Dim addressText As String
    Dim matchStart As Long

    Set precedents = New Collection
    If sourceCell Is Nothing Then
        Set FAS_DrillParsePrecedentRanges = precedents
        Exit Function
    End If

    formulaText = FAS_DrillBlankQuotedStrings(sourceCell.Formula)
    If FAS_DrillFormulaHasUnsupportedReferences(formulaText) Then
        unsupportedCount = unsupportedCount + 1
        Set FAS_DrillParsePrecedentRanges = precedents
        Exit Function
    End If

    Set crossSheetRegex = CreateObject("VBScript.RegExp")
    crossSheetRegex.Global = True
    crossSheetRegex.IgnoreCase = True
    crossSheetRegex.Pattern = "('([^']|'')+'|[A-Za-z0-9_ .&-]+)!\$?[A-Z]{1,3}\$?[0-9]+(:\$?[A-Z]{1,3}\$?[0-9]+)?"

    Set matches = crossSheetRegex.Execute(formulaText)
    For Each matchItem In matches
        If InStr(1, matchItem.Value, "[", vbBinaryCompare) = 0 Then
            FAS_DrillSplitSheetReference CStr(matchItem.Value), sheetName, addressText
            FAS_DrillAddPrecedentIfValid precedents, sourceCell.Worksheet.Parent, sheetName, addressText
        End If
    Next matchItem

    formulaText = crossSheetRegex.Replace(formulaText, " ")

    Set sameSheetRegex = CreateObject("VBScript.RegExp")
    sameSheetRegex.Global = True
    sameSheetRegex.IgnoreCase = True
    sameSheetRegex.Pattern = "\$?[A-Z]{1,3}\$?[0-9]+(:\$?[A-Z]{1,3}\$?[0-9]+)?"

    Set matches = sameSheetRegex.Execute(formulaText)
    For Each matchItem In matches
        matchStart = CLng(matchItem.FirstIndex) + 1
        If FAS_DrillIsStandaloneA1Match(formulaText, matchStart, Len(matchItem.Value)) Then
            FAS_DrillAddPrecedentIfValid precedents, sourceCell.Worksheet.Parent, sourceCell.Worksheet.Name, CStr(matchItem.Value)
        End If
    Next matchItem

    Set FAS_DrillParsePrecedentRanges = precedents
End Function

Private Function FAS_TryResolveSumifMatchedCells(ByVal sourceCell As Range, ByRef matchedCells As Collection) As Boolean
    Dim formulaText As String
    Dim functionBody As String
    Dim arguments As Collection
    Dim functionName As String

    If sourceCell Is Nothing Or Not sourceCell.HasFormula Then
        Exit Function
    End If

    formulaText = Trim$(sourceCell.Formula)
    If Left$(formulaText, 1) <> "=" Then
        Exit Function
    End If

    functionBody = Mid$(formulaText, 2)
    functionName = UCase$(FAS_DrillGetLeadingFunctionName(functionBody))
    If functionName <> "SUMIF" And functionName <> "SUMIFS" Then
        Exit Function
    End If

    Set arguments = FAS_DrillExtractFunctionArguments(functionBody)
    If arguments Is Nothing Then
        Exit Function
    End If

    If functionName = "SUMIF" Then
        FAS_TryResolveSumifMatchedCells = FAS_DrillResolveSumif(sourceCell, arguments, matchedCells)
    Else
        FAS_TryResolveSumifMatchedCells = FAS_DrillResolveSumifs(sourceCell, arguments, matchedCells)
    End If
End Function

Private Function FAS_DrillResolveSumif(ByVal sourceCell As Range, ByVal arguments As Collection, ByRef matchedCells As Collection) As Boolean
    Dim criteriaRange As Range
    Dim sumRange As Range
    Dim criteriaText As String
    Dim criteriaToken As String

    If arguments.Count < 2 Or arguments.Count > 3 Then
        Exit Function
    End If

    If Not FAS_DrillTryResolveReferenceRange(sourceCell, CStr(arguments(1)), criteriaRange) Then
        Exit Function
    End If

    criteriaToken = CStr(arguments(2))
    criteriaText = FAS_DrillResolveCriteriaText(sourceCell, criteriaToken)

    If arguments.Count = 3 Then
        If Not FAS_DrillTryResolveReferenceRange(sourceCell, CStr(arguments(3)), sumRange) Then
            Exit Function
        End If
    Else
        Set sumRange = criteriaRange
    End If

    Set matchedCells = New Collection
    FAS_DrillCollectMatchedSumRangeCells criteriaRange, sumRange, criteriaText, matchedCells
    FAS_DrillResolveSumif = (matchedCells.Count > 0)
End Function

Private Function FAS_DrillResolveSumifs(ByVal sourceCell As Range, ByVal arguments As Collection, ByRef matchedCells As Collection) As Boolean
    Dim sumRange As Range
    Dim pairIndex As Long
    Dim criteriaPairs As Collection
    Dim criteriaRange As Range
    Dim criteriaText As String
    Dim criteriaPair As Variant

    If arguments.Count < 3 Then
        Exit Function
    End If

    If ((arguments.Count - 1) Mod 2) <> 0 Then
        Exit Function
    End If

    If Not FAS_DrillTryResolveReferenceRange(sourceCell, CStr(arguments(1)), sumRange) Then
        Exit Function
    End If

    Set criteriaPairs = New Collection
    For pairIndex = 2 To arguments.Count Step 2
        If Not FAS_DrillTryResolveReferenceRange(sourceCell, CStr(arguments(pairIndex)), criteriaRange) Then
            Exit Function
        End If

        criteriaText = FAS_DrillResolveCriteriaText(sourceCell, CStr(arguments(pairIndex + 1)))
        criteriaPairs.Add Array(criteriaRange, criteriaText)
    Next pairIndex

    Set matchedCells = New Collection
    FAS_DrillCollectMatchedSumifsCells sumRange, criteriaPairs, matchedCells
    FAS_DrillResolveSumifs = (matchedCells.Count > 0)
End Function

Private Sub FAS_DrillCollectMatchedSumRangeCells(ByVal criteriaRange As Range, ByVal sumRange As Range, ByVal criteriaText As String, ByVal matchedCells As Collection)
    Dim rowsCount As Long
    Dim columnsCount As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim scanned As Long
    Dim criteriaCell As Range
    Dim sumCell As Range

    rowsCount = criteriaRange.Rows.Count
    columnsCount = criteriaRange.Columns.Count

    For rowIndex = 1 To rowsCount
        For columnIndex = 1 To columnsCount
            scanned = scanned + 1
            If scanned > FAS_DRILLDOWN_MAX_SUMIF_SCAN Then
                Exit Sub
            End If

            Set criteriaCell = criteriaRange.Cells(rowIndex, columnIndex)
            If FAS_DrillCriteriaMatches(criteriaCell.Value2, criteriaText) Then
                Set sumCell = FAS_DrillGetAlignedCell(sumRange, rowIndex, columnIndex)
                If Not sumCell Is Nothing Then
                    FAS_DrillAddUniqueCell matchedCells, sumCell
                End If
            End If
        Next columnIndex
    Next rowIndex
End Sub

Private Sub FAS_DrillCollectMatchedSumifsCells(ByVal sumRange As Range, ByVal criteriaPairs As Collection, ByVal matchedCells As Collection)
    Dim primaryRange As Range
    Dim rowsCount As Long
    Dim columnsCount As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim pairItem As Variant
    Dim criteriaRange As Range
    Dim criteriaText As String
    Dim scanCount As Long
    Dim isMatch As Boolean
    Dim sumCell As Range

    If criteriaPairs.Count = 0 Then
        Exit Sub
    End If

    pairItem = criteriaPairs(1)
    Set primaryRange = pairItem(0)
    rowsCount = primaryRange.Rows.Count
    columnsCount = primaryRange.Columns.Count

    For rowIndex = 1 To rowsCount
        For columnIndex = 1 To columnsCount
            scanCount = scanCount + 1
            If scanCount > FAS_DRILLDOWN_MAX_SUMIF_SCAN Then
                Exit Sub
            End If

            isMatch = True
            For Each pairItem In criteriaPairs
                Set criteriaRange = pairItem(0)
                criteriaText = CStr(pairItem(1))
                If rowIndex > criteriaRange.Rows.Count Or columnIndex > criteriaRange.Columns.Count Then
                    isMatch = False
                    Exit For
                End If

                If Not FAS_DrillCriteriaMatches(criteriaRange.Cells(rowIndex, columnIndex).Value2, criteriaText) Then
                    isMatch = False
                    Exit For
                End If
            Next pairItem

            If isMatch Then
                Set sumCell = FAS_DrillGetAlignedCell(sumRange, rowIndex, columnIndex)
                If Not sumCell Is Nothing Then
                    FAS_DrillAddUniqueCell matchedCells, sumCell
                End If
            End If
        Next columnIndex
    Next rowIndex
End Sub

Private Function FAS_DrillGetAlignedCell(ByVal baseRange As Range, ByVal rowOffset As Long, ByVal columnOffset As Long) As Range
    If baseRange Is Nothing Then
        Exit Function
    End If

    If rowOffset < 1 Or columnOffset < 1 Then
        Exit Function
    End If

    If rowOffset > baseRange.Rows.Count Or columnOffset > baseRange.Columns.Count Then
        Exit Function
    End If

    Set FAS_DrillGetAlignedCell = baseRange.Cells(rowOffset, columnOffset)
End Function

Private Function FAS_DrillTryResolveReferenceRange(ByVal sourceCell As Range, ByVal referenceText As String, ByRef resolvedRange As Range) As Boolean
    Dim token As String
    Dim bangPosition As Long
    Dim sheetToken As String
    Dim addressToken As String
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet

    token = Trim$(referenceText)
    If Len(token) = 0 Then
        Exit Function
    End If

    If Left$(token, 1) = "(" And Right$(token, 1) = ")" Then
        token = Mid$(token, 2, Len(token) - 2)
    End If

    token = Replace(token, "$", "")
    Set targetWorkbook = sourceCell.Worksheet.Parent

    bangPosition = InStrRev(token, "!")
    If bangPosition > 0 Then
        sheetToken = Left$(token, bangPosition - 1)
        addressToken = Mid$(token, bangPosition + 1)
        sheetToken = FAS_DrillUnquoteSheetName(sheetToken)

        On Error Resume Next
        Set targetSheet = targetWorkbook.Worksheets(sheetToken)
        On Error GoTo 0
        If targetSheet Is Nothing Then
            Exit Function
        End If

        On Error Resume Next
        Set resolvedRange = targetSheet.Range(addressToken)
        On Error GoTo 0
    Else
        On Error Resume Next
        Set resolvedRange = sourceCell.Worksheet.Range(token)
        On Error GoTo 0
    End If

    FAS_DrillTryResolveReferenceRange = Not (resolvedRange Is Nothing)
End Function

Private Function FAS_DrillResolveCriteriaText(ByVal sourceCell As Range, ByVal criteriaToken As String) As String
    Dim token As String
    Dim criteriaCell As Range

    token = Trim$(criteriaToken)
    If Len(token) = 0 Then
        Exit Function
    End If

    If Left$(token, 1) = Chr$(34) And Right$(token, 1) = Chr$(34) Then
        token = Mid$(token, 2, Len(token) - 2)
        FAS_DrillResolveCriteriaText = Replace(token, Chr$(34) & Chr$(34), Chr$(34))
        Exit Function
    End If

    If FAS_DrillTryResolveReferenceRange(sourceCell, token, criteriaCell) Then
        FAS_DrillResolveCriteriaText = CStr(criteriaCell.Cells(1, 1).Value2)
    Else
        FAS_DrillResolveCriteriaText = token
    End If
End Function

Private Function FAS_DrillCriteriaMatches(ByVal valueToCheck As Variant, ByVal criteriaText As String) As Boolean
    Dim token As String
    Dim opText As String
    Dim rhsText As String
    Dim compareValue As Double
    Dim compareTarget As Double
    Dim wildcardPattern As String
    Dim normalizedLeft As String
    Dim normalizedRight As String

    token = Trim$(criteriaText)
    If Len(token) = 0 Then
        Exit Function
    End If

    opText = "="
    rhsText = token
    If Len(token) >= 2 Then
        Select Case Left$(token, 2)
            Case ">=", "<=", "<>"
                opText = Left$(token, 2)
                rhsText = Mid$(token, 3)
        End Select
    End If
    If opText = "=" And Len(token) >= 1 Then
        Select Case Left$(token, 1)
            Case "=", "<", ">"
                opText = Left$(token, 1)
                rhsText = Mid$(token, 2)
        End Select
    End If

    rhsText = Trim$(rhsText)
    normalizedLeft = CStr(valueToCheck)
    normalizedRight = rhsText

    If (InStr(1, rhsText, "*", vbBinaryCompare) > 0 Or InStr(1, rhsText, "?", vbBinaryCompare) > 0) And (opText = "=" Or opText = "<>") Then
        wildcardPattern = Replace(rhsText, "~", "")
        FAS_DrillCriteriaMatches = (normalizedLeft Like wildcardPattern)
        If opText = "<>" Then
            FAS_DrillCriteriaMatches = Not FAS_DrillCriteriaMatches
        End If
        Exit Function
    End If

    If IsNumeric(valueToCheck) And IsNumeric(rhsText) Then
        compareValue = CDbl(valueToCheck)
        compareTarget = CDbl(rhsText)

        Select Case opText
            Case "="
                FAS_DrillCriteriaMatches = (compareValue = compareTarget)
            Case "<>"
                FAS_DrillCriteriaMatches = (compareValue <> compareTarget)
            Case ">"
                FAS_DrillCriteriaMatches = (compareValue > compareTarget)
            Case "<"
                FAS_DrillCriteriaMatches = (compareValue < compareTarget)
            Case ">="
                FAS_DrillCriteriaMatches = (compareValue >= compareTarget)
            Case "<="
                FAS_DrillCriteriaMatches = (compareValue <= compareTarget)
        End Select
        Exit Function
    End If

    Select Case opText
        Case "="
            FAS_DrillCriteriaMatches = (StrComp(normalizedLeft, normalizedRight, vbTextCompare) = 0)
        Case "<>"
            FAS_DrillCriteriaMatches = (StrComp(normalizedLeft, normalizedRight, vbTextCompare) <> 0)
        Case Else
            FAS_DrillCriteriaMatches = False
    End Select
End Function

Private Function FAS_DrillGetLeadingFunctionName(ByVal formulaBody As String) As String
    Dim index As Long
    Dim currentChar As String
    Dim resultText As String

    formulaBody = Trim$(formulaBody)
    For index = 1 To Len(formulaBody)
        currentChar = Mid$(formulaBody, index, 1)
        If currentChar Like "[A-Za-z0-9_.]" Then
            resultText = resultText & currentChar
        ElseIf currentChar = "(" Then
            FAS_DrillGetLeadingFunctionName = resultText
            Exit Function
        Else
            Exit Function
        End If
    Next index
End Function

Private Function FAS_DrillExtractFunctionArguments(ByVal formulaBody As String) As Collection
    Dim openParen As Long
    Dim closeParen As Long
    Dim argsText As String

    openParen = InStr(1, formulaBody, "(", vbBinaryCompare)
    closeParen = InStrRev(formulaBody, ")", -1, vbBinaryCompare)
    If openParen = 0 Or closeParen <= openParen Then
        Exit Function
    End If

    If InStr(closeParen + 1, formulaBody, "+", vbBinaryCompare) > 0 Or _
        InStr(closeParen + 1, formulaBody, "-", vbBinaryCompare) > 0 Or _
        InStr(closeParen + 1, formulaBody, "*", vbBinaryCompare) > 0 Or _
        InStr(closeParen + 1, formulaBody, "/", vbBinaryCompare) > 0 Then
        Exit Function
    End If

    argsText = Mid$(formulaBody, openParen + 1, closeParen - openParen - 1)
    Set FAS_DrillExtractFunctionArguments = FAS_DrillSplitArguments(argsText)
End Function

Private Function FAS_DrillSplitArguments(ByVal argsText As String) As Collection
    Dim parts As Collection
    Dim index As Long
    Dim currentChar As String
    Dim inString As Boolean
    Dim depth As Long
    Dim currentToken As String

    Set parts = New Collection

    For index = 1 To Len(argsText)
        currentChar = Mid$(argsText, index, 1)
        If currentChar = Chr$(34) Then
            inString = Not inString
            currentToken = currentToken & currentChar
        ElseIf Not inString And currentChar = "(" Then
            depth = depth + 1
            currentToken = currentToken & currentChar
        ElseIf Not inString And currentChar = ")" Then
            depth = depth - 1
            currentToken = currentToken & currentChar
        ElseIf Not inString And depth = 0 And currentChar = "," Then
            parts.Add Trim$(currentToken)
            currentToken = vbNullString
        Else
            currentToken = currentToken & currentChar
        End If
    Next index

    If Len(Trim$(currentToken)) > 0 Or Len(argsText) > 0 Then
        parts.Add Trim$(currentToken)
    End If

    Set FAS_DrillSplitArguments = parts
End Function

Private Function FAS_GetSampleCellsFromRange(ByVal sourceRange As Range) As Collection
    Dim sampledCells As Collection
    Dim candidateCell As Range
    Dim sampledCount As Long

    Set sampledCells = New Collection
    If sourceRange Is Nothing Then
        Set FAS_GetSampleCellsFromRange = sampledCells
        Exit Function
    End If

    If sourceRange.Cells.CountLarge = 1 Then
        sampledCells.Add sourceRange.Cells(1, 1)
        Set FAS_GetSampleCellsFromRange = sampledCells
        Exit Function
    End If

    For Each candidateCell In sourceRange.Cells
        If candidateCell.HasFormula Or Len(Trim$(CStr(candidateCell.Text))) > 0 Then
            sampledCells.Add candidateCell
            sampledCount = sampledCount + 1
            If sampledCount >= FAS_DRILLDOWN_MAX_RANGE_SAMPLE Then
                Exit For
            End If
        End If
    Next candidateCell

    If sampledCells.Count = 0 Then
        sampledCells.Add sourceRange.Cells(1, 1)
    End If

    Set FAS_GetSampleCellsFromRange = sampledCells
End Function

Private Sub FAS_AddCellTraceLine(ByVal traceLines As Collection, ByVal sourceCell As Range, ByVal depth As Long)
    Dim prefixText As String
    Dim valueText As String

    If traceLines Is Nothing Or sourceCell Is Nothing Then
        Exit Sub
    End If

    If traceLines.Count >= FAS_DRILLDOWN_MAX_TRACE_LINES Then
        Exit Sub
    End If

    prefixText = String$(depth * 2, " ")
    valueText = Trim$(CStr(sourceCell.Text))
    If Len(valueText) > 44 Then
        valueText = Left$(valueText, 44) & "..."
    End If

    If Len(valueText) = 0 Then
        traceLines.Add prefixText & sourceCell.Worksheet.Name & "!" & sourceCell.Address(False, False)
    Else
        traceLines.Add prefixText & sourceCell.Worksheet.Name & "!" & sourceCell.Address(False, False) & " = " & valueText
    End If
End Sub

Private Sub FAS_AddCellValueTerms(ByVal terms As Collection, ByVal sourceCell As Range)
    Dim cellText As String

    If sourceCell Is Nothing Then
        Exit Sub
    End If

    cellText = Trim$(CStr(sourceCell.Text))
    If FAS_IsUsefulDrilldownText(cellText) Then
        FAS_AddDrilldownTerm terms, cellText
        FAS_AddAccountLikeTerms terms, cellText
    End If
End Sub

Private Function FAS_BuildDrilldownText(ByVal targetWorkbook As Workbook, ByVal sourceCell As Range, ByVal terms As Collection, ByVal traceLines As Collection, ByVal traceNodes As Long, ByVal traceTruncated As Long, ByVal unsupportedCount As Long, ByRef truncatedCount As Long) As String
    Dim resultText As String
    Dim termText As String
    Dim traceText As String
    Dim lineItem As Variant

    termText = FAS_JoinCollection(terms, ", ")
    traceText = vbNullString
    For Each lineItem In traceLines
        traceText = traceText & "  " & CStr(lineItem) & vbCrLf
    Next lineItem

    resultText = "FAS Drilldown - " & sourceCell.Worksheet.Name & "!" & sourceCell.Address(False, False) & vbCrLf & _
        "Terms: " & termText & vbCrLf & _
        "Trace nodes: " & CStr(traceNodes)
    If traceTruncated > 0 Then
        resultText = resultText & " (+" & CStr(traceTruncated) & " truncated)"
    End If
    If unsupportedCount > 0 Then
        resultText = resultText & " | unsupported refs: " & CStr(unsupportedCount)
    End If

    resultText = resultText & vbCrLf & String$(48, "-") & vbCrLf
    If Len(traceText) > 0 Then
        resultText = resultText & "Trace sample:" & vbCrLf & traceText & String$(48, "-") & vbCrLf
    End If

    resultText = resultText & FAS_BuildHitsSection(targetWorkbook, "MAPPING", "Mapping hits", terms, FAS_DRILLDOWN_MAX_MAPPING, truncatedCount)
    resultText = resultText & FAS_BuildHitsSection(targetWorkbook, "TB", "TB hits", terms, FAS_DRILLDOWN_MAX_TB, truncatedCount)
    resultText = resultText & FAS_BuildHitsSection(targetWorkbook, "GL", "GL hits", terms, FAS_DRILLDOWN_MAX_GL, truncatedCount)

    FAS_BuildDrilldownText = resultText
End Function

Private Function FAS_BuildHitsSection(ByVal targetWorkbook As Workbook, ByVal groupKey As String, ByVal sectionTitle As String, ByVal terms As Collection, ByVal maxHits As Long, ByRef truncatedCount As Long) As String
    Dim targetSheet As Worksheet
    Dim hits As Collection
    Dim hitItem As Variant
    Dim resultText As String
    Dim hitIndex As Long
    Dim sectionTruncated As Long

    Set targetSheet = FAS_GetResolvedGroupSheet(targetWorkbook, groupKey, 1)
    If targetSheet Is Nothing Then
        FAS_BuildHitsSection = sectionTitle & ": no mapped/detected sheet" & vbCrLf & vbCrLf
        Exit Function
    End If

    Set hits = FAS_FindDrilldownHits(targetSheet, terms, maxHits, sectionTruncated)
    truncatedCount = truncatedCount + sectionTruncated

    resultText = sectionTitle & " (" & targetSheet.Name & "):"
    If hits.Count = 0 Then
        FAS_BuildHitsSection = resultText & " none" & vbCrLf & vbCrLf
        Exit Function
    End If

    resultText = resultText & vbCrLf
    For hitIndex = 1 To hits.Count
        hitItem = hits(hitIndex)
        resultText = resultText & "  r" & CStr(hitItem(0)) & " [" & CStr(hitItem(1)) & "] " & CStr(hitItem(2)) & vbCrLf
    Next hitIndex

    If sectionTruncated > 0 Then
        resultText = resultText & "  + " & sectionTruncated & " more" & vbCrLf
    End If

    FAS_BuildHitsSection = resultText & vbCrLf
End Function

Private Function FAS_FindDrilldownHits(ByVal targetSheet As Worksheet, ByVal terms As Collection, ByVal maxHits As Long, ByRef truncatedCount As Long) As Collection
    Dim hits As Collection
    Dim seenRows As Object
    Dim searchRange As Range
    Dim termItem As Variant
    Dim foundCell As Range
    Dim firstAddress As String
    Dim rowKey As String

    Set hits = New Collection
    Set seenRows = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set searchRange = targetSheet.UsedRange
    On Error GoTo 0

    If searchRange Is Nothing Then
        Set FAS_FindDrilldownHits = hits
        Exit Function
    End If

    For Each termItem In terms
        If Len(Trim$(CStr(termItem))) > 0 Then
            Application.FindFormat.Clear
            Set foundCell = searchRange.Find( _
                What:=CStr(termItem), _
                After:=searchRange.Cells(searchRange.Cells.Count), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False, _
                SearchFormat:=False)

            If Not foundCell Is Nothing Then
                firstAddress = foundCell.Address(False, False)
                Do
                    rowKey = CStr(foundCell.Row)
                    If Not seenRows.Exists(rowKey) Then
                        seenRows.Add rowKey, True
                        If hits.Count < maxHits Then
                            hits.Add Array(foundCell.Row, CStr(termItem), FAS_GetRowSnippet(targetSheet, foundCell.Row))
                        Else
                            truncatedCount = truncatedCount + 1
                        End If
                    End If

                    Set foundCell = searchRange.FindNext(foundCell)
                    If foundCell Is Nothing Then
                        Exit Do
                    End If
                Loop While StrComp(foundCell.Address(False, False), firstAddress, vbTextCompare) <> 0
            End If
        End If
    Next termItem

    Set FAS_FindDrilldownHits = hits
End Function

Private Function FAS_GetRowSnippet(ByVal targetSheet As Worksheet, ByVal rowIndex As Long) As String
    Dim usedRange As Range
    Dim firstColumn As Long
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim valueText As String
    Dim resultText As String
    Dim itemCount As Long

    Set usedRange = targetSheet.UsedRange
    firstColumn = usedRange.Column
    lastColumn = usedRange.Column + usedRange.Columns.Count - 1
    If lastColumn > firstColumn + 8 Then
        lastColumn = firstColumn + 8
    End If

    For columnIndex = firstColumn To lastColumn
        valueText = Trim$(CStr(targetSheet.Cells(rowIndex, columnIndex).Text))
        If Len(valueText) > 0 Then
            If Len(resultText) > 0 Then
                resultText = resultText & " | "
            End If
            resultText = resultText & Left$(valueText, 28)
            itemCount = itemCount + 1
            If itemCount >= 5 Then
                Exit For
            End If
        End If
    Next columnIndex

    If Len(resultText) = 0 Then
        resultText = "(blank row snippet)"
    End If

    FAS_GetRowSnippet = Left$(resultText, 160)
End Function

Private Sub FAS_RenderDrilldownShape(ByVal sourceCell As Range, ByVal resultText As String)
    Dim targetSheet As Worksheet
    Dim targetShape As Shape
    Dim shapeLeft As Double
    Dim shapeTop As Double
    Dim shapeWidth As Double
    Dim shapeHeight As Double
    Dim lineCount As Long

    Set targetSheet = sourceCell.Worksheet
    FAS_DeleteDrilldownShape targetSheet

    shapeWidth = 520
    lineCount = UBound(Split(resultText, vbCrLf)) + 1
    shapeHeight = 72 + (lineCount * 12)
    If shapeHeight > 360 Then
        shapeHeight = 360
    End If

    shapeLeft = sourceCell.Left
    shapeTop = sourceCell.Top - shapeHeight - 8
    If shapeTop < 5 Then
        shapeTop = sourceCell.Top + sourceCell.Height + 8
    End If

    Set targetShape = targetSheet.Shapes.AddShape(msoShapeRoundedRectangle, shapeLeft, shapeTop, shapeWidth, shapeHeight)
    targetShape.Name = FAS_DRILLDOWN_SHAPE_NAME
    targetShape.Placement = xlFreeFloating
    targetShape.Fill.ForeColor.RGB = RGB(255, 255, 230)
    targetShape.Line.ForeColor.RGB = RGB(120, 120, 90)
    targetShape.Line.Weight = 1
    targetShape.TextFrame.Characters.Text = resultText
    targetShape.TextFrame.MarginLeft = 8
    targetShape.TextFrame.MarginRight = 8
    targetShape.TextFrame.MarginTop = 6
    targetShape.TextFrame.MarginBottom = 6
    targetShape.TextFrame.AutoSize = False
    targetShape.TextFrame.Characters.Font.Name = "Arial"
    targetShape.TextFrame.Characters.Font.Size = 8
    targetShape.TextFrame.Characters.Font.Color = RGB(40, 40, 40)
    targetShape.ZOrder msoBringToFront
End Sub

Private Sub FAS_DeleteDrilldownShape(ByVal targetSheet As Worksheet)
    On Error Resume Next
    targetSheet.Shapes(FAS_DRILLDOWN_SHAPE_NAME).Delete
    On Error GoTo 0
End Sub

Private Function FAS_IsUsefulDrilldownText(ByVal valueText As String) As Boolean
    valueText = Trim$(valueText)
    If Len(valueText) < 3 Then
        Exit Function
    End If

    If IsNumeric(valueText) Then
        Exit Function
    End If

    FAS_IsUsefulDrilldownText = True
End Function

Private Sub FAS_AddDrilldownTerm(ByVal terms As Collection, ByVal termText As String)
    Dim item As Variant

    termText = Trim$(termText)
    If Len(termText) < 3 Then
        Exit Sub
    End If

    For Each item In terms
        If StrComp(CStr(item), termText, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next item

    terms.Add termText
    If terms.Count > FAS_DRILLDOWN_MAX_TERMS Then
        terms.Remove terms.Count
    End If
End Sub

Private Sub FAS_AddAccountLikeTerms(ByVal terms As Collection, ByVal sourceText As String)
    Dim regex As Object
    Dim matches As Object
    Dim matchItem As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[0-9][0-9A-Za-z._/\-]{2,}"

    Set matches = regex.Execute(sourceText)
    For Each matchItem In matches
        FAS_AddDrilldownTerm terms, CStr(matchItem.Value)
    Next matchItem
End Sub

Private Function FAS_JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim resultText As String
    Dim item As Variant

    For Each item In items
        If Len(resultText) > 0 Then
            resultText = resultText & delimiter
        End If
        resultText = resultText & CStr(item)
    Next item

    FAS_JoinCollection = resultText
End Function

Private Function FAS_BuildCellKey(ByVal targetCell As Range) As String
    FAS_BuildCellKey = LCase$(targetCell.Worksheet.Parent.FullName) & "|" & LCase$(targetCell.Worksheet.Name) & "|" & targetCell.Address(False, False)
End Function

Private Sub FAS_DrillAddUniqueCell(ByVal cells As Collection, ByVal targetCell As Range)
    Dim cellItem As Variant
    Dim keyToAdd As String
    Dim itemKey As String

    keyToAdd = FAS_BuildCellKey(targetCell)
    For Each cellItem In cells
        itemKey = FAS_BuildCellKey(cellItem)
        If StrComp(itemKey, keyToAdd, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next cellItem

    cells.Add targetCell
End Sub

Private Function FAS_DrillBlankQuotedStrings(ByVal formulaText As String) As String
    Dim index As Long
    Dim currentChar As String
    Dim inString As Boolean
    Dim result As String

    For index = 1 To Len(formulaText)
        currentChar = Mid$(formulaText, index, 1)
        If currentChar = Chr$(34) Then
            inString = Not inString
            result = result & " "
        ElseIf inString Then
            result = result & " "
        Else
            result = result & currentChar
        End If
    Next index

    FAS_DrillBlankQuotedStrings = result
End Function

Private Function FAS_DrillFormulaHasUnsupportedReferences(ByVal formulaText As String) As Boolean
    FAS_DrillFormulaHasUnsupportedReferences = (InStr(1, formulaText, "[", vbBinaryCompare) > 0 Or _
        InStr(1, formulaText, "INDIRECT(", vbTextCompare) > 0 Or _
        InStr(1, formulaText, "OFFSET(", vbTextCompare) > 0 Or _
        InStr(1, formulaText, "#", vbBinaryCompare) > 0)
End Function

Private Sub FAS_DrillSplitSheetReference(ByVal referenceText As String, ByRef sheetName As String, ByRef addressText As String)
    Dim bangPosition As Long

    bangPosition = InStrRev(referenceText, "!")
    sheetName = Left$(referenceText, bangPosition - 1)
    addressText = Mid$(referenceText, bangPosition + 1)
    sheetName = FAS_DrillUnquoteSheetName(sheetName)
End Sub

Private Function FAS_DrillUnquoteSheetName(ByVal sheetName As String) As String
    sheetName = Trim$(sheetName)
    If Left$(sheetName, 1) = "'" And Right$(sheetName, 1) = "'" Then
        sheetName = Mid$(sheetName, 2, Len(sheetName) - 2)
        sheetName = Replace(sheetName, "''", "'")
    End If

    FAS_DrillUnquoteSheetName = sheetName
End Function

Private Sub FAS_DrillAddPrecedentIfValid(ByVal precedents As Collection, ByVal targetWorkbook As Workbook, ByVal sheetName As String, ByVal addressText As String)
    Dim targetSheet As Worksheet
    Dim targetRange As Range
    Dim uniqueKey As String
    Dim item As Variant

    If targetWorkbook Is Nothing Then
        Exit Sub
    End If

    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets(sheetName)
    Set targetRange = targetSheet.Range(addressText)
    On Error GoTo 0

    If targetRange Is Nothing Then
        Exit Sub
    End If

    uniqueKey = targetSheet.Name & "!" & targetRange.Address(False, False)
    For Each item In precedents
        If StrComp(item.Worksheet.Name & "!" & item.Address(False, False), uniqueKey, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next item

    precedents.Add targetRange
End Sub

Private Function FAS_DrillIsStandaloneA1Match(ByVal formulaText As String, ByVal startPosition As Long, ByVal matchLength As Long) As Boolean
    Dim previousChar As String
    Dim nextChar As String

    If startPosition > 1 Then
        previousChar = Mid$(formulaText, startPosition - 1, 1)
        If previousChar Like "[A-Za-z0-9_!.]" Then
            Exit Function
        End If
    End If

    If startPosition + matchLength <= Len(formulaText) Then
        nextChar = Mid$(formulaText, startPosition + matchLength, 1)
        If nextChar Like "[A-Za-z0-9_]" Then
            Exit Function
        End If
    End If

    FAS_DrillIsStandaloneA1Match = True
End Function
