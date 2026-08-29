Attribute VB_Name = "modGroupByReport"
Option Explicit

Private Const FAS_GROUPBY_KEY_SEPARATOR As String = "||"
Private Const FAS_GROUPBY_BLANK_KEY As String = "{blank}"

' Ordinary formulas recalculate when source values change. New categories are
' structural, so rerun the command when the source grouping itself changes.
Public Sub FAS_CreateGroupByReport()
    FAS_CreateGroupByReportCore False
End Sub

Public Sub FAS_CreateGroupByReportWithCount()
    FAS_CreateGroupByReportCore True
End Sub

Public Sub FAS_CreateCapexSummaryFromGroupBy()
    Dim sourceBlock As Range
    Dim sourceSheet As Worksheet
    Dim headerRow As Long
    Dim lastDataRow As Long
    Dim classOneColumn As Long
    Dim classTwoColumn As Long
    Dim descriptionColumn As Long
    Dim acquisitionCostColumn As Long
    Dim netBookValueColumn As Long
    Dim rowFieldRanges As Collection
    Dim valueRanges As Collection

    Set sourceBlock = FAS_GB_GetCapexSourceBlock()
    If sourceBlock Is Nothing Then Exit Sub
    If Not FAS_GB_FindCapexHeaders(sourceBlock, headerRow, classOneColumn, classTwoColumn, _
        descriptionColumn, acquisitionCostColumn, netBookValueColumn) Then
        FAS_SetStatus "CAPEX Group By needs Class I, Class II, Desc, Acquisition Cost, and Net Book Value."
        Exit Sub
    End If

    Set sourceSheet = sourceBlock.Worksheet
    lastDataRow = FAS_GB_FindCapexLastDataRow(sourceBlock, headerRow, classOneColumn, classTwoColumn, _
        descriptionColumn, acquisitionCostColumn, netBookValueColumn)
    If lastDataRow <= headerRow Then
        FAS_SetStatus "No CAPEX data rows were found below the headers."
        Exit Sub
    End If

    Set rowFieldRanges = FAS_GB_GetColumnRanges(sourceSheet, headerRow + 1, lastDataRow, _
        Array(classOneColumn, classTwoColumn, descriptionColumn))
    Set valueRanges = FAS_GB_GetColumnRanges(sourceSheet, headerRow + 1, lastDataRow, _
        Array(acquisitionCostColumn, netBookValueColumn))
    FAS_CreateGroupByReportFromCollections sourceSheet, rowFieldRanges, valueRanges, _
        "SUM", True, False, "Group By - CAPEX", "Group By"
End Sub

Public Function FAS_TryRunGroupByCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String

    normalizedCommand = LCase$(FAS_NormalizeCommandText(commandText))
    Select Case normalizedCommand
        Case "groupby", "group", "groupbyreport"
            FAS_CreateGroupByReport
        Case "groupbycount", "groupcount"
            FAS_CreateGroupByReportWithCount
        Case Else
            Exit Function
    End Select
    FAS_TryRunGroupByCommand = True
End Function

Public Function FAS_IsGroupByCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String

    normalizedCommand = LCase$(FAS_NormalizeCommandText(commandText))
    FAS_IsGroupByCommand = normalizedCommand = "groupby" Or normalizedCommand = "group" Or _
        normalizedCommand = "groupbyreport" Or normalizedCommand = "groupbycount" Or _
        normalizedCommand = "groupcount"
End Function

' The first three GROUPBY arguments reproduce the editable report. Optional
' GROUPBY arguments are accepted and ignored because sorting is materialized.
Public Function FAS_ParseGroupByFormula(ByVal formulaText As String, ByRef rowFieldsText As String, _
    ByRef valueFieldsText As String, ByRef functionName As String) As Boolean
    Dim arguments() As String

    If Not FAS_GB_TryParseArguments(formulaText, arguments) Then Exit Function
    If UBound(arguments) < 3 Then Exit Function
    rowFieldsText = arguments(1)
    valueFieldsText = arguments(2)
    functionName = UCase$(Trim$(arguments(3)))
    If Not FAS_GB_IsFunction(functionName) Then Exit Function
    FAS_ParseGroupByFormula = True
End Function

Private Sub FAS_CreateGroupByReportCore(ByVal forceCount As Boolean)
    Dim rowFields As Range
    Dim valueFields As Range
    Dim rowFieldRanges As Collection
    Dim valueRanges As Collection
    Dim functionName As String
    Dim parsedFromFormula As Boolean
    Dim includeCount As Boolean
    Dim sourceSheet As Worksheet

    parsedFromFormula = FAS_GB_TryGetSelectedInputs(rowFields, valueFields, functionName)
    If Not parsedFromFormula Then
        Set rowFields = FAS_GB_PickRange("Select the row-field range (one or more adjacent columns):")
        If rowFields Is Nothing Then Exit Sub
        Set valueFields = FAS_GB_PickRange("Select the value range (one or more adjacent columns):")
        If valueFields Is Nothing Then Exit Sub
        functionName = FAS_GB_RequestFunction()
        If Len(functionName) = 0 Then Exit Sub
    End If

    If Not FAS_GB_ValidateInputs(rowFields, valueFields) Then Exit Sub
    Set rowFieldRanges = FAS_GB_SplitRange(rowFields)
    Set valueRanges = FAS_GB_SplitRange(valueFields)
    includeCount = forceCount
    If functionName = "COUNT" Then includeCount = True
    If Not forceCount And Not parsedFromFormula Then includeCount = FAS_GB_RequestCount()

    Set sourceSheet = rowFields.Worksheet
    FAS_CreateGroupByReportFromCollections sourceSheet, rowFieldRanges, valueRanges, _
        functionName, includeCount, True, "Group By Report", "Group By"
End Sub

Private Sub FAS_CreateGroupByReportFromCollections(ByVal sourceSheet As Worksheet, _
    ByVal rowFieldRanges As Collection, ByVal valueRanges As Collection, _
    ByVal functionName As String, ByVal includeCount As Boolean, _
    ByVal includeIntermediateRows As Boolean, ByVal titleText As String, ByVal sheetBaseName As String)

    Dim targetSheet As Worksheet
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If sourceSheet Is Nothing Or rowFieldRanges Is Nothing Or valueRanges Is Nothing Then Exit Sub
    If rowFieldRanges.Count = 0 Or valueRanges.Count = 0 Then Exit Sub
    functionName = UCase$(Trim$(functionName))
    If Not FAS_GB_IsFunction(functionName) Then Exit Sub

    On Error GoTo BuildFailed
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set targetSheet = sourceSheet.Parent.Worksheets.Add(After:=sourceSheet)
    targetSheet.Name = FAS_GB_UniqueSheetName(sourceSheet.Parent, sheetBaseName)
    FAS_GB_BuildReport targetSheet, rowFieldRanges, valueRanges, functionName, includeCount, _
        includeIntermediateRows, titleText

    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    FAS_UndoRegisterCreatedSheet targetSheet, "Create Group By report"
    targetSheet.Activate
    targetSheet.Range("A1").Select
    FAS_SetStatus "Created editable Group By report on '" & targetSheet.Name & "'."
    Exit Sub

BuildFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If Not targetSheet Is Nothing Then targetSheet.Delete
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
    If Not Application.Visible Then Err.Raise errorNumber, errorSource, errorDescription
    FAS_SetStatus "Could not create the Group By report: " & errorDescription
End Sub

Private Sub FAS_GB_BuildReport(ByVal targetSheet As Worksheet, ByVal rowFieldRanges As Collection, _
    ByVal valueRanges As Collection, ByVal functionName As String, ByVal includeCount As Boolean, _
    ByVal includeIntermediateRows As Boolean, ByVal titleText As String)

    Dim childrenByParent As Object
    Dim pathValuesByKey As Object
    Dim sortTotals As Object
    Dim fieldValues As Collection
    Dim firstMeasureValues As Variant
    Dim sourceRefs As Collection
    Dim valueRefs As Collection
    Dim sourceRange As Range
    Dim sourceValue As Variant
    Dim rowIndex As Long
    Dim fieldIndex As Long
    Dim pathKey As String
    Dim parentKey As String
    Dim keyPart As String
    Dim rowCount As Long
    Dim outputRow As Long
    Dim finalColumn As Long

    Set childrenByParent = FAS_GB_NewDictionary()
    Set pathValuesByKey = FAS_GB_NewDictionary()
    Set sortTotals = FAS_GB_NewDictionary()
    Set fieldValues = New Collection
    Set sourceRefs = New Collection
    Set valueRefs = New Collection

    For Each sourceRange In rowFieldRanges
        fieldValues.Add sourceRange.Value2
        sourceRefs.Add FAS_GB_SourceReference(sourceRange)
    Next sourceRange
    For Each sourceRange In valueRanges
        valueRefs.Add FAS_GB_SourceReference(sourceRange)
    Next sourceRange

    rowCount = rowFieldRanges(1).Rows.Count
    firstMeasureValues = valueRanges(1).Value2
    For rowIndex = 1 To rowCount
        parentKey = vbNullString
        For fieldIndex = 1 To rowFieldRanges.Count
            sourceValue = FAS_GB_ArrayValue(fieldValues(fieldIndex), rowIndex)
            keyPart = FAS_GB_Key(sourceValue)
            If Len(parentKey) = 0 Then
                pathKey = keyPart
            Else
                pathKey = parentKey & FAS_GROUPBY_KEY_SEPARATOR & keyPart
            End If
            If Not sortTotals.Exists(pathKey) Then
                sortTotals.Add pathKey, 0#
                FAS_GB_AddPath pathValuesByKey, pathKey, parentKey, sourceValue
                FAS_GB_AddChild childrenByParent, parentKey, pathKey
            End If
            sortTotals(pathKey) = CDbl(sortTotals(pathKey)) + _
                FAS_GB_Number(FAS_GB_ArrayValue(firstMeasureValues, rowIndex))
            parentKey = pathKey
        Next fieldIndex
    Next rowIndex

    FAS_GB_WriteHeadings targetSheet, rowFieldRanges, valueRanges, includeCount, titleText
    outputRow = 3
    FAS_GB_WriteTree targetSheet, vbNullString, 1, rowFieldRanges.Count, valueRanges.Count, _
        functionName, includeCount, includeIntermediateRows, childrenByParent, pathValuesByKey, _
        sortTotals, sourceRefs, valueRefs, outputRow
    FAS_GB_WriteTotal targetSheet, outputRow, rowFieldRanges.Count, valueRanges.Count, functionName, _
        includeCount, sourceRefs, valueRefs

    finalColumn = rowFieldRanges.Count + valueRanges.Count + IIf(includeCount, 1, 0)
    FAS_GB_FormatReport targetSheet, outputRow, rowFieldRanges.Count, finalColumn
End Sub

Private Sub FAS_GB_WriteTree(ByVal targetSheet As Worksheet, ByVal parentKey As String, _
    ByVal currentLevel As Long, ByVal rowFieldCount As Long, ByVal valueCount As Long, _
    ByVal functionName As String, ByVal includeCount As Boolean, ByVal includeIntermediateRows As Boolean, _
    ByVal childrenByParent As Object, ByVal pathValuesByKey As Object, ByVal sortTotals As Object, _
    ByVal sourceRefs As Collection, ByVal valueRefs As Collection, ByRef outputRow As Long)

    Dim childOrder As Collection
    Dim sortedOrder As Collection
    Dim nodeKey As Variant
    Dim pathValues As Collection

    If Not childrenByParent.Exists(parentKey) Then Exit Sub
    Set childOrder = childrenByParent(parentKey)
    Set sortedOrder = FAS_GB_Sort(childOrder, sortTotals)
    For Each nodeKey In sortedOrder
        Set pathValues = pathValuesByKey(CStr(nodeKey))
        If includeIntermediateRows Or currentLevel = 1 Or currentLevel = rowFieldCount Then
            FAS_GB_WriteDataRow targetSheet, outputRow, pathValues, currentLevel, rowFieldCount, _
                valueCount, functionName, includeCount, sourceRefs, valueRefs
            outputRow = outputRow + 1
        End If
        If currentLevel < rowFieldCount Then
            FAS_GB_WriteTree targetSheet, CStr(nodeKey), currentLevel + 1, rowFieldCount, valueCount, _
                functionName, includeCount, includeIntermediateRows, childrenByParent, pathValuesByKey, _
                sortTotals, sourceRefs, valueRefs, outputRow
        End If
    Next nodeKey
End Sub

Private Sub FAS_GB_WriteDataRow(ByVal targetSheet As Worksheet, ByVal targetRow As Long, _
    ByVal pathValues As Collection, ByVal level As Long, ByVal rowFieldCount As Long, _
    ByVal valueCount As Long, ByVal functionName As String, ByVal includeCount As Boolean, _
    ByVal sourceRefs As Collection, ByVal valueRefs As Collection)

    Dim fieldIndex As Long
    Dim valueIndex As Long
    Dim targetColumn As Long
    Dim criteriaText As String

    For fieldIndex = 1 To rowFieldCount
        If fieldIndex <= level Then
            targetSheet.Cells(targetRow, fieldIndex).Value = FAS_GB_Display(pathValues(fieldIndex))
        Else
            targetSheet.Cells(targetRow, fieldIndex).ClearContents
        End If
    Next fieldIndex

    criteriaText = FAS_GB_Criteria(targetSheet, targetRow, level, sourceRefs)
    targetColumn = rowFieldCount + 1
    For valueIndex = 1 To valueCount
        targetSheet.Cells(targetRow, targetColumn + valueIndex - 1).Formula = _
            FAS_GB_MeasureFormula(functionName, CStr(valueRefs(valueIndex)), criteriaText)
    Next valueIndex
    If includeCount Then
        targetSheet.Cells(targetRow, targetColumn + valueCount).Formula = "=COUNTIFS(" & criteriaText & ")"
    End If
End Sub

Private Sub FAS_GB_WriteTotal(ByVal targetSheet As Worksheet, ByVal targetRow As Long, _
    ByVal rowFieldCount As Long, ByVal valueCount As Long, ByVal functionName As String, _
    ByVal includeCount As Boolean, ByVal sourceRefs As Collection, ByVal valueRefs As Collection)

    Dim valueIndex As Long
    Dim targetColumn As Long
    Dim totalCriteria As String
    Dim finalColumn As Long

    targetSheet.Cells(targetRow, 1).Value = "Total"
    totalCriteria = CStr(sourceRefs(1)) & "," & Chr$(34) & "<>" & Chr$(34)
    targetColumn = rowFieldCount + 1
    For valueIndex = 1 To valueCount
        targetSheet.Cells(targetRow, targetColumn + valueIndex - 1).Formula = _
            FAS_GB_MeasureFormula(functionName, CStr(valueRefs(valueIndex)), totalCriteria)
    Next valueIndex
    If includeCount Then targetSheet.Cells(targetRow, targetColumn + valueCount).Formula = _
        "=COUNTIFS(" & totalCriteria & ")"
    finalColumn = targetColumn + valueCount - 1 + IIf(includeCount, 1, 0)
    targetSheet.Range(targetSheet.Cells(targetRow, 1), targetSheet.Cells(targetRow, finalColumn)).Font.Bold = True
    targetSheet.Range(targetSheet.Cells(targetRow, 1), targetSheet.Cells(targetRow, finalColumn)).Borders(xlEdgeTop).LineStyle = xlDouble
End Sub

Private Function FAS_GB_MeasureFormula(ByVal functionName As String, ByVal valueReference As String, _
    ByVal criteriaText As String) As String
    If UCase$(functionName) = "COUNT" Then
        FAS_GB_MeasureFormula = "=COUNTIFS(" & valueReference & "," & Chr$(34) & "<>" & Chr$(34) & "," & criteriaText & ")"
    Else
        FAS_GB_MeasureFormula = "=SUMIFS(" & valueReference & "," & criteriaText & ")"
    End If
End Function

Private Function FAS_GB_Criteria(ByVal targetSheet As Worksheet, ByVal targetRow As Long, _
    ByVal level As Long, ByVal sourceRefs As Collection) As String
    Dim fieldIndex As Long
    Dim metaAddress As String

    For fieldIndex = 1 To level
        If Len(FAS_GB_Criteria) > 0 Then FAS_GB_Criteria = FAS_GB_Criteria & ","
        metaAddress = targetSheet.Cells(targetRow, fieldIndex).Address(False, True)
        FAS_GB_Criteria = FAS_GB_Criteria & CStr(sourceRefs(fieldIndex)) & "," & metaAddress
    Next fieldIndex
End Function

Private Sub FAS_GB_WriteHeadings(ByVal targetSheet As Worksheet, ByVal rowFieldRanges As Collection, _
    ByVal valueRanges As Collection, ByVal includeCount As Boolean, ByVal titleText As String)
    Dim sourceRange As Range
    Dim columnIndex As Long
    Dim finalColumn As Long

    targetSheet.Cells(1, 1).Value = titleText
    For Each sourceRange In rowFieldRanges
        targetSheet.Cells(2, columnIndex + 1).Value = FAS_GB_Header(sourceRange, "Field " & (columnIndex + 1))
        columnIndex = columnIndex + 1
    Next sourceRange
    For Each sourceRange In valueRanges
        targetSheet.Cells(2, columnIndex + 1).Value = FAS_GB_Header(sourceRange, "Measure " & (columnIndex - rowFieldRanges.Count + 1))
        columnIndex = columnIndex + 1
    Next sourceRange
    If includeCount Then targetSheet.Cells(2, columnIndex + 1).Value = "Count"
    finalColumn = rowFieldRanges.Count + valueRanges.Count + IIf(includeCount, 1, 0)

    With targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(1, finalColumn))
        .Merge
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
    End With
    With targetSheet.Range(targetSheet.Cells(2, 1), targetSheet.Cells(2, finalColumn))
        .Font.Name = "Arial"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .WrapText = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
End Sub

Private Sub FAS_GB_FormatReport(ByVal targetSheet As Worksheet, ByVal finalRow As Long, _
    ByVal rowFieldCount As Long, ByVal finalColumn As Long)
    Dim fieldIndex As Long
    Dim firstMeasureColumn As Long

    firstMeasureColumn = rowFieldCount + 1
    targetSheet.Range(targetSheet.Cells(3, 1), targetSheet.Cells(finalRow, finalColumn)).Font.Name = "Arial"
    targetSheet.Range(targetSheet.Cells(3, firstMeasureColumn), targetSheet.Cells(finalRow, finalColumn)).NumberFormat = "#,##0;(#,##0);"
    targetSheet.Range(targetSheet.Cells(3, firstMeasureColumn), targetSheet.Cells(finalRow, finalColumn)).HorizontalAlignment = xlRight
    For fieldIndex = 1 To rowFieldCount
        targetSheet.Columns(fieldIndex).ColumnWidth = 24
    Next fieldIndex
    targetSheet.Range(targetSheet.Cells(1, firstMeasureColumn), targetSheet.Cells(1, finalColumn)).EntireColumn.ColumnWidth = 14
    targetSheet.Rows("1:2").AutoFit
    targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(finalRow, finalColumn)).VerticalAlignment = xlCenter
End Sub

Private Function FAS_GB_TryGetSelectedInputs(ByRef rowFields As Range, ByRef valueFields As Range, _
    ByRef functionName As String) As Boolean
    Dim sourceCell As Range
    Dim formulaText As String
    Dim rowFieldsText As String
    Dim valueFieldsText As String

    If TypeName(Selection) <> "Range" Then Exit Function
    Set sourceCell = ActiveCell
    If sourceCell Is Nothing Or Not sourceCell.HasFormula Then Exit Function
    formulaText = CStr(sourceCell.Formula)
    If Not FAS_ParseGroupByFormula(formulaText, rowFieldsText, valueFieldsText, functionName) Then
        formulaText = CStr(sourceCell.FormulaLocal)
        If Not FAS_ParseGroupByFormula(formulaText, rowFieldsText, valueFieldsText, functionName) Then Exit Function
    End If
    Set rowFields = FAS_GB_ResolveRange(rowFieldsText, sourceCell.Worksheet)
    Set valueFields = FAS_GB_ResolveRange(valueFieldsText, sourceCell.Worksheet)
    If rowFields Is Nothing Or valueFields Is Nothing Then Exit Function
    FAS_GB_TryGetSelectedInputs = True
End Function

Private Function FAS_GB_ResolveRange(ByVal referenceText As String, ByVal contextSheet As Worksheet) As Range
    Dim workbookObject As Workbook
    Dim sheetText As String
    Dim addressText As String
    Dim bangPosition As Long
    Dim targetSheet As Worksheet

    referenceText = Trim$(referenceText)
    If Len(referenceText) = 0 Or InStr(1, referenceText, "[", vbBinaryCompare) > 0 Then Exit Function
    Set workbookObject = contextSheet.Parent
    bangPosition = InStrRev(referenceText, "!")
    If bangPosition > 0 Then
        sheetText = Left$(referenceText, bangPosition - 1)
        addressText = Mid$(referenceText, bangPosition + 1)
        If Left$(sheetText, 1) = "'" And Right$(sheetText, 1) = "'" Then
            sheetText = Mid$(sheetText, 2, Len(sheetText) - 2)
            sheetText = Replace(sheetText, "''", "'", 1, -1, vbBinaryCompare)
        End If
        On Error Resume Next
        Set targetSheet = workbookObject.Worksheets(sheetText)
        If Not targetSheet Is Nothing Then Set FAS_GB_ResolveRange = targetSheet.Range(addressText)
        On Error GoTo 0
    Else
        On Error Resume Next
        Set FAS_GB_ResolveRange = contextSheet.Range(referenceText)
        On Error GoTo 0
    End If
End Function

Private Function FAS_GB_TryParseArguments(ByVal formulaText As String, ByRef arguments() As String) As Boolean
    Dim normalizedFormula As String
    Dim openPosition As Long
    Dim closePosition As Long
    Dim delimiter As String
    Dim depth As Long
    Dim inString As Boolean
    Dim characterIndex As Long
    Dim argumentStart As Long
    Dim currentChar As String
    Dim argumentText As String
    Dim functionText As String

    normalizedFormula = Trim$(formulaText)
    If Left$(normalizedFormula, 1) <> "=" Then Exit Function
    openPosition = InStr(2, normalizedFormula, "(", vbBinaryCompare)
    If openPosition <= 2 Then Exit Function
    functionText = UCase$(Trim$(Mid$(normalizedFormula, 2, openPosition - 2)))
    If Left$(functionText, 6) = "_XLFN." Or Left$(functionText, 6) = "_XLWS." Then functionText = Mid$(functionText, 7)
    If functionText <> "GROUPBY" Then Exit Function
    For characterIndex = openPosition To Len(normalizedFormula)
        currentChar = Mid$(normalizedFormula, characterIndex, 1)
        If currentChar = Chr$(34) Then
            inString = Not inString
        ElseIf Not inString Then
            If currentChar = "(" Then depth = depth + 1
            If currentChar = ")" Then
                depth = depth - 1
                If depth = 0 Then closePosition = characterIndex: Exit For
            End If
        End If
    Next characterIndex
    If closePosition = 0 Or Len(Trim$(Mid$(normalizedFormula, closePosition + 1))) > 0 Then Exit Function
    delimiter = FAS_GB_Delimiter(normalizedFormula, openPosition + 1, closePosition - 1)
    If Len(delimiter) = 0 Then Exit Function

    ReDim arguments(1 To 1)
    argumentStart = openPosition + 1
    depth = 0
    inString = False
    For characterIndex = argumentStart To closePosition - 1
        currentChar = Mid$(normalizedFormula, characterIndex, 1)
        If currentChar = Chr$(34) Then
            inString = Not inString
        ElseIf Not inString Then
            If currentChar = "(" Then depth = depth + 1
            If currentChar = ")" Then depth = depth - 1
            If currentChar = delimiter And depth = 0 Then
                argumentText = Trim$(Mid$(normalizedFormula, argumentStart, characterIndex - argumentStart))
                If Len(argumentText) = 0 Then Exit Function
                arguments(UBound(arguments)) = argumentText
                ReDim Preserve arguments(1 To UBound(arguments) + 1)
                argumentStart = characterIndex + 1
            End If
        End If
    Next characterIndex
    argumentText = Trim$(Mid$(normalizedFormula, argumentStart, closePosition - argumentStart))
    If Len(argumentText) = 0 Then Exit Function
    arguments(UBound(arguments)) = argumentText
    FAS_GB_TryParseArguments = UBound(arguments) >= 3
End Function

Private Function FAS_GB_Delimiter(ByVal formulaText As String, ByVal firstPosition As Long, _
    ByVal lastPosition As Long) As String
    Dim characterIndex As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim currentChar As String

    For characterIndex = firstPosition To lastPosition
        currentChar = Mid$(formulaText, characterIndex, 1)
        If currentChar = Chr$(34) Then
            inString = Not inString
        ElseIf Not inString Then
            If currentChar = "(" Then depth = depth + 1
            If currentChar = ")" Then depth = depth - 1
            If depth = 0 And (currentChar = "," Or currentChar = ";") Then
                FAS_GB_Delimiter = currentChar
                Exit Function
            End If
        End If
    Next characterIndex
End Function

Private Function FAS_GB_PickRange(ByVal promptText As String) As Range
    On Error Resume Next
    Set FAS_GB_PickRange = Application.InputBox(Prompt:=promptText, Title:="FAS Group By", Type:=8)
    On Error GoTo 0
End Function

Private Function FAS_GB_RequestFunction() As String
    Dim enteredValue As Variant
    Dim functionName As String

    enteredValue = Application.InputBox(Prompt:="Aggregation: SUM or COUNT", Title:="FAS Group By", Default:="SUM", Type:=2)
    If VarType(enteredValue) = vbBoolean Then Exit Function
    functionName = UCase$(Trim$(CStr(enteredValue)))
    If FAS_GB_IsFunction(functionName) Then FAS_GB_RequestFunction = functionName
End Function

Private Function FAS_GB_RequestCount() As Boolean
    If Application.Visible Then FAS_GB_RequestCount = MsgBox("Add a COUNTIFS column?", vbQuestion + vbYesNo, "FAS Group By") = vbYes
End Function

Private Function FAS_GB_ValidateInputs(ByVal rowFields As Range, ByVal valueFields As Range) As Boolean
    Dim rowWorkbook As Workbook
    Dim valueWorkbook As Workbook

    If rowFields Is Nothing Or valueFields Is Nothing Then Exit Function
    If rowFields.Areas.Count <> 1 Or valueFields.Areas.Count <> 1 Then
        FAS_SetStatus "Group By needs one contiguous row-field range and one contiguous value range."
        Exit Function
    End If
    If rowFields.Rows.Count <> valueFields.Rows.Count Then
        FAS_SetStatus "Row-field and value ranges must have the same number of rows."
        Exit Function
    End If
    Set rowWorkbook = rowFields.Worksheet.Parent
    Set valueWorkbook = valueFields.Worksheet.Parent
    If Not rowWorkbook Is valueWorkbook Then
        FAS_SetStatus "Row-field and value ranges must be in the same workbook."
        Exit Function
    End If
    FAS_GB_ValidateInputs = True
End Function

Private Function FAS_GB_SplitRange(ByVal sourceRange As Range) As Collection
    Dim result As Collection
    Dim columnIndex As Long

    Set result = New Collection
    For columnIndex = 1 To sourceRange.Columns.Count
        result.Add sourceRange.Columns(columnIndex)
    Next columnIndex
    Set FAS_GB_SplitRange = result
End Function

Private Function FAS_GB_GetColumnRanges(ByVal sourceSheet As Worksheet, ByVal firstRow As Long, _
    ByVal lastRow As Long, ByVal columnNumbers As Variant) As Collection
    Dim result As Collection
    Dim columnValue As Variant

    Set result = New Collection
    For Each columnValue In columnNumbers
        result.Add sourceSheet.Range(sourceSheet.Cells(firstRow, CLng(columnValue)), _
            sourceSheet.Cells(lastRow, CLng(columnValue)))
    Next columnValue
    Set FAS_GB_GetColumnRanges = result
End Function

Private Function FAS_GB_ArrayValue(ByVal sourceValues As Variant, ByVal rowIndex As Long) As Variant
    If IsArray(sourceValues) Then
        FAS_GB_ArrayValue = sourceValues(rowIndex, 1)
    Else
        FAS_GB_ArrayValue = sourceValues
    End If
End Function

Private Function FAS_GB_Number(ByVal sourceValue As Variant) As Double
    Dim valueText As String
    Dim dotPosition As Long
    Dim commaPosition As Long

    If IsError(sourceValue) Or IsNull(sourceValue) Or IsEmpty(sourceValue) Then Exit Function
    If IsNumeric(sourceValue) Then
        FAS_GB_Number = CDbl(sourceValue)
    Else
        valueText = Replace(CStr(sourceValue), Chr$(160), vbNullString)
        valueText = Replace(valueText, " ", vbNullString)
        dotPosition = InStrRev(valueText, ".")
        commaPosition = InStrRev(valueText, ",")
        If dotPosition > 0 And commaPosition > 0 Then
            If commaPosition > dotPosition Then
                valueText = Replace(valueText, ".", vbNullString)
                valueText = Replace(valueText, ",", ".")
            Else
                valueText = Replace(valueText, ",", vbNullString)
            End If
        Else
            valueText = Replace(valueText, ",", ".")
        End If
        FAS_GB_Number = Val(valueText)
    End If
End Function

Private Function FAS_GB_ValueText(ByVal sourceValue As Variant) As String
    If IsError(sourceValue) Or IsNull(sourceValue) Or IsEmpty(sourceValue) Then Exit Function
    FAS_GB_ValueText = CStr(sourceValue)
End Function

Private Function FAS_GB_Key(ByVal sourceValue As Variant) As String
    Dim valueText As String

    valueText = FAS_GB_ValueText(sourceValue)
    If Len(Trim$(valueText)) = 0 Then
        FAS_GB_Key = FAS_GROUPBY_BLANK_KEY
    Else
        ' Keep nonblank source text exact so the visible criteria cell remains
        ' faithful to the source and the generated SUMIFS matches every row.
        FAS_GB_Key = LCase$(valueText)
    End If
End Function

Private Function FAS_GB_Display(ByVal sourceValue As Variant) As String
    FAS_GB_Display = FAS_GB_ValueText(sourceValue)
End Function

Private Sub FAS_GB_AddPath(ByVal pathValuesByKey As Object, ByVal pathKey As String, _
    ByVal parentKey As String, ByVal sourceValue As Variant)
    Dim pathValues As Collection
    Dim parentValues As Collection
    Dim itemIndex As Long

    Set pathValues = New Collection
    If Len(parentKey) > 0 Then
        Set parentValues = pathValuesByKey(parentKey)
        For itemIndex = 1 To parentValues.Count
            pathValues.Add parentValues(itemIndex)
        Next itemIndex
    End If
    pathValues.Add sourceValue
    pathValuesByKey.Add pathKey, pathValues
End Sub

Private Sub FAS_GB_AddChild(ByVal childrenByParent As Object, ByVal parentKey As String, ByVal pathKey As String)
    Dim childOrder As Collection

    If Not childrenByParent.Exists(parentKey) Then
        Set childOrder = New Collection
        childrenByParent.Add parentKey, childOrder
    End If
    Set childOrder = childrenByParent(parentKey)
    childOrder.Add pathKey
End Sub

Private Function FAS_GB_Sort(ByVal sourceOrder As Collection, ByVal totals As Object) As Collection
    Dim sortedKeys() As String
    Dim result As Collection
    Dim firstIndex As Long
    Dim secondIndex As Long
    Dim temporaryKey As String
    Dim firstTotal As Double
    Dim secondTotal As Double

    Set result = New Collection
    If sourceOrder Is Nothing Or sourceOrder.Count = 0 Then
        Set FAS_GB_Sort = result
        Exit Function
    End If
    ReDim sortedKeys(1 To sourceOrder.Count)
    For firstIndex = 1 To sourceOrder.Count
        sortedKeys(firstIndex) = CStr(sourceOrder(firstIndex))
    Next firstIndex
    For firstIndex = 2 To UBound(sortedKeys)
        temporaryKey = sortedKeys(firstIndex)
        secondIndex = firstIndex - 1
        Do While secondIndex >= 1
            firstTotal = CDbl(totals(sortedKeys(secondIndex)))
            secondTotal = CDbl(totals(temporaryKey))
            If firstTotal >= secondTotal Then Exit Do
            sortedKeys(secondIndex + 1) = sortedKeys(secondIndex)
            secondIndex = secondIndex - 1
        Loop
        sortedKeys(secondIndex + 1) = temporaryKey
    Next firstIndex
    For firstIndex = 1 To UBound(sortedKeys)
        result.Add sortedKeys(firstIndex)
    Next firstIndex
    Set FAS_GB_Sort = result
End Function

Private Function FAS_GB_IsFunction(ByVal functionName As String) As Boolean
    FAS_GB_IsFunction = UCase$(Trim$(functionName)) = "SUM" Or UCase$(Trim$(functionName)) = "COUNT"
End Function

Private Function FAS_GB_SourceReference(ByVal sourceRange As Range) As String
    FAS_GB_SourceReference = "'" & Replace(sourceRange.Worksheet.Name, "'", "''") & "'!" & _
        sourceRange.Address(True, True, xlA1)
End Function

Private Function FAS_GB_Header(ByVal sourceRange As Range, ByVal fallbackText As String) As String
    Dim headerValue As Variant
    Dim headerRow As Long

    headerRow = sourceRange.Row - 1
    If headerRow >= 1 Then headerValue = sourceRange.Worksheet.Cells(headerRow, sourceRange.Column).Value2
    If IsError(headerValue) Or IsNull(headerValue) Or IsEmpty(headerValue) Then
        FAS_GB_Header = fallbackText
    ElseIf Len(Trim$(CStr(headerValue))) = 0 Then
        FAS_GB_Header = fallbackText
    Else
        FAS_GB_Header = CStr(headerValue)
    End If
End Function

Private Function FAS_GB_UniqueSheetName(ByVal targetWorkbook As Workbook, ByVal baseName As String) As String
    Dim candidateName As String
    Dim suffixNumber As Long

    candidateName = baseName
    suffixNumber = 2
    Do While FAS_GB_SheetExists(targetWorkbook, candidateName)
        candidateName = Left$(baseName, 27) & " (" & suffixNumber & ")"
        suffixNumber = suffixNumber + 1
    Loop
    FAS_GB_UniqueSheetName = candidateName
End Function

Private Function FAS_GB_SheetExists(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Boolean
    Dim candidateSheet As Object

    On Error Resume Next
    Set candidateSheet = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    FAS_GB_SheetExists = Not candidateSheet Is Nothing
End Function

Private Function FAS_GB_NewDictionary() As Object
    Dim dictionary As Object

    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbTextCompare
    Set FAS_GB_NewDictionary = dictionary
End Function

Private Function FAS_GB_GetCapexSourceBlock() As Range
    Dim selectedRange As Range
    Dim usedSelection As Range

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select the CAPEX source table first."
        Exit Function
    End If
    Set selectedRange = Selection
    If selectedRange.Areas.Count <> 1 Then
        FAS_SetStatus "CAPEX Group By needs one contiguous source selection."
        Exit Function
    End If
    If selectedRange.CountLarge = 1 Then
        Set selectedRange = selectedRange.Worksheet.UsedRange
    Else
        Set usedSelection = Intersect(selectedRange, selectedRange.Worksheet.UsedRange)
        If Not usedSelection Is Nothing Then Set selectedRange = usedSelection
    End If
    Set FAS_GB_GetCapexSourceBlock = selectedRange
End Function

Private Function FAS_GB_FindCapexHeaders(ByVal sourceBlock As Range, ByRef headerRow As Long, _
    ByRef classOneColumn As Long, ByRef classTwoColumn As Long, ByRef descriptionColumn As Long, _
    ByRef acquisitionCostColumn As Long, ByRef netBookValueColumn As Long) As Boolean
    Dim candidateRow As Long
    Dim candidateColumn As Long
    Dim firstRow As Long
    Dim finalRow As Long
    Dim firstColumn As Long
    Dim finalColumn As Long
    Dim headerText As String

    firstRow = sourceBlock.Row
    finalRow = Application.Min(sourceBlock.Row + sourceBlock.Rows.Count - 1, firstRow + 24)
    firstColumn = sourceBlock.Column
    finalColumn = sourceBlock.Column + sourceBlock.Columns.Count - 1
    For candidateRow = firstRow To finalRow
        classOneColumn = 0
        classTwoColumn = 0
        descriptionColumn = 0
        acquisitionCostColumn = 0
        netBookValueColumn = 0
        For candidateColumn = firstColumn To finalColumn
            headerText = FAS_NormalizeCapexHeader(sourceBlock.Worksheet.Cells(candidateRow, candidateColumn).Value2)
            Select Case headerText
                Case "class i", "class 1": classOneColumn = candidateColumn
                Case "class ii", "class 2": classTwoColumn = candidateColumn
                Case "desc", "description": descriptionColumn = candidateColumn
                Case "acquisition cost", "acq. cost", "acq cost": acquisitionCostColumn = candidateColumn
                Case "net book value", "nbv", "book value", "bv": netBookValueColumn = candidateColumn
            End Select
        Next candidateColumn
        If classOneColumn > 0 And classTwoColumn > 0 And descriptionColumn > 0 And _
            acquisitionCostColumn > 0 And netBookValueColumn > 0 Then
            headerRow = candidateRow
            FAS_GB_FindCapexHeaders = True
            Exit Function
        End If
    Next candidateRow
End Function

Private Function FAS_GB_FindCapexLastDataRow(ByVal sourceBlock As Range, ByVal headerRow As Long, _
    ByVal classOneColumn As Long, ByVal classTwoColumn As Long, ByVal descriptionColumn As Long, _
    ByVal acquisitionCostColumn As Long, ByVal netBookValueColumn As Long) As Long
    Dim columnsToCheck As Variant
    Dim columnValue As Variant
    Dim candidateBottom As Long
    Dim columnLastRow As Long
    Dim searchRange As Range
    Dim lastPopulatedCell As Range

    columnsToCheck = Array(classOneColumn, classTwoColumn, descriptionColumn, acquisitionCostColumn, netBookValueColumn)
    candidateBottom = sourceBlock.Row + sourceBlock.Rows.Count - 1
    If candidateBottom <= headerRow Then
        candidateBottom = sourceBlock.Worksheet.UsedRange.Row + sourceBlock.Worksheet.UsedRange.Rows.Count - 1
    End If
    If candidateBottom <= headerRow Then Exit Function
    FAS_GB_FindCapexLastDataRow = headerRow
    For Each columnValue In columnsToCheck
        Set searchRange = sourceBlock.Worksheet.Range( _
            sourceBlock.Worksheet.Cells(headerRow + 1, CLng(columnValue)), _
            sourceBlock.Worksheet.Cells(candidateBottom, CLng(columnValue)))
        Set lastPopulatedCell = Nothing
        ' xlValues excludes formulas that currently evaluate to blank.
        Set lastPopulatedCell = searchRange.Find(What:="*", After:=searchRange.Cells(1, 1), _
            LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, MatchCase:=False)
        If Not lastPopulatedCell Is Nothing Then
            columnLastRow = lastPopulatedCell.Row
            If columnLastRow > FAS_GB_FindCapexLastDataRow Then FAS_GB_FindCapexLastDataRow = columnLastRow
        End If
    Next columnValue
End Function
