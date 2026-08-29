Attribute VB_Name = "modIFSTools"
Option Explicit

Private mCapturedIFSFormula As String
Private mCapturedIFSFormulaR1C1 As String
Private mCapturedIFSFunction As String
Private mCapturedIFSArgumentCount As Long
Private mKnownIFSMeasureReference As String

Public Sub FAS_BuildSUMIFS()
    FAS_BuildIFSFormula "SUMIFS"
End Sub

Public Sub FAS_BuildCOUNTIFS()
    FAS_BuildIFSFormula "COUNTIFS"
End Sub

Public Sub FAS_CaptureIFSPattern()
    Dim sourceCell As Range
    Dim functionName As String
    Dim arguments() As String

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a direct SUMIFS, COUNTIFS, AVERAGEIFS, MINIFS, or MAXIFS formula first."
        Exit Sub
    End If

    Set sourceCell = ActiveCell
    If Not FAS_TryParseIFSFormula(CStr(sourceCell.Formula), functionName, arguments) Then
        FAS_SetStatus "Active cell must contain a direct IFS formula."
        Exit Sub
    End If

    mCapturedIFSFormula = CStr(sourceCell.Formula)
    mCapturedIFSFormulaR1C1 = CStr(sourceCell.FormulaR1C1)
    mCapturedIFSFunction = functionName
    mCapturedIFSArgumentCount = UBound(arguments)
    If FAS_IsMeasureIFSFunction(functionName) Then mKnownIFSMeasureReference = arguments(1)

    FAS_SetStatus "Captured " & functionName & " pattern with " & mCapturedIFSArgumentCount & " arguments."
End Sub

Public Sub FAS_FillIFSPattern()
    Dim selectedRange As Range
    Dim currentCell As Range
    Dim eligibleCount As Long
    Dim skippedCount As Long
    Dim undoStarted As Boolean

    If Len(mCapturedIFSFormulaR1C1) = 0 Then
        FAS_SetStatus "Capture an IFS pattern before filling."
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select the cells to fill with the captured IFS pattern."
        Exit Sub
    End If

    Set selectedRange = Selection
    For Each currentCell In selectedRange.Cells
        If FAS_IsBlankIFSFillTarget(currentCell) Or FAS_IsMatchingIFSFillTarget(currentCell) Then
            eligibleCount = eligibleCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
    Next currentCell

    If eligibleCount = 0 Then
        FAS_SetStatus "No blank or matching IFS formula cells were found. " & skippedCount & " cell(s) skipped."
        Exit Sub
    End If

    On Error GoTo FillFailed
    If Not FAS_UndoBeginRange(selectedRange, "Fill IFS pattern") Then Exit Sub
    undoStarted = True
    For Each currentCell In selectedRange.Cells
        If FAS_IsBlankIFSFillTarget(currentCell) Or FAS_IsMatchingIFSFillTarget(currentCell) Then
            currentCell.FormulaR1C1 = mCapturedIFSFormulaR1C1
        End If
    Next currentCell

    FAS_UndoCommit
    FAS_SetStatus "Filled " & eligibleCount & " cell(s) with " & mCapturedIFSFunction & ". " & skippedCount & " cell(s) protected."
    Exit Sub

FillFailed:
    If undoStarted Then FAS_UndoCancel
    If Not Application.Visible Then Err.Raise Err.Number, Err.Source, Err.Description
    FAS_SetStatus "Could not fill the captured IFS pattern."
End Sub

Public Sub FAS_SwitchToSUMIFS()
    FAS_SwitchSelectedIFSFunction "SUMIFS"
End Sub

Public Sub FAS_SwitchToCOUNTIFS()
    FAS_SwitchSelectedIFSFunction "COUNTIFS"
End Sub

Public Sub FAS_SwitchToAVERAGEIFS()
    FAS_SwitchSelectedIFSFunction "AVERAGEIFS"
End Sub

Public Sub FAS_SwitchToMINIFS()
    FAS_SwitchSelectedIFSFunction "MINIFS"
End Sub

Public Sub FAS_SwitchToMAXIFS()
    FAS_SwitchSelectedIFSFunction "MAXIFS"
End Sub

Public Sub FAS_SwitchSelectedIFSFunction(ByVal targetFunction As String)
    Dim selectedRange As Range
    Dim currentCell As Range
    Dim functionName As String
    Dim arguments() As String
    Dim convertedFormula As String
    Dim measureReference As String
    Dim pickedRange As Range
    Dim eligibleCount As Long
    Dim changedCount As Long
    Dim skippedCount As Long
    Dim requiresMeasure As Boolean
    Dim response As VbMsgBoxResult
    Dim undoStarted As Boolean

    targetFunction = UCase$(Trim$(targetFunction))
    If Not FAS_IsSupportedIFSFunction(targetFunction) Then Exit Sub
    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select direct IFS formulas to switch."
        Exit Sub
    End If

    Set selectedRange = Selection
    measureReference = mKnownIFSMeasureReference
    For Each currentCell In selectedRange.Cells
        If FAS_IsWritableIFSFormulaCell(currentCell) And FAS_TryParseIFSFormula(CStr(currentCell.Formula), functionName, arguments) Then
            eligibleCount = eligibleCount + 1
            If FAS_IsMeasureIFSFunction(functionName) And Len(measureReference) = 0 Then measureReference = arguments(1)
            If functionName = "COUNTIFS" And FAS_IsMeasureIFSFunction(targetFunction) Then requiresMeasure = True
        Else
            skippedCount = skippedCount + 1
        End If
    Next currentCell

    If eligibleCount = 0 Then
        FAS_SetStatus "No direct IFS formulas were found. " & skippedCount & " cell(s) skipped."
        Exit Sub
    End If

    If requiresMeasure And Len(measureReference) = 0 Then
        Set pickedRange = FAS_PickIFSRange("Select the measure range for COUNTIFS conversion:")
        If pickedRange Is Nothing Then Exit Sub
        measureReference = FAS_GetIFSRangeReference(pickedRange)
    End If

    For Each currentCell In selectedRange.Cells
        If FAS_IsWritableIFSFormulaCell(currentCell) And FAS_TryParseIFSFormula(CStr(currentCell.Formula), functionName, arguments) Then
            If FAS_ConvertIFSFormula(CStr(currentCell.Formula), targetFunction, measureReference, convertedFormula) Then
                If StrComp(CStr(currentCell.Formula), convertedFormula, vbBinaryCompare) <> 0 Then changedCount = changedCount + 1
            End If
        End If
    Next currentCell

    If changedCount = 0 Then
        FAS_SetStatus "Selected IFS formulas already use " & targetFunction & "."
        Exit Sub
    End If
    If Application.Visible Then
        response = MsgBox("Switch " & changedCount & " formula cell(s) to " & targetFunction & "?", _
            vbQuestion + vbOKCancel, "FAS IFS Tools")
        If response <> vbOK Then Exit Sub
    End If

    On Error GoTo SwitchFailed
    If Not FAS_UndoBeginRange(selectedRange, "Switch IFS function") Then Exit Sub
    undoStarted = True
    For Each currentCell In selectedRange.Cells
        If FAS_IsWritableIFSFormulaCell(currentCell) And FAS_TryParseIFSFormula(CStr(currentCell.Formula), functionName, arguments) Then
            If FAS_ConvertIFSFormula(CStr(currentCell.Formula), targetFunction, measureReference, convertedFormula) Then
                If StrComp(CStr(currentCell.Formula), convertedFormula, vbBinaryCompare) <> 0 Then currentCell.Formula = convertedFormula
            End If
        End If
    Next currentCell

    If FAS_IsMeasureIFSFunction(targetFunction) And Len(measureReference) > 0 Then mKnownIFSMeasureReference = measureReference
    FAS_UndoCommit
    FAS_SetStatus "Switched " & changedCount & " formula cell(s) to " & targetFunction & ". " & skippedCount & " cell(s) skipped."
    Exit Sub

SwitchFailed:
    If undoStarted Then FAS_UndoCancel
    If Not Application.Visible Then Err.Raise Err.Number, Err.Source, Err.Description
    FAS_SetStatus "Could not switch the selected IFS formulas."
End Sub

Public Function FAS_TryRunIFSCommand(ByVal commandText As String) As Boolean
    Select Case LCase$(FAS_NormalizeCommandText(commandText))
        Case "sumifs", "buildsumifs", "ifssum"
            FAS_BuildSUMIFS
        Case "countifs", "buildcountifs", "ifscount"
            FAS_BuildCOUNTIFS
        Case "captureifs", "captureifspattern", "ifscapture"
            FAS_CaptureIFSPattern
        Case "fillifs", "fillifspattern", "ifsfill"
            FAS_FillIFSPattern
        Case "switchsumifs", "switchifssum"
            FAS_SwitchToSUMIFS
        Case "switchcountifs", "switchifscount"
            FAS_SwitchToCOUNTIFS
        Case "switchaverageifs", "switchifsaverage"
            FAS_SwitchToAVERAGEIFS
        Case "switchminifs", "switchifsmin"
            FAS_SwitchToMINIFS
        Case "switchmaxifs", "switchifsmax"
            FAS_SwitchToMAXIFS
        Case Else
            Exit Function
    End Select
    FAS_TryRunIFSCommand = True
End Function

Public Function FAS_IsIFSCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String

    normalizedCommand = LCase$(FAS_NormalizeCommandText(commandText))
    FAS_IsIFSCommand = normalizedCommand = "sumifs" Or normalizedCommand = "buildsumifs" Or normalizedCommand = "ifssum" Or _
        normalizedCommand = "countifs" Or normalizedCommand = "buildcountifs" Or normalizedCommand = "ifscount" Or _
        normalizedCommand = "captureifs" Or normalizedCommand = "captureifspattern" Or normalizedCommand = "ifscapture" Or _
        normalizedCommand = "fillifs" Or normalizedCommand = "fillifspattern" Or normalizedCommand = "ifsfill" Or _
        normalizedCommand = "switchsumifs" Or normalizedCommand = "switchifssum" Or _
        normalizedCommand = "switchcountifs" Or normalizedCommand = "switchifscount" Or _
        normalizedCommand = "switchaverageifs" Or normalizedCommand = "switchifsaverage" Or _
        normalizedCommand = "switchminifs" Or normalizedCommand = "switchifsmin" Or _
        normalizedCommand = "switchmaxifs" Or normalizedCommand = "switchifsmax"
End Function

Public Function FAS_ParseIFSFormula(ByVal formulaText As String, ByRef functionName As String, ByRef argumentCount As Long) As Boolean
    Dim arguments() As String

    If Not FAS_TryParseIFSFormula(formulaText, functionName, arguments) Then Exit Function
    argumentCount = UBound(arguments)
    FAS_ParseIFSFormula = True
End Function

Public Function FAS_ConvertIFSFormula(ByVal formulaText As String, ByVal targetFunction As String, _
    ByVal measureReference As String, ByRef convertedFormula As String) As Boolean
    Dim sourceFunction As String
    Dim sourceArguments() As String
    Dim targetArguments() As String
    Dim argumentCount As Long
    Dim argumentIndex As Long

    targetFunction = UCase$(Trim$(targetFunction))
    If Not FAS_IsSupportedIFSFunction(targetFunction) Then Exit Function
    If Not FAS_TryParseIFSFormula(formulaText, sourceFunction, sourceArguments) Then Exit Function
    If sourceFunction = targetFunction Then
        convertedFormula = formulaText
        FAS_ConvertIFSFormula = True
        Exit Function
    End If

    argumentCount = UBound(sourceArguments)
    If targetFunction = "COUNTIFS" Then
        ReDim targetArguments(1 To argumentCount - 1)
        For argumentIndex = 2 To argumentCount
            targetArguments(argumentIndex - 1) = sourceArguments(argumentIndex)
        Next argumentIndex
    ElseIf sourceFunction = "COUNTIFS" Then
        If Len(Trim$(measureReference)) = 0 Then Exit Function
        ReDim targetArguments(1 To argumentCount + 1)
        targetArguments(1) = Trim$(measureReference)
        For argumentIndex = 1 To argumentCount
            targetArguments(argumentIndex + 1) = sourceArguments(argumentIndex)
        Next argumentIndex
    Else
        ReDim targetArguments(1 To argumentCount)
        For argumentIndex = 1 To argumentCount
            targetArguments(argumentIndex) = sourceArguments(argumentIndex)
        Next argumentIndex
    End If

    convertedFormula = "=" & targetFunction & "(" & FAS_JoinIFSArguments(targetArguments) & ")"
    FAS_ConvertIFSFormula = True
End Function

Private Sub FAS_BuildIFSFormula(ByVal functionName As String)
    Dim targetCell As Range
    Dim measureRange As Range
    Dim criteriaRange As Range
    Dim criteriaCell As Range
    Dim formulaText As String
    Dim undoStarted As Boolean

    Set targetCell = ActiveCell
    If targetCell Is Nothing Then Exit Sub
    If functionName <> "COUNTIFS" Then
        Set measureRange = FAS_PickIFSRange("Select the measure/sum range:")
        If measureRange Is Nothing Then Exit Sub
    End If
    Set criteriaRange = FAS_PickIFSRange("Select the criteria range:")
    If criteriaRange Is Nothing Then Exit Sub
    Set criteriaCell = FAS_PickIFSRange("Select the criteria cell:")
    If criteriaCell Is Nothing Then Exit Sub
    If criteriaCell.CountLarge <> 1 Then
        FAS_SetStatus "Criteria cell selection must contain one cell."
        Exit Sub
    End If
    If functionName <> "COUNTIFS" Then
        If Not FAS_IFSRangesMatch(measureRange, criteriaRange) Then Exit Sub
        formulaText = "=" & functionName & "(" & FAS_GetIFSRangeReference(measureRange) & ","
        mKnownIFSMeasureReference = FAS_GetIFSRangeReference(measureRange)
    Else
        formulaText = "=" & functionName & "("
    End If
    formulaText = formulaText & FAS_GetIFSRangeReference(criteriaRange) & "," & _
        FAS_GetIFSRelativeCriteriaReference(criteriaCell) & ")"

    On Error GoTo BuildFailed
    If Not FAS_UndoBeginRange(targetCell, "Build " & functionName) Then Exit Sub
    undoStarted = True
    targetCell.Formula = formulaText
    FAS_UndoCommit
    FAS_SetStatus "Built " & functionName & " in " & targetCell.Address(False, False) & "."
    Exit Sub

BuildFailed:
    If undoStarted Then FAS_UndoCancel
    If Not Application.Visible Then Err.Raise Err.Number, Err.Source, Err.Description
    FAS_SetStatus "Could not build the " & functionName & " formula."
End Sub

Private Function FAS_PickIFSRange(ByVal promptText As String) As Range
    On Error Resume Next
    Set FAS_PickIFSRange = Application.InputBox(Prompt:=promptText, Title:="FAS IFS Tools", Type:=8)
    On Error GoTo 0
End Function

Private Function FAS_GetIFSRangeReference(ByVal targetRange As Range) As String
    FAS_GetIFSRangeReference = targetRange.Address(True, True, xlA1, True)
End Function

Private Function FAS_GetIFSRelativeCriteriaReference(ByVal criteriaCell As Range) As String
    ' Keep the criteria column fixed while the row moves when the formula is filled.
    FAS_GetIFSRelativeCriteriaReference = criteriaCell.Address(False, True, xlA1, True)
End Function

Private Function FAS_IFSRangesMatch(ByVal firstRange As Range, ByVal secondRange As Range) As Boolean
    If firstRange.Rows.Count <> secondRange.Rows.Count Or firstRange.Columns.Count <> secondRange.Columns.Count Then
        FAS_SetStatus "Measure and criteria ranges must have the same shape."
        Exit Function
    End If
    FAS_IFSRangesMatch = True
End Function

Private Function FAS_IsBlankIFSFillTarget(ByVal currentCell As Range) As Boolean
    If currentCell.HasFormula Then Exit Function
    If IsError(currentCell.Value2) Then Exit Function
    FAS_IsBlankIFSFillTarget = IsEmpty(currentCell.Value2)
End Function

Private Function FAS_IsMatchingIFSFillTarget(ByVal currentCell As Range) As Boolean
    Dim functionName As String
    Dim arguments() As String

    On Error Resume Next
    If currentCell.HasArray Or currentCell.HasSpill Then Exit Function
    On Error GoTo 0
    If Not FAS_TryParseIFSFormula(CStr(currentCell.Formula), functionName, arguments) Then Exit Function
    FAS_IsMatchingIFSFillTarget = functionName = mCapturedIFSFunction And UBound(arguments) = mCapturedIFSArgumentCount
End Function

Private Function FAS_IsWritableIFSFormulaCell(ByVal currentCell As Range) As Boolean
    On Error Resume Next
    FAS_IsWritableIFSFormulaCell = Not currentCell.HasArray And Not currentCell.HasSpill
    On Error GoTo 0
End Function

Private Function FAS_TryParseIFSFormula(ByVal formulaText As String, ByRef functionName As String, _
    ByRef arguments() As String) As Boolean
    Dim normalizedFormula As String
    Dim openPosition As Long
    Dim closePosition As Long
    Dim argumentStart As Long
    Dim argumentIndex As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim currentChar As String
    Dim argumentText As String

    normalizedFormula = Trim$(formulaText)
    If Left$(normalizedFormula, 1) <> "=" Then Exit Function
    openPosition = InStr(2, normalizedFormula, "(", vbBinaryCompare)
    If openPosition <= 2 Then Exit Function
    functionName = UCase$(Trim$(Mid$(normalizedFormula, 2, openPosition - 2)))
    If Not FAS_IsSupportedIFSFunction(functionName) Then Exit Function

    depth = 0
    argumentIndex = openPosition
    Do While argumentIndex <= Len(normalizedFormula)
        currentChar = Mid$(normalizedFormula, argumentIndex, 1)
        If currentChar = Chr$(34) Then
            If inString And argumentIndex < Len(normalizedFormula) And Mid$(normalizedFormula, argumentIndex + 1, 1) = Chr$(34) Then
                argumentIndex = argumentIndex + 1
            Else
                inString = Not inString
            End If
        ElseIf Not inString Then
            If currentChar = "(" Then
                depth = depth + 1
            ElseIf currentChar = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    closePosition = argumentIndex
                    Exit Do
                End If
                If depth < 0 Then Exit Function
            End If
        End If
        argumentIndex = argumentIndex + 1
    Loop

    If closePosition = 0 Or Len(Trim$(Mid$(normalizedFormula, closePosition + 1))) > 0 Then Exit Function
    If Len(Trim$(Mid$(normalizedFormula, openPosition + 1, closePosition - openPosition - 1))) = 0 Then Exit Function

    ReDim arguments(1 To 1)
    argumentStart = openPosition + 1
    depth = 0
    inString = False
    argumentIndex = argumentStart
    Do While argumentIndex < closePosition
        currentChar = Mid$(normalizedFormula, argumentIndex, 1)
        If currentChar = Chr$(34) Then
            If inString And argumentIndex < closePosition - 1 And Mid$(normalizedFormula, argumentIndex + 1, 1) = Chr$(34) Then
                argumentIndex = argumentIndex + 1
            Else
                inString = Not inString
            End If
        ElseIf Not inString Then
            If currentChar = "(" Then
                depth = depth + 1
            ElseIf currentChar = ")" Then
                depth = depth - 1
                If depth < 0 Then Exit Function
            ElseIf currentChar = "," And depth = 0 Then
                argumentText = Trim$(Mid$(normalizedFormula, argumentStart, argumentIndex - argumentStart))
                If Len(argumentText) = 0 Then Exit Function
                arguments(UBound(arguments)) = argumentText
                ReDim Preserve arguments(1 To UBound(arguments) + 1)
                argumentStart = argumentIndex + 1
            End If
        End If
        argumentIndex = argumentIndex + 1
    Loop

    argumentText = Trim$(Mid$(normalizedFormula, argumentStart, closePosition - argumentStart))
    If Len(argumentText) = 0 Then Exit Function
    arguments(UBound(arguments)) = argumentText
    If Not FAS_IsValidIFSArgumentCount(functionName, UBound(arguments)) Then Exit Function
    FAS_TryParseIFSFormula = True
End Function

Private Function FAS_IsValidIFSArgumentCount(ByVal functionName As String, ByVal argumentCount As Long) As Boolean
    If functionName = "COUNTIFS" Then
        FAS_IsValidIFSArgumentCount = argumentCount >= 2 And argumentCount Mod 2 = 0
    Else
        FAS_IsValidIFSArgumentCount = argumentCount >= 3 And argumentCount Mod 2 = 1
    End If
End Function

Private Function FAS_IsSupportedIFSFunction(ByVal functionName As String) As Boolean
    Select Case UCase$(functionName)
        Case "SUMIFS", "COUNTIFS", "AVERAGEIFS", "MINIFS", "MAXIFS"
            FAS_IsSupportedIFSFunction = True
    End Select
End Function

Private Function FAS_IsMeasureIFSFunction(ByVal functionName As String) As Boolean
    FAS_IsMeasureIFSFunction = UCase$(functionName) <> "COUNTIFS" And FAS_IsSupportedIFSFunction(functionName)
End Function

Private Function FAS_JoinIFSArguments(ByRef arguments() As String) As String
    Dim argumentIndex As Long

    For argumentIndex = LBound(arguments) To UBound(arguments)
        If Len(FAS_JoinIFSArguments) > 0 Then FAS_JoinIFSArguments = FAS_JoinIFSArguments & ","
        FAS_JoinIFSArguments = FAS_JoinIFSArguments & arguments(argumentIndex)
    Next argumentIndex
End Function
