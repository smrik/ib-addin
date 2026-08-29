Attribute VB_Name = "modFormulaTools"
Option Explicit

Public Sub FAS_ReplaceInSelectedFormulas()
    Dim selectedRange As Range
    Dim formulaCells As Range
    Dim currentCell As Range
    Dim findText As String
    Dim replaceText As String
    Dim proposedFormula As String
    Dim changeCount As Long
    Dim skippedArrayCount As Long
    Dim response As VbMsgBoxResult

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a range before replacing formula text."
        Exit Sub
    End If

    Set selectedRange = Selection

    findText = InputBox("Find this text inside formulas in the current selection:", "FAS Replace In Formulas")
    If Len(findText) = 0 Then
        Exit Sub
    End If

    replaceText = InputBox("Replace it with:", "FAS Replace In Formulas")
    If StrPtr(replaceText) = 0 Then
        Exit Sub
    End If

    On Error Resume Next
    Set formulaCells = selectedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaCells Is Nothing Then
        FAS_SetStatus "Selection contains no formulas."
        Exit Sub
    End If

    For Each currentCell In formulaCells.Cells
        If currentCell.HasArray Then
            skippedArrayCount = skippedArrayCount + 1
        Else
            proposedFormula = Replace(currentCell.Formula, findText, replaceText, 1, -1, vbTextCompare)
            If StrComp(proposedFormula, currentCell.Formula, vbBinaryCompare) <> 0 Then
                changeCount = changeCount + 1
            End If
        End If
    Next currentCell

    If changeCount = 0 Then
        FAS_SetStatus "No selected formulas contain '" & findText & "'."
        Exit Sub
    End If

    response = MsgBox( _
        "Replace '" & findText & "' with '" & replaceText & "' in " & changeCount & " formula cell(s)?" & vbCrLf & _
        IIf(skippedArrayCount > 0, vbCrLf & skippedArrayCount & " array formula cell(s) will be skipped.", vbNullString), _
        vbQuestion + vbOKCancel, _
        "FAS Replace In Formulas")

    If response <> vbOK Then
        Exit Sub
    End If

    changeCount = 0
    For Each currentCell In formulaCells.Cells
        If Not currentCell.HasArray Then
            proposedFormula = Replace(currentCell.Formula, findText, replaceText, 1, -1, vbTextCompare)
            If StrComp(proposedFormula, currentCell.Formula, vbBinaryCompare) <> 0 Then
                currentCell.Formula = proposedFormula
                changeCount = changeCount + 1
            End If
        End If
    Next currentCell

    FAS_SetStatus changeCount & " formula cell(s) updated." & IIf(skippedArrayCount > 0, " " & skippedArrayCount & " array formula cell(s) skipped.", vbNullString)
End Sub

' Run a formula wrapper command. An explicit round value becomes the next default.
Public Function FAS_TryRunFormulaWrapperCommand(ByVal commandText As String) As Boolean
    Dim rawText As String
    Dim normalizedText As String
    Dim restText As String
    Dim decimalPlaces As Long

    rawText = Trim$(commandText)
    normalizedText = LCase$(FAS_NormalizeCommandText(rawText))

    If Left$(LCase$(rawText), 6) = "round " Then
        restText = Trim$(Mid$(rawText, 7))
        If Not FAS_TryParseWholeNumber(restText, decimalPlaces) Then Exit Function
        FAS_WrapSelectedFormulasRound decimalPlaces
        FAS_TryRunFormulaWrapperCommand = True
        Exit Function
    End If

    Select Case normalizedText
        Case "round", "roundformula", "wrapround"
            FAS_WrapSelectedFormulasRound
            FAS_TryRunFormulaWrapperCommand = True
        Case "iferror", "errorblank", "wrapiferror"
            FAS_WrapSelectedFormulasIfError
            FAS_TryRunFormulaWrapperCommand = True
    End Select
End Function

Public Function FAS_IsFormulaWrapperCommand(ByVal commandText As String) As Boolean
    Dim normalizedText As String
    normalizedText = LCase$(FAS_NormalizeCommandText(commandText))
    FAS_IsFormulaWrapperCommand = (normalizedText = "round" Or normalizedText = "roundformula" Or normalizedText = "wrapround" Or _
        normalizedText = "iferror" Or normalizedText = "errorblank" Or normalizedText = "wrapiferror" Or Left$(LCase$(Trim$(commandText)), 6) = "round ")
End Function

Public Sub FAS_WrapSelectedFormulasRound(Optional ByVal decimalPlaces As Long = -1)
    Dim formulaCells As Range
    Dim currentCell As Range
    Dim formulaText As String
    Dim wrappedFormula As String
    Dim changedCount As Long
    Dim skippedCount As Long
    Dim response As VbMsgBoxResult

    If decimalPlaces < 0 Then
        decimalPlaces = FAS_GetRoundDigits()
    Else
        If decimalPlaces > 15 Then decimalPlaces = 15
        FAS_SetSetting "Formula.RoundDigits", CStr(decimalPlaces)
    End If

    Set formulaCells = FAS_GetSelectedFormulaCells()
    If formulaCells Is Nothing Then Exit Sub

    For Each currentCell In formulaCells.Cells
        If currentCell.HasArray Then
            skippedCount = skippedCount + 1
        Else
            formulaText = CStr(currentCell.Formula)
            wrappedFormula = FAS_WrapFormula(formulaText, "ROUND", CStr(decimalPlaces))
            If StrComp(formulaText, wrappedFormula, vbBinaryCompare) <> 0 Then changedCount = changedCount + 1
        End If
    Next currentCell

    If changedCount = 0 Then
        FAS_SetStatus "No selected formulas needed ROUND."
        Exit Sub
    End If

    If Application.Visible Then
        response = MsgBox("Wrap " & changedCount & " formula cell(s) with ROUND(..., " & decimalPlaces & ")?", vbQuestion + vbOKCancel, "FAS Round Formulas")
        If response <> vbOK Then Exit Sub
    End If

    changedCount = 0
    For Each currentCell In formulaCells.Cells
        If Not currentCell.HasArray Then
            formulaText = CStr(currentCell.Formula)
            wrappedFormula = FAS_WrapFormula(formulaText, "ROUND", CStr(decimalPlaces))
            If StrComp(formulaText, wrappedFormula, vbBinaryCompare) <> 0 Then
                currentCell.Formula = wrappedFormula
                changedCount = changedCount + 1
            End If
        End If
    Next currentCell

    FAS_SetStatus changedCount & " formula cell(s) wrapped with ROUND." & FAS_SkippedFormulaText(skippedCount)
End Sub

Public Sub FAS_WrapSelectedFormulasIfError()
    Dim formulaCells As Range
    Dim currentCell As Range
    Dim formulaText As String
    Dim wrappedFormula As String
    Dim changedCount As Long
    Dim skippedCount As Long
    Dim response As VbMsgBoxResult

    Set formulaCells = FAS_GetSelectedFormulaCells()
    If formulaCells Is Nothing Then Exit Sub

    For Each currentCell In formulaCells.Cells
        If currentCell.HasArray Then
            skippedCount = skippedCount + 1
        Else
            formulaText = CStr(currentCell.Formula)
            wrappedFormula = FAS_WrapFormula(formulaText, "IFERROR", Chr$(34) & Chr$(34))
            If StrComp(formulaText, wrappedFormula, vbBinaryCompare) <> 0 Then changedCount = changedCount + 1
        End If
    Next currentCell

    If changedCount = 0 Then
        FAS_SetStatus "No selected formulas needed IFERROR."
        Exit Sub
    End If

    If Application.Visible Then
        response = MsgBox("Wrap " & changedCount & " formula cell(s) with IFERROR(..., """")?", vbQuestion + vbOKCancel, "FAS IFERROR Formulas")
        If response <> vbOK Then Exit Sub
    End If

    changedCount = 0
    For Each currentCell In formulaCells.Cells
        If Not currentCell.HasArray Then
            formulaText = CStr(currentCell.Formula)
            wrappedFormula = FAS_WrapFormula(formulaText, "IFERROR", Chr$(34) & Chr$(34))
            If StrComp(formulaText, wrappedFormula, vbBinaryCompare) <> 0 Then
                currentCell.Formula = wrappedFormula
                changedCount = changedCount + 1
            End If
        End If
    Next currentCell

    FAS_SetStatus changedCount & " formula cell(s) wrapped with IFERROR." & FAS_SkippedFormulaText(skippedCount)
End Sub

Private Function FAS_GetSelectedFormulaCells() As Range
    Dim selectedRange As Range

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a range containing formulas first."
        Exit Function
    End If

    Set selectedRange = Selection
    On Error Resume Next
    Set FAS_GetSelectedFormulaCells = selectedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If FAS_GetSelectedFormulaCells Is Nothing Then
        FAS_SetStatus "Selection contains no formulas."
    End If
End Function

Private Function FAS_GetRoundDigits() As Long
    Dim savedValue As String
    Dim parsedValue As Long

    savedValue = FAS_GetSetting("Formula.RoundDigits", "0")
    If FAS_TryParseWholeNumber(savedValue, parsedValue) Then
        FAS_GetRoundDigits = parsedValue
    Else
        FAS_GetRoundDigits = 0
    End If
End Function

Private Function FAS_TryParseWholeNumber(ByVal valueText As String, ByRef parsedValue As Long) As Boolean
    Dim numericValue As Double

    If Len(Trim$(valueText)) = 0 Then Exit Function
    If Not IsNumeric(valueText) Then Exit Function
    numericValue = CDbl(valueText)
    If numericValue <> Fix(numericValue) Then Exit Function
    If numericValue < 0 Or numericValue > 15 Then Exit Function
    parsedValue = CLng(numericValue)
    FAS_TryParseWholeNumber = True
End Function

Private Function FAS_WrapFormula(ByVal formulaText As String, ByVal wrapperName As String, ByVal secondArgument As String) As String
    Dim innerFormula As String
    Dim expectedPrefix As String

    If Len(formulaText) = 0 Or Left$(formulaText, 1) <> "=" Then
        FAS_WrapFormula = formulaText
        Exit Function
    End If

    expectedPrefix = "=" & wrapperName & "("
    If UCase$(Left$(formulaText, Len(expectedPrefix))) = expectedPrefix Then
        FAS_WrapFormula = formulaText
        Exit Function
    End If

    innerFormula = Mid$(formulaText, 2)
    FAS_WrapFormula = expectedPrefix & innerFormula & "," & secondArgument & ")"
End Function

Private Function FAS_SkippedFormulaText(ByVal skippedCount As Long) As String
    If skippedCount > 0 Then FAS_SkippedFormulaText = " " & skippedCount & " array formula cell(s) skipped."
End Function
