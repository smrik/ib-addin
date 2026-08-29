Attribute VB_Name = "modFormulaTools"
Option Explicit

Private Const FAS_ROUND_DIGITS_SETTING As String = "Formula.RoundDigits"

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

    For Each currentCell In formulaCells.cells
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
    For Each currentCell In formulaCells.cells
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

' Palette entry point for the formula wrappers. Supports a bare command
' ("round", "iferror") and a parameterized round ("round 2").
Public Function FAS_TryRunFormulaWrapperCommand(ByVal commandText As String) As Boolean
    Dim rawText As String
    Dim normalizedText As String
    Dim restText As String
    Dim decimalPlaces As Long

    rawText = Trim$(commandText)
    normalizedText = LCase$(FAS_NormalizeCommandText(rawText))

    If LCase$(Left$(rawText, 6)) = "round " Then
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
    Dim rawText As String
    Dim decimalPlaces As Long

    rawText = Trim$(commandText)
    If LCase$(Left$(rawText, 6)) = "round " Then
        FAS_IsFormulaWrapperCommand = FAS_TryParseWholeNumber(Trim$(Mid$(rawText, 7)), decimalPlaces)
        Exit Function
    End If

    Select Case LCase$(FAS_NormalizeCommandText(rawText))
        Case "round", "roundformula", "wrapround", "iferror", "errorblank", "wrapiferror"
            FAS_IsFormulaWrapperCommand = True
    End Select
End Function

' decimalPlaces = -1 means "use the saved Formula.RoundDigits setting".
Public Sub FAS_WrapSelectedFormulasRound(Optional ByVal decimalPlaces As Long = -1)
    Dim roundDigits As Long

    If decimalPlaces >= 0 Then
        roundDigits = decimalPlaces
        ' Remember an explicit "round 2" so a later bare "round" reuses it.
        FAS_SetSetting FAS_ROUND_DIGITS_SETTING, CStr(roundDigits)
    Else
        roundDigits = Val(FAS_GetSetting(FAS_ROUND_DIGITS_SETTING, "0"))
    End If

    FAS_WrapSelectedFormulas "ROUND", CStr(roundDigits)
End Sub

Private Function FAS_TryParseWholeNumber(ByVal valueText As String, ByRef parsedValue As Long) As Boolean
    Dim trimmedText As String
    Dim characterIndex As Long
    Dim currentCharacter As String

    trimmedText = Trim$(valueText)
    If Len(trimmedText) = 0 Then Exit Function

    For characterIndex = 1 To Len(trimmedText)
        currentCharacter = Mid$(trimmedText, characterIndex, 1)
        If currentCharacter < "0" Or currentCharacter > "9" Then Exit Function
    Next characterIndex

    parsedValue = CLng(trimmedText)
    FAS_TryParseWholeNumber = True
End Function

Public Sub FAS_WrapSelectedFormulasIfError()
    ' Six quotes produce the two-character literal "" so the result is
    ' =IFERROR(<formula>,""). Four quotes yield a single " and build a
    ' malformed formula that Excel rejects.
    FAS_WrapSelectedFormulas "IFERROR", """"""
End Sub

Private Sub FAS_WrapSelectedFormulas(ByVal wrapperName As String, ByVal secondArgument As String)
    Dim selectedRange As Range
    Dim formulaCells As Range
    Dim currentCell As Range
    Dim formulaText As String
    Dim wrappedCount As Long

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select formula cells first."
        Exit Sub
    End If

    Set selectedRange = Selection
    On Error Resume Next
    Set formulaCells = selectedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaCells Is Nothing Then
        FAS_SetStatus "Selection contains no formulas."
        Exit Sub
    End If

    If Not FAS_UndoBeginRange(selectedRange, "Wrap selected formulas") Then Exit Sub

    On Error GoTo WrapFailed
    For Each currentCell In formulaCells.Cells
        If Not currentCell.HasArray Then
            formulaText = CStr(currentCell.Formula)
            If UCase$(Left$(formulaText, Len(wrapperName) + 2)) <> "=" & wrapperName & "(" Then
                currentCell.Formula = "=" & wrapperName & "(" & Mid$(formulaText, 2) & "," & secondArgument & ")"
                wrappedCount = wrappedCount + 1
            End If
        End If
    Next currentCell

    FAS_UndoCommit
    FAS_SetStatus wrappedCount & " formula cell(s) wrapped with " & wrapperName & "."
    Exit Sub

WrapFailed:
    FAS_UndoCancel
    FAS_SetStatus "Could not wrap the selected formulas."
End Sub
