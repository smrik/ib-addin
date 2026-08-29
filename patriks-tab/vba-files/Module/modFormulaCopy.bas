Attribute VB_Name = "modFormulaCopy"
Option Explicit

Public Sub FAS_CopyFormulasSmart()
    Dim destinationCell As Range

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a source range before smart formula copy."
        Exit Sub
    End If

    On Error Resume Next
    Set destinationCell = Application.InputBox("Select the top-left destination cell:", "FAS Smart Formula Copy", Type:=8)
    On Error GoTo 0

    If destinationCell Is Nothing Then
        Exit Sub
    End If

    If FAS_CopySelectionFormulasSmartToRange(destinationCell.cells(1, 1)) Then
        FAS_SetStatus "Smart formula copy complete."
    End If
End Sub

Public Function FAS_TryParseFormulaCopyCommand(ByVal commandText As String, ByRef destinationAddress As String) As Boolean
    Dim rawText As String
    Dim normalizedText As String

    rawText = Trim$(commandText)
    normalizedText = FAS_NormalizeSheetText(rawText)

    If Left$(normalizedText, 6) = "fcopy " Then
        destinationAddress = FAS_GetFormulaCopyTextAfterPrefix(rawText, "fcopy")
    ElseIf Left$(normalizedText, 10) = "smartcopy " Then
        destinationAddress = FAS_GetFormulaCopyTextAfterPrefix(rawText, "smartcopy")
    ElseIf Left$(normalizedText, 12) = "formulacopy " Then
        destinationAddress = FAS_GetFormulaCopyTextAfterPrefix(rawText, "formulacopy")
    End If

    FAS_TryParseFormulaCopyCommand = (Len(Trim$(destinationAddress)) > 0)
End Function

Public Function FAS_CopySelectionFormulasSmartToAddress(ByVal destinationAddress As String) As Boolean
    Dim destinationCell As Range

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select a source range before smart formula copy."
        Exit Function
    End If

    On Error Resume Next
    Set destinationCell = ActiveSheet.Range(destinationAddress).cells(1, 1)
    On Error GoTo 0

    If destinationCell Is Nothing Then
        FAS_SetStatus "Invalid smart-copy destination: " & destinationAddress
        Exit Function
    End If

    FAS_CopySelectionFormulasSmartToAddress = FAS_CopySelectionFormulasSmartToRange(destinationCell)
End Function

Public Function FAS_CopySelectionFormulasSmartToRange(ByVal destinationCell As Range) As Boolean
    Dim sourceRange As Range
    Dim destinationRange As Range
    Dim sourceCell As Range
    Dim targetCell As Range
    Dim rowOffset As Long
    Dim columnOffset As Long
    Dim response As VbMsgBoxResult

    On Error GoTo CopyFailed

    If TypeName(Selection) <> "Range" Or destinationCell Is Nothing Then
        FAS_SetStatus "Select a source range and valid destination first."
        Exit Function
    End If

    Set sourceRange = Selection
    If sourceRange.Areas.Count > 1 Then
        If Application.Visible Then
            FAS_SetStatus "Smart formula copy needs one contiguous source range."
        End If
        Exit Function
    End If

    Set destinationCell = destinationCell.cells(1, 1)
    If destinationCell.Row + sourceRange.Rows.Count - 1 > destinationCell.Worksheet.Rows.Count Or _
        destinationCell.Column + sourceRange.Columns.Count - 1 > destinationCell.Worksheet.Columns.Count Then
        If Application.Visible Then
            FAS_SetStatus "Smart-copy destination exceeds worksheet bounds."
        End If
        Exit Function
    End If

    Set destinationRange = destinationCell.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)

    If sourceRange.Worksheet Is destinationRange.Worksheet Then
        If Not Intersect(sourceRange, destinationRange) Is Nothing Then
            If Application.Visible Then
                FAS_SetStatus "Smart-copy destination overlaps source range."
            End If
            Exit Function
        End If
    End If

    If Not FAS_FormulasAreSafeForSmartCopy(sourceRange) Then
        If Application.Visible Then
            FAS_SetStatus "Smart copy skipped: unsupported formula references."
        End If
        Exit Function
    End If

    If Application.WorksheetFunction.CountA(destinationRange) > 0 And Application.Visible Then
        response = MsgBox("Overwrite existing cells in " & destinationRange.Address(False, False) & "?", vbQuestion + vbOKCancel, "FAS Formula Copy")
        If response <> vbOK Then
            Exit Function
        End If
    End If

    For Each sourceCell In sourceRange.cells
        rowOffset = sourceCell.Row - sourceRange.Row
        columnOffset = sourceCell.Column - sourceRange.Column
        Set targetCell = destinationCell.Offset(rowOffset, columnOffset)

        targetCell.NumberFormat = sourceCell.NumberFormat
        targetCell.Interior.Color = sourceCell.Interior.Color
        targetCell.Font.Bold = sourceCell.Font.Bold
        targetCell.Font.Italic = sourceCell.Font.Italic
        targetCell.Font.Color = sourceCell.Font.Color
        If sourceCell.HasFormula Then
            targetCell.Formula = FAS_RewriteFormulaForSmartCopy(sourceCell, sourceRange, destinationCell)
        Else
            targetCell.value = sourceCell.value
        End If
    Next sourceCell

    FAS_CopySelectionFormulasSmartToRange = True
    FAS_SetStatus "Smart-copied " & sourceRange.Address(False, False) & " to " & destinationRange.Address(False, False) & "."
    Exit Function

CopyFailed:
    If Not Application.Visible Then
        Err.Raise Err.Number, Err.Source, Err.description
    End If

    FAS_SetStatus "Smart formula copy failed: " & Err.description
End Function

Private Function FAS_GetFormulaCopyTextAfterPrefix(ByVal rawText As String, ByVal prefixText As String) As String
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

    FAS_GetFormulaCopyTextAfterPrefix = result
End Function

Private Function FAS_FormulasAreSafeForSmartCopy(ByVal sourceRange As Range) As Boolean
    Dim sourceCell As Range
    Dim formulaText As String
    Dim cleanedFormula As String

    FAS_FormulasAreSafeForSmartCopy = True
    For Each sourceCell In sourceRange.cells
        If sourceCell.HasFormula Then
            formulaText = sourceCell.Formula
            cleanedFormula = FAS_BlankFormulaStringsForCopy(formulaText)
            If InStr(1, cleanedFormula, "[", vbBinaryCompare) > 0 Or _
                InStr(1, cleanedFormula, "#", vbBinaryCompare) > 0 Or _
                InStr(1, cleanedFormula, "INDIRECT(", vbTextCompare) > 0 Or _
                InStr(1, cleanedFormula, "OFFSET(", vbTextCompare) > 0 Or _
                FAS_FormulaHasStructuredReference(cleanedFormula) Or _
                FAS_FormulaHas3DReference(cleanedFormula) Then
                FAS_FormulasAreSafeForSmartCopy = False
                Exit Function
            End If
        End If
    Next sourceCell
End Function

Private Function FAS_RewriteFormulaForSmartCopy(ByVal sourceCell As Range, ByVal sourceRange As Range, ByVal destinationTopLeft As Range) As String
    Dim formulaText As String
    Dim regex As Object
    Dim matches As Object
    Dim matchIndex As Long
    Dim matchItem As Object
    Dim matchStart As Long
    Dim replacementText As String

    formulaText = sourceCell.Formula
    formulaText = FAS_RewriteCrossSheetReferencesForSmartCopy(formulaText, sourceCell.Worksheet)

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\$?[A-Z]{1,3}\$?[0-9]+(:\$?[A-Z]{1,3}\$?[0-9]+)?"

    Set matches = regex.Execute(FAS_BlankFormulaStringsForCopy(formulaText))
    For matchIndex = matches.Count - 1 To 0 Step -1
        Set matchItem = matches.item(matchIndex)
        matchStart = CLng(matchItem.FirstIndex) + 1
        If FAS_IsStandaloneCopyReference(formulaText, matchStart, Len(matchItem.value)) Then
            replacementText = FAS_GetSmartCopyReplacement(sourceCell, sourceRange, destinationTopLeft, matchItem.value)
            If Len(replacementText) > 0 Then
                formulaText = Left$(formulaText, matchStart - 1) & replacementText & Mid$(formulaText, matchStart + Len(matchItem.value))
            End If
        End If
    Next matchIndex

    FAS_RewriteFormulaForSmartCopy = formulaText
End Function

Private Function FAS_RewriteCrossSheetReferencesForSmartCopy(ByVal formulaText As String, ByVal sourceSheet As Worksheet) As String
    Dim regex As Object
    Dim matches As Object
    Dim matchIndex As Long
    Dim matchItem As Object
    Dim sheetName As String
    Dim addressText As String
    Dim targetRange As Range
    Dim replacementText As String
    Dim matchStart As Long

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "('([^']|'')+'|[A-Za-z0-9_ .&-]+)!\$?[A-Z]{1,3}\$?[0-9]+(:\$?[A-Z]{1,3}\$?[0-9]+)?"

    Set matches = regex.Execute(FAS_BlankFormulaStringsForCopy(formulaText))
    For matchIndex = matches.Count - 1 To 0 Step -1
        Set matchItem = matches.item(matchIndex)
        FAS_SplitFormulaCopySheetReference matchItem.value, sheetName, addressText

        On Error Resume Next
        Set targetRange = sourceSheet.Parent.Worksheets(sheetName).Range(addressText)
        On Error GoTo 0

        If Not targetRange Is Nothing Then
            replacementText = FAS_QuoteSheetName(targetRange.Worksheet.Name) & "!" & targetRange.Address(True, True)
            matchStart = CLng(matchItem.FirstIndex) + 1
            formulaText = Left$(formulaText, matchStart - 1) & replacementText & Mid$(formulaText, matchStart + Len(matchItem.value))
        End If
        Set targetRange = Nothing
    Next matchIndex

    FAS_RewriteCrossSheetReferencesForSmartCopy = formulaText
End Function

Private Function FAS_GetSmartCopyReplacement(ByVal sourceCell As Range, ByVal sourceRange As Range, ByVal destinationTopLeft As Range, ByVal referenceText As String) As String
    Dim referencedRange As Range
    Dim movedRange As Range
    Dim rowOffset As Long
    Dim columnOffset As Long

    On Error Resume Next
    Set referencedRange = sourceCell.Worksheet.Range(referenceText)
    On Error GoTo 0

    If referencedRange Is Nothing Then
        Exit Function
    End If

    If FAS_RangeIsFullyInside(referencedRange, sourceRange) Then
        rowOffset = referencedRange.Row - sourceRange.Row
        columnOffset = referencedRange.Column - sourceRange.Column
        Set movedRange = destinationTopLeft.Worksheet.Range( _
            destinationTopLeft.Offset(rowOffset, columnOffset), _
            destinationTopLeft.Offset(rowOffset + referencedRange.Rows.Count - 1, columnOffset + referencedRange.Columns.Count - 1))
        FAS_GetSmartCopyReplacement = movedRange.Address(False, False)
    ElseIf destinationTopLeft.Worksheet Is sourceCell.Worksheet Then
        FAS_GetSmartCopyReplacement = referencedRange.Address(True, True)
    Else
        FAS_GetSmartCopyReplacement = FAS_QuoteSheetName(sourceCell.Worksheet.Name) & "!" & referencedRange.Address(True, True)
    End If
End Function

Private Sub FAS_SplitFormulaCopySheetReference(ByVal referenceText As String, ByRef sheetName As String, ByRef addressText As String)
    Dim bangPosition As Long

    bangPosition = InStrRev(referenceText, "!")
    sheetName = Left$(referenceText, bangPosition - 1)
    addressText = Mid$(referenceText, bangPosition + 1)

    If Left$(sheetName, 1) = "'" And Right$(sheetName, 1) = "'" Then
        sheetName = Mid$(sheetName, 2, Len(sheetName) - 2)
        sheetName = Replace(sheetName, "''", "'")
    End If
End Sub

Private Function FAS_FormulaHasStructuredReference(ByVal formulaText As String) As Boolean
    FAS_FormulaHasStructuredReference = (InStr(1, formulaText, "[#", vbBinaryCompare) > 0 Or _
        InStr(1, formulaText, "[@", vbBinaryCompare) > 0 Or _
        InStr(1, formulaText, "[[", vbBinaryCompare) > 0)
End Function

Private Function FAS_FormulaHas3DReference(ByVal formulaText As String) As Boolean
    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "('([^']|'')+'|[A-Za-z0-9_ .&-]+):('([^']|'')+'|[A-Za-z0-9_ .&-]+)!"

    FAS_FormulaHas3DReference = regex.Test(formulaText)
End Function

Private Function FAS_RangeIsFullyInside(ByVal candidateRange As Range, ByVal containerRange As Range) As Boolean
    If candidateRange.Worksheet Is Nothing Or containerRange.Worksheet Is Nothing Then
        Exit Function
    End If

    If Not (candidateRange.Worksheet Is containerRange.Worksheet) Then
        Exit Function
    End If

    FAS_RangeIsFullyInside = (candidateRange.Row >= containerRange.Row And _
        candidateRange.Column >= containerRange.Column And _
        candidateRange.Row + candidateRange.Rows.Count - 1 <= containerRange.Row + containerRange.Rows.Count - 1 And _
        candidateRange.Column + candidateRange.Columns.Count - 1 <= containerRange.Column + containerRange.Columns.Count - 1)
End Function

Private Function FAS_IsStandaloneCopyReference(ByVal formulaText As String, ByVal startPosition As Long, ByVal matchLength As Long) As Boolean
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

    FAS_IsStandaloneCopyReference = True
End Function

Private Function FAS_BlankFormulaStringsForCopy(ByVal formulaText As String) As String
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

    FAS_BlankFormulaStringsForCopy = result
End Function

Private Function FAS_QuoteSheetName(ByVal sheetName As String) As String
    FAS_QuoteSheetName = "'" & Replace(sheetName, "'", "''") & "'"
End Function
