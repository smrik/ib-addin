Attribute VB_Name = "modPrecedentTrace"
Option Explicit

Private mTraceOriginWorkbookFullName As String
Private mTraceOriginSheetName As String
Private mTraceOriginAddress As String
Private mCycleSourceWorkbookFullName As String
Private mCycleSourceSheetName As String
Private mCycleSourceAddress As String
Private mCycleSourceFormula As String
Private mCycleLastTargetWorkbookFullName As String
Private mCycleLastTargetSheetName As String
Private mCycleLastTargetAddress As String
Private mCyclePrecedents As Collection
Private mCycleIndex As Long

Public Sub FAS_ShowPrecedents()
    FAS_ShowPrecedentsForSelection
End Sub

Public Function FAS_ShowPrecedentsForSelection() As Boolean
    Dim sourceCell As Range
    Dim precedents As Collection
    Dim prompt As String
    Dim userInput As String
    Dim selectedIndex As Long

    Set sourceCell = FAS_GetTraceSourceCell(True)
    If sourceCell Is Nothing Then
        Exit Function
    End If

    Set precedents = FAS_ParsePrecedentRanges(sourceCell)
    If precedents.Count = 0 Then
        FAS_SetStatus "No supported direct precedents for " & sourceCell.Address(False, False) & "."
        Exit Function
    End If

    prompt = FAS_BuildPrecedentPrompt(sourceCell, precedents)
    userInput = InputBox(prompt, "FAS Precedents")
    If Len(Trim$(userInput)) = 0 Then
        Exit Function
    End If

    If Not IsNumeric(userInput) Then
        FAS_SetStatus "Enter a precedent number from the list."
        Exit Function
    End If

    selectedIndex = CLng(userInput)
    If selectedIndex < 1 Or selectedIndex > precedents.Count Then
        FAS_SetStatus "Precedent number is outside the list."
        Exit Function
    End If

    FAS_SaveTraceOrigin sourceCell
    FAS_ActivatePrecedentRange precedents(selectedIndex)
    FAS_ShowPrecedentsForSelection = True
End Function

Public Function FAS_CountPrecedentsInActiveCell() As Long
    Dim sourceCell As Range
    Dim precedents As Collection

    Set sourceCell = FAS_GetTraceSourceCell(False)
    If sourceCell Is Nothing Then
        Exit Function
    End If

    Set precedents = FAS_ParsePrecedentRanges(sourceCell)
    FAS_CountPrecedentsInActiveCell = precedents.Count
End Function

Public Sub FAS_CycleDirectPrecedent()
    Dim sourceCell As Range
    Dim targetRange As Range

    On Error GoTo CycleFailed

    If Not FAS_CanContinuePrecedentCycle() Then
        Set sourceCell = FAS_GetTraceSourceCell(True)
        If sourceCell Is Nothing Then
            FAS_ResetPrecedentCycle
            Exit Sub
        End If

        FAS_StartPrecedentCycle sourceCell
        If mCyclePrecedents Is Nothing Then Exit Sub
        If mCyclePrecedents.Count = 0 Then Exit Sub
    End If

    If Not FAS_IsCycleSourceCurrent() Then
        FAS_SetStatus "The formula changed; start the precedent trace again."
        FAS_ResetPrecedentCycle
        Exit Sub
    End If

    mCycleIndex = mCycleIndex + 1
    If mCycleIndex > mCyclePrecedents.Count Then mCycleIndex = 1

    Set targetRange = mCyclePrecedents(mCycleIndex)
    FAS_ActivatePrecedentRange targetRange
    FAS_SaveCycleLastTarget targetRange
    FAS_SetStatus "Precedent " & mCycleIndex & " of " & mCyclePrecedents.Count & ": " & _
        targetRange.Worksheet.Name & "!" & targetRange.Address(False, False)
    Exit Sub

CycleFailed:
    FAS_ResetPrecedentCycle
    FAS_SetStatus "The precedent trace became stale. Select the formula cell and try again."
End Sub

Public Function FAS_JumpToPrecedentByIndex(ByVal precedentIndex As Long) As Boolean
    Dim sourceCell As Range
    Dim precedents As Collection

    Set sourceCell = FAS_GetTraceSourceCell(False)
    If sourceCell Is Nothing Then
        Exit Function
    End If

    Set precedents = FAS_ParsePrecedentRanges(sourceCell)
    If precedentIndex < 1 Or precedentIndex > precedents.Count Then
        Exit Function
    End If

    FAS_SaveTraceOrigin sourceCell
    FAS_ActivatePrecedentRange precedents(precedentIndex)
    FAS_SetStatus "Precedent " & precedentIndex & " of " & precedents.Count & ": " & precedents(precedentIndex).Worksheet.Name & "!" & precedents(precedentIndex).Address(False, False)
    FAS_JumpToPrecedentByIndex = True
End Function

Public Function FAS_TryParsePrecedentCommand(ByVal commandText As String, ByRef precedentIndex As Long) As Boolean
    Dim normalizedCommand As String
    Dim digitText As String

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    If normalizedCommand = "pre" Or normalizedCommand = "precedent" Or normalizedCommand = "precedents" Then
        precedentIndex = 1
        FAS_TryParsePrecedentCommand = True
        Exit Function
    End If

    If Left$(normalizedCommand, 3) = "pre" Then
        digitText = Mid$(normalizedCommand, 4)
    ElseIf Left$(normalizedCommand, 9) = "precedent" Then
        digitText = Mid$(normalizedCommand, 10)
    End If

    If Len(digitText) = 0 Or Not IsNumeric(digitText) Then
        Exit Function
    End If

    precedentIndex = CLng(digitText)
    If precedentIndex < 1 Then
        Exit Function
    End If

    FAS_TryParsePrecedentCommand = True
End Function

Public Function FAS_GoBackToTraceOrigin() As Boolean
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim targetRange As Range

    If Len(mTraceOriginWorkbookFullName) = 0 Or Len(mTraceOriginSheetName) = 0 Or Len(mTraceOriginAddress) = 0 Then
        FAS_SetStatus "No precedent trace origin stored yet."
        Exit Function
    End If

    Set targetWorkbook = FAS_GetOpenWorkbookByFullName(mTraceOriginWorkbookFullName)
    If targetWorkbook Is Nothing Then
        FAS_SetStatus "Trace origin workbook is no longer open."
        Exit Function
    End If

    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets(mTraceOriginSheetName)
    Set targetRange = targetSheet.Range(mTraceOriginAddress)
    On Error GoTo 0

    If targetRange Is Nothing Then
        FAS_SetStatus "Trace origin cell could not be found."
        Exit Function
    End If

    targetWorkbook.Activate
    targetSheet.Activate
    targetRange.Select
    FAS_SetStatus "Returned to " & targetSheet.Name & "!" & targetRange.Address(False, False) & "."
    FAS_ResetPrecedentCycle
    FAS_GoBackToTraceOrigin = True
End Function

Public Sub FAS_BackToTraceOrigin()
    FAS_GoBackToTraceOrigin
End Sub

Private Function FAS_GetTraceSourceCell(ByVal showMessages As Boolean) As Range
    Dim selectedRange As Range
    Dim sourceCell As Range

    If TypeName(Selection) <> "Range" Then
        If showMessages Then
            FAS_SetStatus "Select a formula cell before tracing precedents."
        End If
        Exit Function
    End If

    Set selectedRange = Selection
    Set sourceCell = ActiveCell
    If sourceCell Is Nothing Then
        Set sourceCell = selectedRange.Cells(1, 1)
    ElseIf Intersect(sourceCell, selectedRange) Is Nothing Then
        Set sourceCell = selectedRange.Cells(1, 1)
    End If

    If sourceCell.Cells.CountLarge > 1 Then
        Set sourceCell = sourceCell.Cells(1, 1)
    End If

    If Not sourceCell.HasFormula Then
        If showMessages Then
            FAS_SetStatus "Active cell does not contain a formula."
        End If
        Exit Function
    End If

    Set FAS_GetTraceSourceCell = sourceCell
End Function

Private Function FAS_ParsePrecedentRanges(ByVal sourceCell As Range) As Collection
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
        Set FAS_ParsePrecedentRanges = precedents
        Exit Function
    End If

    formulaText = FAS_BlankQuotedStrings(sourceCell.Formula)
    If FAS_FormulaHasUnsupportedPrecedents(formulaText) Then
        Set FAS_ParsePrecedentRanges = precedents
        Exit Function
    End If

    Set crossSheetRegex = CreateObject("VBScript.RegExp")
    crossSheetRegex.Global = True
    crossSheetRegex.IgnoreCase = True
    crossSheetRegex.Pattern = "('([^']|'')+'|[A-Za-z0-9_ .&-]+)!\$?[A-Z]{1,3}\$?[0-9]+(:\$?[A-Z]{1,3}\$?[0-9]+)?"

    Set matches = crossSheetRegex.Execute(formulaText)
    For Each matchItem In matches
        If InStr(1, matchItem.Value, "[", vbBinaryCompare) = 0 Then
            FAS_SplitSheetReference matchItem.Value, sheetName, addressText
            FAS_AddPrecedentIfValid precedents, sourceCell.Worksheet.Parent, sheetName, addressText
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
        If FAS_IsStandaloneA1Match(formulaText, matchStart, Len(matchItem.Value)) Then
            FAS_AddPrecedentIfValid precedents, sourceCell.Worksheet.Parent, sourceCell.Worksheet.Name, matchItem.Value
        End If
    Next matchItem

    Set FAS_ParsePrecedentRanges = precedents
End Function

Private Function FAS_FormulaHasUnsupportedPrecedents(ByVal formulaText As String) As Boolean
    FAS_FormulaHasUnsupportedPrecedents = (InStr(1, formulaText, "[", vbBinaryCompare) > 0 Or _
        InStr(1, formulaText, "INDIRECT(", vbTextCompare) > 0 Or _
        InStr(1, formulaText, "OFFSET(", vbTextCompare) > 0 Or _
        InStr(1, formulaText, "#", vbBinaryCompare) > 0)
End Function

Private Sub FAS_SplitSheetReference(ByVal referenceText As String, ByRef sheetName As String, ByRef addressText As String)
    Dim bangPosition As Long

    bangPosition = InStrRev(referenceText, "!")
    sheetName = Left$(referenceText, bangPosition - 1)
    addressText = Mid$(referenceText, bangPosition + 1)

    If Left$(sheetName, 1) = "'" And Right$(sheetName, 1) = "'" Then
        sheetName = Mid$(sheetName, 2, Len(sheetName) - 2)
        sheetName = Replace(sheetName, "''", "'")
    End If
End Sub

Private Sub FAS_AddPrecedentIfValid(ByVal precedents As Collection, ByVal targetWorkbook As Workbook, ByVal sheetName As String, ByVal addressText As String)
    Dim targetSheet As Worksheet
    Dim targetRange As Range
    Dim uniqueKey As String
    Dim item As Range

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

Private Function FAS_IsStandaloneA1Match(ByVal formulaText As String, ByVal startPosition As Long, ByVal matchLength As Long) As Boolean
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

    FAS_IsStandaloneA1Match = True
End Function

Private Function FAS_BlankQuotedStrings(ByVal formulaText As String) As String
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

    FAS_BlankQuotedStrings = result
End Function

Private Function FAS_BuildPrecedentPrompt(ByVal sourceCell As Range, ByVal precedents As Collection) As String
    Dim prompt As String
    Dim index As Long
    Dim precedentRange As Range

    prompt = "Precedents for " & sourceCell.Worksheet.Name & "!" & sourceCell.Address(False, False) & ":" & vbCrLf & vbCrLf
    For index = 1 To precedents.Count
        Set precedentRange = precedents(index)
        prompt = prompt & index & ". " & precedentRange.Worksheet.Name & "!" & precedentRange.Address(False, False) & vbCrLf
    Next index

    FAS_BuildPrecedentPrompt = prompt & vbCrLf & "Enter a number to jump. Use command 'back' to return."
End Function

Private Sub FAS_SaveTraceOrigin(ByVal sourceCell As Range)
    mTraceOriginWorkbookFullName = sourceCell.Worksheet.Parent.FullName
    mTraceOriginSheetName = sourceCell.Worksheet.Name
    mTraceOriginAddress = sourceCell.Address(False, False)
End Sub

Private Sub FAS_StartPrecedentCycle(ByVal sourceCell As Range)
    FAS_ResetPrecedentCycle
    Set mCyclePrecedents = FAS_ParsePrecedentRanges(sourceCell)

    If mCyclePrecedents.Count = 0 Then
        FAS_SetStatus "No supported direct precedents for " & sourceCell.Address(False, False) & "."
        Set mCyclePrecedents = Nothing
        Exit Sub
    End If

    FAS_SaveTraceOrigin sourceCell
    mCycleSourceWorkbookFullName = sourceCell.Worksheet.Parent.FullName
    mCycleSourceSheetName = sourceCell.Worksheet.Name
    mCycleSourceAddress = sourceCell.Address(False, False)
    mCycleSourceFormula = CStr(sourceCell.Formula)
    mCycleIndex = 0
End Sub

Private Function FAS_CanContinuePrecedentCycle() As Boolean
    Dim activeRange As Range

    If mCyclePrecedents Is Nothing Then Exit Function
    If mCyclePrecedents.Count = 0 Then Exit Function
    If TypeName(Selection) <> "Range" Then Exit Function

    Set activeRange = ActiveCell
    If activeRange Is Nothing Then Exit Function

    FAS_CanContinuePrecedentCycle = _
        StrComp(activeRange.Worksheet.Parent.FullName, mCycleLastTargetWorkbookFullName, vbTextCompare) = 0 And _
        StrComp(activeRange.Worksheet.Name, mCycleLastTargetSheetName, vbTextCompare) = 0 And _
        StrComp(activeRange.Address(False, False), mCycleLastTargetAddress, vbTextCompare) = 0
End Function

Private Function FAS_IsCycleSourceCurrent() As Boolean
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceCell As Range

    Set sourceWorkbook = FAS_GetOpenWorkbookByFullName(mCycleSourceWorkbookFullName)
    If sourceWorkbook Is Nothing Then Exit Function

    On Error Resume Next
    Set sourceSheet = sourceWorkbook.Worksheets(mCycleSourceSheetName)
    Set sourceCell = sourceSheet.Range(mCycleSourceAddress)
    On Error GoTo 0
    If sourceCell Is Nothing Then Exit Function

    FAS_IsCycleSourceCurrent = sourceCell.HasFormula And CStr(sourceCell.Formula) = mCycleSourceFormula
End Function

Private Sub FAS_SaveCycleLastTarget(ByVal targetRange As Range)
    mCycleLastTargetWorkbookFullName = targetRange.Worksheet.Parent.FullName
    mCycleLastTargetSheetName = targetRange.Worksheet.Name
    mCycleLastTargetAddress = targetRange.Address(False, False)
End Sub

Private Sub FAS_ResetPrecedentCycle()
    mCycleSourceWorkbookFullName = vbNullString
    mCycleSourceSheetName = vbNullString
    mCycleSourceAddress = vbNullString
    mCycleSourceFormula = vbNullString
    mCycleLastTargetWorkbookFullName = vbNullString
    mCycleLastTargetSheetName = vbNullString
    mCycleLastTargetAddress = vbNullString
    Set mCyclePrecedents = Nothing
    mCycleIndex = 0
End Sub

Private Sub FAS_ActivatePrecedentRange(ByVal targetRange As Range)
    targetRange.Worksheet.Parent.Activate
    targetRange.Worksheet.Activate
    targetRange.Select
End Sub

Private Function FAS_GetOpenWorkbookByFullName(ByVal workbookFullName As String) As Workbook
    Dim currentWorkbook As Workbook

    For Each currentWorkbook In Application.Workbooks
        If StrComp(currentWorkbook.FullName, workbookFullName, vbTextCompare) = 0 Then
            Set FAS_GetOpenWorkbookByFullName = currentWorkbook
            Exit Function
        End If
    Next currentWorkbook
End Function
