Attribute VB_Name = "modNavigatorData"
Option Explicit

Private Const FAS_SCORE_NONE As Long = 0
Private Const FAS_SCORE_CONTAINS As Long = 1
Private Const FAS_SCORE_STARTS_WITH As Long = 2
Private Const FAS_SCORE_EXACT As Long = 3
Private Const FAS_SEPARATOR_TEXT As String = "------------------------------"

Public Function FAS_NormalizeSheetText(ByVal value As String) As String
    Dim result As String
    Dim replacements As Variant
    Dim item As Variant

    result = LCase$(Trim$(value))
    If Len(result) = 0 Then
        Exit Function
    End If

    result = Replace(result, "&", " and ")
    result = FAS_RemoveDiacritics(result)

    replacements = Array(".", ",", ";", ":", "-", "_", "/", "\", "(", ")", "[", "]", "{", "}", "'", Chr$(34), "+", "!", "?", vbTab, vbCr, vbLf)
    For Each item In replacements
        result = Replace(result, CStr(item), " ")
    Next item

    result = Application.WorksheetFunction.Trim(result)
    FAS_NormalizeSheetText = result
End Function

Public Function FAS_GetBestMatchScoreForText(ByVal candidateText As String, ByVal keyword As String) As Long
    Dim normalizedCandidate As String
    Dim normalizedKeyword As String

    normalizedCandidate = FAS_NormalizeSheetText(candidateText)
    normalizedKeyword = FAS_NormalizeSheetText(keyword)

    If Len(normalizedCandidate) = 0 Or Len(normalizedKeyword) = 0 Then
        Exit Function
    End If

    If StrComp(normalizedCandidate, normalizedKeyword, vbTextCompare) = 0 Then
        FAS_GetBestMatchScoreForText = FAS_SCORE_EXACT
    ElseIf Left$(normalizedCandidate, Len(normalizedKeyword)) = normalizedKeyword Then
        FAS_GetBestMatchScoreForText = FAS_SCORE_STARTS_WITH
    ElseIf InStr(1, normalizedCandidate, normalizedKeyword, vbTextCompare) > 0 Then
        FAS_GetBestMatchScoreForText = FAS_SCORE_CONTAINS
    Else
        FAS_GetBestMatchScoreForText = FAS_SCORE_NONE
    End If
End Function

Public Function FAS_FindBestSheetByKeywords(ByVal targetWorkbook As Workbook, ByVal keywords As Variant) As Worksheet
    Dim currentSheet As Worksheet
    Dim keyword As Variant
    Dim currentScore As Long
    Dim bestScore As Long

    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    For Each currentSheet In targetWorkbook.Worksheets
        currentScore = FAS_GetHighestScoreForName(currentSheet.Name, keywords)
        If currentScore > bestScore Then
            bestScore = currentScore
            Set FAS_FindBestSheetByKeywords = currentSheet
        End If
    Next currentSheet
End Function

Public Function FAS_FindBestSheetByQuery(ByVal targetWorkbook As Workbook, ByVal query As String) As Worksheet
    Dim normalizedQuery As String
    Dim currentSheet As Worksheet
    Dim currentScore As Long
    Dim bestScore As Long
    Dim groupKey As Variant

    normalizedQuery = FAS_NormalizeSheetText(query)
    If targetWorkbook Is Nothing Or Len(normalizedQuery) = 0 Then
        Exit Function
    End If

    For Each currentSheet In targetWorkbook.Worksheets
        currentScore = FAS_GetBestMatchScoreForText(currentSheet.Name, normalizedQuery)
        If currentScore > bestScore Then
            bestScore = currentScore
            Set FAS_FindBestSheetByQuery = currentSheet
        End If
    Next currentSheet

    If Not FAS_FindBestSheetByQuery Is Nothing Then
        Exit Function
    End If

    For Each groupKey In FAS_GetPinnedGroupOrder()
        If FAS_QueryMatchesGroup(normalizedQuery, CStr(groupKey)) Then
            Set FAS_FindBestSheetByQuery = FAS_FindBestSheetForCommonGroup(targetWorkbook, CStr(groupKey))
            Exit Function
        End If
    Next groupKey
End Function

Public Function FAS_FindBestSheetForCommonGroup(ByVal targetWorkbook As Workbook, ByVal groupKey As String) As Worksheet
    Set FAS_FindBestSheetForCommonGroup = FAS_FindBestSheetByKeywords(targetWorkbook, FAS_GetKeywordsForGroup(groupKey))
End Function

Public Function FAS_GetRankedSheetsForCommonGroup(ByVal targetWorkbook As Workbook, ByVal groupKey As String) As Collection
    Dim rankedSheets As Collection
    Dim currentSheet As Worksheet
    Dim currentScore As Long

    Set rankedSheets = New Collection
    If targetWorkbook Is Nothing Then
        Set FAS_GetRankedSheetsForCommonGroup = rankedSheets
        Exit Function
    End If

    For Each currentSheet In targetWorkbook.Worksheets
        currentScore = FAS_GetSheetMatchScoreForGroup(currentSheet.Name, groupKey)
        If currentScore > FAS_SCORE_NONE Then
            FAS_InsertSheetByScore rankedSheets, currentSheet, currentScore, groupKey
        End If
    Next currentSheet

    Set FAS_GetRankedSheetsForCommonGroup = rankedSheets
End Function

Public Function FAS_GetSheetMatchScoreForGroup(ByVal sheetName As String, ByVal groupKey As String) As Long
    FAS_GetSheetMatchScoreForGroup = FAS_GetHighestScoreForName(sheetName, FAS_GetKeywordsForGroup(groupKey))
End Function

Public Function FAS_IsPinnedSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Boolean
    Dim groupKey As Variant
    Dim bestSheet As Worksheet

    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    For Each groupKey In FAS_GetPinnedGroupOrder()
        Set bestSheet = FAS_FindBestSheetForCommonGroup(targetWorkbook, CStr(groupKey))
        If Not bestSheet Is Nothing Then
            If StrComp(bestSheet.Name, sheetName, vbTextCompare) = 0 Then
                FAS_IsPinnedSheet = True
                Exit Function
            End If
        End If
    Next groupKey
End Function

Public Function FAS_GetNavigatorItems(ByVal targetWorkbook As Workbook, Optional ByVal searchText As String = vbNullString) As Collection
    Dim items As Collection
    Dim groupKey As Variant
    Dim currentSheet As Worksheet
    Dim resolvedSheet As Worksheet
    Dim normalizedSearch As String
    Dim pinnedCount As Long
    Dim allCount As Long
    Dim slotIndex As Long
    Dim displaySlotCount As Long

    Set items = New Collection
    If targetWorkbook Is Nothing Then
        Set FAS_GetNavigatorItems = items
        Exit Function
    End If

    normalizedSearch = FAS_NormalizeSheetText(searchText)

    For Each groupKey In FAS_GetPinnedGroupOrder()
        displaySlotCount = FAS_GetDisplaySlotCount(targetWorkbook, CStr(groupKey))
        For slotIndex = 1 To displaySlotCount
            Set resolvedSheet = FAS_GetResolvedGroupSheet(targetWorkbook, CStr(groupKey), slotIndex)
            If Not resolvedSheet Is Nothing Then
                If FAS_SheetMatchesSearch(resolvedSheet.Name, normalizedSearch, CStr(groupKey)) Or InStr(1, FAS_GetPrimaryJumpAlias(CStr(groupKey)) & CStr(slotIndex), Replace(normalizedSearch, " ", vbNullString), vbTextCompare) > 0 Then
                    items.Add Array("[" & UCase$(FAS_GetPrimaryJumpAlias(CStr(groupKey))) & CStr(slotIndex) & "] " & resolvedSheet.Name, resolvedSheet.Name, False)
                    pinnedCount = pinnedCount + 1
                End If
            End If
        Next slotIndex
    Next groupKey

    For Each currentSheet In targetWorkbook.Worksheets
        If FAS_SheetMatchesSearch(currentSheet.Name, normalizedSearch, vbNullString) Then
            allCount = allCount + 1
        End If
    Next currentSheet

    If pinnedCount > 0 And allCount > 0 Then
        items.Add Array(FAS_SEPARATOR_TEXT, vbNullString, True)
    End If

    For Each currentSheet In targetWorkbook.Worksheets
        If FAS_SheetMatchesSearch(currentSheet.Name, normalizedSearch, vbNullString) Then
            items.Add Array(currentSheet.Name, currentSheet.Name, False)
        End If
    Next currentSheet

    Set FAS_GetNavigatorItems = items
End Function

Public Function FAS_GetPinnedGroupOrder() As Variant
    FAS_GetPinnedGroupOrder = Array("PL", "BS", "TB", "GL", "MAPPING")
End Function

Public Function FAS_GetGroupLabel(ByVal groupKey As String) As String
    Select Case UCase$(Trim$(groupKey))
        Case "PL"
            FAS_GetGroupLabel = "P&L"
        Case "BS"
            FAS_GetGroupLabel = "BS"
        Case "TB"
            FAS_GetGroupLabel = "TB"
        Case "GL"
            FAS_GetGroupLabel = "GL"
        Case "MAPPING"
            FAS_GetGroupLabel = "Mapping"
        Case Else
            FAS_GetGroupLabel = groupKey
    End Select
End Function

Public Function FAS_IsNavigatorSeparator(ByVal displayText As String, ByVal targetSheetName As String) As Boolean
    FAS_IsNavigatorSeparator = (Len(targetSheetName) = 0 And displayText = FAS_SEPARATOR_TEXT)
End Function

Private Function FAS_GetKeywordsForGroup(ByVal groupKey As String) As Variant
    Select Case UCase$(Trim$(groupKey))
        Case "PL"
            FAS_GetKeywordsForGroup = Array("P&L", "P and L", "PL", "Profit", "Income", "Izkaz")
        Case "BS"
            FAS_GetKeywordsForGroup = Array("BS", "Balance", "Bilanca", "Statement of Financial Position")
        Case "TB"
            FAS_GetKeywordsForGroup = Array("TB", "Trial Balance", "Bruto bilanca", "BB")
        Case "GL"
            FAS_GetKeywordsForGroup = Array("GL", "General Ledger", "Dnevnik", "Dnevnik knjizenja", "Knjizenja")
        Case "MAPPING"
            FAS_GetKeywordsForGroup = Array("Mapping", "Klasifikacija", "Classification", "Map")
        Case Else
            FAS_GetKeywordsForGroup = Array()
    End Select
End Function

Private Function FAS_GetHighestScoreForName(ByVal sheetName As String, ByVal keywords As Variant) As Long
    Dim keyword As Variant
    Dim score As Long

    For Each keyword In keywords
        score = FAS_GetBestMatchScoreForText(sheetName, CStr(keyword))
        If score > FAS_GetHighestScoreForName Then
            FAS_GetHighestScoreForName = score
        End If
    Next keyword
End Function

Private Sub FAS_InsertSheetByScore(ByVal rankedSheets As Collection, ByVal targetSheet As Worksheet, ByVal targetScore As Long, ByVal groupKey As String)
    Dim insertIndex As Long
    Dim currentSheet As Worksheet
    Dim currentScore As Long

    If rankedSheets.Count = 0 Then
        rankedSheets.Add targetSheet
        Exit Sub
    End If

    insertIndex = rankedSheets.Count + 1
    For insertIndex = 1 To rankedSheets.Count
        Set currentSheet = rankedSheets(insertIndex)
        currentScore = FAS_GetSheetMatchScoreForGroup(currentSheet.Name, groupKey)
        If targetScore > currentScore Then
            rankedSheets.Add targetSheet, Before:=insertIndex
            Exit Sub
        End If
    Next insertIndex

    rankedSheets.Add targetSheet
End Sub

Private Function FAS_SheetMatchesSearch(ByVal sheetName As String, ByVal normalizedSearch As String, ByVal groupKey As String) As Boolean
    Dim normalizedName As String

    If Len(normalizedSearch) = 0 Then
        FAS_SheetMatchesSearch = True
        Exit Function
    End If

    normalizedName = FAS_NormalizeSheetText(sheetName)
    If InStr(1, normalizedName, normalizedSearch, vbTextCompare) > 0 Then
        FAS_SheetMatchesSearch = True
        Exit Function
    End If

    If Len(groupKey) > 0 Then
        If InStr(1, FAS_NormalizeSheetText(FAS_GetGroupLabel(groupKey)), normalizedSearch, vbTextCompare) > 0 Then
            FAS_SheetMatchesSearch = True
            Exit Function
        End If

        FAS_SheetMatchesSearch = (InStr(1, FAS_NormalizeCommandText(FAS_GetPrimaryJumpAlias(groupKey)), Replace(normalizedSearch, " ", vbNullString), vbTextCompare) > 0)
    End If
End Function

Private Function FAS_QueryMatchesGroup(ByVal normalizedQuery As String, ByVal groupKey As String) As Boolean
    Dim keyword As Variant

    If FAS_GetBestMatchScoreForText(FAS_GetGroupLabel(groupKey), normalizedQuery) > FAS_SCORE_NONE Then
        FAS_QueryMatchesGroup = True
        Exit Function
    End If

    For Each keyword In FAS_GetKeywordsForGroup(groupKey)
        If FAS_GetBestMatchScoreForText(CStr(keyword), normalizedQuery) > FAS_SCORE_NONE Then
            FAS_QueryMatchesGroup = True
            Exit Function
        End If
    Next keyword
End Function

Private Function FAS_RemoveDiacritics(ByVal value As String) As String
    Dim sourceChars As Variant
    Dim targetChars As Variant
    Dim index As Long

    sourceChars = Array(ChrW$(353), ChrW$(382), ChrW$(269), ChrW$(263), ChrW$(273), ChrW$(225), ChrW$(224), ChrW$(228), ChrW$(226), ChrW$(233), ChrW$(232), ChrW$(235), ChrW$(234), ChrW$(237), ChrW$(236), ChrW$(239), ChrW$(238), ChrW$(243), ChrW$(242), ChrW$(246), ChrW$(244), ChrW$(250), ChrW$(249), ChrW$(252), ChrW$(251), ChrW$(241))
    targetChars = Array("s", "z", "c", "c", "d", "a", "a", "a", "a", "e", "e", "e", "e", "i", "i", "i", "i", "o", "o", "o", "o", "u", "u", "u", "u", "n")

    For index = LBound(sourceChars) To UBound(sourceChars)
        value = Replace(value, sourceChars(index), targetChars(index))
    Next index

    FAS_RemoveDiacritics = value
End Function
