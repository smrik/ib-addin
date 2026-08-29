Attribute VB_Name = "modTBSearch"
Option Explicit

Private Const FAS_TB_SEARCH_MAX_MATCHES As Long = 30

Public Sub FAS_FindTBAccount()
    Dim searchText As String

    searchText = InputBox("Search the mapped TB/BB sheet for account text or account number:", "FAS TB Search")
    If Len(Trim$(searchText)) = 0 Then
        Exit Sub
    End If

    If Not FAS_GoToTBAccount(searchText) Then
        FAS_SetStatus "No TB match found for '" & searchText & "'."
    End If
End Sub

Public Function FAS_GoToTBAccount(ByVal searchText As String) As Boolean
    FAS_GoToTBAccount = FAS_GoToCommonSheetMatch("TB", searchText)
End Function

Public Function FAS_GoToCommonSheetMatch(ByVal groupKey As String, ByVal searchText As String) As Boolean
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim matches As Collection
    Dim selectedIndex As Long
    Dim userInput As String

    searchText = Trim$(searchText)
    If Len(searchText) = 0 Then
        Exit Function
    End If

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    Set targetSheet = FAS_GetResolvedGroupSheet(targetWorkbook, groupKey, 1)
    If targetSheet Is Nothing Then
        If Application.Visible Then
            FAS_SetStatus "No " & FAS_GetGroupLabel(groupKey) & " sheet mapped or detectable."
        End If
        Exit Function
    End If

    Set matches = FAS_FindTBMatches(targetSheet, searchText, FAS_TB_SEARCH_MAX_MATCHES)
    If matches.Count = 0 Then
        FAS_SetStatus "No " & FAS_GetGroupLabel(groupKey) & " match for '" & searchText & "'."
        Exit Function
    End If

    selectedIndex = 1
    FAS_ActivateTBMatch matches(selectedIndex)
    FAS_SetStatus FAS_GetGroupLabel(groupKey) & " match 1 of " & matches.Count & ": " & matches(selectedIndex).Address(False, False) & " - " & Left$(matches(selectedIndex).Text, 60)
    FAS_GoToCommonSheetMatch = True
End Function

Public Function FAS_CountTBAccountMatches(ByVal searchText As String) As Long
    FAS_CountTBAccountMatches = FAS_CountCommonSheetMatches("TB", searchText)
End Function

Public Function FAS_CountCommonSheetMatches(ByVal groupKey As String, ByVal searchText As String) As Long
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim matches As Collection

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    Set targetSheet = FAS_GetResolvedGroupSheet(targetWorkbook, groupKey, 1)
    If targetSheet Is Nothing Then
        Exit Function
    End If

    Set matches = FAS_FindTBMatches(targetSheet, searchText, FAS_TB_SEARCH_MAX_MATCHES)
    FAS_CountCommonSheetMatches = matches.Count
End Function

Public Function FAS_TryParseTBSearchCommand(ByVal commandText As String, ByRef searchText As String) As Boolean
    Dim groupKey As String

    FAS_TryParseTBSearchCommand = FAS_TryParseSheetLookupCommand(commandText, groupKey, searchText)
    FAS_TryParseTBSearchCommand = (FAS_TryParseTBSearchCommand And groupKey = "TB")
End Function

Public Function FAS_TryParseSheetLookupCommand(ByVal commandText As String, ByRef groupKey As String, ByRef searchText As String) As Boolean
    Dim rawText As String
    Dim normalizedText As String

    rawText = Trim$(commandText)
    normalizedText = FAS_NormalizeSheetText(rawText)
    If Len(normalizedText) = 0 Then
        Exit Function
    End If

    If Left$(normalizedText, 3) = "tb " Then
        groupKey = "TB"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "tb")
    ElseIf Left$(normalizedText, 3) = "bb " Then
        groupKey = "TB"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "bb")
    ElseIf Left$(normalizedText, 5) = "acct " Then
        groupKey = "TB"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "acct")
    ElseIf Left$(normalizedText, 8) = "account " Then
        groupKey = "TB"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "account")
    ElseIf Left$(normalizedText, 3) = "gl " Then
        groupKey = "GL"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "gl")
    ElseIf Left$(normalizedText, 4) = "map " Then
        groupKey = "MAPPING"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "map")
    ElseIf Left$(normalizedText, 8) = "mapping " Then
        groupKey = "MAPPING"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "mapping")
    ElseIf Left$(normalizedText, 6) = "class " Then
        groupKey = "MAPPING"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "class")
    ElseIf Left$(normalizedText, 5) = "klas " Then
        groupKey = "MAPPING"
        searchText = FAS_GetLookupTextAfterPrefix(rawText, "klas")
    End If

    FAS_TryParseSheetLookupCommand = (Len(groupKey) > 0 And Len(Trim$(searchText)) > 0)
End Function

Private Function FAS_GetLookupTextAfterPrefix(ByVal rawText As String, ByVal prefixText As String) As String
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

    FAS_GetLookupTextAfterPrefix = result
End Function

Private Function FAS_FindTBMatches(ByVal tbSheet As Worksheet, ByVal searchText As String, ByVal maxMatches As Long) As Collection
    Dim matches As Collection
    Dim searchRange As Range
    Dim foundCell As Range
    Dim firstAddress As String

    Set matches = New Collection
    searchText = Trim$(searchText)
    If tbSheet Is Nothing Or Len(searchText) = 0 Then
        Set FAS_FindTBMatches = matches
        Exit Function
    End If

    On Error Resume Next
    Set searchRange = tbSheet.UsedRange
    On Error GoTo 0

    If searchRange Is Nothing Then
        Set FAS_FindTBMatches = matches
        Exit Function
    End If

    Application.FindFormat.Clear
    Set foundCell = searchRange.Find( _
        What:=searchText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)

    If foundCell Is Nothing Then
        Set FAS_FindTBMatches = matches
        Exit Function
    End If

    firstAddress = foundCell.Address(False, False)
    Do
        matches.Add foundCell
        If matches.Count >= maxMatches Then
            Exit Do
        End If

        Set foundCell = searchRange.FindNext(foundCell)
        If foundCell Is Nothing Then
            Exit Do
        End If
    Loop While StrComp(foundCell.Address(False, False), firstAddress, vbTextCompare) <> 0

    Set FAS_FindTBMatches = matches
End Function

Private Function FAS_BuildTBMatchPrompt(ByVal tbSheet As Worksheet, ByVal searchText As String, ByVal matches As Collection) As String
    Dim prompt As String
    Dim index As Long
    Dim matchCell As Range

    prompt = "Matches for '" & searchText & "' on " & tbSheet.Name & ":" & vbCrLf & vbCrLf
    For index = 1 To matches.Count
        Set matchCell = matches(index)
        prompt = prompt & index & ". " & matchCell.Address(False, False) & " - " & Left$(matchCell.Text, 80) & vbCrLf
    Next index

    If matches.Count >= FAS_TB_SEARCH_MAX_MATCHES Then
        prompt = prompt & vbCrLf & "Showing first " & FAS_TB_SEARCH_MAX_MATCHES & " matches."
    End If

    FAS_BuildTBMatchPrompt = prompt & vbCrLf & "Enter a number to jump."
End Function

Private Sub FAS_ActivateTBMatch(ByVal matchCell As Range)
    matchCell.Worksheet.Parent.Activate
    matchCell.Worksheet.Activate
    matchCell.Select
End Sub
