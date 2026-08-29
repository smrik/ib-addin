Attribute VB_Name = "modSheetAssignments"
Option Explicit

Private Const FAS_DEFAULT_VISIBLE_SLOT_COUNT As Long = 2
Private Const FAS_MAPPINGS_FOLDER_NAME As String = "FAS_Toolkit\Mappings"
Private Const FAS_TEXT_FOR_READING As Long = 1
Private Const FAS_TEXT_FOR_WRITING As Long = 2
Private Const FAS_TEXT_FORMAT_UNICODE As Long = -1
Public Const FAS_MAX_SLOT_INDEX As Long = 9

Public Function FAS_SaveSheetAssignment(ByVal targetWorkbook As Workbook, ByVal groupKey As String, ByVal slotIndex As Long, ByVal sheetName As String) As Boolean
    Dim lines As Collection
    Dim lineIndex As Long
    Dim fields As Variant
    Dim assignmentKey As String
    Dim normalizedGroupKey As String
    Dim updated As Boolean
    Dim targetSheet As Worksheet

    normalizedGroupKey = FAS_NormalizeGroupKey(groupKey)
    If Not FAS_CanPersistAssignments(targetWorkbook) Then
        Exit Function
    End If

    If slotIndex < 1 Or slotIndex > FAS_MAX_SLOT_INDEX Then
        Exit Function
    End If

    Set targetSheet = FAS_GetWorksheetByName(targetWorkbook, sheetName)
    If targetSheet Is Nothing Then
        Exit Function
    End If

    Set lines = FAS_ReadMappingLines(targetWorkbook)
    assignmentKey = normalizedGroupKey & "|" & CStr(slotIndex)

    For lineIndex = 1 To lines.Count
        fields = Split(CStr(lines(lineIndex)), "|", 3)
        If UBound(fields) >= 2 Then
            If StrComp(fields(0) & "|" & fields(1), assignmentKey, vbTextCompare) = 0 Then
                lines.Remove lineIndex
                lines.Add FAS_BuildMappingLine(normalizedGroupKey, slotIndex, sheetName), Before:=lineIndex
                updated = True
                Exit For
            End If
        End If
    Next lineIndex

    If Not updated Then
        lines.Add FAS_BuildMappingLine(normalizedGroupKey, slotIndex, sheetName)
    End If

    FAS_WriteMappingLines targetWorkbook, lines
    FAS_SaveSheetAssignment = True
End Function

Public Function FAS_ClearSheetAssignment(ByVal targetWorkbook As Workbook, ByVal groupKey As String, ByVal slotIndex As Long) As Boolean
    Dim lines As Collection
    Dim lineIndex As Long
    Dim fields As Variant
    Dim assignmentKey As String

    If Not FAS_CanPersistAssignments(targetWorkbook) Then
        Exit Function
    End If

    If slotIndex < 1 Or slotIndex > FAS_MAX_SLOT_INDEX Then
        Exit Function
    End If

    Set lines = FAS_ReadMappingLines(targetWorkbook)
    assignmentKey = FAS_NormalizeGroupKey(groupKey) & "|" & CStr(slotIndex)

    For lineIndex = lines.Count To 1 Step -1
        fields = Split(CStr(lines(lineIndex)), "|", 3)
        If UBound(fields) >= 2 Then
            If StrComp(fields(0) & "|" & fields(1), assignmentKey, vbTextCompare) = 0 Then
                lines.Remove lineIndex
                FAS_ClearSheetAssignment = True
            End If
        End If
    Next lineIndex

    FAS_WriteMappingLines targetWorkbook, lines
End Function

Public Function FAS_GetAssignedSheet(ByVal targetWorkbook As Workbook, ByVal groupKey As String, Optional ByVal slotIndex As Long = 1) As Worksheet
    Dim assignedName As String

    assignedName = FAS_GetAssignedSheetName(targetWorkbook, groupKey, slotIndex)
    If Len(assignedName) = 0 Then
        Exit Function
    End If

    Set FAS_GetAssignedSheet = FAS_GetWorksheetByName(targetWorkbook, assignedName)
End Function

Public Function FAS_GetAssignedSheetName(ByVal targetWorkbook As Workbook, ByVal groupKey As String, Optional ByVal slotIndex As Long = 1) As String
    Dim lines As Collection

    If Not FAS_CanPersistAssignments(targetWorkbook) Then
        Exit Function
    End If

    If slotIndex < 1 Or slotIndex > FAS_MAX_SLOT_INDEX Then
        Exit Function
    End If

    Set lines = FAS_ReadMappingLines(targetWorkbook)
    FAS_GetAssignedSheetName = FAS_GetAssignedSheetNameFromLines(lines, groupKey, slotIndex)
End Function

Public Function FAS_GetResolvedGroupSheet(ByVal targetWorkbook As Workbook, ByVal groupKey As String, Optional ByVal slotIndex As Long = 1) As Worksheet
    Dim lines As Collection
    Dim resolvedSheets As Collection
    Dim assignedName As String
    Dim assignedSheet As Worksheet
    Dim fuzzySheets As Collection
    Dim currentSheet As Worksheet
    Dim currentSlot As Long

    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    If slotIndex < 1 Or slotIndex > FAS_MAX_SLOT_INDEX Then
        Exit Function
    End If

    Set lines = FAS_ReadMappingLines(targetWorkbook)

    assignedName = FAS_GetAssignedSheetNameFromLines(lines, groupKey, slotIndex)
    If Len(assignedName) > 0 Then
        Set assignedSheet = FAS_GetWorksheetByName(targetWorkbook, assignedName)
        If Not assignedSheet Is Nothing Then
            Set FAS_GetResolvedGroupSheet = assignedSheet
            Exit Function
        End If
    End If

    Set resolvedSheets = New Collection
    For currentSlot = 1 To FAS_MAX_SLOT_INDEX
        assignedName = FAS_GetAssignedSheetNameFromLines(lines, groupKey, currentSlot)
        If Len(assignedName) > 0 Then
            Set assignedSheet = FAS_GetWorksheetByName(targetWorkbook, assignedName)
            If Not assignedSheet Is Nothing Then
                FAS_AddUniqueSheet resolvedSheets, assignedSheet
            End If
        End If
    Next currentSlot

    Set fuzzySheets = FAS_GetRankedSheetsForCommonGroup(targetWorkbook, groupKey)
    For Each currentSheet In fuzzySheets
        FAS_AddUniqueSheet resolvedSheets, currentSheet
    Next currentSheet

    If slotIndex <= resolvedSheets.Count Then
        Set FAS_GetResolvedGroupSheet = resolvedSheets(slotIndex)
    End If
End Function

Public Function FAS_GetResolvedGroupSheets(ByVal targetWorkbook As Workbook, ByVal groupKey As String) As Collection
    Dim resolvedSheets As Collection
    Dim fuzzySheets As Collection
    Dim currentSheet As Worksheet
    Dim slotIndex As Long
    Dim lines As Collection
    Dim assignedName As String

    Set resolvedSheets = New Collection
    If targetWorkbook Is Nothing Then
        Set FAS_GetResolvedGroupSheets = resolvedSheets
        Exit Function
    End If

    Set lines = FAS_ReadMappingLines(targetWorkbook)
    For slotIndex = 1 To FAS_MAX_SLOT_INDEX
        assignedName = FAS_GetAssignedSheetNameFromLines(lines, groupKey, slotIndex)
        If Len(assignedName) > 0 Then
            Set currentSheet = FAS_GetWorksheetByName(targetWorkbook, assignedName)
            If Not currentSheet Is Nothing Then
                FAS_AddUniqueSheet resolvedSheets, currentSheet
            End If
        End If
    Next slotIndex

    Set fuzzySheets = FAS_GetRankedSheetsForCommonGroup(targetWorkbook, groupKey)
    For Each currentSheet In fuzzySheets
        FAS_AddUniqueSheet resolvedSheets, currentSheet
    Next currentSheet

    Set FAS_GetResolvedGroupSheets = resolvedSheets
End Function

Public Function FAS_GetDisplaySlotCount(ByVal targetWorkbook As Workbook, ByVal groupKey As String) As Long
    Dim resolvedSheets As Collection
    Dim highestAssignedSlot As Long

    highestAssignedSlot = FAS_GetHighestAssignedSlot(targetWorkbook, groupKey)
    Set resolvedSheets = FAS_GetResolvedGroupSheets(targetWorkbook, groupKey)

    If resolvedSheets Is Nothing Then
        Exit Function
    End If

    FAS_GetDisplaySlotCount = highestAssignedSlot
    If FAS_GetDisplaySlotCount < FAS_DEFAULT_VISIBLE_SLOT_COUNT Then
        FAS_GetDisplaySlotCount = FAS_DEFAULT_VISIBLE_SLOT_COUNT
    End If

    If FAS_GetDisplaySlotCount > resolvedSheets.Count Then
        FAS_GetDisplaySlotCount = resolvedSheets.Count
    End If
End Function

Public Function FAS_GetHighestAssignedSlot(ByVal targetWorkbook As Workbook, ByVal groupKey As String) As Long
    Dim slotIndex As Long
    Dim lines As Collection

    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    Set lines = FAS_ReadMappingLines(targetWorkbook)
    For slotIndex = FAS_MAX_SLOT_INDEX To 1 Step -1
        If Len(FAS_GetAssignedSheetNameFromLines(lines, groupKey, slotIndex)) > 0 Then
            FAS_GetHighestAssignedSlot = slotIndex
            Exit Function
        End If
    Next slotIndex
End Function

Public Function FAS_ParseJumpCommand(ByVal commandText As String, ByRef groupKey As String, ByRef slotIndex As Long) As Boolean
    Dim normalizedCommand As String
    Dim aliasText As String
    Dim digitText As String
    Dim charIndex As Long
    Dim currentChar As String

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    If Len(normalizedCommand) = 0 Then
        Exit Function
    End If

    For charIndex = 1 To Len(normalizedCommand)
        currentChar = Mid$(normalizedCommand, charIndex, 1)
        If currentChar Like "#" Then
            digitText = Mid$(normalizedCommand, charIndex)
            Exit For
        End If
        aliasText = aliasText & currentChar
    Next charIndex

    If Len(aliasText) = 0 Then
        Exit Function
    End If

    groupKey = FAS_GetGroupKeyFromAlias(aliasText)
    If Len(groupKey) = 0 Then
        Exit Function
    End If

    If Len(digitText) = 0 Then
        slotIndex = 1
    ElseIf IsNumeric(digitText) Then
        slotIndex = CLng(digitText)
    Else
        Exit Function
    End If

    If slotIndex < 1 Or slotIndex > FAS_MAX_SLOT_INDEX Then
        Exit Function
    End If

    FAS_ParseJumpCommand = True
End Function

Public Function FAS_NormalizeCommandText(ByVal commandText As String) As String
    FAS_NormalizeCommandText = Replace(FAS_NormalizeSheetText(commandText), " ", vbNullString)
End Function

Public Function FAS_GetGroupKeyFromAlias(ByVal aliasText As String) As String
    Select Case FAS_NormalizeCommandText(aliasText)
        Case "p", "pl"
            FAS_GetGroupKeyFromAlias = "PL"
        Case "b", "bs"
            FAS_GetGroupKeyFromAlias = "BS"
        Case "t", "tb", "bb"
            FAS_GetGroupKeyFromAlias = "TB"
        Case "g", "gl"
            FAS_GetGroupKeyFromAlias = "GL"
        Case "m", "map", "mapping"
            FAS_GetGroupKeyFromAlias = "MAPPING"
    End Select
End Function

Public Function FAS_GetPrimaryJumpAlias(ByVal groupKey As String) As String
    Select Case FAS_NormalizeGroupKey(groupKey)
        Case "PL"
            FAS_GetPrimaryJumpAlias = "p"
        Case "BS"
            FAS_GetPrimaryJumpAlias = "bs"
        Case "TB"
            FAS_GetPrimaryJumpAlias = "t"
        Case "GL"
            FAS_GetPrimaryJumpAlias = "g"
        Case "MAPPING"
            FAS_GetPrimaryJumpAlias = "m"
    End Select
End Function

Public Function FAS_CanPersistAssignments(ByVal targetWorkbook As Workbook) As Boolean
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    FAS_CanPersistAssignments = (Len(Trim$(targetWorkbook.FullName)) > 0)
End Function

Public Function FAS_SaveActiveWorkbookAssignment(ByVal groupKey As String, ByVal slotIndex As Long, ByVal sheetName As String) As Boolean
    FAS_SaveActiveWorkbookAssignment = FAS_SaveSheetAssignment(FAS_GetTargetWorkbook(), groupKey, slotIndex, sheetName)
End Function

Public Sub FAS_ShowActiveWorkbookMappings()
    Dim mappingText As String

    mappingText = FAS_GetActiveWorkbookMappingsText()
    If Len(mappingText) = 0 Then
        FAS_SetStatus "No local sheet mappings for this workbook."
        Exit Sub
    End If

    FAS_SetStatus Left$(Replace(mappingText, vbCrLf, " | "), 240)
End Sub

Public Function FAS_GetActiveWorkbookMappingsText() As String
    Dim targetWorkbook As Workbook
    Dim lines As Collection
    Dim lineItem As Variant
    Dim fields As Variant
    Dim result As String

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    Set lines = FAS_ReadMappingLines(targetWorkbook)
    If lines.Count = 0 Then
        Exit Function
    End If

    result = "Local mappings for:" & vbCrLf & targetWorkbook.FullName & vbCrLf & vbCrLf
    For Each lineItem In lines
        fields = Split(CStr(lineItem), "|", 3)
        If UBound(fields) >= 2 Then
            result = result & fields(0) & fields(1) & " -> " & fields(2) & vbCrLf
        End If
    Next lineItem

    result = result & vbCrLf & "Stored at:" & vbCrLf & FAS_GetMappingFilePath(targetWorkbook)
    FAS_GetActiveWorkbookMappingsText = result
End Function

Public Function FAS_GetActiveWorkbookMappingFilePath() As String
    FAS_GetActiveWorkbookMappingFilePath = FAS_GetMappingFilePath(FAS_GetTargetWorkbook())
End Function

Public Function FAS_ResetActiveWorkbookMappings() As Boolean
    Dim targetWorkbook As Workbook
    Dim fileSystem As Object
    Dim mappingPath As String
    Dim response As VbMsgBoxResult

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    mappingPath = FAS_GetMappingFilePath(targetWorkbook)
    If Len(mappingPath) = 0 Then
        Exit Function
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(mappingPath) Then
        If Application.Visible Then
            FAS_SetStatus "No local sheet mappings for this workbook."
        End If
        Exit Function
    End If

    If Application.Visible Then
        response = MsgBox( _
            "Clear local FAS sheet mappings for this workbook?" & vbCrLf & vbCrLf & targetWorkbook.FullName, _
            vbQuestion + vbOKCancel, _
            "FAS Sheet Mappings")
        If response <> vbOK Then
            Exit Function
        End If
    End If

    fileSystem.DeleteFile mappingPath, True
    FAS_ResetActiveWorkbookMappings = True

    If Application.Visible Then
        FAS_SetStatus "Local sheet mappings cleared."
    End If
End Function

Private Function FAS_ReadMappingLines(ByVal targetWorkbook As Workbook) As Collection
    Dim lines As Collection
    Dim fileSystem As Object
    Dim textStream As Object
    Dim mappingPath As String
    Dim rawLine As String

    Set lines = New Collection
    If targetWorkbook Is Nothing Then
        Set FAS_ReadMappingLines = lines
        Exit Function
    End If

    mappingPath = FAS_GetMappingFilePath(targetWorkbook)
    If Len(mappingPath) = 0 Then
        Set FAS_ReadMappingLines = lines
        Exit Function
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(mappingPath) Then
        Set FAS_ReadMappingLines = lines
        Exit Function
    End If

    Set textStream = fileSystem.OpenTextFile(mappingPath, FAS_TEXT_FOR_READING, False, FAS_TEXT_FORMAT_UNICODE)
    Do While Not textStream.AtEndOfStream
        rawLine = textStream.ReadLine
        If Len(Trim$(rawLine)) > 0 Then
            lines.Add rawLine
        End If
    Loop
    textStream.Close

    Set FAS_ReadMappingLines = lines
End Function

Private Sub FAS_WriteMappingLines(ByVal targetWorkbook As Workbook, ByVal lines As Collection)
    Dim fileSystem As Object
    Dim textStream As Object
    Dim mappingFolder As String
    Dim mappingPath As String
    Dim lineItem As Variant

    If targetWorkbook Is Nothing Then
        Exit Sub
    End If

    mappingFolder = FAS_GetMappingsFolderPath()
    If Len(mappingFolder) = 0 Then
        Exit Sub
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    FAS_EnsureFolderExists fileSystem, mappingFolder

    mappingPath = FAS_GetMappingFilePath(targetWorkbook)
    If lines.Count = 0 Then
        If fileSystem.FileExists(mappingPath) Then
            fileSystem.DeleteFile mappingPath, True
        End If
        Exit Sub
    End If

    Set textStream = fileSystem.CreateTextFile(mappingPath, True, True)
    For Each lineItem In lines
        textStream.WriteLine CStr(lineItem)
    Next lineItem
    textStream.Close
End Sub

Private Function FAS_GetMappingFilePath(ByVal targetWorkbook As Workbook) As String
    Dim workbookKey As String

    If Not FAS_CanPersistAssignments(targetWorkbook) Then
        Exit Function
    End If

    workbookKey = LCase$(Trim$(targetWorkbook.FullName))
    FAS_GetMappingFilePath = FAS_GetMappingsFolderPath() & "\" & FAS_SanitizeFileName(targetWorkbook.Name) & "_" & FAS_GetStableHash(workbookKey) & ".txt"
End Function

Private Function FAS_GetMappingsFolderPath() As String
    Dim appDataPath As String

    appDataPath = Environ$("APPDATA")
    If Len(appDataPath) = 0 Then
        Exit Function
    End If

    FAS_GetMappingsFolderPath = appDataPath & "\" & FAS_MAPPINGS_FOLDER_NAME
End Function

Private Function FAS_BuildMappingLine(ByVal groupKey As String, ByVal slotIndex As Long, ByVal sheetName As String) As String
    FAS_BuildMappingLine = FAS_NormalizeGroupKey(groupKey) & "|" & CStr(slotIndex) & "|" & sheetName
End Function

Private Function FAS_GetAssignedSheetNameFromLines(ByVal lines As Collection, ByVal groupKey As String, ByVal slotIndex As Long) As String
    Dim lineIndex As Long
    Dim fields As Variant
    Dim assignmentKey As String

    If lines Is Nothing Then
        Exit Function
    End If

    assignmentKey = FAS_NormalizeGroupKey(groupKey) & "|" & CStr(slotIndex)

    For lineIndex = 1 To lines.Count
        fields = Split(CStr(lines(lineIndex)), "|", 3)
        If UBound(fields) >= 2 Then
            If StrComp(fields(0) & "|" & fields(1), assignmentKey, vbTextCompare) = 0 Then
                FAS_GetAssignedSheetNameFromLines = fields(2)
                Exit Function
            End If
        End If
    Next lineIndex
End Function

Private Function FAS_NormalizeGroupKey(ByVal groupKey As String) As String
    FAS_NormalizeGroupKey = UCase$(Trim$(groupKey))
End Function

Private Function FAS_GetWorksheetByName(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim currentSheet As Worksheet

    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    For Each currentSheet In targetWorkbook.Worksheets
        If StrComp(currentSheet.Name, sheetName, vbTextCompare) = 0 Then
            Set FAS_GetWorksheetByName = currentSheet
            Exit Function
        End If
    Next currentSheet
End Function

Private Sub FAS_AddUniqueSheet(ByVal sheets As Collection, ByVal targetSheet As Worksheet)
    Dim currentSheet As Worksheet

    If targetSheet Is Nothing Then
        Exit Sub
    End If

    For Each currentSheet In sheets
        If StrComp(currentSheet.Name, targetSheet.Name, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next currentSheet

    sheets.Add targetSheet
End Sub

Private Function FAS_GetStableHash(ByVal value As String) As String
    Dim index As Long
    Dim hashValue As Double
    Dim charValue As Long

    hashValue = 0
    For index = 1 To Len(value)
        charValue = AscW(Mid$(value, index, 1))
        If charValue < 0 Then
            charValue = charValue + 65536
        End If

        hashValue = (hashValue * 131#) + charValue
        hashValue = hashValue - (Fix(hashValue / 2147483647#) * 2147483647#)
    Next index

    FAS_GetStableHash = Right$("00000000" & Hex$(CLng(hashValue)), 8)
End Function

Private Function FAS_SanitizeFileName(ByVal value As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In invalidChars
        value = Replace(value, CStr(item), "_")
    Next item

    If Len(value) = 0 Then
        value = "Workbook"
    End If

    FAS_SanitizeFileName = value
End Function

Private Sub FAS_EnsureFolderExists(ByVal fileSystem As Object, ByVal folderPath As String)
    Dim parentPath As String

    If fileSystem.FolderExists(folderPath) Then
        Exit Sub
    End If

    parentPath = fileSystem.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then
        If Not fileSystem.FolderExists(parentPath) Then
            FAS_EnsureFolderExists fileSystem, parentPath
        End If
    End If

    fileSystem.CreateFolder folderPath
End Sub
