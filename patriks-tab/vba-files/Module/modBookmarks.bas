Attribute VB_Name = "modBookmarks"
Option Explicit

Private Const FAS_BOOKMARKS_FOLDER_NAME As String = "FAS_Toolkit\Bookmarks"
Private Const FAS_TEXT_FOR_READING As Long = 1
Private Const FAS_TEXT_FORMAT_UNICODE As Long = -1

Public Function FAS_RunBookmarkCommand(ByVal commandText As String) As Boolean
    Dim normalizedText As String
    Dim bookmarkName As String

    normalizedText = FAS_NormalizeSheetText(commandText)

    Select Case normalizedText
        Case "bm", "bms", "bookmarks"
            FAS_ShowActiveWorkbookBookmarks
            FAS_RunBookmarkCommand = True
            Exit Function
        Case "bmreset", "resetbookmarks"
            FAS_RunBookmarkCommand = FAS_ResetActiveWorkbookBookmarks()
            Exit Function
    End Select

    If Left$(normalizedText, 7) = "bm set " Then
        bookmarkName = Trim$(Mid$(normalizedText, 8))
        FAS_RunBookmarkCommand = FAS_SaveActiveCellBookmark(bookmarkName)
    ElseIf Left$(normalizedText, 5) = "mark " Then
        bookmarkName = Trim$(Mid$(normalizedText, 6))
        FAS_RunBookmarkCommand = FAS_SaveActiveCellBookmark(bookmarkName)
    ElseIf Left$(normalizedText, 3) = "bm " Then
        bookmarkName = Trim$(Mid$(normalizedText, 4))
        FAS_RunBookmarkCommand = FAS_GoToBookmark(bookmarkName)
    End If
End Function

Public Function FAS_IsBookmarkCommand(ByVal commandText As String) As Boolean
    Dim normalizedText As String

    normalizedText = FAS_NormalizeSheetText(commandText)
    FAS_IsBookmarkCommand = (normalizedText = "bm" Or normalizedText = "bms" Or normalizedText = "bookmarks" Or _
        normalizedText = "bmreset" Or normalizedText = "resetbookmarks" Or _
        Left$(normalizedText, 7) = "bm set " Or Left$(normalizedText, 5) = "mark " Or Left$(normalizedText, 3) = "bm ")
End Function

Public Function FAS_SaveActiveCellBookmark(ByVal bookmarkName As String) As Boolean
    Dim targetWorkbook As Workbook
    Dim targetCell As Range
    Dim lines As Collection
    Dim newLines As Collection
    Dim lineItem As Variant
    Dim fields As Variant
    Dim normalizedName As String

    bookmarkName = Trim$(bookmarkName)
    normalizedName = FAS_NormalizeBookmarkName(bookmarkName)
    If Len(normalizedName) = 0 Then
        Exit Function
    End If

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Or TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Set targetCell = ActiveCell
    If targetCell Is Nothing Then
        Set targetCell = Selection.cells(1, 1)
    End If

    Set lines = FAS_ReadBookmarkLines(targetWorkbook)
    Set newLines = New Collection
    For Each lineItem In lines
        fields = Split(CStr(lineItem), "|", 4)
        If UBound(fields) >= 3 Then
            If StrComp(FAS_DecodeBookmarkField(CStr(fields(0))), normalizedName, vbTextCompare) <> 0 Then
                newLines.Add CStr(lineItem)
            End If
        End If
    Next lineItem

    newLines.Add FAS_EncodeBookmarkField(normalizedName) & "|" & FAS_EncodeBookmarkField(bookmarkName) & "|" & FAS_EncodeBookmarkField(targetCell.Worksheet.Name) & "|" & FAS_EncodeBookmarkField(targetCell.Address(False, False))
    FAS_WriteBookmarkLines targetWorkbook, newLines
    FAS_SaveActiveCellBookmark = True

    If Application.Visible Then
        FAS_SetStatus "Bookmark saved: " & bookmarkName & " -> " & targetCell.Worksheet.Name & "!" & targetCell.Address(False, False)
    End If
End Function

Public Function FAS_GoToBookmark(ByVal bookmarkName As String) As Boolean
    Dim targetWorkbook As Workbook
    Dim lines As Collection
    Dim lineItem As Variant
    Dim fields As Variant
    Dim normalizedName As String
    Dim targetSheet As Worksheet
    Dim targetRange As Range

    normalizedName = FAS_NormalizeBookmarkName(bookmarkName)
    If Len(normalizedName) = 0 Then
        Exit Function
    End If

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    Set lines = FAS_ReadBookmarkLines(targetWorkbook)
    For Each lineItem In lines
        fields = Split(CStr(lineItem), "|", 4)
        If UBound(fields) >= 3 Then
            If StrComp(FAS_DecodeBookmarkField(CStr(fields(0))), normalizedName, vbTextCompare) = 0 Then
                On Error Resume Next
                Set targetSheet = targetWorkbook.Worksheets(FAS_DecodeBookmarkField(CStr(fields(2))))
                Set targetRange = targetSheet.Range(FAS_DecodeBookmarkField(CStr(fields(3))))
                On Error GoTo 0

                If Not targetRange Is Nothing Then
                    targetWorkbook.Activate
                    targetSheet.Activate
                    targetRange.Select
                    FAS_GoToBookmark = True
                End If
                Exit Function
            End If
        End If
    Next lineItem

    If Application.Visible Then
        FAS_SetStatus "Bookmark not found: " & bookmarkName
    End If
End Function

Public Sub FAS_ShowActiveWorkbookBookmarks()
    Dim bookmarkText As String

    bookmarkText = FAS_GetActiveWorkbookBookmarksText()
    If Len(bookmarkText) = 0 Then
        FAS_SetStatus "No local bookmarks for this workbook."
        Exit Sub
    End If

    FAS_SetStatus Left$(Replace(bookmarkText, vbCrLf, " | "), 240)
End Sub

Public Function FAS_GetActiveWorkbookBookmarksText() As String
    Dim targetWorkbook As Workbook
    Dim lines As Collection
    Dim lineItem As Variant
    Dim fields As Variant
    Dim result As String

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    Set lines = FAS_ReadBookmarkLines(targetWorkbook)
    If lines.Count = 0 Then
        Exit Function
    End If

    result = "Local bookmarks for:" & vbCrLf & targetWorkbook.FullName & vbCrLf & vbCrLf
    For Each lineItem In lines
        fields = Split(CStr(lineItem), "|", 4)
        If UBound(fields) >= 3 Then
            result = result & FAS_DecodeBookmarkField(CStr(fields(1))) & " -> " & FAS_DecodeBookmarkField(CStr(fields(2))) & "!" & FAS_DecodeBookmarkField(CStr(fields(3))) & vbCrLf
        End If
    Next lineItem

    FAS_GetActiveWorkbookBookmarksText = result
End Function

Public Function FAS_CountActiveWorkbookBookmarks() As Long
    Dim targetWorkbook As Workbook
    Dim lines As Collection

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    Set lines = FAS_ReadBookmarkLines(targetWorkbook)
    FAS_CountActiveWorkbookBookmarks = lines.Count
End Function

Public Function FAS_ResetActiveWorkbookBookmarks() As Boolean
    Dim targetWorkbook As Workbook
    Dim fileSystem As Object
    Dim bookmarkPath As String
    Dim response As VbMsgBoxResult

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    bookmarkPath = FAS_GetBookmarksFilePath(targetWorkbook)
    If Len(bookmarkPath) = 0 Then
        Exit Function
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(bookmarkPath) Then
        Exit Function
    End If

    If Application.Visible Then
        response = MsgBox("Clear local FAS bookmarks for this workbook?", vbQuestion + vbOKCancel, "FAS Bookmarks")
        If response <> vbOK Then
            Exit Function
        End If
    End If

    fileSystem.DeleteFile bookmarkPath, True
    FAS_ResetActiveWorkbookBookmarks = True
End Function

Private Function FAS_ReadBookmarkLines(ByVal targetWorkbook As Workbook) As Collection
    Dim lines As Collection
    Dim fileSystem As Object
    Dim textStream As Object
    Dim bookmarkPath As String
    Dim rawLine As String

    Set lines = New Collection
    bookmarkPath = FAS_GetBookmarksFilePath(targetWorkbook)
    If Len(bookmarkPath) = 0 Then
        Set FAS_ReadBookmarkLines = lines
        Exit Function
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(bookmarkPath) Then
        Set FAS_ReadBookmarkLines = lines
        Exit Function
    End If

    Set textStream = fileSystem.OpenTextFile(bookmarkPath, FAS_TEXT_FOR_READING, False, FAS_TEXT_FORMAT_UNICODE)
    Do While Not textStream.AtEndOfStream
        rawLine = textStream.ReadLine
        If Len(Trim$(rawLine)) > 0 Then
            lines.Add rawLine
        End If
    Loop
    textStream.Close

    Set FAS_ReadBookmarkLines = lines
End Function

Private Sub FAS_WriteBookmarkLines(ByVal targetWorkbook As Workbook, ByVal lines As Collection)
    Dim fileSystem As Object
    Dim textStream As Object
    Dim bookmarkFolder As String
    Dim bookmarkPath As String
    Dim lineItem As Variant

    bookmarkFolder = FAS_GetBookmarksFolderPath()
    If Len(bookmarkFolder) = 0 Then
        Exit Sub
    End If

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    FAS_EnsureBookmarkFolderExists fileSystem, bookmarkFolder

    bookmarkPath = FAS_GetBookmarksFilePath(targetWorkbook)
    Set textStream = fileSystem.CreateTextFile(bookmarkPath, True, True)
    For Each lineItem In lines
        textStream.WriteLine CStr(lineItem)
    Next lineItem
    textStream.Close
End Sub

Private Function FAS_GetBookmarksFilePath(ByVal targetWorkbook As Workbook) As String
    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    If Len(Trim$(targetWorkbook.FullName)) = 0 Then
        Exit Function
    End If

    FAS_GetBookmarksFilePath = FAS_GetBookmarksFolderPath() & "\" & FAS_SanitizeBookmarkFileName(targetWorkbook.Name) & "_" & FAS_GetBookmarkStableHash(LCase$(Trim$(targetWorkbook.FullName))) & ".txt"
End Function

Private Function FAS_GetBookmarksFolderPath() As String
    Dim appDataPath As String

    appDataPath = Environ$("APPDATA")
    If Len(appDataPath) = 0 Then
        Exit Function
    End If

    FAS_GetBookmarksFolderPath = appDataPath & "\" & FAS_BOOKMARKS_FOLDER_NAME
End Function

Private Function FAS_NormalizeBookmarkName(ByVal bookmarkName As String) As String
    FAS_NormalizeBookmarkName = Replace(FAS_NormalizeSheetText(bookmarkName), "|", vbNullString)
End Function

Private Function FAS_EncodeBookmarkField(ByVal value As String) As String
    value = Replace(value, "%", "%25")
    value = Replace(value, "|", "%7C")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    FAS_EncodeBookmarkField = value
End Function

Private Function FAS_DecodeBookmarkField(ByVal value As String) As String
    value = Replace(value, "%7C", "|")
    value = Replace(value, "%25", "%")
    FAS_DecodeBookmarkField = value
End Function

Private Function FAS_GetBookmarkStableHash(ByVal value As String) As String
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

    FAS_GetBookmarkStableHash = Right$("00000000" & Hex$(CLng(hashValue)), 8)
End Function

Private Function FAS_SanitizeBookmarkFileName(ByVal value As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In invalidChars
        value = Replace(value, CStr(item), "_")
    Next item

    If Len(value) = 0 Then
        value = "Workbook"
    End If

    FAS_SanitizeBookmarkFileName = value
End Function

Private Sub FAS_EnsureBookmarkFolderExists(ByVal fileSystem As Object, ByVal folderPath As String)
    Dim parentPath As String

    If fileSystem.FolderExists(folderPath) Then
        Exit Sub
    End If

    parentPath = fileSystem.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then
        If Not fileSystem.FolderExists(parentPath) Then
            FAS_EnsureBookmarkFolderExists fileSystem, parentPath
        End If
    End If

    fileSystem.CreateFolder folderPath
End Sub
