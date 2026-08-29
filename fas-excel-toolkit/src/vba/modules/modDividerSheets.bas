Attribute VB_Name = "modDividerSheets"
Option Explicit

Public Sub FAS_InsertDividerSheet()
    Dim targetWorkbook As Workbook
    Dim dividerTitle As String

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        FAS_SetStatus "No active user workbook available."
        Exit Sub
    End If

    dividerTitle = InputBox("Divider title:", "FAS Divider Sheet", "P&Ls >")
    If Len(Trim$(dividerTitle)) = 0 Then
        Exit Sub
    End If

    FAS_CreateDividerSheet targetWorkbook, dividerTitle, ActiveSheet.Index
End Sub

Public Function FAS_CreateDividerSheet(ByVal targetWorkbook As Workbook, ByVal dividerTitle As String, Optional ByVal beforeSheetIndex As Long = 0) As Worksheet
    Dim dividerSheet As Worksheet
    Dim sheetName As String

    If targetWorkbook Is Nothing Then
        Exit Function
    End If

    dividerTitle = Trim$(dividerTitle)
    If Len(dividerTitle) = 0 Then
        Exit Function
    End If

    sheetName = FAS_GetUniqueDividerSheetName(targetWorkbook, dividerTitle)
    If beforeSheetIndex >= 1 And beforeSheetIndex <= targetWorkbook.Worksheets.Count Then
        Set dividerSheet = targetWorkbook.Worksheets.Add(Before:=targetWorkbook.Worksheets(beforeSheetIndex))
    Else
        Set dividerSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
    End If

    dividerSheet.Name = sheetName
    dividerSheet.Tab.Color = RGB(31, 78, 121)
    dividerSheet.Range("A1:Z40").Interior.Color = RGB(31, 78, 121)
    dividerSheet.Range("B4").Value = dividerTitle
    dividerSheet.Range("B4").Font.Color = RGB(255, 255, 255)
    dividerSheet.Range("B4").Font.Bold = True
    dividerSheet.Range("B4").Font.Size = 28
    dividerSheet.Range("B4").Font.Name = "Arial"
    dividerSheet.Rows("4").RowHeight = 50
    dividerSheet.Activate
    FAS_SetStatus "Inserted divider sheet: " & dividerSheet.Name

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    Set FAS_CreateDividerSheet = dividerSheet
End Function

Public Function FAS_CreateActiveWorkbookDividerSheet(ByVal dividerTitle As String, Optional ByVal beforeSheetIndex As Long = 0) As Boolean
    Dim dividerSheet As Worksheet

    Set dividerSheet = FAS_CreateDividerSheet(FAS_GetTargetWorkbook(), dividerTitle, beforeSheetIndex)
    FAS_CreateActiveWorkbookDividerSheet = Not dividerSheet Is Nothing
End Function

Private Function FAS_GetUniqueDividerSheetName(ByVal targetWorkbook As Workbook, ByVal dividerTitle As String) As String
    Dim baseName As String
    Dim candidateName As String
    Dim suffix As Long

    baseName = Left$(FAS_SanitizeDividerSheetName(dividerTitle), 27)
    If Len(baseName) = 0 Then
        baseName = "Divider"
    End If

    candidateName = baseName
    suffix = 1
    Do While FAS_WorksheetNameExists(targetWorkbook, candidateName)
        suffix = suffix + 1
        candidateName = Left$(baseName, 27) & " " & CStr(suffix)
    Loop

    FAS_GetUniqueDividerSheetName = candidateName
End Function

Private Function FAS_SanitizeDividerSheetName(ByVal value As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array("\", "/", ":", "*", "?", "[", "]")
    For Each item In invalidChars
        value = Replace(value, CStr(item), " ")
    Next item

    FAS_SanitizeDividerSheetName = Application.WorksheetFunction.Trim(value)
End Function

Private Function FAS_WorksheetNameExists(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Boolean
    Dim targetSheet As Worksheet

    For Each targetSheet In targetWorkbook.Worksheets
        If StrComp(targetSheet.Name, sheetName, vbTextCompare) = 0 Then
            FAS_WorksheetNameExists = True
            Exit Function
        End If
    Next targetSheet
End Function
