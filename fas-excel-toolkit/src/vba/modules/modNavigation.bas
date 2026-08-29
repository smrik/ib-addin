Attribute VB_Name = "modNavigation"
Option Explicit

Public Sub FAS_JumpToSheet()
    FAS_ShowCommandPalette vbNullString
End Sub

Public Sub FAS_GoToPL()
    FAS_ActivateBestSheetByGroup "PL"
End Sub

Public Sub FAS_GoToBS()
    FAS_ActivateBestSheetByGroup "BS"
End Sub

Public Sub FAS_GoToTB()
    FAS_ActivateBestSheetByGroup "TB"
End Sub

Public Sub FAS_GoToGL()
    FAS_ActivateBestSheetByGroup "GL"
End Sub

Public Sub FAS_GoToMapping()
    FAS_ActivateBestSheetByGroup "MAPPING"
End Sub

Public Sub FAS_ShowWorkbookTabsMenu()
    Dim targetWorkbook As Workbook
    Dim prompt As String
    Dim currentSheet As Worksheet
    Dim userInput As String
    Dim selectedIndex As Long

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        FAS_SetStatus "No active user workbook available."
        Exit Sub
    End If

    prompt = "Enter the sheet number to activate:" & vbCrLf & vbCrLf
    selectedIndex = 0

    For Each currentSheet In targetWorkbook.Worksheets
        selectedIndex = selectedIndex + 1
        prompt = prompt & selectedIndex & ". " & currentSheet.Name & vbCrLf
    Next currentSheet

    userInput = InputBox(prompt, "FAS Workbook Tabs")
    If Len(Trim$(userInput)) = 0 Then
        Exit Sub
    End If

    If Not IsNumeric(userInput) Then
        FAS_SetStatus "Enter a valid sheet number."
        Exit Sub
    End If

    selectedIndex = CLng(userInput)
    If selectedIndex < 1 Or selectedIndex > targetWorkbook.Worksheets.Count Then
        FAS_SetStatus "Sheet number is outside the list."
        Exit Sub
    End If

    targetWorkbook.Worksheets(selectedIndex).Activate
End Sub

Private Sub FAS_ActivateBestSheetByGroup(ByVal groupKey As String)
    Dim targetWorkbook As Workbook
    Dim bestSheet As Worksheet

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        FAS_SetStatus "No active user workbook available."
        Exit Sub
    End If

    Set bestSheet = FAS_GetResolvedGroupSheet(targetWorkbook, groupKey, 1)
    If bestSheet Is Nothing Then
        FAS_SetStatus "No " & FAS_GetGroupLabel(groupKey) & " sheet found."
        Exit Sub
    End If

    bestSheet.Activate
    FAS_SetStatus "Activated " & bestSheet.Name & "."
End Sub

Public Function FAS_RunJumpCommand(ByVal userInput As String) As Boolean
    FAS_RunJumpCommand = FAS_RunCommand(userInput)
End Function
