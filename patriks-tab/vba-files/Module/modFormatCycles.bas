Attribute VB_Name = "modFormatCycles"
Option Explicit

' Format lists are stored as one custom Excel format per line in the settings form.
' The settings file uses "|" as the line separator so semicolons remain valid format syntax.
Private Const FAS_NUMBER_FORMAT_KEY As String = "Format.Number.Options"
Private Const FAS_PERCENT_FORMAT_KEY As String = "Format.Percent.Options"
Private Const FAS_CURRENCY_FORMAT_KEY As String = "Format.Currency.Options"

Private Const FAS_NUMBER_FORMAT_DEFAULT As String = "#,##0;(#,##0);-|#,##0.0;(#,##0.0);-|#,##0.00;(#,##0.00);-"
Private Const FAS_PERCENT_FORMAT_DEFAULT As String = "0%;(0%);-|0.0%;(0.0%);-|0.00%;(0.00%);-"
Private Const FAS_CURRENCY_FORMAT_DEFAULT As String = "#,##0;(#,##0);-|#,##0.00;(#,##0.00);-|€ #,##0;[Red](€ #,##0);-"

Public Sub FAS_CycleNumberFormat()
    FAS_CycleSelectedFormat "number", "Number"
End Sub

Public Sub FAS_CyclePercentFormat()
    FAS_CycleSelectedFormat "percent", "Percent"
End Sub

Public Sub FAS_CycleCurrencyFormat()
    FAS_CycleSelectedFormat "currency", "Currency"
End Sub

Public Function FAS_TryRunFormatCycleCommand(ByVal commandText As String) As Boolean
    Select Case LCase$(Trim$(commandText))
        Case "format", "formats", "format settings", "formatsettings", "number formats"
            FAS_ShowSettings
        Case "format number", "number format", "number", "numberformat", "fmt1"
            FAS_CycleNumberFormat
        Case "format percent", "percent format", "percent", "percentage", "percentformat", "fmt5"
            FAS_CyclePercentFormat
        Case "format currency", "currency format", "currency", "currencyformat", "fmt4"
            FAS_CycleCurrencyFormat
        Case Else
            Exit Function
    End Select

    FAS_TryRunFormatCycleCommand = True
End Function

Public Function FAS_IsFormatCycleCommand(ByVal commandText As String) As Boolean
    Select Case LCase$(Trim$(commandText))
        Case "format", "formats", "format settings", "formatsettings", "number formats", _
             "format number", "number format", "number", "numberformat", "fmt1", _
             "format percent", "percent format", "percent", "percentage", "percentformat", "fmt5", _
             "format currency", "currency format", "currency", "currencyformat", "fmt4"
            FAS_IsFormatCycleCommand = True
    End Select
End Function

Public Function FAS_GetFormatOptionsText(ByVal formatKind As String) As String
    Dim storedText As String

    storedText = FAS_GetSetting(FAS_FormatSettingKey(formatKind), FAS_DefaultFormatList(formatKind))
    FAS_GetFormatOptionsText = Replace(storedText, "|", vbCrLf)
End Function

Public Function FAS_SaveFormatOptions(ByVal formatKind As String, ByVal multilineText As String, Optional ByRef errorText As String) As Boolean
    Dim normalizedText As String
    Dim parts As Variant
    Dim partIndex As Long
    Dim formatText As String
    Dim savedText As String
    Dim optionCount As Long

    errorText = vbNullString
    If Len(FAS_FormatSettingKey(formatKind)) = 0 Then
        errorText = "Unknown format family."
        Exit Function
    End If

    normalizedText = Replace(multilineText, vbCrLf, "|")
    normalizedText = Replace(normalizedText, vbCr, "|")
    normalizedText = Replace(normalizedText, vbLf, "|")
    parts = Split(normalizedText, "|")

    For partIndex = LBound(parts) To UBound(parts)
        formatText = Trim$(CStr(parts(partIndex)))
        If Len(formatText) > 0 Then
            If InStr(1, formatText, "|", vbBinaryCompare) > 0 Then
                errorText = "Use one Excel format per line; the pipe character is reserved."
                Exit Function
            End If
            If Len(savedText) > 0 Then savedText = savedText & "|"
            savedText = savedText & formatText
            optionCount = optionCount + 1
        End If
    Next partIndex

    If optionCount = 0 Then
        errorText = "Enter at least one custom Excel format."
        Exit Function
    End If

    FAS_SetSetting FAS_FormatSettingKey(formatKind), savedText
    FAS_SaveFormatOptions = True
End Function

Public Sub FAS_ResetFormatSettings()
    FAS_DeleteSetting FAS_NUMBER_FORMAT_KEY
    FAS_DeleteSetting FAS_PERCENT_FORMAT_KEY
    FAS_DeleteSetting FAS_CURRENCY_FORMAT_KEY
End Sub

Private Sub FAS_CycleSelectedFormat(ByVal formatKind As String, ByVal displayName As String)
    Dim targetRange As Range
    Dim formatOptions As Collection
    Dim currentValue As Variant
    Dim currentText As String
    Dim nextText As String

    If TypeName(Selection) <> "Range" Then
        FAS_SetStatus "Select cells before cycling a " & LCase$(displayName) & " format."
        Exit Sub
    End If

    Set targetRange = Selection
    If Not FAS_UndoBeginRange(targetRange, "Cycle " & displayName & " format") Then Exit Sub

    On Error GoTo FormatFailed
    Set formatOptions = FAS_GetFormatOptions(formatKind)
    currentValue = targetRange.NumberFormatLocal
    If IsNull(currentValue) Then
        currentText = vbNullString
    Else
        currentText = CStr(currentValue)
    End If

    nextText = FAS_GetNextFormatText(formatOptions, currentText)
    targetRange.NumberFormatLocal = nextText
    FAS_UndoCommit
    FAS_SetStatus displayName & " format: " & nextText
    Exit Sub

FormatFailed:
    FAS_UndoCancel
    FAS_SetStatus "Could not apply the selected " & LCase$(displayName) & " format."
End Sub

Private Function FAS_GetFormatOptions(ByVal formatKind As String) As Collection
    Dim options As Collection
    Dim parts As Variant
    Dim partIndex As Long
    Dim formatText As String

    Set options = New Collection
    parts = Split(FAS_GetSetting(FAS_FormatSettingKey(formatKind), FAS_DefaultFormatList(formatKind)), "|")
    For partIndex = LBound(parts) To UBound(parts)
        formatText = Trim$(CStr(parts(partIndex)))
        If Len(formatText) > 0 Then options.Add formatText
    Next partIndex

    If options.Count = 0 Then options.Add FAS_DefaultFormatList(formatKind)
    Set FAS_GetFormatOptions = options
End Function

Private Function FAS_GetNextFormatText(ByVal options As Collection, ByVal currentText As String) As String
    Dim optionIndex As Long

    For optionIndex = 1 To options.Count
        If StrComp(CStr(options(optionIndex)), currentText, vbTextCompare) = 0 Then
            If optionIndex = options.Count Then
                FAS_GetNextFormatText = CStr(options(1))
            Else
                FAS_GetNextFormatText = CStr(options(optionIndex + 1))
            End If
            Exit Function
        End If
    Next optionIndex

    FAS_GetNextFormatText = CStr(options(1))
End Function

Private Function FAS_FormatSettingKey(ByVal formatKind As String) As String
    Select Case LCase$(Trim$(formatKind))
        Case "number": FAS_FormatSettingKey = FAS_NUMBER_FORMAT_KEY
        Case "percent": FAS_FormatSettingKey = FAS_PERCENT_FORMAT_KEY
        Case "currency": FAS_FormatSettingKey = FAS_CURRENCY_FORMAT_KEY
    End Select
End Function

Private Function FAS_DefaultFormatList(ByVal formatKind As String) As String
    Select Case LCase$(Trim$(formatKind))
        Case "number": FAS_DefaultFormatList = FAS_NUMBER_FORMAT_DEFAULT
        Case "percent": FAS_DefaultFormatList = FAS_PERCENT_FORMAT_DEFAULT
        Case "currency": FAS_DefaultFormatList = FAS_CURRENCY_FORMAT_DEFAULT
    End Select
End Function
