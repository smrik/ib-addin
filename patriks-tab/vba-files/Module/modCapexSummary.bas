Attribute VB_Name = "modCapexSummary"
Option Explicit

' Kept as the approved public CAPEX command entry point. The implementation is
' the generic editable Group By engine; CAPEX only supplies source columns.
Public Sub FAS_CreateCapexSummary()
    FAS_CreateCapexSummaryFromGroupBy
End Sub

Public Function FAS_TryRunCapexCommand(ByVal commandText As String) As Boolean
    Select Case LCase$(FAS_NormalizeCommandText(commandText))
        Case "capex", "capexsummary"
            FAS_CreateCapexSummary
        Case Else
            Exit Function
    End Select
    FAS_TryRunCapexCommand = True
End Function

Public Function FAS_IsCapexCommand(ByVal commandText As String) As Boolean
    Select Case LCase$(FAS_NormalizeCommandText(commandText))
        Case "capex", "capexsummary"
            FAS_IsCapexCommand = True
    End Select
End Function

' Header normalization is public so the xVBA unit tests can verify aliases.
Public Function FAS_NormalizeCapexHeader(ByVal headerValue As Variant) As String
    Dim headerText As String

    If IsError(headerValue) Or IsNull(headerValue) Or IsEmpty(headerValue) Then Exit Function
    headerText = LCase$(Trim$(CStr(headerValue)))
    headerText = Replace(headerText, Chr$(160), " ")
    Do While InStr(1, headerText, "  ", vbBinaryCompare) > 0
        headerText = Replace(headerText, "  ", " ")
    Loop
    FAS_NormalizeCapexHeader = headerText
End Function

' Returns the first plausible four-digit year found in a title.
Public Function FAS_ExtractCapexYear(ByVal titleText As String) As String
    Dim characterIndex As Long
    Dim candidateText As String
    Dim candidateYear As Long

    For characterIndex = 1 To Len(titleText) - 3
        candidateText = Mid$(titleText, characterIndex, 4)
        If IsNumeric(candidateText) Then
            candidateYear = CLng(candidateText)
            If candidateYear >= 1900 And candidateYear <= 2100 Then
                FAS_ExtractCapexYear = candidateText
                Exit Function
            End If
        End If
    Next characterIndex
End Function
