Attribute VB_Name = "modWorkbookTarget"
Option Explicit

Public Function FAS_GetTargetWorkbook() As Workbook
    Dim candidate As Workbook

    On Error Resume Next
    Set candidate = ActiveWorkbook
    On Error GoTo 0

    If candidate Is Nothing Then
        Exit Function
    End If

    If FAS_IsAddInWorkbook(candidate) Then
        Exit Function
    End If

    Set FAS_GetTargetWorkbook = candidate
End Function

Private Function FAS_IsAddInWorkbook(ByVal candidate As Workbook) As Boolean
    If candidate Is Nothing Then
        Exit Function
    End If

    If candidate Is ThisWorkbook Then
        FAS_IsAddInWorkbook = True
        Exit Function
    End If

    On Error Resume Next
    If candidate.IsAddin Then
        FAS_IsAddInWorkbook = True
        Exit Function
    End If
    On Error GoTo 0

    If LenB(ThisWorkbook.FullName) > 0 Then
        FAS_IsAddInWorkbook = (StrComp(candidate.FullName, ThisWorkbook.FullName, vbTextCompare) = 0)
    End If
End Function
