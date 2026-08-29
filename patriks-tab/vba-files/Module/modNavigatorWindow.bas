Attribute VB_Name = "modNavigatorWindow"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongA Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Public Sub FAS_EnableResizableForm(ByVal formCaption As String)
#If VBA7 Then
    Dim formHandle As LongPtr
    Dim windowStyle As LongPtr
#Else
    Dim formHandle As Long
    Dim windowStyle As Long
#End If

    If Len(formCaption) = 0 Then
        Exit Sub
    End If

    formHandle = FindWindowA("ThunderDFrame", formCaption)
    If formHandle = 0 Then
        Exit Sub
    End If

    windowStyle = FAS_GetWindowStyle(formHandle)
    windowStyle = windowStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    FAS_SetWindowStyle formHandle, windowStyle
    SetWindowPos formHandle, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

#If VBA7 Then
Private Function FAS_GetWindowStyle(ByVal formHandle As LongPtr) As LongPtr
    FAS_GetWindowStyle = GetWindowLongPtrA(formHandle, GWL_STYLE)
End Function

Private Sub FAS_SetWindowStyle(ByVal formHandle As LongPtr, ByVal windowStyle As LongPtr)
    SetWindowLongPtrA formHandle, GWL_STYLE, windowStyle
End Sub
#Else
Private Function FAS_GetWindowStyle(ByVal formHandle As Long) As Long
    FAS_GetWindowStyle = GetWindowLongA(formHandle, GWL_STYLE)
End Function

Private Sub FAS_SetWindowStyle(ByVal formHandle As Long, ByVal windowStyle As Long)
    SetWindowLongA formHandle, GWL_STYLE, windowStyle
End Sub
#End If
