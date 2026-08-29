Attribute VB_Name = "modGridlines"
Option Explicit

Public Const FAS_DEFAULT_HIDE_GRIDLINES As Boolean = True

Private Const FAS_HIDE_GRIDLINES_SETTING As String = "View.HideGridlines"

Private mFASInitializedGridlineWindows As Object

' Keep the application event sink alive while the add-in is loaded.
Public Sub FAS_StartApplicationEvents()
    ThisWorkbook.StartFASApplicationEvents
End Sub

Public Sub FAS_StopApplicationEvents()
    ThisWorkbook.StopFASApplicationEvents
End Sub

Public Function FAS_GetHideGridlines() As Boolean
    FAS_GetHideGridlines = FAS_GetBooleanSetting(FAS_HIDE_GRIDLINES_SETTING, FAS_DEFAULT_HIDE_GRIDLINES)
End Function

Public Sub FAS_SetHideGridlines(ByVal hideGridlines As Boolean)
    FAS_SetBooleanSetting FAS_HIDE_GRIDLINES_SETTING, hideGridlines
    FAS_StartApplicationEvents
    FAS_ApplyGridlinePreferenceToAllOpenWorkbooks

    If hideGridlines Then
        FAS_SetStatus "Gridlines hidden across open workbooks, new worksheets, and new windows."
    Else
        FAS_SetStatus "Gridlines shown across open workbooks, new worksheets, and new windows."
    End If
End Sub

Public Sub FAS_ApplyGridlinePreferenceToAllOpenWorkbooks()
    FAS_ApplyGridlinePreferenceBulk
End Sub

Public Sub FAS_ApplyGridlinePreferenceToWorkbook(ByVal targetWorkbook As Workbook)
    FAS_ApplyGridlinePreferenceBulk targetWorkbook
End Sub

Public Sub FAS_EnsureGridlinePreferenceForWindow(ByVal targetWorkbook As Workbook, ByVal targetWindow As Object)
    Dim windowKey As String

    If targetWorkbook Is Nothing Then Exit Sub
    If targetWindow Is Nothing Then Exit Sub

    windowKey = FAS_GetGridlineWindowKey(targetWorkbook, targetWindow)
    If FAS_IsGridlineWindowInitialized(windowKey) Then Exit Sub

    ' WindowActivate is also raised for ordinary window switches. Only a window
    ' we have not seen before needs the worksheet-by-worksheet initialization.
    FAS_ApplyGridlinePreferenceBulk targetWorkbook, targetWindow
End Sub

Private Sub FAS_ApplyGridlinePreferenceBulk(Optional ByVal onlyWorkbook As Workbook, Optional ByVal onlyWindow As Object)
    Dim originalWorkbook As Workbook
    Dim originalWindow As Window
    Dim originalSheet As Object
    Dim eventsWereEnabled As Boolean
    Dim screenWasUpdating As Boolean
    Dim targetWorkbook As Workbook
    Dim errorText As String

    On Error GoTo BulkFailed

    Set originalWorkbook = Application.ActiveWorkbook
    Set originalWindow = Application.ActiveWindow
    Set originalSheet = Application.ActiveSheet
    eventsWereEnabled = Application.EnableEvents
    screenWasUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If Not onlyWindow Is Nothing Then
        FAS_ApplyGridlinePreferenceToWindowInternal onlyWorkbook, onlyWindow
    ElseIf onlyWorkbook Is Nothing Then
        For Each targetWorkbook In Application.Workbooks
            FAS_ApplyGridlinePreferenceToWorkbookInternal targetWorkbook
        Next targetWorkbook
    Else
        FAS_ApplyGridlinePreferenceToWorkbookInternal onlyWorkbook
    End If

BulkExit:
    On Error Resume Next
    If Not originalWorkbook Is Nothing Then originalWorkbook.Activate
    If Not originalWindow Is Nothing Then originalWindow.Activate
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Application.ScreenUpdating = screenWasUpdating
    Application.EnableEvents = eventsWereEnabled
    On Error GoTo 0

    If Len(errorText) > 0 Then
        On Error Resume Next
        If Application.Visible Then FAS_SetStatus "Gridlines could not be updated for every worksheet."
        On Error GoTo 0
    End If
    Exit Sub

BulkFailed:
    errorText = Err.Description
    Resume BulkExit
End Sub

Private Sub FAS_ApplyGridlinePreferenceToWorkbookInternal(ByVal targetWorkbook As Workbook)
    Dim targetWindow As Object

    If targetWorkbook Is Nothing Then Exit Sub

    For Each targetWindow In targetWorkbook.Windows
        FAS_ApplyGridlinePreferenceToWindowInternal targetWorkbook, targetWindow
    Next targetWindow
End Sub

Private Sub FAS_ApplyGridlinePreferenceToWindowInternal(ByVal targetWorkbook As Workbook, ByVal targetWindow As Object)
    Dim targetSheet As Worksheet
    Dim windowWasActivated As Boolean

    If targetWorkbook Is Nothing Then Exit Sub
    If targetWindow Is Nothing Then Exit Sub

    On Error Resume Next
    Err.Clear
    If targetWindow.Visible Then
        targetWindow.Activate
        windowWasActivated = (Err.Number = 0)

        If windowWasActivated Then
            For Each targetSheet In targetWorkbook.Worksheets
                If targetSheet.Visible = xlSheetVisible Then
                    Err.Clear
                    targetSheet.Activate
                    If Err.Number = 0 Then
                        targetWindow.DisplayGridlines = Not FAS_GetHideGridlines()
                    End If
                End If
            Next targetSheet
        End If
    End If
    On Error GoTo 0

    If windowWasActivated Then FAS_MarkGridlineWindowInitialized targetWorkbook, targetWindow
End Sub

Private Function FAS_GetGridlineWindowKey(ByVal targetWorkbook As Workbook, ByVal targetWindow As Object) As String
    On Error Resume Next
    ' Window captions distinguish Book1:1, Book1:2, and so on. A later rename may
    ' cause one harmless reinitialization, which is safer than a Win32 dependency.
    FAS_GetGridlineWindowKey = targetWorkbook.Name & "|" & targetWindow.Caption
    On Error GoTo 0
End Function

Private Function FAS_IsGridlineWindowInitialized(ByVal windowKey As String) As Boolean
    If Len(windowKey) = 0 Then Exit Function
    If mFASInitializedGridlineWindows Is Nothing Then Exit Function

    FAS_IsGridlineWindowInitialized = mFASInitializedGridlineWindows.Exists(windowKey)
End Function

Private Sub FAS_MarkGridlineWindowInitialized(ByVal targetWorkbook As Workbook, ByVal targetWindow As Object)
    Dim windowKey As String

    windowKey = FAS_GetGridlineWindowKey(targetWorkbook, targetWindow)
    If Len(windowKey) = 0 Then Exit Sub

    If mFASInitializedGridlineWindows Is Nothing Then
        Set mFASInitializedGridlineWindows = CreateObject("Scripting.Dictionary")
    End If

    mFASInitializedGridlineWindows(windowKey) = True
End Sub

Public Sub FAS_ApplyGridlinePreference(Optional ByVal targetWindow As Object, Optional ByVal targetSheet As Object)
    Dim gridlineWindow As Object
    Dim gridlineSheet As Object

    If Not targetWindow Is Nothing Then Set gridlineWindow = targetWindow
    If gridlineWindow Is Nothing Then Set gridlineWindow = Application.ActiveWindow
    If gridlineWindow Is Nothing Then Exit Sub
    If Not targetSheet Is Nothing Then Set gridlineSheet = targetSheet
    If gridlineSheet Is Nothing Then Set gridlineSheet = gridlineWindow.ActiveSheet
    If gridlineSheet Is Nothing Then Set gridlineSheet = Application.ActiveSheet
    If TypeName(gridlineSheet) <> "Worksheet" Then Exit Sub

    On Error GoTo ApplyFailed
    gridlineWindow.DisplayGridlines = Not FAS_GetHideGridlines()
    Exit Sub

ApplyFailed:
    On Error Resume Next
    If Application.Visible Then FAS_SetStatus "Could not update worksheet gridlines for this window."
    On Error GoTo 0
End Sub
