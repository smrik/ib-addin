VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mFASApplication As Application
Private mFASApplicationEventsStarted As Boolean

Private Sub Workbook_Open()
    FAS_RegisterConfiguredShortcuts
    FAS_StartApplicationEvents
    FAS_ApplyGridlinePreferenceToAllOpenWorkbooks
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    FAS_StopApplicationEvents
    FAS_UnregisterConfiguredShortcuts
End Sub

' Start the application-level event sink once. A WithEvents reference must be
' retained by the add-in; a local variable would be released immediately.
Public Sub StartFASApplicationEvents()
    If mFASApplicationEventsStarted Then Exit Sub

    Set mFASApplication = Application
    mFASApplicationEventsStarted = True
End Sub

Public Sub StopFASApplicationEvents()
    mFASApplicationEventsStarted = False
    Set mFASApplication = Nothing
End Sub

' A new window gets its own DisplayGridlines setting. Initialize it once when
' that window is activated, including windows created with Alt+W, N.
Private Sub mFASApplication_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    FAS_EnsureGridlinePreferenceForWindow Wb, Wn
End Sub

' Sheet changes are frequent, so update only the active window rather than
' rescanning every worksheet on every activation.
Private Sub mFASApplication_SheetActivate(ByVal Sh As Object)
    If TypeName(Sh) = "Worksheet" Then
        FAS_ApplyGridlinePreference Application.ActiveWindow, Sh
    End If
End Sub

' Apply the preference to a workbook as soon as it is opened.
Private Sub mFASApplication_WorkbookOpen(ByVal Wb As Workbook)
    FAS_ApplyGridlinePreferenceToWorkbook Wb
End Sub

' New worksheets can appear without a normal sheet-switch sequence. Applying
' the preference to the workbook also covers any additional workbook windows.
Private Sub mFASApplication_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
    If TypeName(Sh) = "Worksheet" Then
        FAS_ApplyGridlinePreferenceToWorkbook Wb
    End If
End Sub

