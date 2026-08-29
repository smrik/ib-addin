VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFASSettings 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFASSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFASSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "FAS Settings"
    Me.txtNumberShortcut.Text = FAS_GetConfiguredShortcut("numberformat")
    Me.txtPercentShortcut.Text = FAS_GetConfiguredShortcut("percentformat")
    Me.txtCurrencyShortcut.Text = FAS_GetConfiguredShortcut("currencyformat")
    Me.txtNumberFormats.Text = FAS_GetFormatOptionsText("number")
    Me.txtPercentFormats.Text = FAS_GetFormatOptionsText("percent")
    Me.txtCurrencyFormats.Text = FAS_GetFormatOptionsText("currency")
    ConfigureAppearance
End Sub

Private Sub UserForm_Terminate()
    FAS_ClearSettingsReference
End Sub

Private Sub cmdSave_Click()
    SaveSettings
End Sub

Private Sub cmdReset_Click()
    If MsgBox("Restore the default shortcuts and format lists?", vbQuestion + vbYesNo, "FAS Settings") <> vbYes Then Exit Sub

    FAS_ResetFormatSettings
    FAS_ResetShortcuts
    FAS_RegisterConfiguredShortcuts
    LoadSettings
    FAS_SetStatus "FAS settings restored to defaults."
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub LoadSettings()
    Me.txtNumberShortcut.Text = FAS_GetConfiguredShortcut("numberformat")
    Me.txtPercentShortcut.Text = FAS_GetConfiguredShortcut("percentformat")
    Me.txtCurrencyShortcut.Text = FAS_GetConfiguredShortcut("currencyformat")
    Me.txtNumberFormats.Text = FAS_GetFormatOptionsText("number")
    Me.txtPercentFormats.Text = FAS_GetFormatOptionsText("percent")
    Me.txtCurrencyFormats.Text = FAS_GetFormatOptionsText("currency")
End Sub

Private Sub SaveSettings()
    Dim errorText As String

    If Not FAS_SaveFormatOptions("number", Me.txtNumberFormats.Text, errorText) Then
        MsgBox errorText, vbExclamation, "FAS Settings"
        Exit Sub
    End If
    If Not FAS_SaveFormatOptions("percent", Me.txtPercentFormats.Text, errorText) Then
        MsgBox errorText, vbExclamation, "FAS Settings"
        Exit Sub
    End If
    If Not FAS_SaveFormatOptions("currency", Me.txtCurrencyFormats.Text, errorText) Then
        MsgBox errorText, vbExclamation, "FAS Settings"
        Exit Sub
    End If

    If Not FAS_SaveShortcut("numberformat", Me.txtNumberShortcut.Text, errorText) Then
        MsgBox errorText, vbExclamation, "FAS Settings"
        Exit Sub
    End If
    If Not FAS_SaveShortcut("percentformat", Me.txtPercentShortcut.Text, errorText) Then
        MsgBox errorText, vbExclamation, "FAS Settings"
        Exit Sub
    End If
    If Not FAS_SaveShortcut("currencyformat", Me.txtCurrencyShortcut.Text, errorText) Then
        MsgBox errorText, vbExclamation, "FAS Settings"
        Exit Sub
    End If

    FAS_RegisterConfiguredShortcuts
    FAS_SetStatus "FAS settings saved."
End Sub

Private Sub ConfigureAppearance()
    On Error Resume Next
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 10
    Me.lblTitle.Font.Bold = True
    Me.lblTitle.Font.Size = 14
    Me.lblShortcuts.Font.Bold = True
    Me.lblFormats.Font.Bold = True
    Me.lblShortcutHint.ForeColor = RGB(96, 96, 96)
    Me.lblIntro.ForeColor = RGB(96, 96, 96)
    Me.txtNumberFormats.Font.Name = "Consolas"
    Me.txtPercentFormats.Font.Name = "Consolas"
    Me.txtCurrencyFormats.Font.Name = "Consolas"
    On Error GoTo 0
End Sub
