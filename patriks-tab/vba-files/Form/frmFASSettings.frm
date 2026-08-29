VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFASSettings
   Caption         =   "FAS Settings"
   ClientHeight    =   7600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9200
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle
      Caption         =   "FAS Settings"
      Height          =   300
      Left            =   240
      Top             =   180
      Width           =   3000
   End
   Begin MSForms.Label lblIntro
      Caption         =   "Set shortcuts and custom format cycles. Enter one Excel custom format per line."
      Height          =   300
      Left            =   240
      Top             =   540
      Width           =   8500
   End
   Begin MSForms.Label lblShortcuts
      Caption         =   "Keyboard shortcuts"
      Height          =   240
      Left            =   240
      Top             =   960
      Width           =   2500
   End
   Begin MSForms.Label lblNumberShortcut
      Caption         =   "Number cycle"
      Height          =   240
      Left            =   240
      Top             =   1320
      Width           =   1800
   End
   Begin MSForms.TextBox txtNumberShortcut
      Height          =   285
      Left            =   2100
      Top             =   1290
      Width           =   1800
      TabIndex        =   0
   End
   Begin MSForms.Label lblPercentShortcut
      Caption         =   "Percent cycle"
      Height          =   240
      Left            =   4200
      Top             =   1320
      Width           =   1800
   End
   Begin MSForms.TextBox txtPercentShortcut
      Height          =   285
      Left            =   6060
      Top             =   1290
      Width           =   1800
      TabIndex        =   1
   End
   Begin MSForms.Label lblCurrencyShortcut
      Caption         =   "Currency cycle"
      Height          =   240
      Left            =   240
      Top             =   1680
      Width           =   1800
   End
   Begin MSForms.TextBox txtCurrencyShortcut
      Height          =   285
      Left            =   2100
      Top             =   1650
      Width           =   1800
      TabIndex        =   2
   End
   Begin MSForms.Label lblShortcutHint
      Caption         =   "Leave a shortcut blank to restore its default. Example: Ctrl+Shift+1"
      Height          =   240
      Left            =   4200
      Top             =   1680
      Width           =   4300
   End
   Begin MSForms.Label lblFormats
      Caption         =   "Format cycles"
      Height          =   240
      Left            =   240
      Top             =   2130
      Width           =   2500
   End
   Begin MSForms.Label lblNumberFormats
      Caption         =   "Number formats"
      Height          =   240
      Left            =   240
      Top             =   2490
      Width           =   2600
   End
   Begin MSForms.TextBox txtNumberFormats
      Height          =   1020
      Left            =   240
      Top             =   2760
      Width           =   8500
      MultiLine       =   -1  'True
      EnterKeyBehavior =   -1  'True
      ScrollBars      =   2
      TabIndex        =   3
   End
   Begin MSForms.Label lblPercentFormats
      Caption         =   "Percent formats"
      Height          =   240
      Left            =   240
      Top             =   3900
      Width           =   2600
   End
   Begin MSForms.TextBox txtPercentFormats
      Height          =   1020
      Left            =   240
      Top             =   4170
      Width           =   8500
      MultiLine       =   -1  'True
      EnterKeyBehavior =   -1  'True
      ScrollBars      =   2
      TabIndex        =   4
   End
   Begin MSForms.Label lblCurrencyFormats
      Caption         =   "Currency formats"
      Height          =   240
      Left            =   240
      Top             =   5310
      Width           =   2600
   End
   Begin MSForms.TextBox txtCurrencyFormats
      Height          =   1020
      Left            =   240
      Top             =   5580
      Width           =   8500
      MultiLine       =   -1  'True
      EnterKeyBehavior =   -1  'True
      ScrollBars      =   2
      TabIndex        =   5
   End
   Begin MSForms.CommandButton cmdSave
      Caption         =   "Save"
      Height          =   360
      Left            =   240
      Top             =   6900
      Width           =   1700
      TabIndex        =   6
   End
   Begin MSForms.CommandButton cmdReset
      Caption         =   "Reset defaults"
      Height          =   360
      Left            =   2100
      Top             =   6900
      Width           =   1700
      TabIndex        =   7
   End
   Begin MSForms.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   360
      Left            =   7040
      Top             =   6900
      Width           =   1700
      TabIndex        =   8
   End
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
