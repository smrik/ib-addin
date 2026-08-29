VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFASCommandPalette 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmFASCommandPalette.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFASCommandPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIsLoading As Boolean

Private Sub UserForm_Initialize()
    mIsLoading = True
    Me.Caption = "FAS Command Palette"

    With Me.lstCommands
        .ColumnCount = 3
        .ColumnWidths = "128 pt;300 pt;0 pt"
        .BoundColumn = 1
        .IntegralHeight = False
    End With

    Me.lblEmpty.Caption = "No commands match this search."
    Me.lblEmpty.Visible = False
    mIsLoading = False
    RefreshResults
End Sub

Private Sub UserForm_Activate()
    FocusSearch
End Sub

Private Sub UserForm_Terminate()
    FAS_ClearCommandPaletteReference
End Sub

Public Sub SetInitialQuery(ByVal queryText As String)
    mIsLoading = True
    Me.txtCommand.Text = queryText
    Me.txtCommand.SelStart = Len(Me.txtCommand.Text)
    mIsLoading = False
    RefreshResults
End Sub

Public Sub FocusSearch()
    On Error Resume Next
    Me.txtCommand.SetFocus
    Me.txtCommand.SelStart = Len(Me.txtCommand.Text)
    On Error GoTo 0
End Sub

Private Sub txtCommand_Change()
    If mIsLoading Then Exit Sub
    RefreshResults
End Sub

Private Sub txtCommand_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            MoveSelection 1
            KeyCode = 0
        Case vbKeyUp
            MoveSelection -1
            KeyCode = 0
        Case vbKeyReturn
            ExecuteSelectedCommand
            KeyCode = 0
        Case vbKeyEscape
            Unload Me
            KeyCode = 0
    End Select
End Sub

Private Sub lstCommands_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ExecuteSelectedCommand
End Sub

Private Sub lstCommands_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ExecuteSelectedCommand
            KeyCode = 0
        Case vbKeyEscape
            Unload Me
            KeyCode = 0
    End Select
End Sub

Private Sub RefreshResults()
    Dim matches As Collection
    Dim item As Variant

    Set matches = FAS_FilterCommandCatalog(Me.txtCommand.Text)
    Me.lstCommands.Clear

    For Each item In matches
        Me.lstCommands.AddItem CStr(item(4)) & ": " & CStr(item(0))
        Me.lstCommands.List(Me.lstCommands.ListCount - 1, 1) = CStr(item(3))
        Me.lstCommands.List(Me.lstCommands.ListCount - 1, 2) = CStr(item(0))
    Next item

    Me.lblEmpty.Visible = (Me.lstCommands.ListCount = 0)
    If Me.lstCommands.ListCount > 0 Then
        Me.lstCommands.ListIndex = 0
    End If
End Sub

Private Sub MoveSelection(ByVal direction As Long)
    Dim nextIndex As Long

    If Me.lstCommands.ListCount = 0 Then Exit Sub

    nextIndex = Me.lstCommands.ListIndex
    If nextIndex < 0 Then nextIndex = 0
    nextIndex = nextIndex + direction
    If nextIndex < 0 Then nextIndex = Me.lstCommands.ListCount - 1
    If nextIndex >= Me.lstCommands.ListCount Then nextIndex = 0
    Me.lstCommands.ListIndex = nextIndex
End Sub

Private Sub ExecuteSelectedCommand()
    Dim commandText As String
    Dim placeholderPosition As Long
    Dim commandPrefix As String

    If Me.lstCommands.ListIndex < 0 Then Exit Sub
    commandText = CStr(Me.lstCommands.List(Me.lstCommands.ListIndex, 2))
    placeholderPosition = InStr(1, commandText, "<", vbBinaryCompare)

    If placeholderPosition > 0 Then
        commandPrefix = Trim$(Left$(commandText, placeholderPosition - 1)) & " "
        SetInitialQuery commandPrefix
        FocusSearch
        Exit Sub
    End If

    Unload Me
    If Not FAS_RunCommand(commandText) Then
        FAS_SetStatus "Command did not run: " & commandText
    End If
End Sub
