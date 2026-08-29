VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFASCommandPalette 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
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
    Me.Caption = "FAS Command Center"
    ConfigureAppearance

    With Me.lstCommands
        .columnCount = 3
        .BoundColumn = 1
        .IntegralHeight = False
    End With

    Me.lblEmpty.Caption = "No commands match this search. Try 'help'."
    Me.lblEmpty.Visible = False
    ApplyLayout
    mIsLoading = False
    RefreshResults
End Sub

Private Sub UserForm_Resize()
    If mIsLoading Then Exit Sub
    ApplyLayout
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

Private Sub ConfigureAppearance()
    On Error Resume Next
    Me.Width = 680
    Me.Height = 430
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 10
    Me.BackColor = RGB(24, 31, 45)
    Me.txtCommand.BackColor = RGB(42, 55, 75)
    Me.txtCommand.ForeColor = RGB(245, 247, 250)
    Me.txtCommand.Font.Name = "Segoe UI"
    Me.txtCommand.Font.Size = 12
    Me.lstCommands.BackColor = RGB(32, 43, 59)
    Me.lstCommands.ForeColor = RGB(235, 240, 246)
    Me.lstCommands.Font.Name = "Segoe UI"
    Me.lstCommands.Font.Size = 10
    Me.lblEmpty.ForeColor = RGB(190, 198, 210)
    On Error GoTo 0
End Sub

Private Sub ApplyLayout()
    Dim contentWidth As Single
    Dim listTop As Single
    Dim listHeight As Single

    If Me.Width < 540 Then Me.Width = 540
    If Me.Height < 330 Then Me.Height = 330

    contentWidth = Me.InsideWidth - 24
    listTop = 50
    listHeight = Me.InsideHeight - listTop - 18
    If listHeight < 120 Then listHeight = 120

    Me.txtCommand.Left = 12
    Me.txtCommand.Top = 12
    Me.txtCommand.Width = contentWidth
    Me.txtCommand.Height = 26

    Me.lstCommands.Left = 12
    Me.lstCommands.Top = listTop
    Me.lstCommands.Width = contentWidth
    Me.lstCommands.Height = listHeight
    Me.lstCommands.ColumnWidths = "145 pt;" & CStr(Application.Max(contentWidth - 178, 260)) & " pt;0 pt"

    Me.lblEmpty.Left = 20
    Me.lblEmpty.Top = listTop + 12
    Me.lblEmpty.Width = contentWidth - 16
End Sub
