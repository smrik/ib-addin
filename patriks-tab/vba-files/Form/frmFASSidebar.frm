VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFASSidebar 
   Caption         =   "UserForm3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFASSidebar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFASSidebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const FAS_FORM_MIN_WIDTH As Single = 252
Private Const FAS_FORM_MIN_HEIGHT As Single = 320
Private Const FAS_MARGIN As Single = 12
Private Const FAS_ROW_GAP As Single = 8
Private Const FAS_BUTTON_HEIGHT As Single = 24

Private mIsLoading As Boolean

Private Sub UserForm_Initialize()
    mIsLoading = True
    PopulateGroupPicker
    ConfigureListBox
    PositionForm
    ApplyLayout
    RefreshSheetList
    mIsLoading = False
End Sub

Private Sub UserForm_Activate()
    FAS_EnableResizableForm Me.Caption
End Sub

Private Sub UserForm_Resize()
    If mIsLoading Then
        Exit Sub
    End If

    ApplyLayout
End Sub

Private Sub cmdAssign_Click()
    AssignSelectedSheetToSlot
End Sub

Private Sub cmdClearSlot_Click()
    ClearSelectedSlot
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    RefreshSheetList
End Sub

Private Sub lstSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ActivateSelectedSheet
End Sub

Private Sub lstSheets_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ActivateSelectedSheet
        KeyCode = 0
    End If
End Sub

Private Sub txtSearch_Change()
    If mIsLoading Then
        Exit Sub
    End If

    RefreshSheetList
End Sub

Private Sub UserForm_Terminate()
    FAS_ClearNavigatorReference
End Sub

Private Sub ActivateSelectedSheet()
    Dim targetWorkbook As Workbook
    Dim selectedSheetName As String

    If Me.lstSheets.ListIndex < 0 Then
        Exit Sub
    End If

    selectedSheetName = Me.lstSheets.List(Me.lstSheets.ListIndex, 1)
    If FAS_IsNavigatorSeparator(Me.lstSheets.List(Me.lstSheets.ListIndex, 0), selectedSheetName) Then
        Exit Sub
    End If

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        MsgBox "No active user workbook is available.", vbExclamation, "FAS Toolkit"
        Exit Sub
    End If

    On Error GoTo ActivateFailed
    targetWorkbook.Worksheets(selectedSheetName).Activate
    Exit Sub

ActivateFailed:
    MsgBox "The selected sheet could not be activated.", vbExclamation, "FAS Toolkit"
End Sub

Private Sub ConfigureListBox()
    With Me.lstSheets
        .columnCount = 2
        .ColumnWidths = CStr(Application.Max(Me.InsideWidth - (FAS_MARGIN * 2) - 12, 120)) & " pt;0 pt"
        .BoundColumn = 1
        .IntegralHeight = False
    End With

    With Me.cboGroup
        .columnCount = 2
        .ColumnWidths = "78 pt;0 pt"
        .BoundColumn = 2
        .Style = fmStyleDropDownList
    End With

    Me.txtSearch.ControlTipText = "Search by sheet name or use jump commands like p1, p2, bs1, t2 or g1."
    Me.txtSlot.Text = "1"
End Sub

Private Sub PositionForm()
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = Application.Left + Application.Width - Me.Width - 24
    Me.Top = Application.Top + 90
    On Error GoTo 0
End Sub

Private Sub PopulateGroupPicker()
    Me.cboGroup.Clear
    AddGroupPickerItem "P&L", "PL"
    AddGroupPickerItem "BS", "BS"
    AddGroupPickerItem "TB", "TB"
    AddGroupPickerItem "GL", "GL"
    AddGroupPickerItem "Mapping", "MAPPING"
    Me.cboGroup.ListIndex = 0
End Sub

Private Sub AddGroupPickerItem(ByVal displayText As String, ByVal groupKey As String)
    Me.cboGroup.AddItem displayText
    Me.cboGroup.List(Me.cboGroup.ListCount - 1, 1) = groupKey
End Sub

Private Sub ApplyLayout()
    Dim contentWidth As Single
    Dim listHeight As Single
    Dim assignRowTop As Single
    Dim bottomRowTop As Single

    If Me.Width < FAS_FORM_MIN_WIDTH Then
        Me.Width = FAS_FORM_MIN_WIDTH
    End If

    If Me.Height < FAS_FORM_MIN_HEIGHT Then
        Me.Height = FAS_FORM_MIN_HEIGHT
    End If

    contentWidth = Me.InsideWidth - (FAS_MARGIN * 2)
    If contentWidth < 160 Then
        contentWidth = 160
    End If

    Me.txtSearch.Left = FAS_MARGIN
    Me.txtSearch.Top = FAS_MARGIN
    Me.txtSearch.Width = contentWidth

    assignRowTop = Me.InsideHeight - FAS_MARGIN - FAS_BUTTON_HEIGHT - FAS_ROW_GAP - FAS_BUTTON_HEIGHT
    bottomRowTop = Me.InsideHeight - FAS_MARGIN - FAS_BUTTON_HEIGHT
    listHeight = assignRowTop - Me.txtSearch.Top - Me.txtSearch.Height - FAS_ROW_GAP
    If listHeight < 120 Then
        listHeight = 120
    End If

    Me.lstSheets.Left = FAS_MARGIN
    Me.lstSheets.Top = Me.txtSearch.Top + Me.txtSearch.Height + FAS_ROW_GAP
    Me.lstSheets.Width = contentWidth
    Me.lstSheets.Height = listHeight
    Me.lstSheets.ColumnWidths = CStr(Application.Max(contentWidth - 12, 120)) & " pt;0 pt"

    Me.cboGroup.Left = FAS_MARGIN
    Me.cboGroup.Top = assignRowTop
    Me.cboGroup.Width = 84
    Me.txtSlot.Left = Me.cboGroup.Left + Me.cboGroup.Width + 6
    Me.txtSlot.Top = assignRowTop
    Me.txtSlot.Width = 30
    Me.cmdAssign.Top = assignRowTop
    Me.cmdAssign.Left = Me.txtSlot.Left + Me.txtSlot.Width + 6
    Me.cmdAssign.Width = Application.Max((contentWidth - Me.cmdAssign.Left + FAS_MARGIN - 6) / 2, 60)
    Me.cmdClearSlot.Top = assignRowTop
    Me.cmdClearSlot.Left = Me.cmdAssign.Left + Me.cmdAssign.Width + 6
    Me.cmdClearSlot.Width = contentWidth - (Me.cmdClearSlot.Left - FAS_MARGIN)

    Me.cmdRefresh.Left = FAS_MARGIN
    Me.cmdRefresh.Top = bottomRowTop
    Me.cmdRefresh.Width = (contentWidth - 6) / 2
    Me.cmdClose.Top = bottomRowTop
    Me.cmdClose.Left = Me.cmdRefresh.Left + Me.cmdRefresh.Width + 6
    Me.cmdClose.Width = Me.cmdRefresh.Width
End Sub

Private Sub RefreshSheetList()
    Dim items As Collection
    Dim item As Variant
    Dim targetWorkbook As Workbook

    Set targetWorkbook = FAS_GetTargetWorkbook()

    Me.lstSheets.Clear
    If targetWorkbook Is Nothing Then
        Me.lstSheets.AddItem "(No active workbook)"
        Me.lstSheets.List(Me.lstSheets.ListCount - 1, 1) = vbNullString
        Exit Sub
    End If

    Set items = FAS_GetNavigatorItems(targetWorkbook, Me.txtSearch.Text)
    For Each item In items
        Me.lstSheets.AddItem CStr(item(0))
        Me.lstSheets.List(Me.lstSheets.ListCount - 1, 1) = CStr(item(1))
    Next item
End Sub

Private Sub AssignSelectedSheetToSlot()
    Dim targetWorkbook As Workbook
    Dim selectedSheetName As String
    Dim groupKey As String
    Dim slotIndex As Long

    selectedSheetName = GetSelectedSheetName()
    If Len(selectedSheetName) = 0 Then
        MsgBox "Select a sheet before assigning a slot.", vbExclamation, "FAS Toolkit"
        Exit Sub
    End If

    groupKey = GetSelectedGroupKey()
    slotIndex = GetSelectedSlotIndex()
    If Len(groupKey) = 0 Or slotIndex = 0 Then
        Exit Sub
    End If

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        MsgBox "No active user workbook is available.", vbExclamation, "FAS Toolkit"
        Exit Sub
    End If

    On Error GoTo AssignFailed
    If Not FAS_SaveSheetAssignment(targetWorkbook, groupKey, slotIndex, selectedSheetName) Then
        If Not FAS_CanPersistAssignments(targetWorkbook) Then
            MsgBox "Save the workbook once before creating persistent local sheet mappings.", vbExclamation, "FAS Toolkit"
            Exit Sub
        End If
        GoTo AssignFailed
    End If

    RefreshSheetList
    MsgBox FAS_GetGroupLabel(groupKey) & " slot " & slotIndex & " now points to '" & selectedSheetName & "'." & vbCrLf & _
        "This mapping is stored locally on this PC, not inside the workbook.", vbInformation, "FAS Toolkit"
    Exit Sub

AssignFailed:
    MsgBox "The selected slot could not be updated in this workbook.", vbExclamation, "FAS Toolkit"
End Sub

Private Sub ClearSelectedSlot()
    Dim targetWorkbook As Workbook
    Dim groupKey As String
    Dim slotIndex As Long

    groupKey = GetSelectedGroupKey()
    slotIndex = GetSelectedSlotIndex()
    If Len(groupKey) = 0 Or slotIndex = 0 Then
        Exit Sub
    End If

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        MsgBox "No active user workbook is available.", vbExclamation, "FAS Toolkit"
        Exit Sub
    End If

    If Not FAS_CanPersistAssignments(targetWorkbook) Then
        MsgBox "Save the workbook once before clearing persistent local sheet mappings.", vbExclamation, "FAS Toolkit"
        Exit Sub
    End If

    On Error GoTo ClearFailed
    Call FAS_ClearSheetAssignment(targetWorkbook, groupKey, slotIndex)
    RefreshSheetList
    MsgBox FAS_GetGroupLabel(groupKey) & " slot " & slotIndex & " was cleared." & vbCrLf & _
        "The local mapping for this workbook was removed from this PC.", vbInformation, "FAS Toolkit"
    Exit Sub

ClearFailed:
    MsgBox "The selected slot could not be cleared in this workbook.", vbExclamation, "FAS Toolkit"
End Sub

Private Function GetSelectedSheetName() As String
    If Me.lstSheets.ListIndex < 0 Then
        Exit Function
    End If

    GetSelectedSheetName = Me.lstSheets.List(Me.lstSheets.ListIndex, 1)
    If FAS_IsNavigatorSeparator(Me.lstSheets.List(Me.lstSheets.ListIndex, 0), GetSelectedSheetName) Then
        GetSelectedSheetName = vbNullString
    End If
End Function

Private Function GetSelectedGroupKey() As String
    If Me.cboGroup.ListIndex < 0 Then
        MsgBox "Select a slot group first.", vbExclamation, "FAS Toolkit"
        Exit Function
    End If

    GetSelectedGroupKey = Me.cboGroup.List(Me.cboGroup.ListIndex, 1)
End Function

Private Function GetSelectedSlotIndex() As Long
    Dim slotText As String

    slotText = Trim$(Me.txtSlot.Text)
    If Len(slotText) = 0 Or Not IsNumeric(slotText) Then
        MsgBox "Enter a slot number from 1 to " & FAS_MAX_SLOT_INDEX & ".", vbExclamation, "FAS Toolkit"
        Exit Function
    End If

    GetSelectedSlotIndex = CLng(slotText)
    If GetSelectedSlotIndex < 1 Or GetSelectedSlotIndex > FAS_MAX_SLOT_INDEX Then
        MsgBox "Enter a slot number from 1 to " & FAS_MAX_SLOT_INDEX & ".", vbExclamation, "FAS Toolkit"
        GetSelectedSlotIndex = 0
    End If
End Function

