VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FASEditModeForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "FASEditModeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FASEditModeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoading As Boolean

Private Sub UserForm_Initialize()
    mLoading = True
    FASEditModePopulateGridCombo Me.Controls("cboColumns")
    FASEditModePopulateGridCombo Me.Controls("cboRows")
    FASEditModeRefreshExactCombos
    mLoading = False
    RefreshView
End Sub

Private Sub UserForm_Activate()
    RefreshView
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    FASEditModeExit
End Sub

Public Sub RefreshView()
    Dim gridColumns As Long
    Dim gridRows As Long
    Dim exactColumn As Long
    Dim exactRow As Long

    mLoading = True
    gridColumns = FASEditModeGetGridColumns()
    gridRows = FASEditModeGetGridRows()
    exactColumn = FASEditModeGetExactColumn()
    exactRow = FASEditModeGetExactRow()

    Me.Caption = "FAS Edit Mode"
    Me.Controls("lblMode").Caption = IIf(FASEditModeIsActive(), "EDIT MODE ACTIVE", "Edit Mode inactive")
    Me.Controls("lblZone").Caption = "Zone: " & FASGetActiveZoneName()
    Me.Controls("lblSelection").Caption = FASEditModeGetSelectionSummary()
    Me.Controls("lblStatus").Caption = FASEditModeGetStatusText()
    Me.Controls("cboColumns").Value = CStr(gridColumns)
    Me.Controls("cboRows").Value = CStr(gridRows)

    FASEditModeRefreshExactCombos
    Me.Controls("cboExactColumn").ListIndex = exactColumn
    Me.Controls("cboExactRow").ListIndex = exactRow
    mLoading = False
End Sub

Private Sub cboColumns_Change()
    If mLoading Then Exit Sub
    FASEditModeRefreshExactCombos
    FASEditModeSaveGrid
End Sub

Private Sub cboRows_Change()
    If mLoading Then Exit Sub
    FASEditModeRefreshExactCombos
    FASEditModeSaveGrid
End Sub

Private Sub cboExactColumn_Change()
    If mLoading Then Exit Sub
    FASEditModeSaveGrid
End Sub

Private Sub cboExactRow_Change()
    If mLoading Then Exit Sub
    FASEditModeSaveGrid
End Sub

Private Sub cmdClearExact_Click()
    FASEditModeClearExactCell
    RefreshView
End Sub

Private Sub cmdZoneContent_Click()
    FASSetActiveZoneContent
    FASEditModeSetStatus "Active zone: Content"
End Sub

Private Sub cmdZoneTitle_Click()
    FASSetActiveZoneHeading
    FASEditModeSetStatus "Active zone: Title"
End Sub

Private Sub cmdZoneSubtitle_Click()
    FASSetActiveZoneSubheading
    FASEditModeSetStatus "Active zone: Subtitle"
End Sub

Private Sub cmdExit_Click()
    FASEditModeExit
End Sub

Private Sub FASEditModePopulateGridCombo(ByVal combo As Object)
    combo.Clear
    combo.AddItem "2"
    combo.AddItem "3"
    combo.AddItem "4"
End Sub

Private Sub FASEditModeRefreshExactCombos()
    Dim gridColumns As Long
    Dim gridRows As Long
    Dim oldColumn As Long
    Dim oldRow As Long
    Dim i As Long

    If IsNumeric(Me.Controls("cboColumns").Value) Then
        gridColumns = CLng(Me.Controls("cboColumns").Value)
    Else
        gridColumns = FASEditModeGetGridColumns()
    End If
    If IsNumeric(Me.Controls("cboRows").Value) Then
        gridRows = CLng(Me.Controls("cboRows").Value)
    Else
        gridRows = FASEditModeGetGridRows()
    End If

    oldColumn = Me.Controls("cboExactColumn").ListIndex
    oldRow = Me.Controls("cboExactRow").ListIndex
    Me.Controls("cboExactColumn").Clear
    Me.Controls("cboExactRow").Clear
    Me.Controls("cboExactColumn").AddItem "Keypad anchor"
    Me.Controls("cboExactRow").AddItem "Keypad anchor"

    For i = 1 To gridColumns
        Me.Controls("cboExactColumn").AddItem CStr(i)
    Next i
    For i = 1 To gridRows
        Me.Controls("cboExactRow").AddItem CStr(i)
    Next i

    If oldColumn < 0 Or oldColumn > gridColumns Then oldColumn = 0
    If oldRow < 0 Or oldRow > gridRows Then oldRow = 0
    Me.Controls("cboExactColumn").ListIndex = oldColumn
    Me.Controls("cboExactRow").ListIndex = oldRow
End Sub

Private Sub FASEditModeSaveGrid()
    Dim gridColumns As Long
    Dim gridRows As Long
    Dim exactColumn As Long
    Dim exactRow As Long

    If Not IsNumeric(Me.Controls("cboColumns").Value) Then Exit Sub
    If Not IsNumeric(Me.Controls("cboRows").Value) Then Exit Sub

    gridColumns = CLng(Me.Controls("cboColumns").Value)
    gridRows = CLng(Me.Controls("cboRows").Value)
    exactColumn = Me.Controls("cboExactColumn").ListIndex
    exactRow = Me.Controls("cboExactRow").ListIndex
    FASEditModeSetGridPreferences gridColumns, gridRows, exactColumn, exactRow
End Sub

