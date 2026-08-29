Attribute VB_Name = "ModuleFASEditMode"
'MIT License
'
'Copyright (c) 2021 - 2026 iappyx
'
'Keyboard-first FAS placement, movement, sizing, and ordering for Windows.

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private Const EDIT_APP As String = "InstrumentaFASEditMode"
Private Const EDIT_SECTION_SESSION As String = "Session"
Private Const EDIT_SECTION_PREFERENCES As String = "Preferences"

Private Const EDIT_KEY_ACTIVE As String = "Active"
Private Const EDIT_KEY_POWERPOINT_PID As String = "PowerPointPid"
Private Const EDIT_KEY_SESSION_ID As String = "SessionId"
Private Const EDIT_KEY_GRID_COLUMNS As String = "GridColumns"
Private Const EDIT_KEY_GRID_ROWS As String = "GridRows"
Private Const EDIT_KEY_EXACT_COLUMN As String = "ExactColumn"
Private Const EDIT_KEY_EXACT_ROW As String = "ExactRow"
Private Const EDIT_KEY_FINE_STEP As String = "FineStepPt"
Private Const EDIT_KEY_COARSE_STEP As String = "CoarseStepPt"

Private Const EDIT_DEFAULT_GRID_COLUMNS As Long = 3
Private Const EDIT_DEFAULT_GRID_ROWS As Long = 3
Private Const EDIT_DEFAULT_FINE_STEP As Single = 1!
Private Const EDIT_DEFAULT_COARSE_STEP As Single = 10!
Private Const EDIT_MIN_SIZE As Single = 1!

Private mEvents As FASEditModeEventsClass
Private mLastStatus As String

Public Sub FASEditModeInitialize()
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_ACTIVE, "0"
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_POWERPOINT_PID, "0"
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_SESSION_ID, vbNullString
    mLastStatus = "Ready"
End Sub

Public Sub FASEditModeToggle()
    If FASEditModeIsActive() Then
        FASEditModeExit
    Else
        FASEditModeEnter
    End If
End Sub

Public Sub FASEditModeRibbonToggle(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        FASEditModeEnter
    Else
        FASEditModeExit
    End If
End Sub

Public Function FASEditModeGetPressed(ByVal control As IRibbonControl) As Boolean
    FASEditModeGetPressed = FASEditModeIsActive()
End Function

Public Function FASEditModeIsActive() As Boolean
    FASEditModeIsActive = (GetSetting(EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_ACTIVE, "0") = "1")
End Function

Public Sub FASEditModeEnter()
    FASEditModeActivate True
End Sub

Public Sub FASEditModeEnterForAutomation()
    FASEditModeActivate False
End Sub

Private Sub FASEditModeActivate(ByVal showPalette As Boolean)
    Dim processId As Long

    processId = FASEditModeGetPowerPointProcessId()
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_ACTIVE, "1"
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_POWERPOINT_PID, CStr(processId)
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_SESSION_ID, Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Timer, "0")

    FASEditModeEnsureEvents
    mLastStatus = "Edit Mode active"
    If showPalette Then FASEditModeShowPalette
    FASEditModeInvalidateRibbon
End Sub

Public Sub FASEditModeExit()
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_ACTIVE, "0"
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_POWERPOINT_PID, "0"
    SaveSetting EDIT_APP, EDIT_SECTION_SESSION, EDIT_KEY_SESSION_ID, vbNullString
    mLastStatus = "Edit Mode inactive"

    On Error Resume Next
    FASEditModeForm.Hide
    On Error GoTo 0

    Set mEvents = Nothing
    FASEditModeInvalidateRibbon
End Sub

Public Sub FASEditModeShowPalette()
    If Not FASEditModeIsActive() Then FASEditModeEnter

    On Error Resume Next
    FASEditModeForm.RefreshView
    FASEditModeForm.Show 0
    On Error GoTo 0
End Sub

Public Sub FASEditModeRefreshPalette()
    If Not FASEditModeIsActive() Then Exit Sub

    On Error Resume Next
    FASEditModeForm.RefreshView
    On Error GoTo 0
End Sub

Public Sub FASEditModeSetStatus(ByVal statusText As String)
    mLastStatus = statusText
    FASEditModeRefreshPalette
End Sub

Public Function FASEditModeGetStatusText() As String
    If Len(mLastStatus) = 0 Then mLastStatus = "Ready"
    FASEditModeGetStatusText = mLastStatus
End Function

Public Function FASEditModeGetSelectionSummary() As String
    Dim selectedShapes As ShapeRange

    Set selectedShapes = FASEditModeGetSelectedShapes()
    If selectedShapes Is Nothing Then
        FASEditModeGetSelectionSummary = "No shape selected"
    ElseIf selectedShapes.count = 1 Then
        FASEditModeGetSelectionSummary = "1 shape selected"
    Else
        FASEditModeGetSelectionSummary = CStr(selectedShapes.count) & " shapes selected"
    End If
End Function

Public Function FASEditModeGetGridColumns() As Long
    FASEditModeGetGridColumns = FASEditModeGetLongPreference(EDIT_KEY_GRID_COLUMNS, EDIT_DEFAULT_GRID_COLUMNS, 2, 4)
End Function

Public Function FASEditModeGetGridRows() As Long
    FASEditModeGetGridRows = FASEditModeGetLongPreference(EDIT_KEY_GRID_ROWS, EDIT_DEFAULT_GRID_ROWS, 2, 4)
End Function

Public Function FASEditModeGetExactColumn() As Long
    FASEditModeGetExactColumn = FASEditModeGetLongPreference(EDIT_KEY_EXACT_COLUMN, 0, 0, FASEditModeGetGridColumns())
End Function

Public Function FASEditModeGetExactRow() As Long
    FASEditModeGetExactRow = FASEditModeGetLongPreference(EDIT_KEY_EXACT_ROW, 0, 0, FASEditModeGetGridRows())
End Function

Public Sub FASEditModeSetGridPreferences(ByVal gridColumns As Long, ByVal gridRows As Long, ByVal exactColumn As Long, ByVal exactRow As Long)
    gridColumns = FASEditModeClampLong(gridColumns, 2, 4)
    gridRows = FASEditModeClampLong(gridRows, 2, 4)
    exactColumn = FASEditModeClampLong(exactColumn, 0, gridColumns)
    exactRow = FASEditModeClampLong(exactRow, 0, gridRows)

    SaveSetting EDIT_APP, EDIT_SECTION_PREFERENCES, EDIT_KEY_GRID_COLUMNS, CStr(gridColumns)
    SaveSetting EDIT_APP, EDIT_SECTION_PREFERENCES, EDIT_KEY_GRID_ROWS, CStr(gridRows)
    SaveSetting EDIT_APP, EDIT_SECTION_PREFERENCES, EDIT_KEY_EXACT_COLUMN, CStr(exactColumn)
    SaveSetting EDIT_APP, EDIT_SECTION_PREFERENCES, EDIT_KEY_EXACT_ROW, CStr(exactRow)
    FASEditModeSetStatus "Grid " & gridColumns & " x " & gridRows & FASEditModeExactCellSuffix(exactColumn, exactRow)
End Sub

Public Sub FASEditModeClearExactCell()
    SaveSetting EDIT_APP, EDIT_SECTION_PREFERENCES, EDIT_KEY_EXACT_COLUMN, "0"
    SaveSetting EDIT_APP, EDIT_SECTION_PREFERENCES, EDIT_KEY_EXACT_ROW, "0"
    FASEditModeSetStatus "Using keypad anchors"
End Sub

Public Sub FASEditModeNudgeLeftFine()
    FASEditModeNudge -FASEditModeGetStep(False), 0
End Sub

Public Sub FASEditModeNudgeRightFine()
    FASEditModeNudge FASEditModeGetStep(False), 0
End Sub

Public Sub FASEditModeNudgeUpFine()
    FASEditModeNudge 0, -FASEditModeGetStep(False)
End Sub

Public Sub FASEditModeNudgeDownFine()
    FASEditModeNudge 0, FASEditModeGetStep(False)
End Sub

Public Sub FASEditModeNudgeLeftCoarse()
    FASEditModeNudge -FASEditModeGetStep(True), 0
End Sub

Public Sub FASEditModeNudgeRightCoarse()
    FASEditModeNudge FASEditModeGetStep(True), 0
End Sub

Public Sub FASEditModeNudgeUpCoarse()
    FASEditModeNudge 0, -FASEditModeGetStep(True)
End Sub

Public Sub FASEditModeNudgeDownCoarse()
    FASEditModeNudge 0, FASEditModeGetStep(True)
End Sub

Public Sub FASEditModeResizeWidthIncreaseFine()
    FASEditModeResize True, True, FASEditModeGetStep(False)
End Sub

Public Sub FASEditModeResizeWidthDecreaseFine()
    FASEditModeResize True, False, FASEditModeGetStep(False)
End Sub

Public Sub FASEditModeResizeHeightIncreaseFine()
    FASEditModeResize False, True, FASEditModeGetStep(False)
End Sub

Public Sub FASEditModeResizeHeightDecreaseFine()
    FASEditModeResize False, False, FASEditModeGetStep(False)
End Sub

Public Sub FASEditModeResizeWidthIncreaseCoarse()
    FASEditModeResize True, True, FASEditModeGetStep(True)
End Sub

Public Sub FASEditModeResizeWidthDecreaseCoarse()
    FASEditModeResize True, False, FASEditModeGetStep(True)
End Sub

Public Sub FASEditModeResizeHeightIncreaseCoarse()
    FASEditModeResize False, True, FASEditModeGetStep(True)
End Sub

Public Sub FASEditModeResizeHeightDecreaseCoarse()
    FASEditModeResize False, False, FASEditModeGetStep(True)
End Sub

Public Sub FASEditModePlaceTopLeft()
    FASEditModePlace "left", "top", False
End Sub

Public Sub FASEditModePlaceTop()
    FASEditModePlace "center", "top", False
End Sub

Public Sub FASEditModePlaceTopRight()
    FASEditModePlace "right", "top", False
End Sub

Public Sub FASEditModePlaceLeft()
    FASEditModePlace "left", "middle", False
End Sub

Public Sub FASEditModePlaceCenter()
    FASEditModePlace "center", "middle", False
End Sub

Public Sub FASEditModePlaceRight()
    FASEditModePlace "right", "middle", False
End Sub

Public Sub FASEditModePlaceBottomLeft()
    FASEditModePlace "left", "bottom", False
End Sub

Public Sub FASEditModePlaceBottom()
    FASEditModePlace "center", "bottom", False
End Sub

Public Sub FASEditModePlaceBottomRight()
    FASEditModePlace "right", "bottom", False
End Sub

Public Sub FASEditModeFitTopLeft()
    FASEditModePlace "left", "top", True
End Sub

Public Sub FASEditModeFitTop()
    FASEditModePlace "center", "top", True
End Sub

Public Sub FASEditModeFitTopRight()
    FASEditModePlace "right", "top", True
End Sub

Public Sub FASEditModeFitLeft()
    FASEditModePlace "left", "middle", True
End Sub

Public Sub FASEditModeFitCenter()
    FASEditModePlace "center", "middle", True
End Sub

Public Sub FASEditModeFitRight()
    FASEditModePlace "right", "middle", True
End Sub

Public Sub FASEditModeFitBottomLeft()
    FASEditModePlace "left", "bottom", True
End Sub

Public Sub FASEditModeFitBottom()
    FASEditModePlace "center", "bottom", True
End Sub

Public Sub FASEditModeFitBottomRight()
    FASEditModePlace "right", "bottom", True
End Sub

Public Sub FASEditModeBringForward()
    BringForward
    FASEditModeSetStatus "Brought forward"
End Sub

Public Sub FASEditModeSendBackward()
    SendBackward
    FASEditModeSetStatus "Sent backward"
End Sub

Public Sub FASEditModeBringToFront()
    BringToFront
    FASEditModeSetStatus "Brought to front"
End Sub

Public Sub FASEditModeSendToBack()
    SendToBack
    FASEditModeSetStatus "Sent to back"
End Sub

Private Sub FASEditModePlace(ByVal horizontalAnchor As String, ByVal verticalAnchor As String, ByVal fitShape As Boolean)
    Dim gridColumns As Long
    Dim gridRows As Long
    Dim exactColumn As Long
    Dim exactRow As Long

    If fitShape Then
        gridColumns = FASEditModeGetGridColumns()
        gridRows = FASEditModeGetGridRows()
        exactColumn = FASEditModeGetExactColumn()
        exactRow = FASEditModeGetExactRow()

        If exactColumn = 0 Then
            gridColumns = 3
            exactColumn = FASEditModeAnchorIndex(horizontalAnchor)
        End If
        If exactRow = 0 Then
            gridRows = 3
            exactRow = FASEditModeAnchorIndex(verticalAnchor)
        End If
    End If

    FASApplyEditCell horizontalAnchor, verticalAnchor, fitShape, gridColumns, gridRows, exactColumn, exactRow
End Sub

Private Sub FASEditModeNudge(ByVal requestedX As Single, ByVal requestedY As Single)
    Dim selectedShapes As ShapeRange
    Dim zoneLeft As Single
    Dim zoneTop As Single
    Dim zoneWidth As Single
    Dim zoneHeight As Single
    Dim selectionLeft As Single
    Dim selectionTop As Single
    Dim selectionRight As Single
    Dim selectionBottom As Single
    Dim deltaX As Single
    Dim deltaY As Single

    Set selectedShapes = FASEditModeGetSelectedShapes()
    If selectedShapes Is Nothing Then
        FASEditModeSetStatus "Select one or more shapes"
        Exit Sub
    End If
    If Not FASGetActiveZoneBounds(zoneLeft, zoneTop, zoneWidth, zoneHeight) Then Exit Sub

    FASEditModeGetRangeBounds selectedShapes, selectionLeft, selectionTop, selectionRight, selectionBottom
    If selectionRight - selectionLeft > zoneWidth Or selectionBottom - selectionTop > zoneHeight Then
        FASEditModeSetStatus "Selection is larger than the active zone"
        Exit Sub
    End If
    deltaX = requestedX
    deltaY = requestedY

    If selectionLeft + deltaX < zoneLeft Then deltaX = zoneLeft - selectionLeft
    If selectionRight + deltaX > zoneLeft + zoneWidth Then deltaX = zoneLeft + zoneWidth - selectionRight
    If selectionTop + deltaY < zoneTop Then deltaY = zoneTop - selectionTop
    If selectionBottom + deltaY > zoneTop + zoneHeight Then deltaY = zoneTop + zoneHeight - selectionBottom

    FASEditModeTranslateRange selectedShapes, deltaX, deltaY
    FASEditModeSetStatus "Moved " & FASEditModeFormatStep(deltaX, deltaY)
End Sub

Private Sub FASEditModeResize(ByVal resizeWidth As Boolean, ByVal increaseSize As Boolean, ByVal stepSize As Single)
    Dim selectedShapes As ShapeRange
    Dim slideShape As shape
    Dim zoneLeft As Single
    Dim zoneTop As Single
    Dim zoneWidth As Single
    Dim zoneHeight As Single
    Dim originalLeft As Single
    Dim originalTop As Single
    Dim targetSize As Single

    Set selectedShapes = FASEditModeGetSelectedShapes()
    If selectedShapes Is Nothing Then
        FASEditModeSetStatus "Select one shape to resize"
        Exit Sub
    End If
    If selectedShapes.count <> 1 Then
        FASEditModeSetStatus "Resize requires exactly one shape"
        Exit Sub
    End If
    If Not FASGetActiveZoneBounds(zoneLeft, zoneTop, zoneWidth, zoneHeight) Then Exit Sub

    Set slideShape = selectedShapes(1)
    originalLeft = GetRealLeft(slideShape)
    originalTop = GetRealTop(slideShape)
    If increaseSize Then
        targetSize = stepSize
    Else
        targetSize = -stepSize
    End If

    If resizeWidth Then
        targetSize = FASEditModeClampSingle(GetRealWidth(slideShape) + targetSize, EDIT_MIN_SIZE, zoneLeft + zoneWidth - originalLeft)
        SetRealWidth slideShape, targetSize
    Else
        targetSize = FASEditModeClampSingle(GetRealHeight(slideShape) + targetSize, EDIT_MIN_SIZE, zoneTop + zoneHeight - originalTop)
        SetRealHeight slideShape, targetSize
    End If

    FASEditModeClampShapeToZone slideShape, zoneLeft, zoneTop, zoneWidth, zoneHeight
    FASEditModeSetStatus "Resized to " & Format$(GetRealWidth(slideShape), "0.0") & " x " & Format$(GetRealHeight(slideShape), "0.0") & " pt"
End Sub

Private Sub FASEditModeClampShapeToZone(ByVal slideShape As shape, ByVal zoneLeft As Single, ByVal zoneTop As Single, ByVal zoneWidth As Single, ByVal zoneHeight As Single)
    Dim realLeft As Single
    Dim realTop As Single
    Dim realWidth As Single
    Dim realHeight As Single

    realWidth = GetRealWidth(slideShape)
    realHeight = GetRealHeight(slideShape)
    If realWidth > zoneWidth Then SetRealWidth slideShape, zoneWidth
    If realHeight > zoneHeight Then SetRealHeight slideShape, zoneHeight

    realLeft = FASEditModeClampSingle(GetRealLeft(slideShape), zoneLeft, zoneLeft + zoneWidth - GetRealWidth(slideShape))
    realTop = FASEditModeClampSingle(GetRealTop(slideShape), zoneTop, zoneTop + zoneHeight - GetRealHeight(slideShape))
    SetRealLeft slideShape, realLeft
    SetRealTop slideShape, realTop
End Sub

Private Function FASEditModeGetSelectedShapes() As ShapeRange
    On Error GoTo NoSelection

    If Application.ActiveWindow Is Nothing Then Exit Function

    Select Case Application.ActiveWindow.Selection.Type
        Case ppSelectionShapes
            If Application.ActiveWindow.Selection.HasChildShapeRange Then
                Set FASEditModeGetSelectedShapes = Application.ActiveWindow.Selection.ChildShapeRange
            Else
                Set FASEditModeGetSelectedShapes = Application.ActiveWindow.Selection.ShapeRange
            End If
        Case ppSelectionText
            Set FASEditModeGetSelectedShapes = Application.ActiveWindow.Selection.ShapeRange
    End Select
    Exit Function

NoSelection:
    Set FASEditModeGetSelectedShapes = Nothing
End Function

Private Sub FASEditModeGetRangeBounds(ByVal selectedShapes As ShapeRange, ByRef rangeLeft As Single, ByRef rangeTop As Single, ByRef rangeRight As Single, ByRef rangeBottom As Single)
    Dim slideShape As shape
    Dim shapeLeft As Single
    Dim shapeTop As Single
    Dim shapeRight As Single
    Dim shapeBottom As Single
    Dim isFirst As Boolean

    isFirst = True
    For Each slideShape In selectedShapes
        shapeLeft = GetRealLeft(slideShape)
        shapeTop = GetRealTop(slideShape)
        shapeRight = shapeLeft + GetRealWidth(slideShape)
        shapeBottom = shapeTop + GetRealHeight(slideShape)

        If isFirst Then
            rangeLeft = shapeLeft
            rangeTop = shapeTop
            rangeRight = shapeRight
            rangeBottom = shapeBottom
            isFirst = False
        Else
            If shapeLeft < rangeLeft Then rangeLeft = shapeLeft
            If shapeTop < rangeTop Then rangeTop = shapeTop
            If shapeRight > rangeRight Then rangeRight = shapeRight
            If shapeBottom > rangeBottom Then rangeBottom = shapeBottom
        End If
    Next slideShape
End Sub

Private Sub FASEditModeTranslateRange(ByVal selectedShapes As ShapeRange, ByVal deltaX As Single, ByVal deltaY As Single)
    Dim slideShape As shape

    For Each slideShape In selectedShapes
        If deltaX <> 0 Then SetRealLeft slideShape, GetRealLeft(slideShape) + deltaX
        If deltaY <> 0 Then SetRealTop slideShape, GetRealTop(slideShape) + deltaY
    Next slideShape
End Sub

Private Sub FASEditModeEnsureEvents()
    If mEvents Is Nothing Then
        Set mEvents = New FASEditModeEventsClass
        Set mEvents.app = Application
    End If
End Sub

Private Sub FASEditModeInvalidateRibbon()
    On Error Resume Next
    If Not InstrumentaRibbon Is Nothing Then InstrumentaRibbon.Invalidate
    On Error GoTo 0
End Sub

Private Function FASEditModeGetPowerPointProcessId() As Long
    On Error Resume Next
    FASEditModeGetPowerPointProcessId = GetCurrentProcessId()
    On Error GoTo 0
End Function

Private Function FASEditModeGetStep(ByVal coarseStep As Boolean) As Single
    Dim keyName As String
    Dim defaultValue As Single
    Dim rawValue As String

    If coarseStep Then
        keyName = EDIT_KEY_COARSE_STEP
        defaultValue = EDIT_DEFAULT_COARSE_STEP
    Else
        keyName = EDIT_KEY_FINE_STEP
        defaultValue = EDIT_DEFAULT_FINE_STEP
    End If

    rawValue = GetSetting(EDIT_APP, EDIT_SECTION_PREFERENCES, keyName, CStr(defaultValue))
    If IsNumeric(rawValue) Then
        FASEditModeGetStep = FASEditModeClampSingle(CSng(rawValue), 0.1, 1000)
    Else
        FASEditModeGetStep = defaultValue
    End If
End Function

Private Function FASEditModeGetLongPreference(ByVal keyName As String, ByVal defaultValue As Long, ByVal minimumValue As Long, ByVal maximumValue As Long) As Long
    Dim rawValue As String

    rawValue = GetSetting(EDIT_APP, EDIT_SECTION_PREFERENCES, keyName, CStr(defaultValue))
    If IsNumeric(rawValue) Then
        FASEditModeGetLongPreference = FASEditModeClampLong(CLng(rawValue), minimumValue, maximumValue)
    Else
        FASEditModeGetLongPreference = defaultValue
    End If
End Function

Private Function FASEditModeAnchorIndex(ByVal anchorName As String) As Long
    Select Case LCase$(anchorName)
        Case "left", "top"
            FASEditModeAnchorIndex = 1
        Case "center", "middle"
            FASEditModeAnchorIndex = 2
        Case Else
            FASEditModeAnchorIndex = 3
    End Select
End Function

Private Function FASEditModeClampLong(ByVal value As Long, ByVal minimumValue As Long, ByVal maximumValue As Long) As Long
    If value < minimumValue Then value = minimumValue
    If value > maximumValue Then value = maximumValue
    FASEditModeClampLong = value
End Function

Private Function FASEditModeClampSingle(ByVal value As Single, ByVal minimumValue As Single, ByVal maximumValue As Single) As Single
    If maximumValue < minimumValue Then maximumValue = minimumValue
    If value < minimumValue Then value = minimumValue
    If value > maximumValue Then value = maximumValue
    FASEditModeClampSingle = value
End Function

Private Function FASEditModeExactCellSuffix(ByVal exactColumn As Long, ByVal exactRow As Long) As String
    If exactColumn > 0 Or exactRow > 0 Then
        FASEditModeExactCellSuffix = " - exact cell " & IIf(exactColumn > 0, CStr(exactColumn), "anchor") & ", " & IIf(exactRow > 0, CStr(exactRow), "anchor")
    End If
End Function

Private Function FASEditModeFormatStep(ByVal deltaX As Single, ByVal deltaY As Single) As String
    If deltaX <> 0 Then
        FASEditModeFormatStep = Format$(Abs(deltaX), "0.0") & " pt " & IIf(deltaX < 0, "left", "right")
    Else
        FASEditModeFormatStep = Format$(Abs(deltaY), "0.0") & " pt " & IIf(deltaY < 0, "up", "down")
    End If
End Function
