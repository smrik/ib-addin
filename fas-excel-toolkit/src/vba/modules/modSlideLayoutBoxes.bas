Attribute VB_Name = "modSlideLayoutBoxes"
Option Explicit

Private Const FAS_SLIDE_LEFT_CM As Double = 0.83
Private Const FAS_SLIDE_WIDTH_CM As Double = 32.18
Private Const FAS_TITLE_HEIGHT_CM As Double = 1#
Private Const FAS_SUBTITLE_HEIGHT_CM As Double = 1#
Private Const FAS_CONTENT_HEIGHT_CM As Double = 12.2
Private Const FAS_SLIDE_GUTTER_CM As Double = 0.83
Private Const FAS_BOX_LINE_COLOR As Long = 12611584
Private Const FAS_SLIDE_LAYOUT_VERSION As String = "FAS slide layout boxes 2026-07-09"

Public Sub FAS_InsertSlideContentBox()
    FAS_InsertSlideLayoutBox "Content full", FAS_SLIDE_WIDTH_CM, FAS_CONTENT_HEIGHT_CM
End Sub

Public Sub FAS_InsertSlideTitleBox()
    FAS_InsertSlideLayoutBox "Title", FAS_SLIDE_WIDTH_CM, FAS_TITLE_HEIGHT_CM
End Sub

Public Sub FAS_InsertSlideSubtitleBox()
    FAS_InsertSlideLayoutBox "Subtitle", FAS_SLIDE_WIDTH_CM, FAS_SUBTITLE_HEIGHT_CM
End Sub

Public Sub FAS_InsertSlideHalfBox()
    FAS_InsertSlideLayoutBox "Content 1/2", FAS_GetColumnWidthCm(2), FAS_CONTENT_HEIGHT_CM
End Sub

Public Sub FAS_InsertSlideThirdBox()
    FAS_InsertSlideLayoutBox "Content 1/3", FAS_GetColumnWidthCm(3), FAS_CONTENT_HEIGHT_CM
End Sub

Public Sub FAS_InsertSlideTwoThirdsBox()
    FAS_InsertSlideLayoutBox "Content 2/3", FAS_GetColumnSpanCm(3, 2), FAS_CONTENT_HEIGHT_CM
End Sub

Public Sub FAS_InsertSlideQuarterBox()
    FAS_InsertSlideLayoutBox "Content 1/4", FAS_GetColumnWidthCm(4), FAS_CONTENT_HEIGHT_CM
End Sub

Public Sub FAS_InsertSlideHalfRowBox()
    FAS_InsertSlideLayoutBox "Content row 1/2", FAS_SLIDE_WIDTH_CM, FAS_GetRowHeightCm(2)
End Sub

Public Sub FAS_InsertSlideSmallBlockBox()
    FAS_InsertSlideLayoutBox "Content 1/3 x 1/2", FAS_GetColumnWidthCm(3), FAS_GetRowHeightCm(2)
End Sub

Public Sub FAS_ShowSlideBoxMenu()
    Dim prompt As String
    Dim userInput As String

    prompt = "Insert an outline box:" & vbCrLf & vbCrLf & _
        "1 = content full" & vbCrLf & _
        "2 = content 1/2" & vbCrLf & _
        "3 = content 1/3" & vbCrLf & _
        "4 = content 2/3" & vbCrLf & _
        "5 = content 1/4" & vbCrLf & _
        "6 = content half-row" & vbCrLf & _
        "7 = small block (1/3 x 1/2)" & vbCrLf & _
        "8 = title" & vbCrLf & _
        "9 = subtitle"

    userInput = Trim$(InputBox(prompt, "FAS Slide Boxes", "3"))
    If Len(userInput) = 0 Then Exit Sub

    Select Case LCase$(userInput)
        Case "1", "full", "content", "content full"
            FAS_InsertSlideContentBox
        Case "2", "half", "1/2", "content half"
            FAS_InsertSlideHalfBox
        Case "3", "third", "1/3", "content third"
            FAS_InsertSlideThirdBox
        Case "4", "twothirds", "two thirds", "2/3", "content 2/3"
            FAS_InsertSlideTwoThirdsBox
        Case "5", "quarter", "1/4", "content quarter"
            FAS_InsertSlideQuarterBox
        Case "6", "row", "halfrow", "half row"
            FAS_InsertSlideHalfRowBox
        Case "7", "small", "block", "smallblock"
            FAS_InsertSlideSmallBlockBox
        Case "8", "title"
            FAS_InsertSlideTitleBox
        Case "9", "subtitle"
            FAS_InsertSlideSubtitleBox
        Case Else
            FAS_SetStatus "Unknown slide box preset: " & userInput
    End Select
End Sub

Public Sub FAS_SizeSelectedToSlideContent()
    FAS_SizeSelectedObjects FAS_SLIDE_WIDTH_CM, FAS_CONTENT_HEIGHT_CM, True, True, "content full"
End Sub

Public Sub FAS_SizeSelectedToSlideHalfWidth()
    FAS_SizeSelectedObjects FAS_GetColumnWidthCm(2), 0, True, False, "1/2 width"
End Sub

Public Sub FAS_SizeSelectedToSlideThirdWidth()
    FAS_SizeSelectedObjects FAS_GetColumnWidthCm(3), 0, True, False, "1/3 width"
End Sub

Public Sub FAS_SizeSelectedToSlideTwoThirdsWidth()
    FAS_SizeSelectedObjects FAS_GetColumnSpanCm(3, 2), 0, True, False, "2/3 width"
End Sub

Public Sub FAS_SizeSelectedToSlideQuarterWidth()
    FAS_SizeSelectedObjects FAS_GetColumnWidthCm(4), 0, True, False, "1/4 width"
End Sub

Public Sub FAS_SizeSelectedToSlideHalfRowHeight()
    FAS_SizeSelectedObjects 0, FAS_GetRowHeightCm(2), False, True, "1/2 row height"
End Sub

Public Sub FAS_SizeSelectedToSlideSmallBlock()
    FAS_SizeSelectedObjects FAS_GetColumnWidthCm(3), FAS_GetRowHeightCm(2), True, True, "1/3 x 1/2 block"
End Sub

Public Sub FAS_ShowSlideSizeMenu()
    Dim prompt As String
    Dim userInput As String

    prompt = "Resize selected shapes/charts:" & vbCrLf & vbCrLf & _
        "1 = content full" & vbCrLf & _
        "2 = width 1/2" & vbCrLf & _
        "3 = width 1/3" & vbCrLf & _
        "4 = width 2/3" & vbCrLf & _
        "5 = width 1/4" & vbCrLf & _
        "6 = height 1/2 row" & vbCrLf & _
        "7 = small block (1/3 x 1/2)"

    userInput = Trim$(InputBox(prompt, "FAS Slide Size", "3"))
    If Len(userInput) = 0 Then Exit Sub

    Select Case LCase$(userInput)
        Case "1", "full", "content", "content full"
            FAS_SizeSelectedToSlideContent
        Case "2", "half", "1/2", "width half"
            FAS_SizeSelectedToSlideHalfWidth
        Case "3", "third", "1/3", "width third"
            FAS_SizeSelectedToSlideThirdWidth
        Case "4", "twothirds", "two thirds", "2/3", "width 2/3"
            FAS_SizeSelectedToSlideTwoThirdsWidth
        Case "5", "quarter", "1/4", "width quarter"
            FAS_SizeSelectedToSlideQuarterWidth
        Case "6", "row", "halfrow", "half row", "height half"
            FAS_SizeSelectedToSlideHalfRowHeight
        Case "7", "small", "block", "smallblock"
            FAS_SizeSelectedToSlideSmallBlock
        Case Else
            FAS_SetStatus "Unknown slide size preset: " & userInput
    End Select
End Sub

Public Sub FAS_ShowSlideLayoutDebug()
    Dim messageText As String

    messageText = FAS_SLIDE_LAYOUT_VERSION & " | " & _
        "full " & FAS_FormatCm(FAS_SLIDE_WIDTH_CM) & " x " & FAS_FormatCm(FAS_CONTENT_HEIGHT_CM) & " cm" & _
        " | half " & FAS_FormatCm(FAS_GetColumnWidthCm(2)) & _
        " | third " & FAS_FormatCm(FAS_GetColumnWidthCm(3)) & _
        " | 2/3 " & FAS_FormatCm(FAS_GetColumnSpanCm(3, 2)) & _
        " | quarter " & FAS_FormatCm(FAS_GetColumnWidthCm(4)) & _
        " | half-row " & FAS_FormatCm(FAS_GetRowHeightCm(2))

    FAS_SetStatus messageText
    MsgBox messageText, vbInformation, "FAS Slide Layout"
End Sub

Public Function FAS_RunSlideLayoutCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    If Len(normalizedCommand) = 0 Then Exit Function

    Select Case normalizedCommand
        Case "box", "boxes", "slidebox", "slideboxes"
            FAS_ShowSlideBoxMenu
        Case "boxdebug", "layoutdebug", "slidedebug", "faslayout"
            FAS_ShowSlideLayoutDebug
        Case "boxfull", "boxcontent", "boxcontentfull"
            FAS_InsertSlideContentBox
        Case "boxtitle"
            FAS_InsertSlideTitleBox
        Case "boxsubtitle"
            FAS_InsertSlideSubtitleBox
        Case "boxhalf", "box12"
            FAS_InsertSlideHalfBox
        Case "boxthird", "box13"
            FAS_InsertSlideThirdBox
        Case "box23", "boxtwothirds"
            FAS_InsertSlideTwoThirdsBox
        Case "boxquarter", "box14"
            FAS_InsertSlideQuarterBox
        Case "boxrow", "boxhalfrow"
            FAS_InsertSlideHalfRowBox
        Case "boxsmall", "boxblock", "boxsmallblock"
            FAS_InsertSlideSmallBlockBox
        Case "size", "slidesize", "chartsize"
            FAS_ShowSlideSizeMenu
        Case "sizefull", "sizecontent", "sizecontentfull", "chartfull"
            FAS_SizeSelectedToSlideContent
        Case "sizehalf", "size12", "charthalf", "chart12"
            FAS_SizeSelectedToSlideHalfWidth
        Case "sizethird", "size13", "chartthird", "chart13"
            FAS_SizeSelectedToSlideThirdWidth
        Case "size23", "sizetwothirds", "chart23", "charttwothirds"
            FAS_SizeSelectedToSlideTwoThirdsWidth
        Case "sizequarter", "size14", "chartquarter", "chart14"
            FAS_SizeSelectedToSlideQuarterWidth
        Case "sizerow", "sizehalfrow", "chartrow"
            FAS_SizeSelectedToSlideHalfRowHeight
        Case "sizesmall", "sizeblock", "sizesmallblock", "chartsmall", "chartblock"
            FAS_SizeSelectedToSlideSmallBlock
        Case Else
            Exit Function
    End Select

    FAS_RunSlideLayoutCommand = True
End Function

Public Function FAS_IsSlideLayoutCommand(ByVal commandText As String) As Boolean
    Dim normalizedCommand As String

    normalizedCommand = FAS_NormalizeCommandText(commandText)
    Select Case normalizedCommand
        Case "box", "boxes", "slidebox", "slideboxes", _
            "boxdebug", "layoutdebug", "slidedebug", "faslayout", _
            "boxfull", "boxcontent", "boxcontentfull", "boxtitle", "boxsubtitle", _
            "boxhalf", "box12", "boxthird", "box13", "box23", "boxtwothirds", _
            "boxquarter", "box14", "boxrow", "boxhalfrow", "boxsmall", "boxblock", "boxsmallblock", _
            "size", "slidesize", "chartsize", "sizefull", "sizecontent", "sizecontentfull", "chartfull", _
            "sizehalf", "size12", "charthalf", "chart12", "sizethird", "size13", "chartthird", "chart13", _
            "size23", "sizetwothirds", "chart23", "charttwothirds", _
            "sizequarter", "size14", "chartquarter", "chart14", "sizerow", "sizehalfrow", "chartrow", _
            "sizesmall", "sizeblock", "sizesmallblock", "chartsmall", "chartblock"
            FAS_IsSlideLayoutCommand = True
    End Select
End Function

Private Sub FAS_InsertSlideLayoutBox(ByVal labelText As String, ByVal widthCm As Double, ByVal heightCm As Double)
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim anchorRange As Range
    Dim targetShape As Shape

    Set targetWorkbook = FAS_GetTargetWorkbook()
    If targetWorkbook Is Nothing Then
        FAS_SetStatus "No active user workbook available."
        Exit Sub
    End If

    Set targetSheet = ActiveSheet
    Set anchorRange = FAS_GetBoxAnchorRange()
    If anchorRange Is Nothing Then
        FAS_SetStatus "Select a cell where the outline box should start."
        Exit Sub
    End If

    Set targetShape = targetSheet.Shapes.AddShape(msoShapeRectangle, anchorRange.Left, anchorRange.Top, FAS_CmToPoints(widthCm), FAS_CmToPoints(heightCm))
    targetShape.Name = FAS_BuildShapeName(labelText)
    targetShape.Placement = xlFreeFloating
    targetShape.Fill.Visible = msoFalse
    targetShape.Line.Visible = msoTrue
    targetShape.Line.ForeColor.RGB = FAS_BOX_LINE_COLOR
    targetShape.Line.Weight = 1.25
    targetShape.Line.DashStyle = msoLineDash
    targetShape.TextFrame.Characters.Text = vbNullString

    FAS_SetStatus "Inserted " & labelText & " outline: " & FAS_FormatCm(widthCm) & " x " & FAS_FormatCm(heightCm) & " cm."
End Sub

Private Sub FAS_SizeSelectedObjects(ByVal widthCm As Double, ByVal heightCm As Double, ByVal setWidth As Boolean, ByVal setHeight As Boolean, ByVal labelText As String)
    Dim selectedObjects As Collection
    Dim targetObject As Variant
    Dim changedCount As Long

    Set selectedObjects = FAS_GetSelectedLayoutObjects()
    If selectedObjects.Count = 0 Then
        FAS_SetStatus "Select one or more shapes/charts first."
        Exit Sub
    End If

    For Each targetObject In selectedObjects
        If setWidth Then targetObject.Width = FAS_CmToPoints(widthCm)
        If setHeight Then targetObject.Height = FAS_CmToPoints(heightCm)
        On Error Resume Next
        targetObject.Placement = xlFreeFloating
        On Error GoTo 0
        changedCount = changedCount + 1
    Next targetObject

    FAS_SetStatus "Sized " & changedCount & " selected object(s) to " & labelText & "."
End Sub

Private Function FAS_GetSelectedLayoutObjects() As Collection
    Dim selectedObjects As Collection
    Dim selectedShapeRange As Object
    Dim selectedShape As Variant
    Dim chartParent As Object

    Set selectedObjects = New Collection

    On Error Resume Next
    Set selectedShapeRange = Selection.ShapeRange
    On Error GoTo 0

    If Not selectedShapeRange Is Nothing Then
        For Each selectedShape In selectedShapeRange
            selectedObjects.Add selectedShape
        Next selectedShape
    End If

    If selectedObjects.Count = 0 Then
        On Error Resume Next
        If Not ActiveChart Is Nothing Then
            Set chartParent = ActiveChart.Parent
        End If
        On Error GoTo 0

        If Not chartParent Is Nothing Then
            selectedObjects.Add chartParent
        End If
    End If

    Set FAS_GetSelectedLayoutObjects = selectedObjects
End Function

Private Function FAS_GetBoxAnchorRange() As Range
    If TypeName(Selection) = "Range" Then
        Set FAS_GetBoxAnchorRange = Selection.Cells(1, 1)
    Else
        Set FAS_GetBoxAnchorRange = ActiveCell
    End If
End Function

Private Function FAS_GetColumnWidthCm(ByVal columnCount As Long) As Double
    FAS_GetColumnWidthCm = (FAS_SLIDE_WIDTH_CM - FAS_SLIDE_GUTTER_CM * (columnCount - 1)) / columnCount
End Function

Private Function FAS_GetColumnSpanCm(ByVal columnCount As Long, ByVal spanCount As Long) As Double
    FAS_GetColumnSpanCm = FAS_GetColumnWidthCm(columnCount) * spanCount + FAS_SLIDE_GUTTER_CM * (spanCount - 1)
End Function

Private Function FAS_GetRowHeightCm(ByVal rowCount As Long) As Double
    FAS_GetRowHeightCm = (FAS_CONTENT_HEIGHT_CM - FAS_SLIDE_GUTTER_CM * (rowCount - 1)) / rowCount
End Function

Private Function FAS_CmToPoints(ByVal valueCm As Double) As Double
    FAS_CmToPoints = valueCm * 72 / 2.54
End Function

Private Function FAS_FormatCm(ByVal valueCm As Double) As String
    FAS_FormatCm = Format$(valueCm, "0.00")
End Function

Private Function FAS_BuildShapeName(ByVal labelText As String) As String
    FAS_BuildShapeName = "FAS_SlideBox_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Replace(Replace(labelText, " ", "_"), "/", "-")
End Function
