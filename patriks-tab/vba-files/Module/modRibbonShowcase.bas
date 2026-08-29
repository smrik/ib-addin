Attribute VB_Name = "modRibbonShowcase"
Option Explicit

Private mFASShowcaseLinks As Boolean
Private mFASShowcaseHidden As Boolean

Public Sub Ribbon_ShowcaseSplitDefault(ByVal control As Office.IRibbonControl)
    FAS_ShowcaseSetStatus "UI2 splitButton: primary action selected."
End Sub

Public Sub Ribbon_ShowcaseSplitOption(ByVal control As Office.IRibbonControl)
    Select Case control.Id
        Case "btnFASUI2SplitOptionA"
            FAS_ShowcaseSetStatus "UI2 splitButton: Option A selected."
        Case "btnFASUI2SplitOptionB"
            FAS_ShowcaseSetStatus "UI2 splitButton: Option B selected."
        Case Else
            FAS_ShowcaseSetStatus "UI2 splitButton option selected."
    End Select
End Sub

Public Sub Ribbon_ShowcaseOpenSettings(ByVal control As Office.IRibbonControl)
    FAS_ShowShortcutSettings
End Sub

Public Sub Ribbon_ShowcaseSearchChanged(ByVal control As Office.IRibbonControl, ByVal text As String)
    If Len(Trim$(text)) = 0 Then
        FAS_ShowcaseSetStatus "UI2 editBox cleared."
    Else
        FAS_ShowcaseSetStatus "UI2 editBox: " & text
    End If
End Sub

Public Sub Ribbon_ShowcaseRoundChanged(ByVal control As Office.IRibbonControl, ByVal selectedId As String, ByVal selectedIndex As Integer)
    Select Case selectedId
        Case "itemFASUI2Round0": FAS_ShowcaseSetStatus "UI2 dropDown: 0 digits selected."
        Case "itemFASUI2Round1": FAS_ShowcaseSetStatus "UI2 dropDown: 1 digit selected."
        Case "itemFASUI2Round2": FAS_ShowcaseSetStatus "UI2 dropDown: 2 digits selected."
        Case "itemFASUI2Round3": FAS_ShowcaseSetStatus "UI2 dropDown: 3 digits selected."
        Case Else: FAS_ShowcaseSetStatus "UI2 dropDown selection: " & selectedId
    End Select
End Sub

Public Sub Ribbon_ShowcaseToggleLinks(ByVal control As Office.IRibbonControl, ByVal pressed As Boolean)
    mFASShowcaseLinks = pressed
    FAS_ShowcaseSetStatus IIf(pressed, "UI2 toggleButton pressed: Keep Links.", "UI2 toggleButton released: Keep Links.")
    FAS_InvalidateShowcaseControl "tglFASUI2Links"
End Sub

Public Sub Ribbon_ShowcaseGetLinksPressed(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = mFASShowcaseLinks
End Sub

Public Sub Ribbon_ShowcaseToggleHidden(ByVal control As Office.IRibbonControl, ByVal pressed As Boolean)
    mFASShowcaseHidden = pressed
    FAS_ShowcaseSetStatus IIf(pressed, "UI2 checkBox checked: Include hidden rows.", "UI2 checkBox cleared: Include hidden rows.")
    FAS_InvalidateShowcaseControl "chkFASUI2Hidden"
End Sub

Public Sub Ribbon_ShowcaseGetHiddenPressed(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = mFASShowcaseHidden
End Sub

Public Sub Ribbon_ShowcaseButton(ByVal control As Office.IRibbonControl)
    FAS_ShowcaseSetStatus "UI2 compact control selected: " & control.Id
End Sub

Public Sub Ribbon_ShowcaseGalleryItem(ByVal control As Office.IRibbonControl, ByVal selectedId As String, ByVal selectedIndex As Integer)
    Select Case selectedId
        Case "itemFASUI2Chart": FAS_ShowcaseSetStatus "UI2 gallery: Chart preset selected."
        Case "itemFASUI2Table": FAS_ShowcaseSetStatus "UI2 gallery: Table preset selected."
        Case "itemFASUI2Help": FAS_ShowcaseSetStatus "UI2 gallery: Help preset selected."
        Case Else: FAS_ShowcaseSetStatus "UI2 gallery item selected: " & selectedId
    End Select
End Sub

Public Sub Ribbon_ShowcaseGetSheets(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    Const ribbonNamespace As String = "http://schemas.microsoft.com/office/2009/07/customui"
    Dim xmlText As String
    Dim activeBook As Object
    Dim sheetItem As Object
    Dim sheetIndex As Long

    xmlText = "<menu xmlns=""" & ribbonNamespace & """>"

    On Error GoTo BuildFailed
    Set activeBook = Application.ActiveWorkbook

    If activeBook Is Nothing Then
        xmlText = xmlText & "<button id=""btnFASUI2NoWorkbook"" label=""No active workbook"" enabled=""false""/>"
    Else
        For Each sheetItem In activeBook.Worksheets
            sheetIndex = sheetIndex + 1
            xmlText = xmlText & "<button id=""btnFASUI2Sheet" & CStr(sheetIndex) & """ label=""" _
                & FAS_ShowcaseEscapeXml(CStr(sheetItem.Name)) & """ onAction=""Ribbon_ShowcaseSheetAction""/>"
        Next sheetItem

        If sheetIndex = 0 Then
            xmlText = xmlText & "<button id=""btnFASUI2NoSheets"" label=""No worksheets"" enabled=""false""/>"
        End If
    End If

    xmlText = xmlText & "</menu>"
    returnedVal = xmlText
    Exit Sub

BuildFailed:
    xmlText = "<menu xmlns=""" & ribbonNamespace & """><button id=""btnFASUI2SheetsError"" label=""Could not read workbook sheets"" enabled=""false""/></menu>"
    returnedVal = xmlText
End Sub

Public Sub Ribbon_ShowcaseSheetAction(ByVal control As Office.IRibbonControl)
    Dim sheetIndex As Long
    Dim activeBook As Object

    sheetIndex = Val(Replace(control.Id, "btnFASUI2Sheet", vbNullString))
    If sheetIndex <= 0 Then
        FAS_ShowcaseSetStatus "UI2 dynamicMenu item selected."
        Exit Sub
    End If

    On Error GoTo SelectionFailed
    Set activeBook = Application.ActiveWorkbook
    If activeBook Is Nothing Then
        FAS_ShowcaseSetStatus "UI2 dynamicMenu: workbook changed; reopen the menu."
    ElseIf sheetIndex > activeBook.Worksheets.Count Then
        FAS_ShowcaseSetStatus "UI2 dynamicMenu: workbook changed; reopen the menu."
    Else
        FAS_ShowcaseSetStatus "UI2 dynamicMenu selected: " & CStr(activeBook.Worksheets(sheetIndex).Name)
    End If
    Exit Sub

SelectionFailed:
    FAS_ShowcaseSetStatus "UI2 dynamicMenu selection could not be read."
End Sub

Private Sub FAS_ShowcaseSetStatus(ByVal messageText As String)
    On Error Resume Next
    Application.StatusBar = "FAS: " & messageText
    On Error GoTo 0
End Sub

Private Function FAS_ShowcaseEscapeXml(ByVal textValue As String) As String
    textValue = Replace(textValue, "&", "&amp;")
    textValue = Replace(textValue, "<", "&lt;")
    textValue = Replace(textValue, ">", "&gt;")
    textValue = Replace(textValue, """", "&quot;")
    textValue = Replace(textValue, "'", "&apos;")
    FAS_ShowcaseEscapeXml = textValue
End Function

Private Sub FAS_InvalidateShowcaseControl(ByVal controlId As String)
    FAS_InvalidateRibbonControl controlId
End Sub
