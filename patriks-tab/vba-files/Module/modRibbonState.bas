Attribute VB_Name = "modRibbonState"
Option Explicit

Private gFASRibbon As Office.IRibbonUI

' Keep this signature identical to Office's documented customUI onLoad contract.
' The Ribbon object is supplied when Excel loads the add-in's custom UI.
Public Sub Ribbon_OnLoad(ribbon As Office.IRibbonUI)
    Set gFASRibbon = ribbon
End Sub

Public Sub FAS_InvalidateRibbonControl(ByVal controlId As String)
    On Error Resume Next
    If Not gFASRibbon Is Nothing Then gFASRibbon.InvalidateControl controlId
    On Error GoTo 0
End Sub
