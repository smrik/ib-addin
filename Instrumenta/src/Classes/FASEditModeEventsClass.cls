VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FASEditModeEventsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Private Sub app_WindowSelectionChange(ByVal sel As Selection)
    FASEditModeRefreshPalette
End Sub

Private Sub app_AfterShapeSizeChange(ByVal shp As shape)
    FASEditModeRefreshPalette
End Sub
