VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet903"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
  Cancel = T1M.Action_WorkSheet_ToggleSection
End Sub

Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
  Cancel = T1M.Action_WorkSheet_ShowPopupMenu()
End Sub

Private Sub Worksheet_Deactivate()
  T1M.Action_WorkSheet_ClearStatusMessage
End Sub

