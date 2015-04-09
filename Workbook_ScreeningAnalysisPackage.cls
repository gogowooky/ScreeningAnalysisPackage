VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Rem ******************************************************************************************************
Rem アドインの追加登録と削除
Rem ******************************************************************************************************

Private Sub Workbook_Activate()
   T1M.Action_WorkBook_Initialize
End Sub

Private Sub Workbook_Deactivate()
  T1M.Action_WorkBook_Finalize
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  T1M.Action_WorkBook_AfterSave
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Application.Calculation = xlCalculationManual
  T1M.Action_WorkBook_BeforeSave
End Sub


' 開発完了したらこれより下を削除する。

Private Sub Workbook_Open()
  Application.MacroOptions Macro:="ThisWorkbook.reload_module", ShortcutKey:="r"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Application.MacroOptions Macro:="ThisWorkbook.reload_module", ShortcutKey:=""
End Sub

Public Sub reload_module()
Attribute reload_module.VB_ProcData.VB_Invoke_Func = "r\n14"
  On Error Resume Next
  Dim modules As Variant
  modules = Array("RESOURCE.bas", "T1.bas", "T1M.bas", "TSUKUBA_UTIL.bas", _
                  "AssayResults.cls", "CompoundPlatemap.cls", "PlateAlignment.cls", "PlateAlignments.cls", "Well2RowCol.cls")
  
  If 0 < InStr(Application.OperatingSystem, "Windows") And _
     0 < InStr(ActiveWorkbook.path, "ScreeningAnalysisPackage") Then
    
    Dim mdl As Variant
    With ThisWorkbook.VBProject
      For Each mdl In modules
        .VBComponents.Remove .VBComponents(Left(mdl, Len(mdl) - 4))
        .VBComponents.Remove .VBComponents(Left(mdl, Len(mdl) - 4)) ' Remove前にImportしてしまう対策
        ' 対策してもT1MはT1M1に置き換わってしまうので注意。[2015-04-09]
        .VBComponents.Import ActiveWorkbook.path & "\\" & mdl
      Next
    End With
  End If
End Sub



