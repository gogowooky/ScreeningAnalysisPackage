Attribute VB_Name = "T1M"
Rem ******************************************************************************************************
Rem ƒfƒoƒO—pƒtƒ‰ƒO
Rem ******************************************************************************************************
Public Const debug_full_context_menu_in_template = False ' release‚É‚ÍFalse
Public Const debug_convert_relative_formula_to_absolute = False ' release‚É‚ÍTrue
Public Const debug_force_to_popup_context_menu = "" ' release‚É‚Í""
 
 
Rem ******************************************************************************************************
Rem ƒZƒNƒVƒ‡ƒ“—p‚ÌF
Rem ******************************************************************************************************
Const INFO_SECTION_THEME_COLOR = xlThemeColorLight2 ' Â
Const INFO_SECTION_TINT1_COLOR = 0.3
Const INFO_SECTION_TINT2_COLOR = 0.5
Const DATA_SECTION_THEME_COLOR = xlThemeColorAccent2 ' Ô
Const DATA_SECTION_TINT1_COLOR = -0.3
Const DATA_SECTION_TINT2_COLOR = 0.4
Const ANAL_SECTION_THEME_COLOR = xlThemeColorAccent3 ' —Î
Const ANAL_SECTION_TINT1_COLOR = -0.4
Const ANAL_SECTION_TINT2_COLOR = 0.2
Const TBLE_SECTION_THEME_COLOR = xlThemeColorAccent6 ' ò
Const TBLE_SECTION_TINT1_COLOR = -0.2
Const TBLE_SECTION_TINT2_COLOR = -0.2
Const EXTR_SECTION_THEME_COLOR = xlThemeColorAccent4 ' ‡
Const EXTR_SECTION_TINT1_COLOR = -0.2
Const EXTR_SECTION_TINT2_COLOR = 0.2
Const END_SECTION_THEME_COLOR = xlThemeColorDark1 ' ŠDF
Const END_SECTION_TINT1_COLOR = -0.5
Const END_SECTION_TINT2_COLOR = -0.2
 
 
Rem ******************************************************************************************************
Rem ƒVƒXƒeƒ€İ’è’l
Rem ******************************************************************************************************
Public Const SYSTEM_SUPPORT_PLATE_READER = "FDSS,PHERASTER,EZREADER,ENSPIRE,HTFC,ECHO,TRISTAR,FREE"
Public Const SYSTEM_SUPPORT_PLATE_TYPE = "24,96,384,1536"
Public Const SYSTEM_SUPPORT_PLATE_FORMAT = "PRIMARY,CONFIRMATION,DOSE_RESPONSE,FREE"
Public Const SYSTEM_SUPPORT_REALTIME_PLATE_READER = "FDSS,FLIPR"
Public Const SHEETNAME_ASSAY_SUMMARY = "Plates"
Public Const SHEETNAME_REPORT_QC_RESULT = "QCŒ‹‰Ê"
Public Const SHEETNAME_REPORT_ASSAY_RESULT = "ƒAƒbƒZƒCŒ‹‰Ê"
 
Public Const LABEL_PLATE_TYPE = "PLATE_TYPE"
Public Const LABEL_PLATE_READER = "PLATE_READER"
Public Const LABEL_PLATE_FORMAT = "PLATE_FORMAT"
Public Const LABEL_PLATE_WELL_POSITION = "WELL_POS"
Public Const LABEL_PLATE_WELL_ROLE = "WELL_ROLE"
Public Const LABEL_PLATE_COMPOUND_CONC = "CPD_CONC"
Public Const LABEL_TABLE = "TABLE"
 
Public Const PLATE_TITLE = "ƒvƒŒ[ƒgID,Plate,plate ID,Plate_ID"
Public Const WELL_TITLE = "WELL,well"
Public Const WELLROLE_TITLE = "WELL_ROLE,well_role,ROLE"
Public Const COMPOUND_TITLE = "‰»‡•¨ƒTƒ“ƒvƒ‹ID,Compound_Name"
 
Const PLATESHEET_TITLE_FOR_RAWDATA_COLUMN = "Raw Data Filename"
Const PLATESHEET_TITLE_FOR_PLATEID_COLUMN = "PlateID"
Const PLATESHEET_EXTENSION_FOR_FILE_LISTING = "TXT,SCV,CSV,RST"
 
Private MenuLang As Integer ' 0: “ú–{Œê, 1:‰pŒê
 
 
Rem ******************************************************************************************************
Rem ƒ[ƒNƒuƒbƒNƒCƒxƒ“ƒg
Rem ******************************************************************************************************
 
' •Û‘¶‘OÄŒvZƒIƒt
Public Sub Action_WorkBook_BeforeSave()
  T1M.Action_MainMenu_Maintenance_CalculateOff
End Sub
 
' •Û‘¶ŒãÄŒvZƒIƒ“
Public Sub Action_WorkBook_AfterSave()
  T1M.Action_MainMenu_Maintenance_CalculateOn
End Sub
 
' ƒtƒ@ƒCƒ‹ƒNƒ[ƒY
Public Sub Action_WorkBook_Finalize()
  On Error Resume Next
  Application.CommandBars("Worksheet Menu Bar").Controls(T1.SYSTEM()).Delete  ' ƒƒjƒ…[íœ
End Sub
 
 
' ƒƒjƒ…[‚ğƒŠƒZƒbƒg‚·‚é
Private Sub ResetMenu()
  On Error Resume Next
 
  Application.CommandBars("Worksheet Menu Bar").Controls(T1.SYSTEM()).Delete  ' ‚¢‚Á‚½‚ñƒƒjƒ…[íœ
 
  With Application.CommandBars("Worksheet Menu Bar").Controls.Add(Type:=msoControlPopup)
    .Caption = T1.SYSTEM()
    With .Controls.Add(Type:=msoControlPopup)
      .Caption = "1. " & Split("ƒXƒNƒŠ[ƒjƒ“ƒOƒf[ƒ^‚Ìˆ—,Process screening data", ",")(MenuLang)
      .Enabled = 0 < InStr(T1M.GetAnalysisState(), "Template@")
      With .Controls.Add
        .Caption = "1-1. " & Split("ƒf[ƒ^ƒtƒ@ƒCƒ‹‚ğƒŠƒXƒgƒAƒbƒv‚µAƒvƒŒ[ƒg–¼‚ğ‘Î‰•t‚¯‚é,List data files and associate them to plates", ",")(MenuLang)
        .OnAction = "Action_MainMenu_Binding_RawData_To_PlateName"
      End With
      With .Controls.Add
        .Caption = "1-2. " & Split("ƒAƒbƒZƒCƒf[ƒ^‚Ì©“®‰ğÍˆ—‚ğŠJn‚·‚é,Start data processing", ",")(MenuLang)
        .OnAction = "Action_MainMenu_Data_Analysis"
      End With
      With .Controls.Add
        .Caption = Split("‰ğÍŒ‹‰Ê‚ğÁ‹‚·‚é,Remove processed data", ",")(MenuLang)
        .OnAction = "Action_MainMenu_Clear_All_Analyzed_Data"
      End With
    End With
    With .Controls.Add(Type:=msoControlPopup)
      .Caption = "2. " & Split("‰ğÍƒf[ƒ^‚Ì“‡,Integrate analyzed data", ",")(MenuLang)
      .Enabled = 0 < InStr(T1M.GetAnalysisState(), "Template@")
      With .Controls.Add
      .Caption = "2-a. " & Split("‘S‰ğÍŒ‹‰Ê‚ğPDF‚Éo—Í‚·‚é,Export all data to pdf", ",")(MenmuLang)
        .OnAction = "Action_ContextMenu_Export_PDF"
      End With
      With .Controls.Add
        .Caption = "2-b. " & Split("‰ğÍƒf[ƒ^‚ğ•ñ‘‚É“]‹L‚·‚é,Transfer data to report", ",")(MenuLang)
        .OnAction = "Action_MainMenu_Transfer_Data_To_ReportSheet"
      End With
      With .Controls.Add
        .Caption = "2-c. " & Split("‘SƒV[ƒg‚Ì‰ğÍƒf[ƒ^‚ğCSV‚ÉExport‚·‚é,Export all data to csv", ",")(MenuLang)
        .OnAction = "Action_MainMenu_Convert_All_Sheets_To_CSV"
      End With
      With .Controls.Add
        .Caption = "2-d. " & Split("“¯ˆêƒfƒBƒŒƒNƒgƒŠ“à‚Ì‘Scsvƒtƒ@ƒCƒ‹‚ğƒ}[ƒW‚·‚é,Merge all csv files within the directory", ",")(MenuLang)
        .OnAction = "Action_MainMenu_Merge_All_CSV_Files"
      End With
    End With
    With .Controls.Add(Type:=msoControlPopup)
      .Caption = Split("Misc,Misc", ",")(MenuLang)
      With .Controls.Add
        .Caption = Split("ŠÖ”ƒwƒ‹ƒv,Function Help", ",")(MenuLang)
        .OnAction = "Action_Menu_Show_Help"
      End With
      With .Controls.Add
        .Caption = Split("Package‚Ìƒo[ƒWƒ‡ƒ“A‚»‚Ì‘¼î•ñ,Version info", ",")(MenuLang)
        .OnAction = "Action_Menu_Show_Information"
      End With
      With .Controls.Add(Type:=msoControlPopup)
        .Caption = Split("Maintenance,Maintenance", ",")(MenuLang)
        With .Controls.Add
          .Caption = Split("Change menu to English,Change menu to Japanese", ",")(MenuLang)
          .OnAction = "Action_MainMenu_Change_MenuLang"
        End With
        With .Controls.Add
          .Caption = Split("‘¼ƒuƒbƒN‚ÉŠÖ”‚ğExport‚·‚é,Export functions", ",")(MenuLang)
          .OnAction = "Action_MainMenu_Export_Extended_Functions"
        End With
        With .Controls.Add
          .Caption = Split("‘SƒV[ƒgŒvZƒIƒ“,Turn autocalulate on", ",")(MenuLang)
          .OnAction = "Action_MainMenu_Maintenance_CalculateOn"
        End With
        With .Controls.Add
          .Caption = Split("‘SƒV[ƒgŒvZƒIƒt,Turn autocalculation off", ",")(MenuLang)
          .OnAction = "Action_MainMenu_Maintenance_CalculateOff"
        End With
        With .Controls.Add
          .Caption = Split("“à•”ƒRƒŒƒNƒVƒ‡ƒ“•Ï”ƒŠƒZƒbƒg,Reset inner variables", ",")(MenuLang)
          .OnAction = "Action_MainMenu_Maintenance_ResetCollection"
        End With
        With .Controls.Add
          .Caption = Split("‘SƒvƒŒ[ƒgÄŒvZ,Recalculate", ",")(MenuLang)
          .OnAction = "Action_MainMenu_Maintenance_UpdateAllPlate"
        End With
        With .Controls.Add
          .Caption = Split("ƒ‚ƒWƒ…[ƒ‹Ä“Ç‚İ‚İ,Reload the screening macro modules", ",")(MenuLang)
          .OnAction = "Action_MainMenu_Maintenance_ReloadModule"
        End With
        With .Controls.Add
          .Caption = Split("Template‚ğÄ“K—p,Reapply Template", ",")(MenuLang)
          .OnAction = "Action_CopyTemplate_Paste"
        End With
        With .Controls.Add
          .Caption = Split("ŠÖ”‚ğ”’l‰»,Put function to value", ",")(MenuLang)
          .OnAction = "Action_Function_Value"
        End With
 
      End With
    End With
  End With
End Sub
 
' ƒtƒ@ƒCƒ‹ƒI[ƒvƒ“
Public Sub Action_WorkBook_Initialize()
  On Error Resume Next
  Application.Calculation = xlCalculationManual
 
  MenuLang = 0 ' “ú–{Œê(0)
  ResetMenu
 
End Sub
 
 
 
Rem ******************************************************************************************************
Rem ƒ[ƒNƒV[ƒgƒCƒxƒ“ƒg
Rem ******************************************************************************************************
 
' StatusBar‚ÌƒƒbƒZ[ƒW‚ğÁ‹
Public Sub Action_WorkSheet_ClearStatusMessage()
  Call TSUKUBA_UTIL.ShowStatusMessage("")
End Sub
 
' ƒZƒNƒVƒ‡ƒ“‚ğŠJ•Â‚·‚é
Public Function Action_WorkSheet_ToggleSection()
        If ActiveCell.Interior.ThemeColor <> -4142 Then
                If ActiveCell.Interior.ThemeColor = Cells(ActiveCell.row, 1).Interior.ThemeColor Then
                        If T1M.SECTION(ActiveCell, "hide?") Then
                                Rows(T1M.SECTION(ActiveCell, "inrows")).Hidden = False
                        Else
                                Rows(T1M.SECTION(ActiveCell, "inrows")).Hidden = True
                        End If
                End If
        End If
End Function
 
 
 
Rem ******************************************************************************************************
Rem ƒ_ƒCƒAƒƒOƒCƒxƒ“ƒg
Rem ******************************************************************************************************
' ƒz[ƒ€ƒy[ƒW‚ğŠJ‚­
Public Sub Action_Menu_OpenSite()
  TSUKUBA_UTIL.OpenUrl T1.SYSTEM("homepage")
End Sub
 
' ƒAƒbƒZƒCŒnƒoƒŠƒf[ƒVƒ‡ƒ“•û–@‚É‚Â‚¢‚Ä
Public Sub Action_Menu_OpenAssayValidation()
  TSUKUBA_UTIL.OpenUrl T1.SYSTEM("validation")
End Sub
 
' ‰»‡•¨”z•z‚É‚Â‚¢‚Ä
Public Sub Action_Menu_OpenCompoundDistribution()
  TSUKUBA_UTIL.OpenUrl T1.SYSTEM("cpddistrib")
End Sub
 
' ¿–âƒ[ƒ‹‘—MƒtƒH[ƒ€‚ğŠJ‚­
Public Sub Action_Menu_OpenMail()
  TSUKUBA_UTIL.OpenUrl T1.SYSTEM("mailto")
End Sub
 
' ŠJ”­”ÅƒTƒCƒg‚ğŠJ‚­
Public Sub Action_Menu_OpenGitHub()
  TSUKUBA_UTIL.OpenUrl T1.SYSTEM("original")
End Sub
 
' MITƒ‰ƒCƒZƒ“ƒXƒTƒCƒg‚ğŠJ‚­
Public Sub Action_Menu_OpenMITLisence()
  TSUKUBA_UTIL.OpenUrl "http://opensource.org/licenses/mit-license.php"
End Sub
 
 
 
Rem ******************************************************************************************************
Rem ƒƒCƒ“ƒƒjƒ…[ƒCƒxƒ“ƒg
Rem ******************************************************************************************************
' ŠÖ”ƒwƒ‹ƒv‚ğŠJ‚­
Private Sub Action_Menu_Show_Help()
  If TSUKUBA_UTIL.ExistSheetP("HELP") Then
    Worksheets("HELP").Select
  Else
    TSUKUBA_UTIL.DupulicateHiddenSheetAndShow "TSUKUBA_HELP", "HELP"
  End If
End Sub
 
' ƒo[ƒWƒ‡ƒ“A‚»‚Ì‘¼î•ñ‚Ìƒ_ƒCƒAƒƒO‚ğŠJ‚­
Private Sub Action_Menu_Show_Information()
  Version.Caption = T1.SYSTEM()
  Version.Label1 = T1.SYSTEM("title")
  Version.Label2 = T1.SYSTEM("version")
  Version.Label7 = "last updated at " & T1.SYSTEM("update")
  Version.Label4 = T1.SYSTEM("mail")
  Version.Label5 = T1.SYSTEM("affiliation")
  Version.Label9 = "ƒAƒbƒZƒCŒn‚Ì•]‰¿"
  Version.Label6 = "‰»‡•¨ƒ‰ƒCƒuƒ‰ƒŠ—˜—p\"
  Version.Label11 = "GitHub Site:" & vbCrLf & T1.SYSTEM("original")
  Version.Left = Application.Left + (Application.Width - Version.Width) / 2
  Version.Show
End Sub
 
' ƒƒjƒ…[Œ¾Œê‚ğ•ÏX
Private Sub Action_MainMenu_Change_MenuLang()
  If MenuLang = 1 Then
    MenuLang = 0
  Else
    MenuLang = 1
  End If
  ResetMenu
End Sub
 
' ‘SƒvƒŒ[ƒg‚ğÄŒvZ
Private Sub Action_MainMenu_Maintenance_UpdateAllPlate()
  For Each plt In T1.CSV2ARY(T1.ASSAY("plates"))
    Sheets(plt).Activate
    T1M.Action_Worksheet_Update
  Next
End Sub
 
' “à•”ƒRƒŒƒNƒVƒ‡ƒ“•Ï”‚ğƒŠƒZƒbƒg
Private Sub Action_MainMenu_Maintenance_ResetCollection()
        RESOURCE.RestAssayResult
        RESOURCE.ResetCpdTable
End Sub
 
' ‘SƒV[ƒg‚ğŒvZƒIƒt
Private Sub Action_MainMenu_Maintenance_CalculateOff()
  For Each ws In ActiveWorkbook.Worksheets: ws.EnableCalculation = False
  Next
End Sub
 
' ‘SƒV[ƒgŒvZƒIƒ“
Private Sub Action_MainMenu_Maintenance_CalculateOn()
  For Each ws In ActiveWorkbook.Worksheets: ws.EnableCalculation = True
  Next
End Sub
 
' ƒ‚ƒWƒ…[ƒ‹‚ğÄ“Ç‚İ‚İ
Public Sub Action_MainMenu_Maintenance_ReloadModule()
  On Error Resume Next
  Dim modules As Variant
  modules = Array("RESOURCE.bas", "T1.bas", "T1M.bas", "TSUKUBA_UTIL.bas", _
                  "AssayResults.cls", "CompoundPlatemap.cls", "PlateAlignment.cls", "PlateAlignments.cls", "Well2RowCol.cls")
 
 
  If 0 < InStr(Application.OperatingSystem, "Windows") Then
    Dim mdl As Variant
    With ThisWorkbook.VBProject
 
      Dim cpnt As Object
      For Each cpnt In .VBComponents
        If cpnt.Type = 1 Or cpnt.Type = 2 Then .VBComponents.Remove cpnt
      Next
 
      For Each mdl In modules
        .VBComponents.Remove .VBComponents(Left(mdl, Len(mdl) - 4))
        .VBComponents.Import ActiveWorkbook.path & "€" & mdl
      Next
 
      For Each mdl In modules
        .VBComponents(Left(mdl, Len(mdl) - 4) & "1").Name = Left(mdl, Len(mdl) - 4)
      Next
    End With
  Else
    MsgBox "‚±‚Ì‹@”\‚ÍWindows‚Ì‚İ‚É‚È‚è‚Ü‚·"
  End If
End Sub
 
' ”’l‰»‚µ‚½‰ğÍƒV[ƒg‚ğÄ”®‰»‚·‚é
Sub Action_CopyTemplate_Paste()
  If TSUKUBA_UTIL.ExistValueP(T1.ASSAY("plates"), ActiveSheet.Name) Then
    Application.ScreenUpdating = False
    Worksheets("Template").UsedRange().Copy
    ActiveSheet.UsedRange().PasteSpecial Paste:=xlPasteFormulas
    ActiveSheet.Calculate
    ActiveSheet.Range("A1").Select
    Application.ScreenUpdating = True
  End If
End Sub
 
' ‰ğÍƒV[ƒg‚ğ”’l‰»‚·‚é
Sub Action_Function_Value()
  If TSUKUBA_UTIL.ExistValueP(T1.ASSAY("plates"), ActiveSheet.Name) Then
    Application.ScreenUpdating = False
    ActiveSheet.UsedRange().Copy
    ActiveSheet.UsedRange().PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Calculate
    ActiveSheet.Range("A1").Select
    Application.ScreenUpdating = True
  End If
End Sub
 
 
 
' ‘¼ƒuƒbƒN‚ÉŠÖ”‚ğExport‚·‚é
Private Sub Action_MainMenu_Export_Extended_Functions()
  Dim Code As String: Code = ""
  Dim flag As Boolean: flag = False
  Dim i As Integer
 
  With Workbooks("ScreeningAnalysisPackage.xlsm").VBProject.VBComponents("T1").CodeModule
    For i = 1 To .CountOfLines
      If InStr(.Lines(i, 1), "EXPORT ON") Then
        flag = True: i = i + 1
      ElseIf InStr(.Lines(i, 1), "EXPORT OFF") Then
        flag = False
      End If
      If flag Then Code = Code & vbNewLine & .Lines(i, 1)
    Next
  End With
 
  Dim targetfile As String: targetfile = Application.GetOpenFilename("Microsoft ExcelƒuƒbƒN,*.xls?")
  If targetfile <> "" Then
    Workbooks.Open targetfile
    With ActiveWorkbook.VBProject.VBComponents.Add(1)
      .Name = "T1"
      .CodeModule.AddFromString Code
    End With
  End If
End Sub
 
 
'
' 1.ƒXƒNƒŠ[ƒjƒ“ƒOƒf[ƒ^‚Ìˆ— > 1-1. ƒf[ƒ^ƒtƒ@ƒCƒ‹‚ğƒŠƒXƒgƒAƒbƒv‚µAƒvƒŒ[ƒg–¼‚ğ‘Î‰•t‚¯‚é
'
Private Sub Action_MainMenu_Binding_RawData_To_PlateName()
  If TSUKUBA_UTIL.ExistSheetP(SHEETNAME_ASSAY_SUMMARY) = False Then
    Worksheets.Add
                ActiveSheet.Name = SHEETNAME_ASSAY_SUMMARY
  End If
 
  With Worksheets(SHEETNAME_ASSAY_SUMMARY)
    .Select
    .Range("A1").Value = PLATESHEET_TITLE_FOR_RAWDATA_COLUMN
    .Range("B1").Value = PLATESHEET_TITLE_FOR_PLATEID_COLUMN
 
    Dim fil As String: fil = TSUKUBA_UTIL.WinMacDir(ActiveWorkbook.path, PLATESHEET_EXTENSION_FOR_FILE_LISTING)
    Dim cnt As Integer: cnt = 1
    While fil <> ""
      .Cells(cnt + 1, 1).Value = fil
      fil = TSUKUBA_UTIL.WinMacDir()
      cnt = cnt + 1
    Wend
    .Columns("A:B").AutoFit
  End With
End Sub
 
 
 
'
' 1.ƒXƒNƒŠ[ƒjƒ“ƒOƒf[ƒ^‚Ìˆ— > ‰ğÍŒ‹‰Ê‚ğÁ‹‚·‚é"
'
Private Sub Action_MainMenu_Clear_All_Analyzed_Data()
  On Error GoTo Err_Action_MainMenu_Clear_All_Analyzed_Data
  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
 
  Dim i As Integer
  Dim rawdatas As Variant:  ReDim rawdatas(1)
  Dim templates As Variant: ReDim templates(1)
        With Worksheets(SHEETNAME_ASSAY_SUMMARY).Range("B2")
                i = 0
                Do While .Offset(i, 0).Value <> ""
                        ReDim Preserve rawdatas(i)
                        ReDim Preserve templates(i)
                        rawdatas(i) = .Offset(i, 0).Value
                        templates(i) = "(raw)" + rawdatas(i)
                        i = i + 1
                Loop
        End With
  Worksheets(templates).Select: ActiveWindow.SelectedSheets.Delete
  Worksheets(rawdatas).Select:  ActiveWindow.SelectedSheets.Delete
 
  TSUKUBA_UTIL.ShowStatusMessage "‘S‚Ä‚Ìƒf[ƒ^ƒV[ƒgAƒf[ƒ^ˆ—ƒV[ƒg‚ğíœ‚µ‚Ü‚µ‚½B"
  Worksheets(SHEETNAME_ASSAY_SUMMARY).Activate
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
 
  Exit Sub
 
Err_Action_MainMenu_Clear_All_Analyzed_Data:
  TSUKUBA_UTIL.ShowStatusMessage "ƒGƒ‰[‚Å‚·B@PlatesƒV[ƒg‚ğŠm”F‚µ‚ÄAÄ“xÀs‚µ‚Ä‚­‚¾‚³‚¢B"
  MsgBox "ƒGƒ‰[‚Å‚·B@PlatesƒV[ƒg‚ğŠm”F‚µ‚ÄAÄ“xÀs‚µ‚Ä‚­‚¾‚³‚¢B"
End Sub
 
 
'
' 1.ƒXƒNƒŠ[ƒjƒ“ƒOƒf[ƒ^‚Ìˆ— > 1-2.ƒAƒbƒZƒCƒf[ƒ^‚Ì©“®‰ğÍˆ—‚ğŠJn‚·‚é
'
Private Sub Action_MainMenu_Data_Analysis()
  On Error Resume Next
  Application.Volatile
  Application.DisplayAlerts = False
  Application.ScreenUpdating = True
  TSUKUBA_UTIL.ShowStatusMessage "ƒf[ƒ^ƒtƒ@ƒCƒ‹‚Ì“Ç‚İ‚İ‚Æ”’lˆ—‚ğŠJn‚µ‚Ü‚·B"
 
  Dim i As Integer
  With ThisWorkbook
    .Worksheets(SHEETNAME_ASSAY_SUMMARY).Select
    .Sheets("Template").EnableCalculation = False
 
    Dim plates As Variant: plates = T1.CSV2ARY(T1.ASSAY("plates"))
    Dim dfiles As Variant: dfiles = T1.CSV2ARY(T1.ASSAY("datafiles"))
 
    ' ƒf[ƒ^“Ç‚İ‚İ
    For i = 0 To UBound(plates)
      .Worksheets("Template").Copy After:=.Worksheets("Template")
      .Sheets("Template (2)").Name = plates(i)
      Workbooks.OpenText filename:=.path & Application.PathSeparator & dfiles(i)
      ActiveWorkbook.ActiveSheet.Move Before:=.Worksheets("(raw)Template")
      ActiveWorkbook.ActiveSheet.Name = "(raw)" & plates(i)
      TSUKUBA_UTIL.ShowStatusMessage "ƒf[ƒ^ˆ— [" & dfiles(i) & "] ->[" & plates(i) & "]"
      DoEvents
    Next i
 
    ' ÄŒvZ
    For i = 0 To UBound(plates)
      TSUKUBA_UTIL.ShowStatusMessage "ƒf[ƒ^ÄŒvZ’† [" & plates(i) & "]"
      .Worksheets(plates(i)).EnableCalculation = True
      .Worksheets(plates(i)).Activate
      .Worksheets(plates(i)).UsedRange.Calculate
      RESOURCE.UpdateAssayResult CStr(plates(i))
    Next i
 
    ' ”’l‰»
    Action_Function_Value
  End With
 
  ' ‰ğÍ’l‚ğListup
  With ThisWorkbook.Worksheets(SHEETNAME_ASSAY_SUMMARY)
    .Activate
    Dim lbls As Variant: lbls = T1.CSV2ARY(T1.ASSAY("platelabel"))
    For i = 0 To UBound(lbls)
      .Cells(1, 3 + i).Value = lbls(i)
      .Range(.Cells(2, 3 + i), .Cells(2 + UBound(plates), 3 + i)).Select
      With Selection.FormatConditions
        .AddColorScale ColorScaleType:=3
        .Items(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .Items(1).ColorScaleCriteria(1).FormatColor.Color = 7039480
        .Items(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .Item2(1).ColorScaleCriteria(2).Value = 50
        .Items(1).ColorScaleCriteria(2).FormatColor.Color = 8711167
        .Items(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .Items(1).ColorScaleCriteria(3).FormatColor.Color = 8109667
      End With
    Next i
    .Range(Cells(2, 3), Cells(2 + UBound(plates), 3 + UBound(lbls))).Formula = "=T1.PLATE($B2,C$1)"
  End With
 
  Range("A1").Activate
  Columns("A:Z").AutoFit
 
  TSUKUBA_UTIL.ShowStatusMessage "Excel•\¦‚ğXV’†"
  ThisWorkbook.Worksheets(SHEETNAME_ASSAY_SUMMARY).Activate
  ThisWorkbook.Worksheets(SHEETNAME_ASSAY_SUMMARY).Calculate
  TSUKUBA_UTIL.ShowStatusMessage "ƒf[ƒ^ƒtƒ@ƒCƒ‹‚Ì“Ç‚İ‚İ‚Æ”’lˆ—‚ğI—¹‚µ‚Ü‚µ‚½B"
 
  Application.DisplayAlerts = True
  Exit Sub
 
End Sub
 
 
 
 
 
 
Rem ******************************************************************************************************
Rem "2. ‰ğÍƒf[ƒ^‚Ì“‡"
Rem ******************************************************************************************************
 
Rem
Rem "2-a. ‘SƒV[ƒg‚Ì‰ğÍƒf[ƒ^‚ğcsv‚ÉExport‚·‚é"
Rem
Public Sub Action_MainMenu_Convert_All_Sheets_To_CSV()
  On Error Resume Next
 
  Dim cpd_filename As String
  Dim plt_filename As String
  Dim wel_filename As String
 
  Dim curpath As String: curpath = ActiveWorkbook.path & Application.PathSeparator
  If TSUKUBA_UTIL.ExistNameP("Template", LABEL_TABLE) Then
    cpd_filename = curpath & "cpd.csv"
    Kill cpd_filename
    Open cpd_filename For Output As #1
    Print #1, GetCpdLabels()
  End If
  plt_filename = curpath & "plate.csv"
  Open plt_filename For Output As #2
  Kill plt_filename
  Print #2, GetPlateLabels()
  wel_filename = curpath & "well.csv"
  Kill wel_filename
  Open wel_filename For Output As #3
  ' Print #3, GetWellLabels() & "," & T1.CSV_SUB(T1M.GetPlateLabels(), "PLATE_NAME") ' v1.1.2
  Print #3, GetWellLabels() & "," & T1M.GetPlateLabels()
 
  Dim csv As String
  Dim pltcsv As String
  Dim cpd_entry As Double: cpd_entry = 0
  Dim plt_entry As Double: plt_entry = 0
  Dim wel_entry As Double: wel_entry = 0
  Dim plt As Variant
  Dim rw As Integer
  Dim lbl As Variant
 
        Application.ScreenUpdating = True
 
  RESOURCE.RestAssayResult
  For Each plt In T1.CSV2ARY(T1.ASSAY("plates")) ' :::: Plate‚ğ‚Ü‚í‚·
 
    Sheets(plt).Activate
 
    TSUKUBA_UTIL.ShowStatusMessage "ƒf[ƒ^ÄŒvZ’† [" & CStr(plt) & "]"
 
    Sheets(plt).EnableCalculation = True
    Sheets(plt).Calculate
    RESOURCE.UpdateAssayResult CStr(plt)
 
    ' cpdƒe[ƒuƒ‹‚Ìo—Í
    If TSUKUBA_UTIL.ExistNameP("Template", LABEL_TABLE) Then
      TSUKUBA_UTIL.ShowStatusMessage "CSVƒGƒNƒXƒ|[ƒgˆ—(cpd) [" & plt & "]"
      Dim cpdlbls As Variant: cpdlbls = T1.CSV2ARY(T1M.GetCpdLabels())
      For rw = 1 To Range(LABEL_TABLE).Rows.COUNT - 1
        csv = ""
        For Each lbl In cpdlbls
          csv = csv & RESOURCE.GetAssayResult(CStr(plt), CStr(lbl), CInt(rw)) & ","
        Next
        Print #1, Left(csv, Len(csv) - 1): cpd_entry = cpd_entry + 1
      Next rw
    End If
 
    ' plateƒf[ƒ^‚Ìo—Í
    TSUKUBA_UTIL.ShowStatusMessage "CSVƒGƒNƒXƒ|[ƒgˆ—(plate) [" & plt & "]"
    csv = ""
    For Each lbl In T1.CSV2ARY(T1M.GetPlateLabels())
      csv = csv & RESOURCE.GetAssayResult(CStr(plt), "", CStr(lbl)) & ","
    Next
    Print #2, Left(csv, Len(csv) - 1): plt_entry = plt_entry + 1
 
    ' wellƒe[ƒuƒ‹‚Ìo—Í
    TSUKUBA_UTIL.ShowStatusMessage "CSVƒGƒNƒXƒ|[ƒgˆ—(well) [" & plt & "]"
    pltcsv = ""
    'For Each lbl In T1.CSV2ARY(T1.CSV_SUB(T1M.GetPlateLabels(), "PLATE_NAME")) ' v1.1.2
     For Each lbl In T1.CSV2ARY(T1M.GetPlateLabels())
      pltcsv = pltcsv & RESOURCE.GetAssayResult(CStr(plt), "", CStr(lbl)) & ","
    Next
    Dim wl As Variant
    Dim lbls As Variant: lbls = T1.CSV2ARY(T1M.GetWellLabels())
    For Each wl In Range(LABEL_PLATE_WELL_POSITION)
      csv = ""
      For Each lbl In lbls
        csv = csv & RESOURCE.GetAssayResult(CStr(plt), CStr(wl.Value), CStr(lbl)) & ","
      Next
      csv = csv & pltcsv
      Print #3, Left(csv, Len(csv) - 1): wel_entry = wel_entry + 1
    Next
  Next
 
        Application.ScreenUpdating = True
  Dim altname As String
  If TSUKUBA_UTIL.ExistNameP("Template", LABEL_TABLE) Then
    Close #1
    altname = curpath & Format(Now(), "YYMMDD") & "-cpd-" & CStr(cpd_entry) & ".csv"
    Kill altname
    FileCopy cpd_filename, altname
    Kill cpd_filename
  End If
 
  Close #2
  altname = curpath & Format(Now(), "YYMMDD") & "-plate-" & CStr(plt_entry) & ".csv"
  Kill altname
  FileCopy plt_filename, altname
  Kill plt_filename
 
  Close #3
  altname = curpath & Format(Now(), "YYMMDD") & "-well-" & CStr(wel_entry) & ".csv"
  Kill altname
  FileCopy wel_filename, altname
  Kill wel_filename
 
  TSUKUBA_UTIL.ShowStatusMessage "CSVƒGƒNƒXƒ|[ƒgˆ—‚ªŠ®—¹‚µ‚Ü‚µ‚½]"
 
End Sub
 
Rem
Rem "2-b. ‰ğÍƒf[ƒ^‚ğ" & T1.SYSTEM("affiliation3") & "•ñ‘‚É“]‹L‚·‚é"
Rem
Sub Action_MainMenu_Transfer_Data_To_ReportSheet()
  On Error Resume Next
  TSUKUBA_UTIL.ShowStatusMessage "•ñ‘‚Ö‚Ì“]‹Lˆ—‚ğŠJn‚µ‚Ü‚·"
  Application.DisplayAlerts = False
  Application.ScreenUpdating = True
 
  Dim plt As String: Dim val As Variant
  Dim wb As String: wb = ActiveWorkbook.Name
  Dim ws As String: ws = ActiveSheet.Name
 
  'Calculate
  '
 
  TSUKUBA_UTIL.ShowStatusMessage "“]‹L‚·‚é•ñ‘‚ğw’è‚µ‚Ä‚­‚¾‚³‚¢"
  Dim rep As String: rep = TSUKUBA_UTIL.WinMacSelectFile(ActiveWorkbook.path)
 
  If rep <> "" Then
 
    RESOURCE.RestAssayResult
    Dim p As Variant
    For Each p In T1.CSV2ARY(T1.ASSAY("plates")) ' :::: Plate‚ğ‚Ü‚í‚·
      TSUKUBA_UTIL.ShowStatusMessage "ƒf[ƒ^ÄŒvZ’† [" & p & "]"
      Sheets(p).Activate
      Sheets(p).EnableCalculation = True
      Sheets(p).Calculate
      RESOURCE.UpdateAssayResult CStr(p)
      Sheets(p).EnableCalculation = False
    Next
 
 
    Workbooks.Open rep
    Dim repwb As String: repwb = ActiveWorkbook.Name
    Dim colplate As Integer
 
    If TSUKUBA_UTIL.ExistSheetP(SHEETNAME_REPORT_QC_RESULT) = True Then
      TSUKUBA_UTIL.ShowStatusMessage "•ñ‘‚Ö‚Ì“]‹Lˆ—: [QCŒ‹‰Ê]"
      Dim colsb As Integer
      Dim colcvpbk As Integer
      Dim colcvpctrl As Integer
      Dim colzprime As Integer
      Dim cl As Variant
      Dim rw As Variant
 
      For Each rw In Sheets(SHEETNAME_REPORT_QC_RESULT).UsedRange.Rows
        If rw.row = 1 Then
          For Each cl In rw.Columns
            If 1 < InStr(" Plate", cl.Value) Then colplate = cl.Column
            If 1 < InStr(" S/B", cl.Value) Then colsb = cl.Column
            If 1 < InStr(" CV (%, Background)", cl.Value) Then colcvpbk = cl.Column
            If 1 < InStr(" CV (%, Control)", cl.Value) Then colcvpctrl = cl.Column
            If 1 < InStr(" Z'", cl.Value) Then colzprime = cl.Column
          Next
        Else
          With Workbooks(repwb).Sheets(SHEETNAME_REPORT_QC_RESULT)
            plt = .Cells(rw.row, colplate).Value
            val = RESOURCE.GetAssayResult(plt, "", "QC_ZPRIME")
            If val <> "" Then
              .Cells(rw.row, colzprime).Value = val
              val = RESOURCE.GetAssayResult(plt, "", "QC_SB"): If val <> "" Then .Cells(rw.row, colsb).Value = val
              val = RESOURCE.GetAssayResult(plt, "", "QC_CVPBK"): If val <> "" Then .Cells(rw.row, colcvpbk).Value = val
              val = RESOURCE.GetAssayResult(plt, "", "QC_CVPCTRL"): If val <> "" Then .Cells(rw.row, colcvpctrl).Value = val
            End If
          End With
        End If
      Next
    End If
 
    If TSUKUBA_UTIL.ExistSheetP(SHEETNAME_REPORT_ASSAY_RESULT) = True Then
      TSUKUBA_UTIL.ShowStatusMessage "•ñ‘‚Ö‚Ì“]‹Lˆ—: [ƒAƒbƒZƒCŒ‹‰Ê]"
      Dim rc As Variant
      Dim colwell As Integer
      Dim colhit As Integer
      Dim colasyname As Integer
      Dim colasyconc As Integer
      Dim colactivity As Integer
      Dim coladditional As Integer
      Dim wellpos As String
 
      For Each rw In Sheets(SHEETNAME_REPORT_ASSAY_RESULT).UsedRange.Rows
        If rw.row = 1 Then
          colplate = 0: colwell = 0: colhit = 0: colasyname = 0: colasyconc = 0: colactivity = 0
          For Each cl In rw.Columns
            If 1 < InStr(" ƒvƒŒ[ƒgID,Plate_ID", cl.Value) And colplate = 0 Then colplate = cl.Column
            If 1 < InStr(" WELL,well", cl.Value) And colwell = 0 Then colwell = cl.Column
            If 1 < InStr(" IDŠJ¦^’Ç‰ÁŠó–],req", cl.Value) And colhit = 0 Then colhit = cl.Column
            If 1 < InStr(" ƒAƒbƒZƒC–¼i—ªÌj", cl.Value) And colasyname = 0 Then colasyname = cl.Column
            If 1 < InStr(" ƒAƒbƒZƒC”Z“x(ƒÊM)", cl.Value) And colasyconc = 0 Then colasyconc = cl.Column
            If 1 < InStr(" Šˆ«’l,Œ‹‰Ê", cl.Value) And colactivity = 0 Then colactivity = cl.Column
            If 1 < InStr(" ”õl", cl.Value) And coladditional = 0 Then coladditional = cl.Column
          Next
        Else
          With Workbooks(repwb).Sheets(SHEETNAME_REPORT_ASSAY_RESULT)
            plt = .Cells(rw.row, colplate).Value
            wellpos = .Cells(rw.row, colwell).Value
            additional = T1.SYSTEM("today") & "’Ç‹L"
 
            val = RESOURCE.GetAssayResult(plt, wellpos, "CPD_RESULT")
            If val <> "" Then
              .Cells(rw.row, colactivity).Value = val
              val = RESOURCE.GetAssayResult(plt, wellpos, "CPD_HIT"): If val <> "" Then .Cells(rw.row, colhit).Value = val
              val = RESOURCE.GetAssayResult(plt, "", "TEST_ASSAY"): If val <> "" Then .Cells(rw.row, colasyname).Value = val
              val = RESOURCE.GetAssayResult(plt, wellpos, LABEL_PLATE_COMPOUND_CONC): If val <> "" Then .Cells(rw.row, colasyconc).Value = val
 
              If Trim(.Cells(rw.row, coladditional).Value) <> "" Then
                additional = additional & ", " & Trim(.Cells(rw.row, coladditional).Value)
              End If
              .Cells(rw.row, coladditional).Value = additional
            End If
          End With
        End If
      Next
    End If
 
 
    Workbooks(repwb).Save
    Workbooks(repwb).Close
  End If
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
 
  TSUKUBA_UTIL.ShowStatusMessage "•ñ‘‚Ö‚Ì“]‹Lˆ—‚ªŠ®—¹‚µ‚Ü‚µ‚½"
 
End Sub
 
Rem
Rem "2-c. “¯ˆêƒfƒBƒŒƒNƒgƒŠ“à‚Ì‘Scsvƒtƒ@ƒCƒ‹‚ğƒ}[ƒW‚·‚é"
Rem
Private Sub Action_MainMenu_Merge_All_CSV_Files()
  On Error GoTo Err_Action_MainMenu_Merge_All_CSV_Files
  T1M.action_mainmenu_merge_csv_files "cpd"
  T1M.action_mainmenu_merge_csv_files "plate"
  T1M.action_mainmenu_merge_csv_files "well"
  Exit Sub
Err_Action_MainMenu_Merge_All_CSV_Files:
        MsgBox "error"
End Sub
 
Private Sub action_mainmenu_merge_csv_files(key As String)
  Dim curpath As String: curpath = ActiveWorkbook.path & Application.PathSeparator
  Dim csvf As String
  Dim outf As String
  Dim entry As Double
  Dim lin As String
  Dim first As Boolean
 
  csvf = TSUKUBA_UTIL.WinMacDir(ActiveWorkbook.path, "CSV")
  outf = curpath & key & ".csv": Open outf For Output As #1
  entry = 0: first = True
  Do While csvf <> ""
    If 0 < InStr(csvf, "-" & key & "-") Then
      Open csvf For Input As #2: Line Input #2, lin
      If first Then
        Print #1, lin
        first = False
      End If
      Do While Not EOF(2)
        Line Input #2, lin: Print #1, lin
        entry = entry + 1
      Loop
      Close #2
    End If
    csvf = TSUKUBA_UTIL.WinMacDir()
  Loop
  Close #1
  Name outf As curpath & Format(Now(), "YYMMDD") & "-" & UCase(key) & "-" & CStr(entry) & ".csv"
 
End Sub
 
Rem
Rem GetPlateLabels(), GetWellLabels(), GetCpdLabels()
Rem
Rem GetPlateData(platename, labelname)
Rem GetWellData(platename, wellpos, labelname)
Rem GetCpdData(platename, recordpos, labelname)
Rem
 
Public Function GetPlateLabels() As Variant
  On Error Resume Next
  Application.ScreenUpdating = False
  Dim sht As Variant: Set sht = ActiveSheet
  Sheets("Template").Activate
  Dim fixed_csv As String:   fixed_csv = "PLATE_NAME,PLATE_DATAFILE,PLATE_EXCELFILE,ANALYZE_DATE,SYSTEM_VERSION"
  Dim default_csv As String: default_csv = "TEST_ASSAY,TEST_DATE,TEST_TIME,QC_ZPRIME,QC_SB,QC_CVPBK,QC_CVPCTRL,PLATE_TYPE,PLATE_FORMAT,PLATE_READER"
  Dim exist_csv As String:   exist_csv = T1M.LabelNames("exist_plate")
  GetPlateLabels = T1.CSV_OR(fixed_csv, T1.CSV_AND(default_csv, exist_csv))
  sht.Activate
  Set sht = Nothing
  Application.ScreenUpdating = True
End Function
 
Private Function GetPlateData(platename As String, labelname As String)
  Select Case labelname
    Case "PLATE_TYPE":      GetPlateData = T1.PLATE(platename, "type")
    Case "PLATE_FORMAT":    GetPlateData = T1.PLATE(platename, "format")
    Case "PLATE_READER":    GetPlateData = T1.PLATE(platename, "reader")
    Case "PLATE_NAME":      GetPlateData = T1.PLATE(platename, "name")
    Case "PLATE_DATAFILE":  GetPlateData = T1.PLATE(platename, "rawdatafile")
    Case "PLATE_EXCELFILE": GetPlateData = ThisWorkbook.Name
    Case "ANALYZE_DATE":    GetPlateData = T1.SYSTEM("today")
    Case "SYSTEM_VERSION":  GetPlateData = T1.SYSTEM("")
    Case Else:              GetPlateData = T1.PLATE(platename, labelname)
                        ' TEST_ASSAY, TEST_DATE, TEST_TIME, QC_ZPRIME, QC_SB, QC_CVPBK, QC_CVPCTRL, ƒ†[ƒU[’è‹`
  End Select
End Function
 
Public Function GetWellLabels() As Variant
  On Error Resume Next
  Application.ScreenUpdating = False
  Dim sht As Variant: Set sht = ActiveSheet
  Sheets("Template").Activate
  Dim fixed_csv As String:   fixed_csv = "WELL_POS,CPD_ID,WELL_POS0,WELL_ROW,WELL_COLUMN,WELL_ROWNUM"
  Dim default_csv As String: default_csv = "WELL_ROLE,CPD_CONC,RAW_DATA,CPD_HIT,CPD_RESULT"
  Dim exist_csv As String:   exist_csv = T1M.LabelNames("exist_well")
  GetWellLabels = T1.CSV_OR(fixed_csv, T1.CSV_OR(default_csv, exist_csv))
  sht.Activate
  Set sht = Nothing
  Application.ScreenUpdating = True
End Function
 
Private Function GetWellData(platename As String, wellpos As String, labelname As String)
  On Error Resume Next
  Select Case labelname
    Case "WELL_POS":        GetWellData = wellpos
    Case "CPD_ID":          GetWellData = RESOURCE.GetCpdID(platename, wellpos)
    Case "WELL_POS0":       GetWellData = RESOURCE.ConvertWellpos(wellpos, "pos0")
    Case "WELL_ROW":        GetWellData = RESOURCE.ConvertWellpos(wellpos, "ROW")
    Case "WELL_COLUMN":     GetWellData = RESOURCE.ConvertWellpos(wellpos, "COLUMN")
    Case "WELL_ROWNUM":     GetWellData = RESOURCE.ConvertWellpos(wellpos, "ROWNUM")
    Case "WELL_ROLE":       GetWellData = T1.well(wellpos, "role")
    Case LABEL_PLATE_COMPOUND_CONC:        GetWellData = T1.well(wellpos, "conc")
    Case Else:              GetWellData = T1.well(wellpos, labelname, "val")
                        ' RAW_DATA, WELL_ROLE, CPD_CONC, CPD_RESULT, ƒ†[ƒU[’è‹`
  End Select
End Function
 
Public Function GetCpdLabels() As Variant
  Dim cl As Variant
  Dim csv As String
  For Each cl In Sheets("Template").Range(LABEL_TABLE).Rows(1).Columns
    If cl.Value <> "" Then csv = csv & cl.Value & ","
  Next
  GetCpdLabels = Left(csv, Len(csv) - 1)
End Function
 
Private Function GetCpdData(platename As String, recordpos As Double, labelname As String)
  Dim col As Integer
        Dim cl As Variant
  For Each cl In Sheets(platename).Range(LABEL_TABLE).Rows(1).Columns
    If cl.Value = labelname Then col = cl.Column: Exit For
  Next
  GetCpdData = Sheets(platename).Cells(Range(LABEL_TABLE).row + recordpos, col).Value
  ' ƒ†[ƒU[’è‹`
End Function
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
Rem ******************************************************************************************************
Rem ƒRƒ“ƒeƒLƒXƒgƒƒjƒ…[ƒCƒxƒ“ƒg
Rem ******************************************************************************************************
 
' Œ‹‰Ê‚ğPDF‚Éo—Í‚·‚é
Private Sub Action_ContextMenu_Export_PDF()
  Worksheets(T1.CSV2ARY(T1.ASSAY("plates"))).Select
  ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=Replace(ThisWorkbook.FullName, "xlsm", "pdf")
End Sub
 
 
' ‚Í‚¸‚ê’l‚ğœŠO‚·‚éB
Public Sub Action_ContextMenu_ExcludeData(flag As String)
  Dim lb As String: lb = GetLabelOnCurPos()
  Dim rol As Range: Set rol = ActiveCell.Offset(Range("WELL_ROLE").row - Range(lb).row, Range("WELL_ROLE").Column - Range(lb).Column)
  Dim strk As Boolean
 
  Select Case flag
    Case "include": rol.Value = Replace(rol.Value, "-", ""): strk = False
    Case "exclude": rol.Value = rol.Value & "-": strk = True
  End Select
 
  TSUKUBA_UTIL.DeleteNonEffectiveNames
  RESOURCE.RestAssayResult
 
  Dim lbl As Variant
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("exist_well"))
                rol.Offset(Range(lbl).row - Range("WELL_ROLE").row, Range(lbl).Column - Range("WELL_ROLE").Column).Font.Strikethrough = strk
        Next
        Set rol = Nothing
End Sub
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
Rem ******************************************************************************************************
Rem "ƒeƒ“ƒvƒŒ[ƒg‚ğƒfƒUƒCƒ“‚·‚é"
Rem ******************************************************************************************************
 
' "Template•ÒW‚Ì‚½‚ß‰¼ƒf[ƒ^‚ğ“Ç‚İ‚Ş"
Private Sub Action_MainMenu_DataImportation_For_Template_Initialization()
        'On Error Resume Next
        Application.Volatile
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
 
        ' ‰Šúİ’è
        Dim DataSheetName As String: DataSheetName = ActiveSheet.Name
        Dim DataBookName As String: DataBookName = ActiveWorkbook.Name
        Dim OpenFileName As String: OpenFileName = TSUKUBA_UTIL.WinMacSelectFile()
        If OpenFileName = "" Then Exit Sub
 
        ' "TSUKUBA"‚ğŠÜ‚Ü‚È‚¢ƒV[ƒg‚ğ‚·‚×‚Äíœ
        Dim sht As Variant
        For Each sht In Sheets
                If InStr(sht.Name, "TSUKUBA") = 0 And sht.Name <> "Template" Then ' Visible‚ÈƒV[ƒg(Template)‚ğc‚·•K—v‚ ‚è
                        sht.Visible = -1 ' xlSheetVisible
                        sht.Delete
                End If
        Next
        TSUKUBA_UTIL.DupulicateHiddenSheetAndShow "TSUKUBA_TEMPLATE", "Template" ' ‚±‚±‚Å‹ŒTemplateƒV[ƒg‚Ííœ
 
        ' ƒf[ƒ^ƒtƒ@ƒCƒ‹“Ç‚İ‚İ
        Workbooks.Open filename:=OpenFileName
        ActiveSheet.Move Before:=Workbooks(DataBookName).Worksheets(1)
        ActiveSheet.Name = "(raw)Template"
 
        With Sheets("Template")
                .Range("5:10000").Delete
                .Range(LABEL_PLATE_TYPE).Value = "384"
                .Range(LABEL_PLATE_READER).Value = "PHERASTER"
                .Range(T1M.LABEL_PLATE_FORMAT).Value = "PRIMARY"
 
                .Activate
                TSUKUBA_UTIL.DeleteNonEffectiveNames "Template"
 
                ' Template‚ÉPullDown•t‰Á
                With .Range(LABEL_PLATE_TYPE).Validation: .Delete: .Add Type:=xlValidateList, Formula1:=SYSTEM_SUPPORT_PLATE_TYPE
                End With
                With .Range(LABEL_PLATE_FORMAT).Validation: .Delete: .Add Type:=xlValidateList, Formula1:=SYSTEM_SUPPORT_PLATE_FORMAT
                End With
                With .Range(LABEL_PLATE_READER).Validation: .Delete: .Add Type:=xlValidateList, Formula1:=SYSTEM_SUPPORT_PLATE_READER
                End With
 
                .EnableCalculation = True
                .Rows(1).Calculate
 
                T1M.Action_ContextMenu_InsertSection "end"
                T1M.InsertInfoSection "384", "PRIMARY"
                Range("A1").Select: Selection.Font.Bold = True
 
                ' temporary.xlsm ‚Æ‚µ‚Ä•Û‘¶
                Application.DisplayAlerts = False
                ThisWorkbook.SaveAs Left(OpenFileName, InStrRev(OpenFileName, Application.PathSeparator)) & "temporary.xlsm"
                'ThisWorkbook.Close
        End With
 
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
 
End Sub
 
 
 
Public Function LabelNames(label_type As String)
        On Error GoTo LABELNAMES_ERR
 
        Const required_plate = "PLATE_TYPE,PLATE_FORMAT,PLATE_READER,TEST_ASSAY,TEST_DATE,TEST_TIME,QC_ZPRIME,QC_SB,QC_CVPBK,QC_CVPCTRL"
        Const required_well = "WELL_POS,WELL_ROLE,CPD_CONC,RAW_DATA,CPD_RESULT,CPD_HIT"
        Const reserved_plate = "PLATE_NAME,PLATE_DATAFILE,PLATE_EXCELFILE,ANALYZE_DATE,SYSTEM_VERSION"
        Const reserved_well = "CPD_ID,WELL_POS_0,WELL_ROW,WELL_COL,WELL_ROWNUM"
        Const reserved_table = LABEL_TABLE
        Dim lbl As String: lbl = ActiveSheet.Name
 
        Select Case label_type
                Case "exist_plate":    LabelNames = T1.PLATE(lbl, "platelabels")
                Case "exist_well":     LabelNames = T1.PLATE(lbl, "welllabels")
                Case "exist_table":    LabelNames = T1.PLATE(lbl, "tablelabel")
                Case "all_exist":      LabelNames = T1.PLATE(lbl, "labels")
                Case "required_plate": LabelNames = required_plate
                Case "required_well":  LabelNames = required_well
                Case "reserved_plate": LabelNames = required_plate & "," & reserved_plate
                Case "reserved_well":  LabelNames = required_well & "," & reserved_well
                Case "reserved_table": LabelNames = reserved_table
                Case "all_required":   LabelNames = required_plate & "," & required_well
                Case "all_reserved":   LabelNames = required_plate & "," & reserved_plate & "," & required_well & "," & reserved_well & "," & reserved_table
                Case "user_plate":
                        LabelNames = T1.CSV_SUB(T1M.LabelNames("exist_plate"), T1M.LabelNames("reserved_plate"))
                Case "user_well":      LabelNames = T1.CSV_SUB(T1M.LabelNames("exist_well"), T1M.LabelNames("reserved_well"))
                Case Else:
                        LabelNames = False
                        Dim nm As Variant
                        For Each nm In ActiveSheet.names
                                If nm.Name = label_type Then
                                        LabelNames = True
                                        Exit Function
                                End If
                        Next
        End Select
        Exit Function
LABELNAMES_ERR:
        LabelNames = ""
 
End Function
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
Public Function GetAnalysisState()
  ' 1) ScreeningAnalysisPackage.xlsm ‚Ì Template, TSUKUBA_TEMPLATE
  ' 2) temporary.xlsm ‚Ì Template
  ' 3) (FinalTemplate).xlsm ‚Ì TemplateƒV[ƒg
  ' 4) (FinalTemplate).xlsm ‚Ì DataƒV[ƒg
  GetAnalysisState = ""
  If debug_force_to_popup_context_menu <> "" Then
    GetAnalysisState = debug_force_to_popup_context_menu
    Exit Function
  End If
 
  If ActiveWorkbook.Name = "ScreeningAnalysisPackage.xlsm" Then
    Select Case ActiveSheet.Name
      Case "Template", "TSUKUBA_TEMPLATE": GetAnalysisState = "Original@Template"
      Case Else: GetAnalysisState = "Original@" & ActiveSheet.Name
    End Select
 
  ElseIf InStr(ActiveWorkbook.Name, "temporary.xlsm") Then
    Select Case ActiveSheet.Name
      Case "Template": GetAnalysisState = "Temporary@Template"
      Case Else: GetAnalysisState = "Temporary@" & ActiveSheet.Name
    End Select
 
  Else
    Select Case ActiveSheet.Name
      Case "Template": GetAnalysisState = "Template@Template"
      Case Else:
        If 0 < InStr(T1.ASSAY("plates"), ActiveSheetName) Then
          GetAnalysisState = "Template@Data"
        Else
          GetAnalysisState = "Template@" & ActiveSheet.Name
        End If
    End Select
  End If
End Function
 
Public Function Action_WorkSheet_ShowPopupMenu()
  On Error GoTo Action_WorkSheet_ShowPopupMenu_ERR
 
  Select Case T1M.GetAnalysisState()
                Case "Original@Template": Action_WorkSheet_ShowPopupMenu = action_worksheet_showpopupmenu_originaltemplate()
                Case "Temporary@Template": Action_WorkSheet_ShowPopupMenu = action_worksheet_showpopupmenu_temporarytemplate()
                Case "Template@Template": Action_WorkSheet_ShowPopupMenu = action_worksheet_showpopupmenu_templatetemplate()
                Case "Template@Data": Action_WorkSheet_ShowPopupMenu = action_worksheet_showpopupmenu_templatedata()
                Case Else: Action_WorkSheet_ShowPopupMenu = True: Exit Function
  End Select
 
  Action_WorkSheet_ShowPopupMenu = True
 
  TSUKUBA_UTIL.DeleteNonEffectiveNames
 
  RESOURCE.InitRoleInfo CStr(ActiveSheet.Name)
  'Dim plat As Variant
  'For Each plat In T1.CSV2ARY(T1.ASSAY("plates"))
  '  RESOURCE.InitRoleInfo CStr(plat)
  'Next
 
Action_WorkSheet_ShowPopupMenu_ERR:
End Function
 
 
Public Function action_worksheet_showpopupmenu_originaltemplate()
  With Application.CommandBars("Cell")
    .reset
 
    With .Controls.Add(Before:=1, Type:=msoControlPopup): .Caption = "Šeíƒwƒ‹ƒv"
                        With .Controls.Add(): .Caption = "‰ğÍ—pŠÖ”‚Ìƒwƒ‹ƒv": .OnAction = "Action_Menu_Show_Help"
                        End With
                        With .Controls.Add(): .Caption = "ƒXƒNƒŠ[ƒjƒ“ƒO‚É‚Â‚¢‚Ä‚Ìî•ñ": .OnAction = "Action_WorkBook_OpenSite": .BeginGroup = True
                        End With
                        With .Controls.Add(): .Caption = "‰»‡•¨ƒ‰ƒCƒuƒ‰ƒŠ‚Ì’ñ‹Ÿ‚ÉŠÖ‚·‚éî•ñ": .OnAction = "Action_WorkBook_OpenCompoundDistribution"
                        End With
                        With .Controls.Add(): .Caption = "ƒAƒbƒZƒC\’z‚ÌŒŸØ€–Ú": .OnAction = "Action_WorkBook_OpenAssayValidation"
                        End With
                        With .Controls.Add(): .Caption = "ƒpƒbƒP[ƒW‚É‚Â‚¢‚Ä‚Ì¿–â": .OnAction = "Action_WorkBook_OpenMail"
                        End With
    End With
    With .Controls.Add(Before:=1): .Caption = "Še€–Ú‚ÌÄŒvZ": .OnAction = "Action_WorkSheet_Update"
    End With
 
    With .Controls.Add(Before:=1, Type:=msoControlPopup): .Caption = "Template‚ğì¬‚ğŠJn‚·‚é"
      With .Controls.Add(): .Caption = "Templateì¬‚Ì‚½‚ß‚Ì‰¼ƒf[ƒ^‚ğ“Ç‚İ‚Ş"
        .OnAction = "Action_MainMenu_DataImportation_For_Template_Initialization"
      End With
    End With
    .Controls(4).BeginGroup = True
    .ShowPopup
    .reset
  End With
End Function
 
 
 
Public Function action_worksheet_showpopupmenu_temporarytemplate()
  Dim mn As Variant
 
  With Application.CommandBars("Cell")
    .reset
    With .Controls.Add(Before:=1, Type:=msoControlPopup): .Caption = "Šeíƒwƒ‹ƒv"
                        With .Controls.Add(): .Caption = "‰ğÍ—pŠÖ”‚Ìƒwƒ‹ƒv": .OnAction = "Action_Menu_Show_Help"
                        End With
                        With .Controls.Add(): .Caption = "ƒXƒNƒŠ[ƒjƒ“ƒO‚É‚Â‚¢‚Ä‚Ìî•ñ": .OnAction = "Action_WorkBook_OpenSite": .BeginGroup = True
                        End With
                        With .Controls.Add(): .Caption = "‰»‡•¨ƒ‰ƒCƒuƒ‰ƒŠ‚Ì’ñ‹Ÿ‚ÉŠÖ‚·‚éî•ñ": .OnAction = "Action_WorkBook_OpenCompoundDistribution"
                        End With
                        With .Controls.Add(): .Caption = "ƒAƒbƒZƒC\’z‚ÌŒŸØ€–Ú": .OnAction = "Action_WorkBook_OpenAssayValidation"
                        End With
                        With .Controls.Add(): .Caption = "ƒpƒbƒP[ƒW‚É‚Â‚¢‚Ä‚Ì¿–â": .OnAction = "Action_WorkBook_OpenMail"
                        End With
    End With
    With .Controls.Add(Before:=1): .Caption = "Še€–Ú‚ÌÄŒvZ": .OnAction = "Action_WorkSheet_Update"
    End With
 
 
    ' ƒ‰ƒxƒ‹ :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    With .Controls.Add(Type:=msoControlPopup, Before:=1): .Caption = "ƒf[ƒ^ƒ‰ƒxƒ‹İ’è"
      ' ƒvƒŒ[ƒg—p ========================================================================================
      With .Controls.Add(Type:=msoControlPopup): .Caption = "ƒvƒŒ[ƒgƒ‰ƒxƒ‹İ’è (1)"
        ' „§ƒvƒŒ[ƒgƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("required_plate"))
          If TSUKUBA_UTIL.ExistNameP(ActiveSheet.Name, CStr(lbl)) Then
            With .Controls.Add(Type:=msoControlPopup): .Caption = "* " & lbl
              With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
              End With
            End With
          Else
            With .Controls.Add(): .Caption = "* " & lbl & " (—v“o˜^)": .OnAction = "'T1M.Action_ContextMenu_CreatePlateLabel """ & lbl & """'"
            End With
          End If
        Next
        ' ƒ†[ƒU[’è‹`ƒvƒŒ[ƒgƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("user_plate"))
          With .Controls.Add(Type:=msoControlPopup): .Caption = lbl
            With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
            End With
          End With
        Next
        With .Controls.Add(): .Caption = "ƒ†[ƒU[İ’è": .OnAction = "'T1M.Action_ContextMenu_CreatePlateLabel """"'": .BeginGroup = True
        End With
      End With
 
      ' ƒEƒFƒ‹—p ========================================================================================
      With .Controls.Add(Type:=msoControlPopup): .Caption = "ƒEƒFƒ‹ƒ‰ƒxƒ‹‚Ìİ’è (" & T1.PLATE("Template", "type") & ")": .BeginGroup = True:
        ' „§ƒEƒFƒ‹ƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("required_well"))
          If TSUKUBA_UTIL.ExistNameP(ActiveSheet.Name, CStr(lbl)) Then
            With .Controls.Add(Type:=msoControlPopup): .Caption = "* " & lbl
              With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
              End With
            End With
          Else
            With .Controls.Add(): .Caption = "* " & lbl & " (—v“o˜^)": .OnAction = "'T1M.Action_ContextMenu_CreateWellLabel """ & lbl & """'"
            End With
          End If
        Next
        ' ƒ†[ƒU[’è‹`ƒEƒFƒ‹ƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("user_well"))
          With .Controls.Add(Type:=msoControlPopup): .Caption = lbl
            With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
            End With
          End With
        Next
        With .Controls.Add(): .Caption = "ƒ†[ƒU[İ’è": .OnAction = " 'T1M.Action_ContextMenu_CreateWellLabel """"'": .BeginGroup = True
        End With
      End With
 
      ' ƒe[ƒuƒ‹—p ========================================================================================
      If TSUKUBA_UTIL.ExistNameP(ActiveSheet.Name, T1.TABLE("name")) Then
        With .Controls.Add(Type:=msoControlPopup)
          .Caption = "‰»‡•¨ƒe[ƒuƒ‹ƒ‰ƒxƒ‹İ’è (" & CStr(UBound(T1.CSV2ARY(T1.TABLE("items"))) + 1) & "x" & T1.TABLE("records") & ")"
          .BeginGroup = True
          With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & T1.TABLE("name") & """'"
          End With
          With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & T1.TABLE("name") & """'"
          End With
        End With
        For Each lbl In T1.CSV2ARY(T1.TABLE("items"))
          With .Controls.Add(): .Caption = lbl: .Enabled = False
          End With
        Next
      Else
        With .Controls.Add(Type:=msoControlPopup)
          .Caption = "ƒe[ƒuƒ‹ƒ‰ƒxƒ‹‚Ìİ’è"
          .BeginGroup = True
          With .Controls.Add(): .Caption = "İ’è": .OnAction = "'T1M.Action_ContextMenu_CreateTableLabel """ & T1.TABLE("name") & """'"
          End With
        End With
      End If
 
      ' ‰»‡•¨ƒe[ƒuƒ‹—p ========================================================================================
      With .Controls.Add():
        .Caption = "‰»‡•¨ƒe[ƒuƒ‹‚ğ“Ç‚İ‚Ş"
        .OnAction = "RESOURCE.LoadCompoundTable"
        .BeginGroup = True
      End With
 
    End With
 
    ' ƒZƒNƒVƒ‡ƒ“ :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    With .Controls.Add(Before:=1, Type:=msoControlPopup): .Caption = "ƒZƒNƒVƒ‡ƒ“İ’è"
      With .Controls.Add(Type:=msoControlPopup): .Caption = "V‹K"
        If 0 < InStr(SYSTEM_SUPPORT_REALTIME_PLATE_READER, Range(LABEL_PLATE_READER)) Then
          With .Controls.Add(): .Caption = "ƒf[ƒ^•\¦(ŠÔw’è)": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""data2""'"
          End With
          With .Controls.Add(): .Caption = "ƒf[ƒ^•\¦(ŠÔ”ÍˆÍw’è)": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""data4""'"
          End With
        Else
          With .Controls.Add(): .Caption = "ƒf[ƒ^•\¦": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""data1""'"
          End With
        End If
        With .Controls.Add(): .Caption = "‰ğÍ1:ƒf[ƒ^‰ğÍ": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""anal1""'"
        End With
        With .Controls.Add(): .Caption = "‰ğÍ2:ƒqƒbƒg”»’è": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""anal2""'"
        End With
 
        If T1M.SECTION(ActiveCell, "color") = DATA_SECTION_THEME_COLOR Or _
           T1M.SECTION(ActiveCell, "color") = ANAL_SECTION_THEME_COLOR Then
          With .Controls.Add(): .Caption = "U•z}1 (" & T1M.SECTION(ActiveCell, "current") & ")": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""graph""'"
          End With
          With .Controls.Add(): .Caption = "U•z}2 (" & T1M.SECTION(ActiveCell, "current") & ")": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""graph2""'"
          End With
        End If
        If Not TSUKUBA_UTIL.ExistNameP(ActiveSheet.Name, T1.TABLE("name")) Then
          With .Controls.Add(): .Caption = "Œ‹‰Êƒe[ƒuƒ‹": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""table""'"
          End With
        End If
        Dim csv As String: csv = T1M.GetExtraSections()
        If csv <> "" Then
          With .Controls.Add(Type:=msoControlPopup): .Caption = "‚»‚Ì‘¼"
            For Each mn In T1.CSV2ARY(csv)
              With .Controls.Add(): .Caption = CStr(mn): .OnAction = "'T1M.Action_ContextMenu_InsertSection """ & CStr(mn) & """'"
              End With
            Next
          End With
        End If
      End With
 
      With .Controls.Add(): .Caption = "íœ (" & T1M.SECTION(ActiveCell, "current") & ")": .OnAction = "T1M.Action_ContextMenu_DeleteCurrentSection"
      End With
      With .Controls.Add(): .Caption = "‘S•\¦": .OnAction = "T1M.Action_ContextMenu_ShowAllSection"
      End With
      With .Controls.Add(): .Caption = "‘S”ñ•\¦": .OnAction = "T1M.Action_ContextMenu_HideAllSection"
      End With
 
    End With
 
    ' ƒvƒŒ[ƒg :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    With .Controls.Add(Type:=msoControlPopup, Before:=1): .Caption = "ƒvƒŒ[ƒgİ’è"
      With .Controls.Add(Type:=msoControlPopup): .Caption = "ƒvƒŒ[ƒgƒ^ƒCƒv: " & Range(LABEL_PLATE_TYPE)
        For Each mn In T1.CSV2ARY(SYSTEM_SUPPORT_PLATE_TYPE)
          With .Controls.Add(): .Caption = CStr(mn): .OnAction = "'T1M.Action_ContextMenu_UpdatePlateProperty """ & mn & """'"
            If CStr(mn) = Range(LABEL_PLATE_TYPE) Then .State = True
            If T1.FIND_ROW(Range("TSUKUBA_TEMPLATE!A:A"), "WELL_ROLE", CStr(mn), Range(LABEL_PLATE_FORMAT)) = 0 Then .Enabled = False
          End With
        Next
      End With
      With .Controls.Add(Type:=msoControlPopup): .Caption = "ƒvƒŒ[ƒgƒtƒH[ƒ}ƒbƒg: " & Range(LABEL_PLATE_FORMAT)
        For Each mn In T1.CSV2ARY(SYSTEM_SUPPORT_PLATE_FORMAT)
          With .Controls.Add(): .Caption = CStr(mn): .OnAction = "'T1M.Action_ContextMenu_UpdatePlateProperty """ & mn & """'"
            If CStr(mn) = Range(LABEL_PLATE_FORMAT) Then .State = True
            If T1.FIND_ROW(Range("TSUKUBA_TEMPLATE!A:A"), "WELL_ROLE", Range(LABEL_PLATE_TYPE), CStr(mn)) = 0 Then .Enabled = False
          End With
        Next
      End With
      With .Controls.Add(Type:=msoControlPopup): .Caption = "ƒvƒŒ[ƒgƒŠ[ƒ_[: " & Range(LABEL_PLATE_READER)
        For Each mn In T1.CSV2ARY(SYSTEM_SUPPORT_PLATE_READER)
          With .Controls.Add(): .Caption = CStr(mn): .OnAction = "'T1M.Action_ContextMenu_UpdatePlateProperty """ & mn & """'"
            If CStr(mn) = Range(LABEL_PLATE_READER) Then .State = True
          End With
        Next
      End With
      With .Controls.Add(): .Caption = "ƒeƒ“ƒvƒŒ[ƒg‰»": .OnAction = "Action_ContextMenu_SaveAsTemplate": .BeginGroup = True
      End With
    End With
    .Controls(6).BeginGroup = True
    .ShowPopup
    .reset
        End With
End Function
 
Public Function action_worksheet_showpopupmenu_templatetemplate()
  With Application.CommandBars("Cell")
    .reset
 
    With .Controls.Add(Before:=1, Type:=msoControlPopup): .Caption = "Šeíƒwƒ‹ƒv"
                        With .Controls.Add(): .Caption = "‰ğÍ—pŠÖ”‚Ìƒwƒ‹ƒv": .OnAction = "Action_Menu_Show_Help"
                        End With
                        With .Controls.Add(): .Caption = "ƒXƒNƒŠ[ƒjƒ“ƒO‚É‚Â‚¢‚Ä‚Ìî•ñ": .OnAction = "Action_WorkBook_OpenSite": .BeginGroup = True
                        End With
                        With .Controls.Add(): .Caption = "‰»‡•¨ƒ‰ƒCƒuƒ‰ƒŠ‚Ì’ñ‹Ÿ‚ÉŠÖ‚·‚éî•ñ": .OnAction = "Action_WorkBook_OpenCompoundDistribution"
                        End With
                        With .Controls.Add(): .Caption = "ƒAƒbƒZƒC\’z‚ÌŒŸØ€–Ú": .OnAction = "Action_WorkBook_OpenAssayValidation"
                        End With
                        With .Controls.Add(): .Caption = "ƒpƒbƒP[ƒW‚É‚Â‚¢‚Ä‚Ì¿–â": .OnAction = "Action_WorkBook_OpenMail"
                        End With
    End With
    With .Controls.Add(Before:=1): .Caption = "Še€–Ú‚ÌÄŒvZ": .OnAction = "Action_WorkSheet_Update"
    End With
 
 
    ' ƒ‰ƒxƒ‹ :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    With .Controls.Add(Type:=msoControlPopup, Before:=1): .Caption = "ƒf[ƒ^ƒ‰ƒxƒ‹İ’è"
      ' ƒvƒŒ[ƒg—p ========================================================================================
      With .Controls.Add(Type:=msoControlPopup): .Caption = "ƒvƒŒ[ƒgƒ‰ƒxƒ‹İ’è (1)"
        ' „§ƒvƒŒ[ƒgƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("required_plate"))
          If TSUKUBA_UTIL.ExistNameP(ActiveSheet.Name, CStr(lbl)) Then
            With .Controls.Add(Type:=msoControlPopup): .Caption = "* " & lbl
              With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
              End With
            End With
          Else
            With .Controls.Add(): .Caption = "* " & lbl & " (—v“o˜^)": .OnAction = "'T1M.Action_ContextMenu_CreatePlateLabel """ & lbl & """'"
            End With
          End If
        Next
        ' ƒ†[ƒU[’è‹`ƒvƒŒ[ƒgƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("user_plate"))
          With .Controls.Add(Type:=msoControlPopup): .Caption = lbl
            With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
            End With
          End With
        Next
        With .Controls.Add(): .Caption = "ƒ†[ƒU[İ’è": .OnAction = "'T1M.Action_ContextMenu_CreatePlateLabel """"'": .BeginGroup = True
        End With
      End With
 
      ' ƒEƒFƒ‹—p ========================================================================================
      With .Controls.Add(Type:=msoControlPopup): .Caption = "ƒEƒFƒ‹ƒ‰ƒxƒ‹‚Ìİ’è (" & T1.PLATE("Template", "type") & ")": .BeginGroup = True:
        ' „§ƒEƒFƒ‹ƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("required_well"))
          If TSUKUBA_UTIL.ExistNameP(ActiveSheet.Name, CStr(lbl)) Then
            With .Controls.Add(Type:=msoControlPopup): .Caption = "* " & lbl
              With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
              End With
              With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
              End With
            End With
          Else
            With .Controls.Add(): .Caption = "* " & lbl & " (—v“o˜^)": .OnAction = "'T1M.Action_ContextMenu_CreateWellLabel """ & lbl & """'"
            End With
          End If
        Next
        ' ƒ†[ƒU[’è‹`ƒEƒFƒ‹ƒ‰ƒxƒ‹
        For Each lbl In T1.CSV2ARY(T1M.LabelNames("user_well"))
          With .Controls.Add(Type:=msoControlPopup): .Caption = lbl
            With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "–¼‘O•ÏX": .OnAction = "'T1M.Action_ContextMenu_ChangeLabelName """ & lbl & """'"
            End With
            With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & lbl & """'"
            End With
          End With
        Next
        With .Controls.Add(): .Caption = "ƒ†[ƒU[İ’è": .OnAction = "'T1M.Action_ContextMenu_CreateWellLabel """"'": .BeginGroup = True
        End With
      End With
 
      ' ƒe[ƒuƒ‹—p ========================================================================================
      If TSUKUBA_UTIL.ExistNameP(ActiveSheet.Name, T1.TABLE("name")) Then
        With .Controls.Add(Type:=msoControlPopup)
          .Caption = "‰»‡•¨ƒe[ƒuƒ‹ƒ‰ƒxƒ‹İ’è (" & CStr(UBound(T1.CSV2ARY(T1.TABLE("items"))) + 1) & "x" & T1.TABLE("records") & ")"
          .BeginGroup = True
          With .Controls.Add(): .Caption = "‘I‘ğ": .OnAction = "'T1M.Action_ContextMenu_SelectLabel """ & T1.TABLE("name") & """'"
          End With
          With .Controls.Add(): .Caption = "íœ": .OnAction = "'T1M.Action_ContextMenu_DeleteLabel """ & T1.TABLE("name") & """'"
          End With
        End With
        For Each lbl In T1.CSV2ARY(T1.TABLE("items"))
          With .Controls.Add(): .Caption = lbl: .Enabled = False
          End With
        Next
 
      Else
        With .Controls.Add(Type:=msoControlPopup)
          .Caption = "ƒe[ƒuƒ‹ƒ‰ƒxƒ‹‚Ìİ’è"
          .BeginGroup = True
          With .Controls.Add(): .Caption = "İ’è": .OnAction = "'T1M.Action_ContextMenu_CreateTableLabel """ & T1.TABLE("name") & """'"
          End With
          With .Controls.Add(): .Caption = "V‹K": .OnAction = "'T1M.Action_ContextMenu_InsertSection ""table""'"
          End With
        End With
      End If
 
      ' ‰»‡•¨ƒe[ƒuƒ‹—p ========================================================================================
      With .Controls.Add():
        .Caption = "‰»‡•¨ƒe[ƒuƒ‹‚ğ“Ç‚İ‚Ş"
        .OnAction = "RESOURCE.LoadCompoundTable"
        .BeginGroup = True
      End With
 
    End With
 
    ' ƒZƒNƒVƒ‡ƒ“ :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    With .Controls.Add(Before:=1, Type:=msoControlPopup): .Caption = "ƒZƒNƒVƒ‡ƒ“İ’è"
      With .Controls.Add(): .Caption = "íœ (" & T1M.SECTION(ActiveCell, "current") & ")": .OnAction = "T1M.Action_ContextMenu_DeleteCurrentSection"
      End With
      With .Controls.Add(): .Caption = "‘S•\¦": .OnAction = "T1M.Action_ContextMenu_ShowAllSection"
      End With
      With .Controls.Add(): .Caption = "‘S”ñ•\¦": .OnAction = "T1M.Action_ContextMenu_HideAllSection"
      End With
    End With
 
    .Controls(5).BeginGroup = True
    .ShowPopup
    .reset
  End With
End Function
 
 
Public Function action_worksheet_showpopupmenu_templatedata()
  With Application.CommandBars("Cell")
    .reset
 
    With .Controls.Add(Before:=1, Type:=msoControlPopup): .Caption = "Šeíƒwƒ‹ƒv"
                        With .Controls.Add(): .Caption = "‰ğÍ—pŠÖ”‚Ìƒwƒ‹ƒv": .OnAction = "Action_Menu_Show_Help"
                        End With
                        With .Controls.Add(): .Caption = "ƒXƒNƒŠ[ƒjƒ“ƒO‚É‚Â‚¢‚Ä‚Ìî•ñ": .OnAction = "Action_WorkBook_OpenSite": .BeginGroup = True
                        End With
                        With .Controls.Add(): .Caption = "‰»‡•¨ƒ‰ƒCƒuƒ‰ƒŠ‚Ì’ñ‹Ÿ‚ÉŠÖ‚·‚éî•ñ": .OnAction = "Action_WorkBook_OpenCompoundDistribution"
                        End With
                        With .Controls.Add(): .Caption = "ƒAƒbƒZƒC\’z‚ÌŒŸØ€–Ú": .OnAction = "Action_WorkBook_OpenAssayValidation"
                        End With
                        With .Controls.Add(): .Caption = "ƒpƒbƒP[ƒW‚É‚Â‚¢‚Ä‚Ì¿–â": .OnAction = "Action_WorkBook_OpenMail"
                        End With
    End With
    With .Controls.Add(Before:=1): .Caption = "Še€–Ú‚ÌÄŒvZ": .OnAction = "Action_WorkSheet_Update"
    End With
 
 
    With .Controls.Add(Type:=msoControlPopup, Before:=1): .Caption = "ƒf[ƒ^‰ğÍ"
      If InStr(T1M.LabelNames("exist_well"), T1M.GetLabelOnCurPos()) Then
        If T1M.ExcludedWellP() Then
          With .Controls.Add(): .Caption = "œŠO’l‚ğ‰ñ•œ‚·‚é": .OnAction = "'T1M.Action_ContextMenu_ExcludeData ""include""'"
          End With
        Else
          With .Controls.Add(): .Caption = "‚Í‚¸‚ê’l‚ğœŠO‚·‚é": .OnAction = "'T1M.Action_ContextMenu_ExcludeData ""exclude""'"
          End With
        End If
      End If
    End With
    .Controls(4).BeginGroup = True
    .ShowPopup
    .reset
  End With
End Function
 
 
Public Sub Action_ContextMenu_InsertSection(param As String)
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Dim rw, cl As Integer: rw = ActiveCell.row: cl = ActiveCell.Column
 
        If ActiveSheet.Name <> "Template" Then MsgBox "Template‚Å‚È‚¢I": Exit Sub
        Dim typ As String: typ = Range(LABEL_PLATE_TYPE)
        Dim fmt As String: fmt = Range(LABEL_PLATE_FORMAT)
        Dim red As String: red = Range(LABEL_PLATE_READER)
        Select Case param
                Case "info":  T1M.InsertInfoSection typ, fmt
                Case "data1": T1M.InsertDataSection typ, "1PARAM"
                Case "data2": T1M.InsertDataSection typ, "2PARAM"
                Case "data4": T1M.InsertDataSection typ, "4PARAM"
                Case "anal1": T1M.InsertAnalSection typ, "CPD_RESULT"
                Case "anal2": T1M.InsertAnalSection typ, "CPD_HIT"
                Case "graph": T1M.InsertGraphSection typ, "DOT"
                Case "graph2": T1M.InsertGraphSection2 typ, "DOT"
                Case "table": T1M.InsertTableSection typ, fmt
                Case "end":   T1M.InsertEndSection
                Case Else:    T1M.InsertExtraSection typ, fmt, param
        End Select
 
        Cells(rw, cl).Select
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
 
End Sub
 
Public Function GetExtraSections() As String
  On Error Resume Next
  Dim typ As String: typ = Range(LABEL_PLATE_TYPE)
  Dim fmt As String: fmt = Range(LABEL_PLATE_FORMAT)
  Dim red As String: red = Range(LABEL_PLATE_READER)
  Dim csv As String: csv = ""
  Dim sht As String
  sht = ActiveSheet.Name
  Sheets("TSUKUBA_TEMPLATE").Activate
  Sheets("TSUKUBA_TEMPLATE").Cells(1, 1).Calculate
  Dim c As Range
  For Each c In Sheets("TSUKUBA_TEMPLATE").UsedRange.Columns(1).Rows
    If c.Interior.ThemeColor = EXTR_SECTION_THEME_COLOR Then
      If InStr(c.Value, typ) And InStr(c.Value, fmt) Then
        csv = csv & Mid(c.Value, InStr(c.Value, ">") + 1) & ","
      End If
    End If
    If 10000 < c.row Or c.Value = "END" Then Exit For
  Next
  If csv <> "" Then csv = Left(csv, Len(csv) - 1)
  Sheets(sht).Activate
  GetExtraSections = csv
End Function
 
 
Private Sub InsertGraphSection(typ As String, param As String)
  On Error Resume Next
  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
  Dim rw As Integer: rw = T1M.SECTION(ActiveCell, "end") + 1
 
  Dim sect As String
  If 1 < rw Then
    Select Case T1M.SECTION(ActiveCell, "color")
      Case DATA_SECTION_THEME_COLOR: sect = "RAW_DATA"
      Case ANAL_SECTION_THEME_COLOR: sect = "CPD_RESULT"
    End Select
 
    Rows(T1M.SECTION(ActiveCell, "inrows")).Hidden = False
 
    Dim LABEL As String: LABEL = T1M.SECTION(ActiveCell, "current")
 
    If CopySection(sect, "GRAPH", "") Then ' ƒOƒ‰ƒt‚ğƒRƒs[
      Rows(rw).Insert Shift:=xlDown
      Rows(T1M.SECTION(Rows(rw), "inrows")).Hidden = False
      Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = LABEL
 
      Dim data_rng As Range
      Set data_rng = Range(LABEL)
      Set data_rng = data_rng.Resize(data_rng.Rows.COUNT + 1, data_rng.Columns.COUNT + 1).Offset(-1, -1)
      Dim grp_rng As Range
      Dim itm As Variant
 
      If T1.SYSTEM("pc") = "Mac" Or T1.SYSTEM("excelver") <= 14 Then
        Set grp_rng = Range(Cells(rw + 2, 2), Cells(T1M.SECTION(Rows(rw), "end") - 2, 26))
 
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlXYScatter
 
        'ActiveChart.SetSourceData Source:=Range(data_rng)
        ActiveChart.ChartArea.Left = grp_rng.Left
        ActiveChart.ChartArea.Top = grp_rng.Top
        ActiveChart.ChartArea.Width = grp_rng.Width
        ActiveChart.ChartArea.Height = grp_rng.Height
        ActiveChart.PlotBy = xlColumns
        ActiveChart.ApplyLayout (3)
        ActiveChart.ChartTitle.Delete
 
        For Each itm In ActiveChart.SeriesCollection
          itm.MarkerStyle = 8
          itm.MarkerSize = 8
          itm.Format.Line.Visible = msoFalse
        Next
      Else
        Set grp_rng = Range(Cells(rw + 2, 2), Cells(T1M.SECTION(Rows(rw), "end") - 1, 26))
 
        ActiveSheet.Shapes.AddChart2(240, xlXYScatter, grp_rng.Left, grp_rng.Top, grp_rng.Width, grp_rng.Height).Select
        Do While 0 < ActiveChart.SeriesCollection.COUNT
          ActiveChart.SeriesCollection(1).Delete
        Loop
 
        Set data_rng = data_rng.Resize(data_rng.Rows.COUNT - 1, data_rng.Columns.COUNT - 1).Offset(1, 1)
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection(1).Name = "=""" & LABEL & """"
        Dim csv As String: csv = ""
        Dim r As Variant
        For Each r In data_rng.Rows
          csv = csv & ActiveSheet.Name & "!" & r.Address & ","
        Next
        ActiveChart.FullSeriesCollection(1).Values = "=(" & Left(csv, Len(csv) - 1) & ")"
        ActiveChart.Axes(xlCategory).MinimumScale = 0
        ActiveChart.Axes(xlCategory).MaximumScale = Application.WorksheetFunction.Ceiling(data_rng.COUNT, 10)
 
        csv = ""
        For Each r In Range(LABEL_PLATE_WELL_POSITION).Rows
          csv = csv & ActiveSheet.Name & "!" & r.Address & ","
        Next
        ActiveChart.FullSeriesCollection(1).XValues = "=(" & Left(csv, Len(csv) - 1) & ")"
 
        ActiveChart.Axes(xlCategory).Select
        Selection.TickLabelPosition = xlLow
        ActiveChart.Legend.Delete
        ActiveChart.ChartTitle.Delete
 
        Dim i As Integer
        Dim xval As Variant: xval = ActiveChart.FullSeriesCollection(1).XValues
 
        For i = 0 To ActiveChart.FullSeriesCollection(1).Points.COUNT
          With ActiveChart.FullSeriesCollection(1).Points(i)
            .MarkerStyle = 8
            .MarkerSize = 7
            .Format.Line.Visible = msoFalse
            If 0 < InStr(T1.well(CStr(xval(i)), "role"), "MIN") Then
              .Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
            ElseIf 0 < InStr(T1.well(CStr(xval(i)), "role"), "MAX") Then
              .Format.Fill.ForeColor.RGB = RGB(200, 0, 0)
            ElseIf 0 < InStr(T1.well(CStr(xval(i)), "role"), "POS") Then
              .Format.Fill.ForeColor.RGB = RGB(0, 200, 0)
            ElseIf 0 < InStr(T1.well(CStr(xval(i)), "role"), "NEG") Then
              .Format.Fill.ForeColor.RGB = RGB(100, 100, 100)
            ElseIf 0 < InStr(T1.well(CStr(xval(i)), "role"), "REF") Then
              .Format.Fill.ForeColor.RGB = RGB(150, 150, 0)
            ElseIf 0 < InStr(T1.well(CStr(xval(i)), "role"), "BLANK") Then
              .Format.Fill.Visible = msoFalse
            End If
          End With
        Next i
      End If
    End If
  End If
  Set data_rng = Nothing
  Set grp_rng = Nothing
 
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
End Sub
 
Private Sub InsertGraphSection2(typ As String, param As String)
  On Error Resume Next
  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
  Dim rw As Integer: rw = T1M.SECTION(ActiveCell, "end") + 1
 
  Dim sect As String
 
  If 1 < rw Then
    Select Case T1M.SECTION(ActiveCell, "color")
      Case DATA_SECTION_THEME_COLOR: sect = "RAW_DATA"
      Case ANAL_SECTION_THEME_COLOR: sect = "CPD_RESULT"
    End Select
 
    Rows(T1M.SECTION(ActiveCell, "inrows")).Hidden = False
 
    Dim LABEL As String: LABEL = T1M.SECTION(ActiveCell, "current")
 
    If CopySection(sect, "GRAPH", "") Then ' ƒOƒ‰ƒt‚ğƒRƒs[
      Rows(rw).Insert Shift:=xlDown
      Rows(T1M.SECTION(Rows(rw), "inrows")).Hidden = False
      Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = LABEL
 
      Dim data_rng As Range
      Set data_rng = Range(LABEL)
      Set data_rng = data_rng.Resize(data_rng.Rows.COUNT + 1, data_rng.Columns.COUNT + 1).Offset(-1, -1)
      Dim grp_rng As Range
      Dim itm As Variant
 
      If T1.SYSTEM("pc") = "Mac" Or T1.SYSTEM("excelver") <= 14 Then
        Set grp_rng = Range(Cells(rw + 2, 2), Cells(T1M.SECTION(Rows(rw), "end") - 2, 26))
 
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlLineMarkers
 
        ActiveChart.SetSourceData Source:=Range(data_rng)
        ActiveChart.ChartArea.Left = grp_rng.Left
        ActiveChart.ChartArea.Top = grp_rng.Top
        ActiveChart.ChartArea.Width = grp_rng.Width
        ActiveChart.ChartArea.Height = grp_rng.Height
        ActiveChart.PlotBy = xlColumns
        ActiveChart.ApplyLayout (3)
        ActiveChart.ChartTitle.Delete
        For Each itm In ActiveChart.SeriesCollection
          itm.MarkerStyle = 8
          itm.MarkerSize = 8
          itm.Format.Line.Visible = msoFalse
        Next
      Else
        Set grp_rng = Range(Cells(rw + 2, 2), Cells(T1M.SECTION(Rows(rw), "end") - 1, 26))
        ActiveSheet.Shapes.AddChart2(332, xlLineMarkers, grp_rng.Left, grp_rng.Top, grp_rng.Width, grp_rng.Height).Select
        ActiveChart.SetSourceData Source:=Range(data_rng.Address)
        ActiveChart.ChartTitle.Delete
        'For Each item In ActiveChart.FullSeriesCollection
        For Each itm In ActiveChart.SeriesCollection
                                        itm.Format.Line.Visible = msoFalse
        Next
 
 
      End If
    End If
  End If
  Set data_rng = Nothing
  Set grp_rng = Nothing
 
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
End Sub
 
Private Sub InsertExtraSection(typ As String, fmt As String, param As String)
  Dim rw As Integer: rw = 0
  Dim c As Variant
 
  For Each c In Range("A:A")
                If c.Interior.ThemeColor = DATA_SECTION_THEME_COLOR Or _
                         c.Interior.ThemeColor = INFO_SECTION_THEME_COLOR Or _
                         c.Interior.ThemeColor = ANAL_SECTION_THEME_COLOR Or _
                         c.Interior.ThemeColor = EXTR_SECTION_THEME_COLOR Or _
                         c.Interior.ThemeColor = TBLE_SECTION_THEME_COLOR Then
                        rw = c.row
                End If
                If 10000 < c.row Or c.Value = "END" Then Exit For
  Next
 
  If 0 < rw Then
    Dim LABEL As String: LABEL = param
    rw = T1M.SECTION(Rows(rw), "end") + 1
 
    If CopySection(fmt, typ, param) Then
      Rows(rw).Insert Shift:=xlDown
      Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
      Rows(T1M.SECTION(Rows(rw), "rows")).Select
      TSUKUBA_UTIL.ConvertSelectionFomulaFromRelatioveToAbsolute
      Selection.Replace What:=param, Replacement:=LABEL
      Cells(rw + 3, 3).Select: Action_ContextMenu_CreateWellLabel LABEL
      T1M.ShowCurrentSection
    End If
  End If
 
End Sub
 
 
Private Sub InsertAnalSection(typ As String, param As String)
  Dim rw As Integer: rw = 0
  Dim c As Variant
  For Each c In Range("A:A")
                If c.Interior.ThemeColor = DATA_SECTION_THEME_COLOR Or _
                         c.Interior.ThemeColor = INFO_SECTION_THEME_COLOR Or _
                         c.Interior.ThemeColor = ANAL_SECTION_THEME_COLOR Then
                        rw = c.row
                End If
                If 10000 < c.row Or c.Value = "END" Then Exit For
  Next
 
  If 0 < rw Then
    Dim LABEL As String: LABEL = param
    rw = T1M.SECTION(Rows(rw), "end") + 1
    Do While LABEL <> "" And ExistNameP("Template", LABEL)
      LABEL = InputBox("[" & LABEL & "] ‚Æ‚ÍˆÙ‚È‚é–¼‘O‚ğ“ü—Í", "ƒ‰ƒxƒ‹–¼‚ğ“ü—Í", LABEL)
    Loop: If LABEL = "" Then Exit Sub
 
    If CopySection(param, typ, "") Then ' ƒvƒŒ[ƒgƒtƒH[ƒ}ƒbƒg‚ğƒRƒs[
      Rows(rw).Insert Shift:=xlDown
      Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
      Rows(T1M.SECTION(Rows(rw), "rows")).Select
      TSUKUBA_UTIL.ConvertSelectionFomulaFromRelatioveToAbsolute
      Selection.Replace What:=param, Replacement:=LABEL
      Cells(rw + 3, 3).Select: Action_ContextMenu_CreateWellLabel LABEL
      T1M.ShowCurrentSection
    End If
  End If
 
End Sub
 
 
Private Sub InsertDataSection(typ As String, param As String)
  Dim rw As Integer: rw = 0
  Dim c As Variant
  For Each c In Range("A:A")
                If c.Interior.ThemeColor = DATA_SECTION_THEME_COLOR Or _
                         c.Interior.ThemeColor = INFO_SECTION_THEME_COLOR Then
                        rw = c.row
                End If
                If 10000 < c.row Or c.Value = "END" Then Exit For
  Next
 
  If 0 < rw Then
    Dim LABEL As String: LABEL = "RAW_DATA"
    rw = T1M.SECTION(Rows(rw), "end") + 1
    Do While LABEL <> "" And ExistNameP("Template", LABEL)
      LABEL = InputBox("[" & LABEL & "] ‚Æ‚ÍˆÙ‚È‚é–¼‘O‚ğ“ü—Í", "ƒ‰ƒxƒ‹–¼‚ğ“ü—Í", LABEL)
    Loop: If LABEL = "" Then Exit Sub
 
    If CopySection("RAW_DATA", typ, param) Then ' ƒvƒŒ[ƒgƒtƒH[ƒ}ƒbƒg‚ğƒRƒs[
      Rows(rw).Insert Shift:=xlDown
      Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
      Rows(T1M.SECTION(Rows(rw), "rows")).Select
      TSUKUBA_UTIL.ConvertSelectionFomulaFromRelatioveToAbsolute
      Selection.Replace What:="RAW_DATA", Replacement:=LABEL
      Cells(rw + 3, 3).Select: Action_ContextMenu_CreateWellLabel LABEL
      T1M.ShowCurrentSection
    End If
  End If
 
End Sub
 
Private Sub InsertInfoSection(typ As String, fmt As String)
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
 
        Dim p As Boolean
        Dim r As Boolean
        Dim c As Boolean
        Dim rw As Double
 
  rw = T1.FIND_ROW(ActiveSheet.Columns(1), "WELL_POS"): p = Rows(rw + 1).Hidden
  If 0 < rw Then: Rows(T1M.SECTION(ActiveSheet.Rows(rw), "rows")).Delete
  TSUKUBA_UTIL.DeleteNonEffectiveNames "Template"
  rw = T1.FIND_ROW(ActiveSheet.Columns(1), "WELL_ROLE"): r = Rows(rw + 1).Hidden
  If 0 < rw Then: Rows(T1M.SECTION(ActiveSheet.Rows(rw), "rows")).Delete
  TSUKUBA_UTIL.DeleteNonEffectiveNames "Template"
  rw = T1.FIND_ROW(ActiveSheet.Columns(1), LABEL_PLATE_COMPOUND_CONC): c = Rows(rw + 1).Hidden
  If 0 < rw Then: Rows(T1M.SECTION(ActiveSheet.Rows(rw), "rows")).Delete
  TSUKUBA_UTIL.DeleteNonEffectiveNames "Template"
 
  If CopySection("WELL_POS", typ, fmt) Then
    rw = 1
    rw = T1M.SECTION(Rows(rw), "end") + 1
    Rows(rw).Insert Shift:=xlDown
    Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
    Cells(rw + 3, 3).Select: Action_ContextMenu_CreateWellLabel "WELL_POS"
    If Not p Then ShowCurrentSection
  End If
 
  If CopySection("WELL_ROLE", typ, fmt) Then
    rw = T1.FIND_ROW(ActiveSheet.Columns(1), "WELL_POS")
    rw = T1M.SECTION(Rows(rw), "end") + 1
    Rows(rw).Insert Shift:=xlDown
    Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
    Cells(rw + 3, 3).Select: Action_ContextMenu_CreateWellLabel "WELL_ROLE"
    If Not r Then ShowCurrentSection
  End If
 
  If CopySection(LABEL_PLATE_COMPOUND_CONC, typ, fmt) Then
    rw = T1.FIND_ROW(ActiveSheet.Columns(1), "WELL_ROLE")
    rw = T1M.SECTION(Rows(rw), "end") + 1
    Rows(rw).Insert Shift:=xlDown
    Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
    Cells(rw + 3, 3).Select: Action_ContextMenu_CreateWellLabel LABEL_PLATE_COMPOUND_CONC
    If Not c Then ShowCurrentSection
  End If
 
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
End Sub
 
Private Sub InsertTableSection(typ As String, fmt As String)
  T1M.Action_ContextMenu_InsertSection "END"
  rw = T1.FIND_ROW(ActiveSheet.Columns(1), LABEL_TABLE)
  If 0 < rw Then
    Rows(T1M.SECTION(ActiveSheet.Rows(rw), "rows")).Delete
    TSUKUBA_UTIL.DeleteNonEffectiveNames "Template"
  End If
  If T1M.CopySection(LABEL_TABLE, typ, fmt) Then
    rw = T1.FIND_ROW(ActiveSheet.Columns(1), "END")
    Rows(rw).Insert Shift:=xlDown
    Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
    Cells(rw + 2, 2).Select: Action_ContextMenu_CreateTableLabel T1.TABLE("name")
    T1M.ShowCurrentSection
  ElseIf T1M.CopySection(LABEL_TABLE, "0", "FREE") Then
    rw = T1.FIND_ROW(ActiveSheet.Columns(1), "END")
    Rows(rw).Insert Shift:=xlDown
    Cells(rw, 1).Font.Bold = True: Cells(rw, 1).Value = Split(Cells(rw, 1).Value, ":")(0)
    Cells(rw + 2, 2).Select: Action_ContextMenu_CreateTableLabel T1.TABLE("name")
    T1M.ShowCurrentSection
  End If
 
End Sub
 
Private Sub InsertEndSection()
  If 0 = T1.FIND_ROW(ActiveSheet.Columns(1), "END") Then
    rw = ActiveSheet.UsedRange.Rows.COUNT
    Rows(rw + 1).Interior.ThemeColor = END_SECTION_THEME_COLOR
    Rows(rw + 1).Interior.TintAndShade = END_SECTION_TINT1_COLOR
    Rows(rw + 1).Cells(1, 1).Value = "END"
    Rows(rw + 1).Cells(1, 1).Font.Bold = True
    Rows(rw + 1).Cells(1, 1).Font.Color = RGB(255, 255, 255)
        End If
End Sub
Private Sub Action_ContextMenu_UpdatePlateProperty(mnu As String)
        If 0 < InStr(SYSTEM_SUPPORT_PLATE_TYPE, mnu) Then 'ƒvƒŒ[ƒgƒ^ƒCƒv•ÏX
                ActiveSheet.Range(LABEL_PLATE_TYPE) = mnu
                Action_ContextMenu_InsertSection "info"
        ElseIf 0 < InStr(SYSTEM_SUPPORT_PLATE_FORMAT, mnu) Then 'ƒvƒŒ[ƒgƒtƒH[ƒ}ƒbƒg•ÏX
                ActiveSheet.Range(LABEL_PLATE_FORMAT) = mnu
                Action_ContextMenu_InsertSection "info"
        ElseIf 0 < InStr(SYSTEM_SUPPORT_PLATE_READER, mnu) Then 'ƒvƒŒ[ƒgƒŠ[ƒ_[•ÏX
                ActiveSheet.Range(LABEL_PLATE_READER) = mnu
        End If
End Sub
 
Private Function CopySection(Name As String, typ As String, fmt As String)
  On Error Resume Next
  Dim sht As String
  sht = ActiveSheet.Name
  Sheets("TSUKUBA_TEMPLATE").Activate
  rw = T1.FIND_ROW(Sheets("TSUKUBA_TEMPLATE").Columns(1), Name, typ, fmt)
  If 0 < rw Then
    Rows(T1M.SECTION(Sheets("TSUKUBA_TEMPLATE").Rows(rw), "rows")).Copy
    Sheets(sht).Activate
    CopySection = True
  Else
    Sheets(sht).Activate
    ' MsgBox typ & ":" & fmt & "—p‚Ì" & name & "‚Íì¬‚Å‚«‚Ü‚¹‚ñB"
    CopySection = False
  End If
End Function
 
 
Public Sub Action_ContextMenu_HideAllSection()
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Dim pos As Integer: pos = 1
        Dim nxt As Integer: nxt = T1M.SECTION(Cells(pos, 1), "end") + 1
        Do While 1 < nxt
                Rows(T1M.SECTION(Cells(pos, 1), "inrows")).Hidden = True
                Do While pos < nxt
                        pos = pos + 1
                Loop
                nxt = T1M.SECTION(Cells(pos, 1), "end") + 1
        Loop
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
End Sub
 
Public Sub Action_ContextMenu_ShowAllSection()
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Dim rng As Range
        Set rng = Application.Selection
        Cells.Select
        Selection.EntireRow.Hidden = False
        rng.Select
        Set rng = Nothing
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
End Sub
 
Public Sub ShowCurrentSection()
  Rows(T1M.SECTION(ActiveCell, "inrows")).Hidden = False
End Sub
 
Public Sub HideCurrentSection()
  Rows(T1M.SECTION(ActiveCell, "inrows")).Hidden = True
End Sub
 
Public Sub Action_ContextMenu_DeleteCurrentSection()
  Rows(T1M.SECTION(ActiveCell, "rows")).Delete
End Sub
 
 
Public Function SECTION(rng As Range, func As String)
        On Error GoTo SECTION_ERR
        SECTION = 0
        Dim names As String: names = UCase(T1M.LabelNames("all_exist")) & ",EXTRA,END,SECTION,Template,TSUKUBA_TEMPLATE,"
        Dim val As Variant
 
        Dim beg_row As Integer
        For beg_row = rng.row To WorksheetFunction.MIN(1, rng.row - 3000) Step -1
                val = Cells(beg_row, 1).Value
                If Not isEmpty(val) Then
                        ttl = Split(val, ":")(0) & ","
                        If InStr(names, ttl) Then Exit For
                End If
        Next
 
        Dim end_row As Integer
        For end_row = beg_row + 1 To beg_row + 3001
                val = Cells(end_row, 1).Value
                If Not isEmpty(val) Then
                        ttl = Split(val, ":")(0) & ","
                        If InStr(names, ttl) Then Exit For
                End If
        Next
 
        If 3000 <= end_row - beg_row Then end_row = 1: beg_row = 0
 
        Select Case func
                Case "beginning": SECTION = beg_row
                Case "end":       SECTION = end_row - 1
                Case "current":   SECTION = Cells(beg_row, 1).Value
                Case "next":      SECTION = Cells(end_row, 1).Value
                Case "inrows":    SECTION = CStr(beg_row + 1) & ":" & CStr(end_row - 2)
                Case "rows":      SECTION = CStr(beg_row) & ":" & CStr(end_row - 1)
                Case "color":     SECTION = Cells(beg_row, 1).Interior.ThemeColor
                Case "tint":      SECTION = Cells(beg_row, 1).Interior.TintAndShade
                Case "hide?":     SECTION = Rows(beg_row + 1).Hidden
        End Select
 
        Exit Function
SECTION_ERR:
        SECTION = CVErr(xlErrRef)
End Function
 
 
Private Sub Action_ContextMenu_SaveAsTemplate()
  Application.DisplayAlerts = False
 
  ThisWorkbook.Save
 
  Dim sht As Variant
  If 0 < InStr(SYSTEM_SUPPORT_PLATE_TYPE, Range(LABEL_PLATE_TYPE)) And _
     0 < InStr(SYSTEM_SUPPORT_PLATE_FORMAT, Range(LABEL_PLATE_FORMAT)) And _
     0 < InStr(SYSTEM_SUPPORT_PLATE_READER, Range(LABEL_PLATE_READER)) Then
    For Each sht In ThisWorkbook.Worksheets
      If 0 < InStr(sht.Name, "TSUKUBA_TEMPLATE") Then
        sht.Visible = xlSheetVisible
        sht.Delete
      End If
    Next sht
 
    ThisWorkbook.SaveAs filename:=ThisWorkbook.path & Application.PathSeparator & Format(Date, "yymmdd") & "_" & Range(LABEL_PLATE_FORMAT) & "_" & Range(LABEL_PLATE_TYPE) & "_" & Range(LABEL_PLATE_READER) & ".xlsm"
 
    T1M.Action_WorkBook_Initialize
  Else
    MsgBox "ƒvƒŒ[ƒgİ’è‚ğŠ®—¹‚µ‚Ä‚­‚¾‚³‚¢"
    TSUKUBA_UTIL.ShowStatusMessage "ƒvƒŒ[ƒgİ’è‚ğŠ®—¹‚µ‚Ä‚­‚¾‚³‚¢"
  End If
 
  Application.DisplayAlerts = True
End Sub
 
Private Sub Action_ContextMenu_CreateWellLabel(labelname As String)
        Dim cur As Range: Set cur = Selection: Selection.CurrentRegion.Select
        Dim sel As Range: Set sel = Selection
        Dim cnt As Integer: cnt = (sel.Rows.COUNT - 1) * (sel.Columns.COUNT - 1)
 
        If CStr(cnt) = T1.PLATE("Template", "type") And sel.Cells(2, 1).Value = "A" And sel.Cells(1, 2).Value = "1" Then
                If labelname = "" Then
                        labelname = sel.Cells(1, 1).Offset(-2, -1).Value
                Else
                        sel.Cells(1, 1).Offset(-2, -1).Value = UCase(labelname)
                End If
                If labelname <> "" Then
                        sel.Resize(sel.Rows.COUNT - 1, sel.Columns.COUNT - 1).Offset(1, 1).Select
                        Selection.Name = "'" & ActiveSheet.Name & "'!" & labelname
                        TSUKUBA_UTIL.ShowStatusMessage "–¼‘O [" & labelname & "] ‚ğì¬‚µ‚Ü‚µ‚½B"
                End If
        Else
                cur.Select
        End If
  Set cur = Nothing
  Set sel = Nothing
 
End Sub
 
Private Sub Action_ContextMenu_CreateTableLabel(labelname As String)
        Dim cur As Range
        Set cur = Selection
        Selection.CurrentRegion.Select
        Dim sel As Range
        Set sel = Selection
        If 1 < sel.COUNT And 2 < sel.Rows.COUNT And sel.Cells(2, 1).Value <> "A" Then
                sel.Cells(1, 1).Offset(-2, -1).Value = labelname
                sel.Name = "'" & ActiveSheet.Name & "'!" & labelname
        Else
                cur.Select
        End If
        Set cur = Nothing
        Set sel = Nothing
End Sub
 
Private Sub Action_ContextMenu_SelectLabel(labelname As String)
        Range(labelname).Select
End Sub
 
Private Sub Action_ContextMenu_DeleteLabel(labelname As String)
        If CStr(Range(labelname).COUNT) = T1.PLATE("Template", "type") Then
                ActiveSheet.Range(labelname).Cells(1, 1).Offset(-3, -2).Value = ""
        ElseIf 10 < Range(labelname).COUNT Then
                ActiveSheet.Range(labelname).Cells(1, 1).Offset(-2, -1).Value = ""
        End If
        TSUKUBA_UTIL.ShowStatusMessage "–¼‘O [" & Replace(labelname, "Template!", "") & "] ‚ğíœ‚µ‚Ü‚µ‚½"
        ActiveWorkbook.Worksheets("Template").names(labelname).Delete
End Sub
 
Private Sub Action_ContextMenu_ChangeLabelName(labelname As String)
        Dim nam As String: nam = InputBox("–¼‘O‚ğ“ü—Í", "–¼‘O‚Ì•ÏX", labelname)
        If nam = "" Then nam = labelname
        Dim cl  As Variant
        For Each cl In ActiveSheet.UsedRange
                If InStr(CStr(cl.Value), labelname) Then cl.Value = Replace(cl.Value, labelname, nam)
                If InStr(cl.Formula, labelname) Then cl.Formula = Replace(cl.Formula, labelname, nam)
        Next
        ActiveSheet.Range(labelname).Name = "'" & ActiveSheet.Name & "'!" & nam
        ActiveSheet.names(labelname).Delete
        TSUKUBA_UTIL.ShowStatusMessage "–¼‘O [" & labelname & "] ‚ğ [" & nam & "] ‚É•ÏX‚µ‚Ü‚µ‚½"
End Sub
 
Private Sub Action_ContextMenu_CreatePlateLabel(labelname As String)
        If 1 < Selection.COUNT Then
                For Each c In Selection
                        c.Select
                        If InStr(c.Formula, "=T1.ROLE(") Or InStr(c.Formula, "=T1.LABEL(") Then Action_ContextMenu_CreatePlateLabel ""
                Next
        Else
                If labelname = "" Then
                        Dim fml As String: fml = ActiveCell.Formula
                        If InStr(fml, "=T1.ROLE(") Then
                                param = Split(Mid(fml, 10, Len(fml) - 10), ",")
                                labelname = ParameterName(param(0)) & "_" & ParameterName(param(1)) & "_" & ParameterName(param(2))
                        ElseIf InStr(ActiveCell.Formula, "=T1.Label(") Then
                                param = Split(Mid(fml, 11, Len(fml) - 11), ",")
                                labelname = ParameterName(param(0)) & "_" & ParameterName(param(1))
                        End If
                        labelname = InputBox("–¼‘O‚ğ“ü—Í", "–¼‘O‚ğ“ü—Í", labelname)
                End If
 
                labelname = Replace(labelname, "*", "")
                labelname = Replace(labelname, "@", "_")
 
                If labelname <> "" Then
                        ActiveCell.Name = "'" & ActiveSheet.Name & "'!" & labelname
                        TSUKUBA_UTIL.ShowStatusMessage "–¼‘O [" & labelname & "] ‚ğì¬‚µ‚Ü‚µ‚½B"
                End If
        End If
End Sub
 
 
Private Function ParameterName(ByVal adr As String)
        If T1M.LabelNames(adr) Then
                ParameterName = adr
        ElseIf VarType(adr) = vbString And Left(adr, 1) = """" Then
                ParameterName = Mid(adr, 2, Len(adr) - 2)
        Else
                ParameterName = ActiveSheet.Range(adr).Value
        End If
        ParameterName = UCase(ParameterName)
End Function
 
 
Public Sub Action_Worksheet_Update()
  On Error GoTo ERR_Action_WorkSheet_Update
 
        TSUKUBA_UTIL.ShowStatusMessage "ƒf[ƒ^ÄŒvZ’† [" & ActiveSheet.Name & "]"
 
  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
 
  ActiveSheet.EnableCalculation = False
 
  TSUKUBA_UTIL.DeleteNonEffectiveNames
  RESOURCE.RestAssayResult
 
  ActiveSheet.EnableCalculation = True
  ActiveSheet.UsedRange.Calculate
 
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
 
  Exit Sub
 
ERR_Action_WorkSheet_Update:
 
End Sub
 
 
Public Function GetLabelOnCurPos()
  Dim lbl As Variant
  For Each lbl In T1.CSV2ARY(T1M.LabelNames("exist_well"))
    If Not Application.Intersect(Range(lbl), ActiveCell) Is Nothing Then
      GetLabelOnCurPos = lbl: Exit Function
    End If
  Next
End Function
 
 
Public Function ExcludedWellP() As Boolean
  ExcludedWellP = ActiveCell.Font.Strikethrough
End Function
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
Rem Mac‚Å‚ÌExcel VBAŠJ”­ã‚Ì’ˆÓ“_
Rem
 
Rem - ƒtƒ@ƒCƒ‹–¼AƒtƒHƒ‹ƒ_–¼‚Í31•¶šˆÈ“à‚Å‚ ‚é‚±‚ÆB
Rem   ‚»‚êˆÈã‚Ìƒtƒ@ƒCƒ‹–¼‚Í31•¶š‚ÉŠÛ‚ß‚Ä‚µ‚Ü‚¤‚Ì‚ÅA–ÀãƒAƒNƒZƒX‚Å‚«‚È‚­‚È‚éB
Rem
Rem - Application.ScreenUpdating = False@ğŒ‚ÅA‘¼‚Ìƒtƒ@ƒCƒ‹‚ğŠJ‚­‚Æ‚»‚Ì“_‚ÅMacro‚ª’â~‚·‚éB
Rem   ‹°‚ç‚­ƒ}ƒNƒÀsŒ ‚Æbook‚ÌFocus‚ÌŠÔ‚É‰½‚©ŠÖ˜A«‚ª‚ ‚é‚ªA
Rem   TrueğŒ‚Å‚ ‚Á‚Ä‚àuƒtƒ@ƒCƒ‹“Ç‚İ‚İ{Movev‚ÌŒJ‚è•Ô‚µ‚ª“r’†’â~‚·‚é‚Ì‚ÅA‘¼‚Ì—vˆö‚à‚ ‚é‚æ‚¤‚¾B
Rem
Rem - ƒtƒHƒ‹ƒ_‹æØ‚è‚Íu:v Application.PathSeparator
Rem
Rem - g‚¦‚È‚¢ŠÖ”‚ª‚ ‚éB
Rem   Dir, GetOpenFile,@‘¼‘½”
Rem
Rem - UserForm‚Ì•¶šƒR[ƒh‘Îô‚ª‚³‚ê‚Ä‚¨‚ç‚¸AwinÌmacˆÚ“®‚·‚é‚Æ•¶š‰»‚¯‚·‚éB
Rem
Rem - UserForm‚ğModeless‚Å•\¦‚Å‚«‚È‚¢B
Rem
Rem - popupmenu‚Éeditform‚ğ•\¦‚·‚é‚±‚Æ‚ª‚Å‚«‚È‚¢B
Rem
Rem - AddInŠi”[—p‚Ìsystem/user folder‚ª–³‚¢B
Rem
Rem - ƒ†[ƒU[ŠÖ”“à‚ÅFindŠÖ”‚ªg‚¦‚È‚¢B
Rem
Rem - Statusbar‚Ö‚Ì•¶š•\¦‚Ìƒ^ƒCƒ~ƒ“ƒO‚ªwindows‚ÆˆÙ‚È‚éB
Rem
 
Rem
Rem “¤’m¯
Rem
 
Rem - VBAƒvƒƒWƒFƒNƒgƒIƒuƒWƒFƒNƒgƒ‚ƒfƒ‹‚Ö‚ÌƒAƒNƒZƒX‚ğ‹–‰Â‚·‚é‚ÆAƒR[ƒh‚ğ•ÒW‚Å‚«‚éB
Rem   Msgbox Application.VBE.ActiveCodePane.CodeModule.Lines(50,3)
Rem
Rem - D&DŒãpassword“ü—Í‚µ‚ÄŠJ‚¯‚ÎAˆá‚¤book‚©‚ç‚àƒAƒNƒZƒX‰Â”\‚É‚È‚éB
Rem   MsgBox Workbooks("ƒRƒs[01342_Lab report_¶—‰»Šw_®’Jæ¶_]ã—l.xlsx").Sheets("QCŒ‹‰Ê").Range("B1").value
Rem
Rem - InStr("sdsd", "") ‚Í 1
Rem

