Attribute VB_Name = "RESOURCE"
Rem ********************************************************************************
Rem Public変数
Rem ********************************************************************************
Rem
Private cpdmap As CompoundPlatemap
Private well2rwcl As Well2RowCol
Private pltalgns As PlateAlignments
Private asyrslts As AssayResults


Rem ********************************************************************************
Rem [用途] アッセイ値
Rem ********************************************************************************
Public Sub RestAssayResult()
        Set RESOURCE.asyrslts = Nothing
        Set RESOURCE.pltalgns = Nothing
End Sub

Public Function GetAssayResult(platename As String, wellpos As String, param As Variant) As String
        If RESOURCE.asyrslts Is Nothing Then
                Set RESOURCE.asyrslts = New AssayResults
        Else
                GetAssayResult = RESOURCE.asyrslts.result(platename, wellpos, param)
        End If
End Function

Public Sub SetAssayResult(platename As String, wellpos As String, param As Variant, res As String)
        If RESOURCE.asyrslts Is Nothing Then Set RESOURCE.asyrslts = New AssayResults
        RESOURCE.asyrslts.result(platename, wellpos, param) = res
End Sub

Public Sub UpdateAssayResult(Optional plt As String = "")
        On Error Resume Next
        If RESOURCE.asyrslts Is Nothing Then Set RESOURCE.asyrslts = New AssayResults
        Dim val As Variant
        Dim platename As String:
        If plt = "" Then
                platename = ActiveSheet.Name
        Else
                platename = plt
        End If
        
        Dim labelname As Variant
        For Each labelname In T1.CSV2ARY(T1M.GetPlateLabels())
                Select Case labelname
                        Case "PLATE_NAME":      val = platename
                        Case "PLATE_DATAFILE":  val = T1.PLATE(platename, "rawdatafile")
                        Case "PLATE_EXCELFILE": val = ThisWorkbook.Name
                        Case "ANALYZE_DATE":    val = T1.SYSTEM("today")
                        Case "SYSTEM_VERSION":  val = T1.SYSTEM("")
                        Case Else:              val = Range(CStr(labelname)).Value
                                ' TEST_ASSAY, TEST_DATE, TEST_TIME, QC_ZPRIME, QC_SB, QC_CVPBK, QC_CVPCTRL, ユーザー定義
                End Select
                RESOURCE.asyrslts.result(platename, "", CStr(labelname)) = val
        Next
        
        Dim wel As Variant
        Dim wellpos As String
        For Each labelname In T1.CSV2ARY(T1M.GetWellLabels())
                For Each wel In Range(T1M.LABEL_PLATE_WELL_POSITION)
                        wellpos = wel.Value
                        Select Case labelname
                                Case "WELL_POS":        val = wellpos
                                Case "CPD_ID":          val = RESOURCE.GetCpdID(platename, wellpos)
                                Case "WELL_POS0":       val = RESOURCE.ConvertWellpos(wellpos, "pos0")
                                Case "WELL_ROW":        val = RESOURCE.ConvertWellpos(wellpos, "ROW")
                                Case "WELL_COLUMN":     val = RESOURCE.ConvertWellpos(wellpos, "COLUMN")
                                Case "WELL_ROWNUM":     val = RESOURCE.ConvertWellpos(wellpos, "ROWNUM")
                                Case "WELL_ROLE":       val = T1.well(wellpos, "role")
                                Case "CPD_CONC":        val = T1.well(wellpos, "conc")
                                Case Else:              val = T1.well(wellpos, CStr(labelname), "val")
                                        ' RAW_DATA, WELL_ROLE, CPD_CONC, CPD_RESULT, ユーザー定義
                        End Select
                        RESOURCE.asyrslts.result(platename, wellpos, CStr(labelname)) = val
                Next
        Next
        
        If TSUKUBA_UTIL.ExistNameP(platename, LABEL_TABLE) Then
                Dim rownum As Integer
                Dim rng As Range: Set rng = Range(LABEL_TABLE)
                Dim rng2 As Range
                Dim cl As Variant
                Dim rw As Variant
                For Each cl In rng.Rows(1).Columns
                        If cl.Value <> "" Then
                                Set rng2 = rng.Resize(rng.Rows.COUNT - 1, rng.Columns.COUNT).Offset(1, 0)
                                For Each rw In rng2.Rows
                                        val = rw.Cells(1, cl.Column - rng2.Column + 1).Value
                                        RESOURCE.asyrslts.result(platename, cl.Value, CInt(rw.row - rng2.row + 1)) = Replace(val, ",", ";")
                                Next
                        End If
                Next
        End If
End Sub



Rem ********************************************************************************
Rem [用途] 各シートのプレートマップ情報
Rem ********************************************************************************
Rem
Rem  RESOURCE.GetRoleRange( sht, lbl, role, (conc) )
Rem  RESOURCE.GetRoleAddress( sht, lbl, role, (conc) )
Rem  RESOURCE.GetRoleWell( sht, role, (conc) )
Rem  RESOURCE.GetHere( sht, rng, (func) )
Rem
Rem  InitRoleInfo( sht )
Rem

Public Function GetRoleRange(sht As String, lbl As String, role As String, Optional conc As String = "") As Range
        On Error Resume Next
        init_pltalgns
        Dim wpos As Range: Set wpos = Sheets(sht).Range(T1M.LABEL_PLATE_WELL_POSITION)
        Dim lbel As Range: Set lbel = Sheets(sht).Range(lbl)
        Dim adr As String: adr = pltalgns.item(sht).GetRoleAddress(role, conc)
        Set GetRoleRange = Sheets(sht).Range(adr).Offset(lbel.row - wpos.row, lbel.Column - wpos.Column)
End Function

Public Function GetRoleAddress(sht As String, lbl As String, role As String, Optional conc As String = "") As String
        On Error Resume Next
        init_pltalgns
        Dim wpos As Range: Set wpos = Sheets(sht).Range(T1M.LABEL_PLATE_WELL_POSITION)
        Dim lbel As Range: Set lbel = Sheets(sht).Range(lbl)
        Dim adr As String: adr = pltalgns.item(sht).GetRoleAddress(role, conc)
        GetRoleAddress = Sheets(sht).Range(adr).Offset(lbel.row - wpos.row, lbel.Column - wpos.Column).Address
End Function

Public Function GetRoleWell(sht As String, role As String, Optional conc As String = "")
        init_pltalgns
        GetRoleWell = pltalgns.item(sht).GetRoleWell(role, conc)
End Function

Public Function GetHere(sht As String, rng As Range, Optional func As String = "")
        init_pltalgns
        GetHere = pltalgns.item(sht).GetHere(rng, func)
End Function

Public Sub InitRoleInfo(sht As String)
        If Not pltalgns Is Nothing Then pltalgns.Remove sht
End Sub

Private Sub init_pltalgns()
   ' If pltalgns Is Nothing Then Set pltalgns = New PlateAlignments
   Set pltalgns = New PlateAlignments
End Sub




Rem ********************************************************************************
Rem [用途] 化合物番号の取得
Rem ********************************************************************************
Rem
Rem  RESOURCE.GetCpdID( platename, wellpos )
Rem  RESOURCE.GetCpdConc( platename, wellpos )
Rem  RESOURCE.GetCpdVol( platename, wellpos )
Rem
Rem  LoadCompoundTable
Rem  ResetCpdTable
Rem

Public Function GetCpdID(platename As String, wellpos As String)
        init_cpdmap
        GetCpdID = cpdmap.GetCpdID(platename, wellpos)
End Function

Public Function GetCpdConc(platename As String, wellpos As String)
        init_cpdmap
        GetCpdConc = cpdmap.GetCpdConc(platename, wellpos)
End Function

Public Function GetCpdVol(platename As String, wellpos As String)
        init_cpdmap
        GetCpdVol = cpdmap.GetCpdVol(platename, wellpos)
End Function

Private Sub init_cpdmap()
        If cpdmap Is Nothing Then Set cpdmap = New CompoundPlatemap
End Sub

Public Sub LoadCompoundTable()
        Set cpdmap = New CompoundPlatemap
        cpdmap.LoadCompoundTable
End Sub

Public Sub ResetCpdTable()
        Set cpdmap = Nothing
End Sub




Rem ********************************************************************************
Rem [用途] Well番号と位置
Rem ********************************************************************************
Rem
Rem  RESOURCE.GetRC( wellpos )
Rem  RESOURCE.GetWellpos( rw, cl, (opt) )
Rem  RESOURCE.ConvertWellpos( wellpos, (opt) )
Rem

Public Function ConvertWellpos(wellpos As String, Optional opt As String = "pos") As Variant
        init_well2rwcl
        ConvertWellpos = well2rwcl.ConvertWellpos(wellpos, opt)
End Function

Public Function GetRC(wellpos As String) As Variant
        init_well2rwcl
        GetRC = well2rwcl.GetRC(wellpos)
End Function

Public Function GetWellpos(rw As Integer, cl As Integer, Optional opt As String = "") As Variant
        init_well2rwcl
        GetWellpos = well2rwcl.GetWellpos(rw, cl, opt)
End Function

Private Sub init_well2rwcl()
        If well2rwcl Is Nothing Then Set well2rwcl = New Well2RowCol
End Sub




Rem ********************************************************************************
Rem [用途] テスト
Rem ********************************************************************************

Public Sub TestPlates()
        MsgBox RESOURCE.GetRoleWell("Template", "MAX")
        MsgBox RESOURCE.GetRoleWell("Template (2)", "MAX")
        MsgBox RESOURCE.GetRoleAddress("TSUKUBA_TEMPLATE", "WELL_ROLE", "MAX")
End Sub



Public Sub TestCompoundPlatemap()
        MsgBox "Template@A4: " & GetCpdID("Template", "A4")
End Sub


Public Sub TestWell2RowCol()
        Set RESOURCE.well2rwcl = Nothing
        Set RESOURCE.well2rwcl = New Well2RowCol
        MsgBox "A1: " & RESOURCE.GetRC("A1")(0) & "," & RESOURCE.GetRC("A1")(1) & "  " & RESOURCE.ConvertWellpos("A1", "pos0") & " " & _
                                 "(2,4): " & RESOURCE.GetWellpos(2, 4)
End Sub








