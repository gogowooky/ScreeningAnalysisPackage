VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompoundPlatemap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem  プレート番号＋WellPos と 化合物IDの変換
Rem
Rem  GetCpdID( platename, wellpos )
Rem  GetCpdConc( platename, wellpos )
Rem  GetCpdVol( platename, wellpos )
Rem
Rem  LoadCompoundTable
Rem

Private cpds As Collection

Rem
Rem 関数
Rem

Public Function GetCpdID(platename As String, wellpos As String) As Variant
  GetCpdID = GetCpds(platename, wellpos)(0)
End Function

Public Function GetCpdConc(platename As String, wellpos As String) As Variant
  GetCpdConc = Sheets(T1M.SHEETNAME_REPORT_ASSAY_RESULT).Cells(GetCpds(platename, wellpos)(1), 5)
End Function

Public Function GetCpdVol(platename As String, wellpos As String) As Variant
  GetCpdVol = Sheets(T1M.SHEETNAME_REPORT_ASSAY_RESULT).Cells(GetCpds(platename, wellpos)(1), 6)
End Function

Private Function GetCpds(platename As String, wellpos As String) As Variant
  If cpds Is Nothing Then
    GetCpds = Array("", 0)
  Else
    GetCpds = cpds.item(platename & "@" & wellpos)
  End If
End Function

Rem
Rem サブルーチン
Rem

Public Sub LoadCompoundTable()
  If cpds Is Nothing Then
    ' あとから報告書ファイルを読み込む
    Dim wb As String: wb = ActiveWorkbook.Name
    Dim ws As String: ws = ActiveSheet.Name
    Dim repf As String: repf = TSUKUBA_UTIL.WinMacSelectFile(ActiveWorkbook.path)
    Workbooks.Open repf
    Dim rep As String: rep = ActiveWorkbook.Name
    Workbooks(rep).Worksheets(T1M.SHEETNAME_REPORT_ASSAY_RESULT).Select
    Workbooks(rep).Worksheets(T1M.SHEETNAME_REPORT_ASSAY_RESULT).Copy After:=Workbooks(wb).Worksheets("Template")
    Workbooks(rep).Close
    Workbooks(wb).Worksheets(ws).Activate
    LoadTable
  End If
End Sub


Private Sub LoadTable()
  ' platename, wellpos, cpd# だけでなく、
  ' platename, wellrole, cpd# のテーブルも読み込めるように変更した
  On Error Resume Next
  If TSUKUBA_UTIL.ExistSheetP(T1M.SHEETNAME_REPORT_ASSAY_RESULT) Then
    Set cpds = New Collection
    Dim first As Boolean: first = True
    Dim pltcol As Integer: pltcol = 0
    Dim welcol As Integer: welcol = 0
    Dim cpdcol As Integer: cpdcol = 0
    Dim welrolecol As Integer: welrolecol = 0
    Dim rw As Variant: Dim cl As Variant
    Dim wells As String: Dim wel As Variant
    
    For Each rw In Sheets(T1M.SHEETNAME_REPORT_ASSAY_RESULT).UsedRange.Rows
      If first Then
        For Each cl In rw.Columns
          If 0 < InStr(T1M.PLATE_TITLE, cl.Value) Then pltcol = cl.Column
          If 0 < InStr(T1M.WELL_TITLE, cl.Value) Then welcol = cl.Column
          If 0 < InStr(T1M.WELLROLE_TITLE, cl.Value) Then welrolecol = cl.Column
          If 0 < InStr(T1M.COMPOUND_TITLE, cl.Value) Then cpdcol = cl.Column
          If 0 < pltcol * welcol * cpdcol Or 0 < pltcol * welrolecol * cpdcol Then Exit For
        Next
        first = False
      Else
        If 0 < welcol Then
            cpds.Add Array(rw.Cells(1, cpdcol).Value, rw.row), rw.Cells(1, pltcol).Value & "@" & rw.Cells(1, welcol).Value
        Else
            wells = T1.SELECT_WELLS(Sheets("Template").Range("WELL_ROLE"), "match", rw.Cells(1, welrolecol).Value)
            For Each wel In T1.CSV2ARY(wells)
                cpds.Add Array(rw.Cells(1, cpdcol).Value, rw.row), rw.Cells(1, pltcol).Value & "@" & CStr(wel)
            Next
        End If
      End If
    Next
  Else
    Set cpds = Nothing
  End If
End Sub

Private Sub Class_Initialize()
  Dim ws As String: ws = ActiveSheet.Name
  Application.ScreenUpdating = False
  LoadTable
  Sheets(ws).Activate
  Application.ScreenUpdating = True
End Sub








