Attribute VB_Name = "TSUKUBA_UTIL"
' WinMacDir�p
Dim wmdir_filelist As Collection
Dim wmdir_filecount As Integer
Dim wmdir_filepos As Integer
Dim wmdir_fileext As String

Rem ******************************************************************************************************
Rem
Rem �֐�
Rem
Rem ******************************************************************************************************

' ExistNameP
Public Function ExistNameP(sht As String, nam As String) As Boolean
  Dim lblnam As Variant
  For Each lblnam In Sheets(sht).names
    ExistNameP = 0 < InStr(lblnam.Name, "!" & nam)
    If ExistNameP Then Exit Function
  Next
End Function

' ExistSheetP
Public Function ExistSheetP(sht As String) As Boolean
  Dim ws As Variant
  For Each ws In ActiveWorkbook.Worksheets
    ExistSheetP = ws.Name = sht
    If ExistSheetP Then Exit Function
  Next ws
End Function


' WinMacSelectFile:   Window/Mac�ŋ��ʂɎg����SelectFile�֐�
Public Function WinMacSelectFile(Optional path As String = "") As String
  If path = "" Then path = ThisWorkbook.path
  If T1.SYSTEM("pc") = "Windows" Then
    ChDrive Left(path, 1)
    ChDir path
    WinMacSelectFile = Application.GetOpenFilename("���ׂẴt�@�C��,*.*")
  Else
    WinMacSelectFile = MacScript("tell application ""Finder"" to set aFol to """ & path & """" & vbNewLine & "choose file default location aFol as alias")
    WinMacSelectFile = Replace(WinMacSelectFile, "alias ", "")
  End If
End Function

' GetFileExt
Public Function GetFileExt(path As String) As String
 GetFileExt = Mid(path, InStrRev(path, ".") + 1)
End Function

' WinMacDir:  Window/Mac�ŋ��ʂɎg����Dir�֐�
Public Function WinMacDir(Optional path As String = "", Optional ext As String = "") As String
  Dim fil As String
  If path <> "" Then
    Set wmdir_filelist = New Collection
    wmdir_fileext = UCase(ext)
    If T1.SYSTEM("pc") = "Windows" Then
      fil = Dir(path & Application.PathSeparator)
      While fil <> ""
        wmdir_filelist.Add fil
        fil = Dir()
      Wend
    Else
      Dim fils As Variant: Dim fl As Variant
      fils = Split(MacScript("set aFol to """ & path & """" & vbNewLine & _
                             "tell application ""Finder""" & vbNewLine & _
                             "tell folder aFol" & vbNewLine & _
                             "set indList to a reference to (every file)" & vbNewLine & _
                             "end tell" & vbNewLine & _
                             "set namList to name of indList" & vbNewLine & _
                             "end tell"), ",")
      For Each fl In fils
        wmdir_filelist.Add Trim(fl)
      Next
    End If
    wmdir_filecount = wmdir_filelist.COUNT
    wmdir_filepos = 1
  End If
    
  Dim filename As String
  Do
    If wmdir_filecount < wmdir_filepos Then
      WinMacDir = ""
      Exit Function
    Else
      filename = wmdir_filelist(wmdir_filepos)
      WinMacDir = filename
      wmdir_filepos = wmdir_filepos + 1
    End If
  Loop While InStr(wmdir_fileext, UCase(TSUKUBA_UTIL.GetFileExt(filename))) = 0
End Function


Rem ******************************************************************************************************
Rem
Rem �v���V�[�W��
Rem
Rem ******************************************************************************************************

' DeleteNonEffectiveNames:   �����Ȗ��O���폜����
Public Sub DeleteNonEffectiveNames(Optional sht As String = "")
  Dim nam As Variant
  If sht = "" Then sht = ActiveSheet.Name
  For Each nam In Sheets(sht).names
    n1 = nam.RefersToLocal
    If "=Template!#REF!" = n1 Then nam.Delete
  Next
End Sub


' �X�e�[�^�X���C���Ƀ��b�Z�[�W�\��
Public Sub ShowStatusMessage(mes As String)
   Application.StatusBar = mes
   DoEvents
End Sub


' �u���E�U��url���J��
Public Sub OpenUrl(url As String)
   If T1.SYSTEM("pc") = "Mac" Then
      MacScript ("tell application ""Safari""" & vbNewLine _
                 & "activate" & vbNewLine _
                 & "open location """ & url & """" & vbNewLine _
                 & "end tell" & vbNewLine)
   Else
      Shell "Explorer.exe " & url, vbNormalFocus
   End If
End Sub



' �B���V�[�g���R�s�[���ĕ\��
Public Sub DupulicateHiddenSheetAndShow(sht As String, altname As String)
   
   If TSUKUBA_UTIL.ExistSheetP(altname) = True Then Sheets(altname).Delete
   
   Sheets(sht).Visible = -1 ' xlSheetVisible
   Sheets(sht).Copy After:=ActiveSheet
   ActiveSheet.Name = altname
   Sheets(sht).Visible = 2 ' xlVeryHidden
   
   'Sheets(altname).Calculate
End Sub

Public Sub ConvertSelectionFomulaFromRelatioveToAbsolute()
  Dim sl As Variant
  If debug_convert_relative_formula_to_absolute Then
    For Each sl In Selection
      If sl.HasFormula Then sl.Formula = Application.ConvertFormula(sl.Formula, xlA1, xlA1, xlAbsolute)
    Next
  End If
End Sub










