VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Well2RowCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem  wellposï∂éöóÒÇÃïœä∑
Rem
Rem  ConvertWellpos( wellpos, (opt))
Rem  GetRC( wellpos )
Rem  GetWellpos( rw, cl, (opt) )
Rem

Private well2rc As Collection  '  Value: ( rw, cl, "A1", "A01", "0A01" ),   Key: "A1"
Private rowletter As Variant

Rem
Rem ä÷êî
Rem

Public Function GetRC(wellpos As String) As Variant
  GetRC = well2rc.item(wellpos)
End Function

Public Function ConvertWellpos(wellpos As String, Optional opt As String = "pos") As Variant
  Select Case opt
    Case "pos", "": ConvertWellpos = well2rc.item(wellpos)(2)
    Case "pos0":    ConvertWellpos = well2rc.item(wellpos)(3)
    Case "pos00":   ConvertWellpos = well2rc.item(wellpos)(4)
    Case "ROW":     ConvertWellpos = CStr(rowletter(well2rc.item(wellpos)(0) - 1))
    Case "COLUMN":  ConvertWellpos = CStr(well2rc.item(wellpos)(1))
    Case "ROWNUM":  ConvertWellpos = CStr(well2rc.item(wellpos)(0))
  End Select
End Function

Public Function GetWellpos(ByVal rw As Integer, ByVal cl As Integer, Optional opt As String = "pos") As Variant
  Select Case opt
    Case "pos", "": GetWellpos = rowletter(rw - 1) & CStr(cl)
    Case "pos0":    GetWellpos = rowletter(rw - 1) & Right("0" & CStr(cl), 2)
    Case "pos00":   GetWellpos = Right("0" & rowletter(rw - 1), 2) & Right("0" & CStr(cl), 2)
    Case "rc":      GetWellpos = Array(rw, cl)
    Case "RC":      GetWellpos = "R" & CStr(rw) & "C" & CStr(cl)
    Case "ROW":     GetWellpos = CStr(rowletter(rw - 1))
    Case "COLUMN":  GetWellpos = CStr(cl)
    Case "ROWNUM":  GetWellpos = CStr(rw)
    Case "row":     GetWellpos = rw
    Case "col":     GetWellpos = cl
  End Select
End Function

Rem
Rem èâä˙âª 
Rem

Private Sub Class_Initialize()
  rowletter = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", _
                    "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                    "AA", "AB", "AC", "AD", "AE", "AF")
  Set well2rc = New Collection
  
  Dim rw As Integer: Dim cl As Integer
  Dim val As Variant
  Dim pos As String: Dim pos0 As String: Dim pos00 As String
  
  For rw = 1 To 32
    For cl = 1 To 48
      pos = Me.GetWellpos(rw, cl)
      pos0 = Me.GetWellpos(rw, cl, "pos0")
      pos00 = Me.GetWellpos(rw, cl, "pos00")
      val = Array(rw, cl, pos, pos0, pos00)

      well2rc.Add val, pos

      If Len(CStr(cl)) = 1 Then well2rc.Add val, pos0
      If Len(rowletter(rw - 1)) = 1 Then well2rc.Add val, pos00
    Next
  Next
End Sub



