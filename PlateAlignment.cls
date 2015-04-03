VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlateAlignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem  wellpos•¶Žš—ñ‚Ì•ÏŠ·
Rem
Rem  GetHere( cel, (func) )
Rem  GetRoleAddress( role, (conc) )
Rem  GetRoleWell( role, (conc) )
Rem
Rem  PlateType
Rem  PlateWidth
Rem  PlateHeight
Rem

Private adrss As Collection
Private wells As Collection
Private labls As Collection
Private shtname As String
Private plate_width As Collection
Private plate_height As Collection

Rem ********************************************************************************
Rem Function
Rem ********************************************************************************
Public Function GetHere(cel As Range, Optional func As String = "") As Variant
  Dim nam As Variant
  For Each nam In labls
    If nam(0) <= cel.row And cel.row < nam(0) + Me.PlateHeight And _
       nam(1) <= cel.Column And cel.Column < nam(1) + Me.PlateWidth Then
      Select Case func
        Case "": GetHere = Array(nam(2), nam(0), nam(1), Me.GetHere(cel, "wellpos"), cel.row - nam(0) + 1, cel.Column - nam(1) + 1)
        Case "label": GetHere = nam(2)
        Case "wellpos": GetHere = RESOURCE.GetWellpos(cel.row - nam(0) + 1, cel.Column - nam(1) + 1)
        Case "rc": GetHere = RESOURCE.GetWellpos(cel.row - nam(0) + 1, cel.Column - nam(1) + 1, "rc")
      End Select
      Exit Function
    End If
  Next
End Function

Public Function GetRoleAddress(role As String, Optional conc As String = "") As Variant
  On Error Resume Next
  GetRoleAddress = False
  If conc <> "" Then role = role & "@" & conc
  GetRoleAddress = adrss.item(role): Exit Function
End Function


Public Function GetRoleWell(role As String, Optional conc As String = "") As Variant
  On Error Resume Next
  GetRoleWell = False
  If conc <> "" Then role = role & "@" & conc
  GetRoleWell = wells.item(role): Exit Function
End Function


Rem ********************************************************************************
Rem Procedure
Rem ********************************************************************************
Public Sub Initialize(ByVal sht As String)
  If adrss Is Nothing Then
    Set adrss = New Collection: Set wells = New Collection: Set labls = New Collection
  End If
  
  shtname = sht
  
  Dim lbl As String
  Dim nam As Variant
  For Each nam In Sheets(sht).names
    If CStr(nam.RefersToRange.COUNT) = Me.PlateType Then
      lbl = Replace(nam.Name, sht & "!", "")
      labls.Add Array(nam.RefersToRange.row, nam.RefersToRange.Column, lbl), lbl
    End If
  Next
  
  Dim rw As Integer: Dim cl As Integer: Dim rol As String: Dim cnc As String
  Dim pos As Variant
  For Each pos In Sheets(sht).Range("WELL_POS")
    rw = pos.row - Sheets(sht).Range("WELL_POS").row + 1
    cl = pos.Column - Sheets(sht).Range("WELL_POS").Column + 1
    rol = Sheets(sht).Range("WELL_ROLE").Cells(rw, cl)
    cnc = CStr(Sheets(sht).Range("CPD_CONC").Cells(rw, cl))

    set_wells_adrss rol, pos ' Role
    If cnc <> "" Then set_wells_adrss rol & "@" & cnc, pos ' Role@Conc
    If InStr(rol, "CPD") Then set_wells_adrss "CPD*", pos ' CPD*
    If InStr(rol, "REF") Then set_wells_adrss "REF*", pos ' REF*
    If InStr(rol, "MAX") Then set_wells_adrss "MAX*", pos ' MAX*
    If InStr(rol, "MIN") Then set_wells_adrss "MIN*", pos ' MIN*
    If InStr(rol, "POS") Then set_wells_adrss "POS*", pos ' POS*
    If InStr(rol, "NEG") Then set_wells_adrss "NEG*", pos ' NEG*
  Next pos

End Sub


Private Sub set_wells_adrss(key As String, pos As Variant)
  On Error Resume Next
  Dim wellpos As String: Dim celladr As String
  If wells.item(key) = False Then
    wellpos = pos.Value: celladr = pos.Address
  Else
    wellpos = wells.item(key) & "," & pos.Value: wells.Remove key
    celladr = Union(Sheets(shtname).Range(adrss.item(key)), pos).Address: adrss.Remove key
  End If
  wells.Add wellpos, key
  adrss.Add celladr, key
End Sub



Private Sub Class_Initialize()
  Set plate_width = New Collection
  plate_width.Add 48, "1536"
  plate_width.Add 24, "384"
  plate_width.Add 12, "96"
  plate_width.Add 6, "24"
  
  Set plate_height = New Collection
  plate_height.Add 32, "1536"
  plate_height.Add 16, "384"
  plate_height.Add 8, "96"
  plate_height.Add 4, "24"

End Sub




Rem ********************************************************************************
Rem Property
Rem ********************************************************************************
Public Property Get Name()
  Name = shtname
End Property

Public Property Get PlateType() As String
  PlateType = CStr(Sheets(shtname).Range(T1.PLATE_TYPE).Value)
End Property

Public Property Get PlateWidth() As Integer
  PlateWidth = plate_width(Me.PlateType)
End Property

Public Property Get PlateHeight() As Integer
  PlateHeight = plate_height(Me.PlateType)
End Property

