VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlateAlignments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem  PlateAlignmentƒNƒ‰ƒX‚ÌCollection
Rem

Private plts As Collection

Private Sub Class_Initialize()
  Set plts = New Collection
End Sub

Public Function item(sht As String) As PlateAlignment
  On Error Resume Next
  
  If plts.item(sht) Is Nothing Then
    Dim plt As New PlateAlignment
    plt.Initialize sht
    plts.Add plt, sht
    Set item = plt
  Else
    Set item = plts.item(sht)
  End If
End Function

Public Sub Remove(sht As String)
  On Error Resume Next
  If plts.item(sht) Is Nothing Then
    a = Nothing
  Else
    plts.Remove sht
  End If
End Sub
