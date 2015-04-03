VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssayResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem

Private asyrslts As Collection  '  Value: "value",   Key: "Platename@Wellpos@Label" / "Platename@@Label" / "Platename#Field#Row" 

Private Sub Class_Initialize()
  Set asyrslts = New Collection
End Sub

Public Property Get result(platename As String, wellpos As String, param As Variant) As String
  On Error Resume Next
  If asyrslts Is Nothing Then
    Set asyrslts = New Collection
  Else
    Select Case TypeName(param)
			Case "String":  result = asyrslts.item(platename & "@" & wellpos & "@" & param)
			Case "Integer": result = asyrslts.item(platename & "#" & wellpos & "#" & CStr(param))
    End Select
  End If
End Property

Public Property Let result(platename As String, wellpos As String, param As Variant, res As String)
  On Error Resume Next
  
  If asyrslts Is Nothing Then Set asyrslts = New Collection
  
  Dim key As String
  Select Case TypeName(param)
    Case "String":  key = platename & "@" & wellpos & "@" & param
    Case "Integer": key = platename & "#" & wellpos & "#" & CStr(param)
  End Select
  
  If asyrslts.item(key) = False Then
    asyrslts.Add res, key
  Else
    asyrslts.Remove key: asyrslts.Add res, key
  End If
End Property







