VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VCOMObject As Object
Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = VCOMObject.Value
End Property
Property Let Value(ByVal Value As Variant)
    VCOMObject.Value = Value
End Property
Property Get ValueDesc() As String
    ValueDesc = VCOMObject.ValueDesc
End Property
Property Get NameDesc() As String
    NameDesc = VCOMObject.NameDesc
End Property
Property Get Name() As String
    Name = VCOMObject.Name
End Property
Property Get ID() As String
    ID = VCOMObject.ID
End Property

