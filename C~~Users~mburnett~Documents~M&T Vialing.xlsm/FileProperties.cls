VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VCOMObject As Object
Property Get Count() As Long
    Count = VCOMObject.Count
End Property
Property Get Item(ByVal IndexOrNameOrID As Variant) As FileProperty
    Set Item = VCOMObject.Item(IndexOrNameOrID)
End Property
Property Get FilePath() As String
    FilePath = VCOMObject.FilePath
End Property

