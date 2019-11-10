VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GPSExifProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VCOMObject As Object
Property Get FilePath() As String
    FilePath = VCOMObject.FilePath
End Property
Property Get DateTimeOriginal() As String
    DateTimeOriginal = VCOMObject.DateTimeOriginal
End Property
Property Get GPSVersionID() As String
    GPSVersionID = VCOMObject.GPSVersionID
End Property
Property Get GPSLatitudeDecimal() As Variant
    Call VCOMObject.AssignVar(GPSLatitudeDecimal, VCOMObject.GPSLatitudeDecimal)
End Property
Property Get GPSLongitudeDecimal() As Variant
    Call VCOMObject.AssignVar(GPSLongitudeDecimal, VCOMObject.GPSLongitudeDecimal)
End Property
Property Get GPSAltitudeDecimal() As Variant
    Call VCOMObject.AssignVar(GPSAltitudeDecimal, VCOMObject.GPSAltitudeDecimal)
End Property
Property Get GPSTimeStamp() As Variant
    Call VCOMObject.AssignVar(GPSTimeStamp, VCOMObject.GPSTimeStamp)
End Property
Property Get GPSSatellites() As String
    GPSSatellites = VCOMObject.GPSSatellites
End Property
Property Get GPSStatus() As String
    GPSStatus = VCOMObject.GPSStatus
End Property
Property Get GPSMeasureMode() As String
    GPSMeasureMode = VCOMObject.GPSMeasureMode
End Property
Property Get GPSDOPDecimal() As Variant
    Call VCOMObject.AssignVar(GPSDOPDecimal, VCOMObject.GPSDOPDecimal)
End Property
Property Get GPSSpeedRef() As String
    GPSSpeedRef = VCOMObject.GPSSpeedRef
End Property
Property Get GPSSpeedDecimal() As Variant
    Call VCOMObject.AssignVar(GPSSpeedDecimal, VCOMObject.GPSSpeedDecimal)
End Property
Property Get GPSTrackRef() As String
    GPSTrackRef = VCOMObject.GPSTrackRef
End Property
Property Get GPSTrackDecimal() As Variant
    Call VCOMObject.AssignVar(GPSTrackDecimal, VCOMObject.GPSTrackDecimal)
End Property
Property Get GPSImgDirectionRef() As String
    GPSImgDirectionRef = VCOMObject.GPSImgDirectionRef
End Property
Property Get GPSImgDirectionDecimal() As Variant
    Call VCOMObject.AssignVar(GPSImgDirectionDecimal, VCOMObject.GPSImgDirectionDecimal)
End Property
Property Get GPSMapDatum() As String
    GPSMapDatum = VCOMObject.GPSMapDatum
End Property
Property Get GPSDestLatitudeDecimal() As Variant
    Call VCOMObject.AssignVar(GPSDestLatitudeDecimal, VCOMObject.GPSDestLatitudeDecimal)
End Property
Property Get GPSDestLongitudeDecimal() As Variant
    Call VCOMObject.AssignVar(GPSDestLongitudeDecimal, VCOMObject.GPSDestLongitudeDecimal)
End Property
Property Get GPSDestBearingRef() As String
    GPSDestBearingRef = VCOMObject.GPSDestBearingRef
End Property
Property Get GPSDestBearingDecimal() As Variant
    Call VCOMObject.AssignVar(GPSDestBearingDecimal, VCOMObject.GPSDestBearingDecimal)
End Property
Property Get GPSDestDistanceRef() As String
    GPSDestDistanceRef = VCOMObject.GPSDestDistanceRef
End Property
Property Get GPSDestDistanceDecimal() As Variant
    Call VCOMObject.AssignVar(GPSDestDistanceDecimal, VCOMObject.GPSDestDistanceDecimal)
End Property
Property Get GPSProcessingMethod() As String
    GPSProcessingMethod = VCOMObject.GPSProcessingMethod
End Property
Property Get GPSAreaInformation() As String
    GPSAreaInformation = VCOMObject.GPSAreaInformation
End Property
Property Get GPSDateStamp() As String
    GPSDateStamp = VCOMObject.GPSDateStamp
End Property
Property Get GPSDifferentialCorrection() As Variant
    Call VCOMObject.AssignVar(GPSDifferentialCorrection, VCOMObject.GPSDifferentialCorrection)
End Property

