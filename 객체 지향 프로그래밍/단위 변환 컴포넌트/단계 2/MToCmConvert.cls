VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MToCmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private nMeter As Integer


Public Property Get Meter() As Variant
    Meter = nMeter
End Property

Public Property Let Meter(ByVal vNewValue As Variant)
    nMeter = vNewValue
End Property

Public Function ChangeToCm() As Long
    Dim lCm As Long
    lCm = CLng(nMeter) * 100
    ChangeToCm = lCm
End Function
