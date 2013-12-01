VERSION 5.00
Begin VB.UserControl MToCmConvert 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1605
   ScaleHeight     =   750
   ScaleWidth      =   1605
   Begin VB.CommandButton cmdConvert 
      Caption         =   "º¯È¯ ¢º"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "MToCmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event Click()
Private nMeter As Integer


Public Property Get Meter() As Variant
    Meter = nMeter
End Property

Public Property Let Meter(ByVal vNewValue As Variant)
    nMeter = vNewValue
    Refresh
End Property

Public Function ChangeToCm() As Long
    Dim lCm As Long
    lCm = CLng(nMeter) * 100
    ChangeToCm = lCm
End Function

Private Sub cmdConvert_Click()
    RaiseEvent Click
End Sub
