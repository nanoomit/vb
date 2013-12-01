VERSION 5.00
Begin VB.UserControl ConvertControl 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   945
   ScaleWidth      =   4800
   Begin VB.TextBox txtMeter 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtCentiMeter 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "변환 ▶"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "㎝"
      Height          =   180
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   180
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   165
   End
End
Attribute VB_Name = "ConvertControl"
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
    Refresh
End Property
Public Function ChangeToCM() As Long
    Dim cm As Long
    cm = CLng(nMeter) * 100
    ChangeToCM = cm
End Function

Private Sub cmdConvert_Click()
    nMeter = CInt(txtMeter.Text)
    
    txtCentiMeter.Text = Format(ChangeToCM)
End Sub
