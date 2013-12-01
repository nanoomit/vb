VERSION 5.00
Object = "{FDFA56D1-A7A0-4ACD-BF53-E3DF43ACB57B}#5.0#0"; "ControlConvert.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows 기본값
   Begin ControlConvert.MToCmConvert MToCmConvert1 
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
   End
   Begin VB.TextBox txtMeter 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtCentiMeter 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "㎝"
      Height          =   180
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   180
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MToCmConvert1_Click()
    MToCmConvert1.Meter = CInt(txtMeter.Text)
    Dim lCm As Long
    lCm = MToCmConvert1.ChangeToCm
    txtCentiMeter.Text = Format(lCm, "#,###")
End Sub

