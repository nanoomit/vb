VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdConvert 
      Caption         =   "변환 ▶"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtCentiMeter 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtMeter 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   180
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "㎝"
      Height          =   180
      Left            =   4080
      TabIndex        =   1
      Top             =   480
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
    Dim Meter As Integer
    Dim CentiMeter As Long
    
    Meter = CInt(txtMeter.Text)
    CentiMeter = CLng(Meter) * 100
    
    txtCentiMeter.Text = Format(CentiMeter, "#,###")
End Sub

