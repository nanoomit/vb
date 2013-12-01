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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1553
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1553
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1553
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "＋"
      Height          =   375
      Left            =   1493
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "값1"
      Height          =   180
      Left            =   713
      TabIndex        =   6
      Top             =   360
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "값2"
      Height          =   180
      Left            =   713
      TabIndex        =   5
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "결과"
      Height          =   180
      Left            =   713
      TabIndex        =   4
      Top             =   2400
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
