VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RGB(색의 조합)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "색의 조합"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  '단일 고정
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "파랑(0~255)"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "녹색(0~255)"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "빨강(0~255)"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MakeColors
End Sub

Public Sub MakeColors()
    '변수 선언
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    '문자열을 바이트로 변환
    Red = CByte(Text1.Text)
    Green = CByte(Text2.Text)
    Blue = CByte(Text3.Text)
    
    '라벨 컨트롤의 배경 색을 지정한다
    Label4.BackColor = RGB(Red, Green, Blue)
End Sub

Private Sub Form_Load()
    MakeColors
End Sub

Private Sub Text1_Change()
    MakeColors
End Sub

Private Sub Text2_Change()
    MakeColors
End Sub

Private Sub Text3_Change()
    MakeColors
End Sub
