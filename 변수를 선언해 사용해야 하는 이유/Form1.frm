VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "변수는 선언해야 한다"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command3 
      Caption         =   "데이터 형 선언"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dim만 선언"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "변수 미선언"
      Height          =   495
      Left            =   173
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   173
      TabIndex        =   3
      Top             =   1560
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    a = Timer               '시작 시간(초)

    For m = 1 To 1000
        For n = 1 To 10000
            da = n * 2
        Next n
    Next m
    
    b = Timer               '종료 시간(초)
    
    Label1.Caption = (b - a) & "초 소요"
End Sub

Private Sub Command2_Click()
   Dim a, b
   Dim m, n, da
   
   a = Timer                '시작 시간(초)

   For m = 1 To 1000
       For n = 1 To 10000
           da = n * 2
       Next n
   Next m
   
   b = Timer                '종료 시간(초)
   
   Label2.Caption = (b - a) & "초 소요"
End Sub

Private Sub Command3_Click()
   Dim a As Single, b As Single
   Dim m As Integer, n As Integer, da As Integer
   
   a = Timer                '시작 시간(초)
    
   For m = 1 To 1000
       For n = 1 To 10000
           da = n * 2
       Next n
   Next m
   
   b = Timer                '종료 시간(초)
   
   Label3.Caption = (b - a) & "초 소요"
End Sub
