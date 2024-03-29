VERSION 5.00
Begin VB.Form frmCalc4 
   BorderStyle     =   1  '단일 고정
   Caption         =   "계산기"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmCalc4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtValue1 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1553
      TabIndex        =   1
      ToolTipText     =   "숫자를 입력합니다"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtValue2 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1553
      TabIndex        =   3
      ToolTipText     =   "숫자를 입력합니다"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtResult 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H8000000B&
      Height          =   375
      Left            =   1553
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "＋"
      Height          =   375
      Left            =   1493
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblValue1 
      AutoSize        =   -1  'True
      Caption         =   "값1"
      Height          =   180
      Left            =   713
      TabIndex        =   0
      Top             =   360
      Width           =   270
   End
   Begin VB.Label lblValue2 
      AutoSize        =   -1  'True
      Caption         =   "값2"
      Height          =   180
      Left            =   713
      TabIndex        =   2
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "결과"
      Height          =   180
      Left            =   713
      TabIndex        =   5
      Top             =   2400
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuExit 
         Caption         =   "종료(&X)"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmCalc4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlus_Click()
    If Trim(txtValue1.Text) = "" Then
        MsgBox "숫자를 입력해야 합니다"
        txtValue1.SelStart = 0
        txtValue1.SelLength = Len(txtValue1.Text)
        txtValue1.SetFocus
        Exit Sub
    End If
    
    If Trim(txtValue2.Text) = "" Then
        MsgBox "숫자를 입력해야 합니다"
        txtValue2.SelStart = 0
        txtValue2.SelLength = Len(txtValue1.Text)
        txtValue2.SetFocus
        Exit Sub
    End If
    
    Dim Value1 As Single    '1 번째 텍스트상자의 단정도 소수 값
    Dim Value2 As Single    '2 번째 텍스트상자의 단정도 소수 값
    Dim Result As Single    '2 개의 단정도 소수 값의 합
    
    Value1 = CSng(txtValue1.Text)
    Value2 = CSng(txtValue2.Text)
    
    Result = Value1 + Value2
        
    txtResult.Text = Format(Result, "#,###.###")
End Sub

Private Sub mnuExit_Click()
    End
End Sub
