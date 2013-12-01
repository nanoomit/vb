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
   Begin VB.CommandButton Command2 
      Caption         =   "메인 화면으로"
      Height          =   375
      Left            =   1433
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "입금"
      Height          =   375
      Left            =   2813
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   653
      TabIndex        =   0
      Text            =   "0"
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'명시적인 변수 선언
Public nMoney As Integer
Public nTotalMoney  As Long

Private Sub Command1_Click()
    nMoney = CInt(Text1.Text)                '입금 금액
    nTotalMoney = nTotalMoney + nMoney       '총 잔액
    
    Print "입금한 금액은 " & nMoney & "원 입니다"
    Print "총 잔액은 " & nTotalMoney & "원 입니다"
End Sub

Private Sub Command2_Click()
    Me.Hide
    Form3.Show vbModal      '메인 폼을 표시한다
End Sub
