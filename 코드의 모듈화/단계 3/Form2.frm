VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
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
   Begin VB.TextBox Text1 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   593
      TabIndex        =   1
      Text            =   "0"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "인출"
      Height          =   375
      Left            =   2753
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Form1.nMoney = CInt(Text1.Text)                            '입금 금액
    Form1.nTotalMoney = Form1.nTotalMoney - Form1.nMoney       '총 잔액
    
    Print "인출한 금액은 " & Form1.nMoney & "원 입니다"
    Print "총 잔액은 " & Form1.nTotalMoney & "원 입니다"
End Sub

Private Sub Command2_Click()
    Me.Hide
    Form3.Show vbModal       '메인 폼을 표시한다
End Sub
