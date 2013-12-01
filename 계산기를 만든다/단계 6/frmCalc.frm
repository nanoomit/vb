VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   1  '단일 고정
   Caption         =   "계산기"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmCalc.frx":0000
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
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlus_Click()
On Error GoTo Error_rtn

    Dim Value1 As Integer  '1 번째 텍스트상자의 정수 값
    Dim Value2 As Integer  '2 번째 텍스트상자의 정수 값
    Dim Result As Integer   '2 개의 정수 값의 합

    Value1 = CInt(txtValue1.Text)
    Value2 = CInt(txtValue2.Text)

    Result = Value1 + Value2
    
    txtResult.Text = Format(Result, "#,###")
    
    Exit Sub
Error_rtn:
    
    MsgBox (Err.Number & " " & Err.Description)
End Sub

Private Sub mnuExit_Click()
    End
End Sub
