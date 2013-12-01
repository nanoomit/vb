VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "서버"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "종료"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "접속 대상"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                  '변수 선언을 강제한다
Dim Word(1, 20) As String        '단어 등록용 배열

Private Sub Command1_Click()
    Winsock1.Close               '네트워크 절단
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    
    Winsock1.LocalPort = 1001    '접속 요구 접수 포토 번호 설정
    Winsock1.Listen              '접속 요구 대기
    '단어 등록
    Word(0, 0) = "english"
    Word(1, 0) = "영어 "
    Word(0, 1) = "apple"
    Word(1, 1) = "사과 "
    Word(0, 2) = "japan"
    Word(1, 2) = "일본 "
    Word(0, 3) = "computer"
    Word(1, 3) = "컴퓨터 "
    Word(0, 4) = "sun"
    Word(1, 4) = "태양 "
    Word(0, 5) = "moon"
    Word(1, 5) = "달"

End Sub

Private Sub Winsock1_Close()
    Text2.Text = ""        '컴퓨터 이름을 지운다
    Winsock1.Close         '접속을 닫는다
    Winsock1.Listen        '접속 요구를 기다린다

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then    'Winsock 상태(닫지 않았다)
        Winsock1.Close    'Winsock를 닫는다
    End If
    Winsock1.Accept requestID    '접속 처리

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String, ans As String
    Dim n As Integer

    Winsock1.GetData dat
    '클라이언트의 컴퓨터 이름 수신
    If Left(dat, 2) = "##" Then    '컴퓨터 이름인지 어떤지를 판정
        Text2.Text = Mid(dat, 3)    '컴퓨터 이름을 표시
        Exit Sub        '프로시저를 빠져 나간다
    End If
    Text1.Text = dat
    For n = 0 To 20
        If dat = Word(0, n) Then
            Winsock1.SendData Word(1, n)
            Exit For
        End If
    Next n
    If n = 21 Then
        Winsock1.SendData "찾을 수 없습니다 "
    End If


End Sub
