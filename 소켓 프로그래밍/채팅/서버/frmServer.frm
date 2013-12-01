VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  '단일 고정
   Caption         =   "채팅 - 서버"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "송신"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtInData 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtDisplay 
      Height          =   1695
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtRComp 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "입력"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "접속 대상"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    '데이터 송신(컴퓨터 이름+입력 데이터)
    Winsock1.SendData Winsock1.LocalHostName & ">" & txtInData.Text
    DisplayData txtInData.Text    '데이터 표시
End Sub

Private Sub Form_Load()
    Winsock1.LocalPort = 1001    '접속 요구 접수 포토 번호 설정
    Winsock1.Listen    '접속 요구 대기
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Winsock1.Close    'Winsock를 닫는다
    Unload Me
    End
End Sub

Private Sub txtInData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then    '리턴 키
        '데이터 송신(컴퓨터 이름+입력 데이터)
        Winsock1.SendData Winsock1.LocalHostName & ">" & txtInData.Text
        DisplayData txtInData.Text    '데이터 표시
    End If
End Sub

Private Sub Winsock1_Close()
    Winsock1.Close    'Winsock를 닫는다
    Winsock1.Listen    '재차, 접속 요구 대기
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then    'Winsock 상태(닫지 않았다)
        Winsock1.Close    'Winsock를 닫는다
    End If
    Winsock1.Accept requestID    '접속 처리
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String

    Winsock1.GetData dat    '데이터 수신
    If Left(dat, 2) = "##" Then    '컴퓨터 이름인가 어떤가 판정
        txtRComp.Text = Mid(dat, 3)    '컴퓨터 이름을 표시
    Else
        DisplayData dat    '데이터 표시
    End If
End Sub

Private Sub DisplayData(msg As String)
'송수신 데이터의 표시
    '입력 데이터 표시
    txtDisplay.Text = txtDisplay.Text & msg & vbCrLf
    '커서를 말미에 이동
    txtDisplay.SelStart = Len(txtDisplay.Text)
End Sub

