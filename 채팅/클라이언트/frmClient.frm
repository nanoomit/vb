VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  '단일 고정
   Caption         =   "채팅 - 클라이언트"
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
   Begin VB.CommandButton cmdConnect 
      Caption         =   "접속"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
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
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtInData 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
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
      TabIndex        =   1
      Top             =   120
      Width           =   2415
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
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
    Winsock1.Close    '접속을 닫는다
    Winsock1.RemoteHost = txtRComp.Text    '서버의 컴퓨터 이름 설정
    Winsock1.RemotePort = 1001    '포토 번호 설정
    Winsock1.Connect    '접속한다
End Sub

Private Sub cmdSend_Click()
    '데이터 송신(컴퓨터 이름+입력 데이터)
    Winsock1.SendData Winsock1.LocalHostName & ">" & txtInData.Text
    DisplayData txtInData.Text    '데이터 표시
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

Private Sub txtRComp_Change()

End Sub

Private Sub Winsock1_Close()
    Winsock1.Close        'Winsock를 닫는다
    txtRComp.Text = ""    '접속처 컴퓨터 이름을 소거
End Sub

Private Sub Winsock1_Connect()
    '접속 완료 후, 서버에 클라이언트의 컴퓨터 이름을 송신
    Winsock1.SendData "##" & Winsock1.LocalHostName
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String

    Winsock1.GetData dat    '서버로부터의 데이터 수신
    DisplayData dat    '데이터 표시
End Sub

Private Sub DisplayData(msg As String)
'송수신 데이터의 표시
    '입력 데이터 표시
    txtDisplay.Text = txtDisplay.Text & msg & vbCrLf
    '커서를 말미에 이동
    txtDisplay.SelStart = Len(txtDisplay.Text)
End Sub


