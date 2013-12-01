VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "서버"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox List1 
      Height          =   2040
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0007
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "종료"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'전역 변수
Option Explicit              '변수 선언을 강제한다
Dim Word(1, 20) As String    '단어 보관용
Dim WsockNum As Integer      'Winsock 컨트롤 배열용
Dim CompName(20) As String   '원격 컴퓨터 이름용


Private Sub Command1_Click()
    Dim x As Integer

    For x = 0 To WsockNum
        Winsock1(x).Close
    Next x
    Unload Me
    End

End Sub

Private Sub Form_Load()
    '초기 설정
    WsockNum = 0                       '컨트롤 배열의 첨자(0 번째)
    Winsock1(0).LocalPort = 1001   '접속용 포트 번호 설정
    Winsock1(0).Listen                  '접속 요구를 기다린다
    
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


Private Sub Winsock1_Close(Index As Integer)
    CompName(Index) = ""      '컴퓨터 이름을 삭제한다
    List1.RemoveItem Index    '리스트에서도 삭제한다
    List1.AddItem "", Index   '공백을 삽입한다
    Winsock1(Index).Close     '접속을 닫는다
End Sub


Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    '컨트롤 배열 0에 접속 요구가 있어
    '0 번째의 윈속은 접속 요구를 받는 용도이므로, 실제 접속에는 사용되지 않는다
    If Index = 0 Then
        WsockNum = WsockNum + 1    '첨자는 1씩 증가
        '신규로 윈속을 추가한다
        Load Winsock1(WsockNum)
        '신규 추가한 윈속 포트 번호를 설정한다(0으로 한다)
        Winsock1(WsockNum).LocalPort = 0
        '신규 추가한 윈속으로 접속 처리를 한다
        Winsock1(WsockNum).Accept requestID
    End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    '인수 Index 번호의 윈속으로부터 데이터를 수신한다
    Dim dat As String, ans As String    '데이터 수신용, 송신용
    Dim n As Integer

    Winsock1(Index).GetData dat    '데이터 수신
    '클라이언트의 컴퓨터 이름을 수신
    If Left(dat, 2) = "##" Then
        CompName(Index) = Mid(dat, 3)        '컴퓨터 이름 등록
        List1.AddItem CompName(Index), Index    'List1에 등록
        Exit Sub        '프로시저를 빠져 나온다
    End If
    '리스트에 수신 데이터를 표시한다
    List1.RemoveItem Index    '리스트로부터 삭제한다
    List1.AddItem CompName(Index) & "=" & dat, Index    '리스트에 데이터 추가
    For n = 0 To 20
        If dat = Word(0, n) Then
            Winsock1(Index).SendData Word(1, n)
            Exit For
        End If
    Next n
    If n = 21 Then
        Winsock1(Index).SendData "찾을 수 없습니다"
    End If
End Sub
