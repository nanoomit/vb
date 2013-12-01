VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "파일의 처리"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6375
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   5520
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4680
      TabIndex        =   17
      Text            =   "Text6"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   3840
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2040
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1200
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   360
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "파일을 닫는다"
      Height          =   180
      Left            =   3600
      TabIndex        =   15
      Top             =   2760
      Width           =   1140
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "데이터 읽기"
      Height          =   180
      Left            =   3600
      TabIndex        =   14
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "파일을 연다(열기 모드)"
      Height          =   180
      Left            =   3600
      TabIndex        =   13
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "데이터 읽기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3600
      TabIndex        =   12
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "파일을 닫는다"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "데이터 쓰기"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "파일을 연다(쓰기 모드)"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "데이터 보관"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "파일 이름 입력"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'①파일을 연다(쓰기 모드)
'  Open 명령
'      파일 이름:Text1.Text
'      모드:Output (쓰기 모드)
'      파일 번호:1 (연 파일에 할당되는 번호)
'
    Open Text1.Text For Output As #1
End Sub

Private Sub Command2_Click()
'②데이터 쓰기
'   Text2~Text4의 내용(Text 속성값)을 파일에 기록한다
'
    Print #1, Text2.Text     '기록 부분
    Print #1, Text3.Text
    Print #1, Text4.Text
End Sub

Private Sub Command3_Click()
'③파일을 닫는다
    Close #1
End Sub

Private Sub Command4_Click()
'①파일을 연다(읽기 모드)
'   Open 명령
'       파일 이름:Text1.Text
'       모드:Input (읽기 모드)
'       파일 번호:1 (연 파일에 할당되는 번호)
'
    Open Text1.Text For Input As #1
End Sub

Private Sub Command5_Click()
'②데이터 읽기
'   파일로부터 읽은 데이터를 Text5~Text7의 텍스트상자에 표시한다
'
    Dim work(2) As String

    Input #1, work(0)       '1번째 데이터의 읽기
    Input #1, work(1)       '2번째 데이터의 읽기
    Input #1, work(2)       '3번째 데이터의 읽기
    Text5.Text = work(0)    '데이터 표시
    Text6.Text = work(1)    '데이터 표시
    Text7.Text = work(2)    '데이터 표시
End Sub

Private Sub Command6_Click()
'③파일을 닫는다
    Close #1
End Sub

Private Sub Form_Load()
'데이터나 객체(부품)의 속성 등을 초기화 한다
    'ChDrive App.Path   '네트워크 드라이브의 경우에는 필요없다
    ChDir App.Path

    Text1.Text = "test.txt"     '파일 이름="test.txt"
    Text2.Text = ""             '텍스트 상자의 내용을 지운다
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
End Sub
