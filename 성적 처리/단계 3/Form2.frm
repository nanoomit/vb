VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "성적 입력"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdScoreAdd 
      Caption         =   "성적 등록"
      Height          =   375
      Left            =   1433
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtSci 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Text            =   "0"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtSahe 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Text            =   "0"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtMath 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtGuk 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtIrum 
      Height          =   375
      IMEMode         =   9  '한글 전자
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "과학"
      Height          =   180
      Left            =   960
      TabIndex        =   8
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "사회"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "수학"
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "국어"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "이름"
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdScoreAdd_Click()
    If Trim(txtIrum.Text) = "" Then         '이름이 입력되어 있지 않으면
        MsgBox "이름을 입력해 주세요"
        Exit Sub
    End If
    
    If IsNumeric(txtGuk.Text) <> True Then
        MsgBox "국어 성적은 숫자여야 합니다"
        Exit Sub
    End If
    
    If IsNumeric(txtMath.Text) <> True Then
        MsgBox "수학 성적은 숫자여야 합니다"
        Exit Sub
    End If
    
    If IsNumeric(txtSahe.Text) <> True Then
        MsgBox "사회 성적은 숫자여야 합니다"
        Exit Sub
    End If
    
    If IsNumeric(txtSci.Text) <> True Then
        MsgBox "과학 성적은 숫자여야 합니다"
        Exit Sub
    End If
    
    '모듈 데이터
    Seq = Seq + 1
    
    ScoreData(Seq).Irum = Trim(txtIrum.Text)
    ScoreData(Seq).Guk = CInt(txtGuk.Text)
    ScoreData(Seq).Math = CInt(txtMath.Text)
    ScoreData(Seq).Sahe = CInt(txtSahe.Text)
    ScoreData(Seq).Sci = CInt(txtSci.Text)
    
    'FlexGrid에 써 넣는다
    Form1.MSFlexGrid1.Row = Seq
    Form1.MSFlexGrid1.Col = 1
    Form1.MSFlexGrid1.Text = CStr(Seq)
    
    Form1.MSFlexGrid1.Col = 2
    Form1.MSFlexGrid1.Text = ScoreData(Seq).Irum
    
    Form1.MSFlexGrid1.Col = 3
    Form1.MSFlexGrid1.Text = ScoreData(Seq).Guk
    
    Form1.MSFlexGrid1.Col = 4
    Form1.MSFlexGrid1.Text = ScoreData(Seq).Math
    
    Form1.MSFlexGrid1.Col = 5
    Form1.MSFlexGrid1.Text = ScoreData(Seq).Sahe

    Form1.MSFlexGrid1.Col = 6
    Form1.MSFlexGrid1.Text = ScoreData(Seq).Sci
    
    Me.Hide
End Sub

