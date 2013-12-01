VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "성적 처리"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdScoreCalc 
      Caption         =   "성적 계산"
      Height          =   495
      Left            =   2348
      TabIndex        =   35
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtAverage4 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtTotal4 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtIrum4 
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtScore41 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1440
      TabIndex        =   31
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtScore42 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtScore43 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtScore44 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3840
      TabIndex        =   28
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtAverage3 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtTotal3 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtIRum3 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtScore31 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1440
      TabIndex        =   24
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtScore32 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtScore33 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtScore34 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtAverage2 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtTotal2 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtIrum2 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtScore21 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtScore22 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtScore23 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtScore24 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtAverage1 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtScore14 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtScore13 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtScore12 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtScore11 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtIrum1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "평균"
      Height          =   180
      Left            =   6000
      TabIndex        =   12
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "총점"
      Height          =   180
      Left            =   5160
      TabIndex        =   10
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "과학"
      Height          =   180
      Left            =   3960
      TabIndex        =   8
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "사회"
      Height          =   180
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "수학"
      Height          =   180
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "국어"
      Height          =   180
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "이름"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdScoreCalc_Click()
    Dim Total1 As Integer
    Dim Total2 As Integer
    Dim Total3 As Integer
    Dim Total4 As Integer
    
    Total1 = CInt(txtScore11.Text) + CInt(txtScore12.Text) _
           + CInt(txtScore13.Text) + CInt(txtScore14.Text)
    Total2 = CInt(txtScore21.Text) + CInt(txtScore22.Text) _
           + CInt(txtScore23.Text) + CInt(txtScore24.Text)
    Total3 = CInt(txtScore31.Text) + CInt(txtScore32.Text) _
           + CInt(txtScore33.Text) + CInt(txtScore34.Text)
    Total4 = CInt(txtScore41.Text) + CInt(txtScore42.Text) _
           + CInt(txtScore43.Text) + CInt(txtScore44.Text)
    
    
    txtTotal1.Text = CStr(Total1)
    txtAverage1.Text = Format(Total1 / 4, "##0.0#")
    
    txtTotal2.Text = CStr(Total2)
    txtAverage2.Text = Format(Total2 / 4, "##0.0#")
    
    txtTotal3.Text = CStr(Total3)
    txtAverage3.Text = Format(Total3 / 4, "##0.0#")
    
    txtTotal4.Text = CStr(Total4)
    txtAverage4.Text = Format(Total4 / 4, "##0.0#")
End Sub
