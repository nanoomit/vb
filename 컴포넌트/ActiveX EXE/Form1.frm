VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text3 
      Alignment       =   2  '가운데 맞춤
      Height          =   375
      Left            =   1935
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2475
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "＋"
      Height          =   420
      Left            =   1965
      TabIndex        =   4
      Top             =   1620
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '가운데 맞춤
      Height          =   375
      Left            =   1950
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   810
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '가운데 맞춤
      Height          =   375
      Left            =   1965
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   165
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "값 2"
      Height          =   180
      Left            =   1275
      TabIndex        =   3
      Top             =   930
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "값 1"
      Height          =   180
      Left            =   1275
      TabIndex        =   1
      Top             =   285
      Width           =   330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim IsSum As New ExeSimpleCom.Class1
    
    Dim Value1 As Integer
    Dim Value2 As Integer
    Dim Result As Integer
    
    Value1 = CInt(Text1.Text)
    Value2 = CInt(Text2.Text)
    
    Result = IsSum.SumTwoValues(Value1, Value2)
    
    Text3.Text = Result

End Sub

