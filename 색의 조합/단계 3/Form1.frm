VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RGB(���� ����)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "���� ����"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "16���� ��:"
      Height          =   180
      Left            =   2640
      TabIndex        =   10
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "10���� ��:"
      Height          =   180
      Left            =   2640
      TabIndex        =   8
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BorderStyle     =   1  '���� ����
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�Ķ�(0~255)"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���(0~255)"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����(0~255)"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MakeColors
End Sub

Public Sub MakeColors()
    '���� ����
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    '���ڿ��� ����Ʈ�� ��ȯ
    Red = CByte(Text1.Text)
    Green = CByte(Text2.Text)
    Blue = CByte(Text3.Text)
    
    '�� ��Ʈ���� ��� ���� �����Ѵ�
    Label4.BackColor = RGB(Red, Green, Blue)
    
    '�� ��Ʈ���� ���� RGB�� ǥ���Ѵ�
    Label4.Caption = "RGB(" & Red & "," & Green & "," & Blue & ")"
    
    'RGB ���� 10������ ǥ���Ѵ�
    Label6.Caption = RGB(Red, Green, Blue)
    'RGB ���� 16������ ǥ���Ѵ�
    Label8.Caption = Hex(RGB(Red, Green, Blue))
End Sub


Private Sub Form_Load()
    MakeColors
End Sub

Private Sub Text1_Change()
    MakeColors
End Sub

Private Sub Text2_Change()
    MakeColors
End Sub

Private Sub Text3_Change()
    MakeColors
End Sub
