VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "ǥ��"
      Height          =   375
      Left            =   3113
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1913
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������ �Է�"
      Height          =   180
      Left            =   593
      TabIndex        =   0
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '���� ����
    Dim nGuGu As Integer
    Dim i As Integer
    Dim strLine As String
    Dim strTotal As String
    
    nGuGu = CInt(Text1.Text)
        
    'Do...Loop ����
    Do
        i = i + 1
        strLine = nGuGu & " �� " & i & " �� " & nGuGu * i & vbCrLf
        strTotal = strTotal & strLine
    Loop While i < 9    '���� ���� �˻�
    
    
    Label2.Caption = strTotal
End Sub
