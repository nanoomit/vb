VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Ŭ���̾�Ʈ"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows �⺻��
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˻�"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "���� ������"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "�۽� ������"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ȣ��Ʈ �̸�"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Winsock1.Close        '������ �ݴ´�
    Winsock1.RemoteHost = Text1.Text    '������ ��ǻ�� �̸� ����
    Winsock1.RemotePort = 1001        '���� ��ȣ ����
    Winsock1.Connect    '�����Ѵ�
End Sub

Private Sub Command2_Click()
    Winsock1.SendData Text2.Text      '�����͸� �۽��Ѵ�
    
End Sub

Private Sub Command3_Click()
    Winsock1.Close         ' ������ �ݴ´�
    Unload Me
    End
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String

    Winsock1.GetData dat    '�����κ����� ������ ����
    Text3.Text = dat

End Sub
