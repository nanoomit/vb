VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������ ó��"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6375
   StartUpPosition =   2  'ȭ�� ���
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
      Caption         =   "������ �ݴ´�"
      Height          =   180
      Left            =   3600
      TabIndex        =   15
      Top             =   2760
      Width           =   1140
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "������ �б�"
      Height          =   180
      Left            =   3600
      TabIndex        =   14
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "������ ����(���� ���)"
      Height          =   180
      Left            =   3600
      TabIndex        =   13
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "������ �б�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������ �ݴ´�"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "������ ����"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "������ ����(���� ���)"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "������ ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���� �̸� �Է�"
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
'�������� ����(���� ���)
'  Open ���
'      ���� �̸�:Text1.Text
'      ���:Output (���� ���)
'      ���� ��ȣ:1 (�� ���Ͽ� �Ҵ�Ǵ� ��ȣ)
'
    Open Text1.Text For Output As #1
End Sub

Private Sub Command2_Click()
'�赥���� ����
'   Text2~Text4�� ����(Text �Ӽ���)�� ���Ͽ� ����Ѵ�
'
    Print #1, Text2.Text     '��� �κ�
    Print #1, Text3.Text
    Print #1, Text4.Text
End Sub

Private Sub Command3_Click()
'�������� �ݴ´�
    Close #1
End Sub

Private Sub Command4_Click()
'�������� ����(�б� ���)
'   Open ���
'       ���� �̸�:Text1.Text
'       ���:Input (�б� ���)
'       ���� ��ȣ:1 (�� ���Ͽ� �Ҵ�Ǵ� ��ȣ)
'
    Open Text1.Text For Input As #1
End Sub

Private Sub Command5_Click()
'�赥���� �б�
'   ���Ϸκ��� ���� �����͸� Text5~Text7�� �ؽ�Ʈ���ڿ� ǥ���Ѵ�
'
    Dim work(2) As String

    Input #1, work(0)       '1��° �������� �б�
    Input #1, work(1)       '2��° �������� �б�
    Input #1, work(2)       '3��° �������� �б�
    Text5.Text = work(0)    '������ ǥ��
    Text6.Text = work(1)    '������ ǥ��
    Text7.Text = work(2)    '������ ǥ��
End Sub

Private Sub Command6_Click()
'�������� �ݴ´�
    Close #1
End Sub

Private Sub Form_Load()
'�����ͳ� ��ü(��ǰ)�� �Ӽ� ���� �ʱ�ȭ �Ѵ�
    'ChDrive App.Path   '��Ʈ��ũ ����̺��� ��쿡�� �ʿ����
    ChDir App.Path

    Text1.Text = "test.txt"     '���� �̸�="test.txt"
    Text2.Text = ""             '�ؽ�Ʈ ������ ������ �����
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
End Sub
