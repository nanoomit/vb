VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "����"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "���� ���"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                  '���� ������ �����Ѵ�
Dim Word(1, 20) As String        '�ܾ� ��Ͽ� �迭

Private Sub Command1_Click()
    Winsock1.Close               '��Ʈ��ũ ����
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    
    Winsock1.LocalPort = 1001    '���� �䱸 ���� ���� ��ȣ ����
    Winsock1.Listen              '���� �䱸 ���
    '�ܾ� ���
    Word(0, 0) = "english"
    Word(1, 0) = "���� "
    Word(0, 1) = "apple"
    Word(1, 1) = "��� "
    Word(0, 2) = "japan"
    Word(1, 2) = "�Ϻ� "
    Word(0, 3) = "computer"
    Word(1, 3) = "��ǻ�� "
    Word(0, 4) = "sun"
    Word(1, 4) = "�¾� "
    Word(0, 5) = "moon"
    Word(1, 5) = "��"

End Sub

Private Sub Winsock1_Close()
    Text2.Text = ""        '��ǻ�� �̸��� �����
    Winsock1.Close         '������ �ݴ´�
    Winsock1.Listen        '���� �䱸�� ��ٸ���

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then    'Winsock ����(���� �ʾҴ�)
        Winsock1.Close    'Winsock�� �ݴ´�
    End If
    Winsock1.Accept requestID    '���� ó��

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String, ans As String
    Dim n As Integer

    Winsock1.GetData dat
    'Ŭ���̾�Ʈ�� ��ǻ�� �̸� ����
    If Left(dat, 2) = "##" Then    '��ǻ�� �̸����� ����� ����
        Text2.Text = Mid(dat, 3)    '��ǻ�� �̸��� ǥ��
        Exit Sub        '���ν����� ���� ������
    End If
    Text1.Text = dat
    For n = 0 To 20
        If dat = Word(0, n) Then
            Winsock1.SendData Word(1, n)
            Exit For
        End If
    Next n
    If n = 21 Then
        Winsock1.SendData "ã�� �� �����ϴ� "
    End If


End Sub
