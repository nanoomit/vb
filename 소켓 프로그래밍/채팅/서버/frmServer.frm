VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  '���� ����
   Caption         =   "ä�� - ����"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "�۽�"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtInData 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtDisplay 
      Height          =   1695
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtRComp 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�Է�"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���� ���"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    '������ �۽�(��ǻ�� �̸�+�Է� ������)
    Winsock1.SendData Winsock1.LocalHostName & ">" & txtInData.Text
    DisplayData txtInData.Text    '������ ǥ��
End Sub

Private Sub Form_Load()
    Winsock1.LocalPort = 1001    '���� �䱸 ���� ���� ��ȣ ����
    Winsock1.Listen    '���� �䱸 ���
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Winsock1.Close    'Winsock�� �ݴ´�
    Unload Me
    End
End Sub

Private Sub txtInData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then    '���� Ű
        '������ �۽�(��ǻ�� �̸�+�Է� ������)
        Winsock1.SendData Winsock1.LocalHostName & ">" & txtInData.Text
        DisplayData txtInData.Text    '������ ǥ��
    End If
End Sub

Private Sub Winsock1_Close()
    Winsock1.Close    'Winsock�� �ݴ´�
    Winsock1.Listen    '����, ���� �䱸 ���
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then    'Winsock ����(���� �ʾҴ�)
        Winsock1.Close    'Winsock�� �ݴ´�
    End If
    Winsock1.Accept requestID    '���� ó��
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String

    Winsock1.GetData dat    '������ ����
    If Left(dat, 2) = "##" Then    '��ǻ�� �̸��ΰ� ��� ����
        txtRComp.Text = Mid(dat, 3)    '��ǻ�� �̸��� ǥ��
    Else
        DisplayData dat    '������ ǥ��
    End If
End Sub

Private Sub DisplayData(msg As String)
'�ۼ��� �������� ǥ��
    '�Է� ������ ǥ��
    txtDisplay.Text = txtDisplay.Text & msg & vbCrLf
    'Ŀ���� ���̿� �̵�
    txtDisplay.SelStart = Len(txtDisplay.Text)
End Sub

