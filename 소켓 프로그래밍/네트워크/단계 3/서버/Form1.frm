VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "����"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.ListBox List1 
      Height          =   2040
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0007
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���� ����
Option Explicit              '���� ������ �����Ѵ�
Dim Word(1, 20) As String    '�ܾ� ������
Dim WsockNum As Integer      'Winsock ��Ʈ�� �迭��
Dim CompName(20) As String   '���� ��ǻ�� �̸���


Private Sub Command1_Click()
    Dim x As Integer

    For x = 0 To WsockNum
        Winsock1(x).Close
    Next x
    Unload Me
    End

End Sub

Private Sub Form_Load()
    '�ʱ� ����
    WsockNum = 0                       '��Ʈ�� �迭�� ÷��(0 ��°)
    Winsock1(0).LocalPort = 1001   '���ӿ� ��Ʈ ��ȣ ����
    Winsock1(0).Listen                  '���� �䱸�� ��ٸ���
    
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


Private Sub Winsock1_Close(Index As Integer)
    CompName(Index) = ""      '��ǻ�� �̸��� �����Ѵ�
    List1.RemoveItem Index    '����Ʈ������ �����Ѵ�
    List1.AddItem "", Index   '������ �����Ѵ�
    Winsock1(Index).Close     '������ �ݴ´�
End Sub


Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    '��Ʈ�� �迭 0�� ���� �䱸�� �־�
    '0 ��°�� ������ ���� �䱸�� �޴� �뵵�̹Ƿ�, ���� ���ӿ��� ������ �ʴ´�
    If Index = 0 Then
        WsockNum = WsockNum + 1    '÷�ڴ� 1�� ����
        '�űԷ� ������ �߰��Ѵ�
        Load Winsock1(WsockNum)
        '�ű� �߰��� ���� ��Ʈ ��ȣ�� �����Ѵ�(0���� �Ѵ�)
        Winsock1(WsockNum).LocalPort = 0
        '�ű� �߰��� �������� ���� ó���� �Ѵ�
        Winsock1(WsockNum).Accept requestID
    End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    '�μ� Index ��ȣ�� �������κ��� �����͸� �����Ѵ�
    Dim dat As String, ans As String    '������ ���ſ�, �۽ſ�
    Dim n As Integer

    Winsock1(Index).GetData dat    '������ ����
    'Ŭ���̾�Ʈ�� ��ǻ�� �̸��� ����
    If Left(dat, 2) = "##" Then
        CompName(Index) = Mid(dat, 3)        '��ǻ�� �̸� ���
        List1.AddItem CompName(Index), Index    'List1�� ���
        Exit Sub        '���ν����� ���� ���´�
    End If
    '����Ʈ�� ���� �����͸� ǥ���Ѵ�
    List1.RemoveItem Index    '����Ʈ�κ��� �����Ѵ�
    List1.AddItem CompName(Index) & "=" & dat, Index    '����Ʈ�� ������ �߰�
    For n = 0 To 20
        If dat = Word(0, n) Then
            Winsock1(Index).SendData Word(1, n)
            Exit For
        End If
    Next n
    If n = 21 Then
        Winsock1(Index).SendData "ã�� �� �����ϴ�"
    End If
End Sub
