VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command2 
      Caption         =   "���� ȭ������"
      Height          =   375
      Left            =   1433
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ա�"
      Height          =   375
      Left            =   2813
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   653
      TabIndex        =   0
      Text            =   "0"
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
'������� ���� ����
Public nMoney As Integer
Public nTotalMoney  As Long

Private Sub Command1_Click()
    nMoney = CInt(Text1.Text)                '�Ա� �ݾ�
    nTotalMoney = nTotalMoney + nMoney       '�� �ܾ�
    
    Print "�Ա��� �ݾ��� " & nMoney & "�� �Դϴ�"
    Print "�� �ܾ��� " & nTotalMoney & "�� �Դϴ�"
End Sub

Private Sub Command2_Click()
    Me.Hide
    Form3.Show vbModal      '���� ���� ǥ���Ѵ�
End Sub
