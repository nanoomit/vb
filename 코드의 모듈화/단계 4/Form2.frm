VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
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
   Begin VB.TextBox Text1 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   593
      TabIndex        =   1
      Text            =   "0"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   2753
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    nMoney = CInt(Text1.Text)                '�Ա� �ݾ�
    nTotalMoney = nTotalMoney - nMoney       '�� �ܾ�
    
    Print "������ �ݾ��� " & nMoney & "�� �Դϴ�"
    Print "�� �ܾ��� " & nTotalMoney & "�� �Դϴ�"
End Sub

Private Sub Command2_Click()
    Me.Hide
    Form3.Show vbModal       '���� ���� ǥ���Ѵ�
End Sub
