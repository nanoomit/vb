VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1260
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ա�"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Form1.Show vbModal  '�Ա� ���� ǥ��
End Sub

Private Sub Command2_Click()
    Me.Hide
    Form2.Show vbModal  '���� ���� ǥ��
End Sub

Private Sub Command3_Click()
    End                 ' ���α׷��� ����
End Sub
