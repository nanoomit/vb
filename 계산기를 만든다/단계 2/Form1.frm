VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '���� ����
   Caption         =   "����"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox Text1 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1553
      TabIndex        =   1
      ToolTipText     =   "���ڸ� �Է��մϴ�"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1553
      TabIndex        =   3
      ToolTipText     =   "���ڸ� �Է��մϴ�"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1553
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��"
      Height          =   375
      Left            =   1493
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��1"
      Height          =   180
      Left            =   713
      TabIndex        =   0
      Top             =   360
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��2"
      Height          =   180
      Left            =   713
      TabIndex        =   2
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   180
      Left            =   713
      TabIndex        =   5
      Top             =   2400
      Width           =   360
   End
   Begin VB.Menu File 
      Caption         =   "����(&F)"
      Begin VB.Menu Exit 
         Caption         =   "����(&X)"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click()
    End
End Sub
