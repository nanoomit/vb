VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   1  '���� ����
   Caption         =   "����"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtValue1 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1553
      TabIndex        =   1
      ToolTipText     =   "���ڸ� �Է��մϴ�"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtValue2 
      Alignment       =   1  '������ ����
      Height          =   375
      Left            =   1553
      TabIndex        =   3
      ToolTipText     =   "���ڸ� �Է��մϴ�"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtResult 
      Alignment       =   1  '������ ����
      BackColor       =   &H8000000B&
      Height          =   375
      Left            =   1553
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "��"
      Height          =   375
      Left            =   1493
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblValue1 
      AutoSize        =   -1  'True
      Caption         =   "��1"
      Height          =   180
      Left            =   713
      TabIndex        =   0
      Top             =   360
      Width           =   270
   End
   Begin VB.Label lblValue2 
      AutoSize        =   -1  'True
      Caption         =   "��2"
      Height          =   180
      Left            =   713
      TabIndex        =   2
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   180
      Left            =   713
      TabIndex        =   5
      Top             =   2400
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Begin VB.Menu mnuExit 
         Caption         =   "����(&X)"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlus_Click()
On Error GoTo Error_rtn

    Dim Value1 As Integer  '1 ��° �ؽ�Ʈ������ ���� ��
    Dim Value2 As Integer  '2 ��° �ؽ�Ʈ������ ���� ��
    Dim Result As Integer   '2 ���� ���� ���� ��

    Value1 = CInt(txtValue1.Text)
    Value2 = CInt(txtValue2.Text)

    Result = Value1 + Value2
    
    txtResult.Text = Format(Result, "#,###")
    
    Exit Sub
Error_rtn:
    
    MsgBox (Err.Number & " " & Err.Description)
End Sub

Private Sub mnuExit_Click()
    End
End Sub
