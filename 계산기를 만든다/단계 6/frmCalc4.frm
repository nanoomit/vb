VERSION 5.00
Begin VB.Form frmCalc4 
   BorderStyle     =   1  '���� ����
   Caption         =   "����"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmCalc4.frx":0000
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
Attribute VB_Name = "frmCalc4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlus_Click()
    If Trim(txtValue1.Text) = "" Then
        MsgBox "���ڸ� �Է��ؾ� �մϴ�"
        txtValue1.SelStart = 0
        txtValue1.SelLength = Len(txtValue1.Text)
        txtValue1.SetFocus
        Exit Sub
    End If
    
    If Trim(txtValue2.Text) = "" Then
        MsgBox "���ڸ� �Է��ؾ� �մϴ�"
        txtValue2.SelStart = 0
        txtValue2.SelLength = Len(txtValue1.Text)
        txtValue2.SetFocus
        Exit Sub
    End If
    
    Dim Value1 As Single    '1 ��° �ؽ�Ʈ������ ������ �Ҽ� ��
    Dim Value2 As Single    '2 ��° �ؽ�Ʈ������ ������ �Ҽ� ��
    Dim Result As Single    '2 ���� ������ �Ҽ� ���� ��
    
    Value1 = CSng(txtValue1.Text)
    Value2 = CSng(txtValue2.Text)
    
    Result = Value1 + Value2
        
    txtResult.Text = Format(Result, "#,###.###")
End Sub

Private Sub mnuExit_Click()
    End
End Sub
