VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0006
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0118
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":022A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":033C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":044E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0560
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0672
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0784
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2715
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   2593
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2005-10-27"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "���� 12:08"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '�� ����
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "���� �����"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "����"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "����"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "�μ�"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "�߶󳻱�"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "����"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "�ٿ��ֱ�"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "�۲�"
            ImageIndex      =   8
            Style           =   5
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "����"
      Begin VB.Menu mnuNew 
         Caption         =   "���� �����(&N)"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "����(&O)..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����(&S)..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "�μ�(&P)"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "����(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "����"
      Begin VB.Menu mnuCut 
         Caption         =   "�߶󳻱�(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "�ٿ��ֱ�(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "��ü ����(&A)"
      End
      Begin VB.Menu mnuEditTime 
         Caption         =   "�ð�/��¥"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "����"
      Begin VB.Menu mnuFont 
         Caption         =   "�۲�(&F)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����"
      Begin VB.Menu mnuAbout 
         Caption         =   "���� �޸����̶�(&A)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    '���� ȭ���� �� ����� ǥ���Ѵ�
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    TestBoxResize
End Sub

Private Sub Form_Resize()
    TestBoxResize
End Sub


Public Sub TestBoxResize()
    '�ؽ�Ʈ ���ڸ� ���� ����� ������ ũ��� �����Ѵ�
    Text1.Left = Me.ScaleLeft
    Text1.Top = Me.ScaleTop + Toolbar1.Height
    Text1.Width = Me.ScaleWidth
    Text1.Height = Me.ScaleHeight - Toolbar1.Height - StatusBar1.Height
End Sub
