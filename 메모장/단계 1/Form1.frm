VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
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
