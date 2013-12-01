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
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Menu mnuFile 
      Caption         =   "파일"
      Begin VB.Menu mnuNew 
         Caption         =   "새로 만들기(&N)"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "열기(&O)..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "저장(&S)..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "인쇄(&P)"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "편집"
      Begin VB.Menu mnuCut 
         Caption         =   "잘라내기(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "복사(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "붙여넣기(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "전체 선택(&A)"
      End
      Begin VB.Menu mnuEditTime 
         Caption         =   "시간/날짜"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "서식"
      Begin VB.Menu mnuFont 
         Caption         =   "글꼴(&F)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말"
      Begin VB.Menu mnuAbout 
         Caption         =   "나의 메모장이란(&A)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
