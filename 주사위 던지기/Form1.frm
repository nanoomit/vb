VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "주사위 던지기"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "시작"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "중지"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2460
      Top             =   600
   End
   Begin VB.Image imgDiceOrg 
      Height          =   720
      Index           =   1
      Left            =   3300
      Picture         =   "Form1.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDiceOrg 
      Height          =   720
      Index           =   2
      Left            =   4080
      Picture         =   "Form1.frx":0502
      Top             =   360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDiceOrg 
      Height          =   720
      Index           =   3
      Left            =   4860
      Picture         =   "Form1.frx":0A04
      Top             =   360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDiceOrg 
      Height          =   720
      Index           =   4
      Left            =   3300
      Picture         =   "Form1.frx":0F06
      Top             =   1140
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDiceOrg 
      Height          =   720
      Index           =   5
      Left            =   4080
      Picture         =   "Form1.frx":1408
      Top             =   1140
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDiceOrg 
      Height          =   720
      Index           =   6
      Left            =   4860
      Picture         =   "Form1.frx":190A
      Top             =   1140
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDice 
      Height          =   720
      Left            =   1320
      Picture         =   "Form1.frx":1E0C
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Timer1.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = True
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
    Command2.Enabled = False
    Command1.Enabled = True
End Sub


Private Sub Timer1_Timer()
   i = Int(Rnd(1) * 6) + 1
   imgDice.Picture = imgDiceOrg(i).Picture
End Sub
