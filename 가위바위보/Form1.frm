VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "가위바위보"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   3360
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   1800
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdBo 
      Caption         =   "보"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdBawi 
      Caption         =   "바위"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdGawi 
      Caption         =   "가위"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGawi_Click()
    Picture1.Picture = LoadPicture(App.Path & "\가위.bmp")
End Sub

Private Sub cmdBawi_Click()
    Picture2.Picture = LoadPicture(App.Path & "\바위.bmp")
End Sub

Private Sub cmdBo_Click()
    Picture3.Picture = LoadPicture(App.Path & "\보.bmp")
End Sub

