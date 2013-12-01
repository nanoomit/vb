VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    '난수를 만든다
    Red = Int(256 * Rnd)
    Green = Int(256 * Rnd)
    Blue = Int(256 * Rnd)
    
    '폼의 배경 색을 변경한다
    Form1.BackColor = RGB(Red, Green, Blue)
End Sub
