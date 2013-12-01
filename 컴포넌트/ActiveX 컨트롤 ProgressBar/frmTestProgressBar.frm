VERSION 5.00
Object = "{41B13205-A4B6-11D2-B62F-CA2D20F8BAA3}#5.0#0"; "ProgressBar.ocx"
Begin VB.Form frmTestProgressBar 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin ProgressBar.Bar Bar1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1508
      Caption         =   "Label1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
   End
End
Attribute VB_Name = "frmTestProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bar1_Change()
    
    Text1.Text = Bar1.Percent
    
End Sub

Private Sub HScroll1_Change()

    Bar1.Percent = HScroll1.Value
    
End Sub


