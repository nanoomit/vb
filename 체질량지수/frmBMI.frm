VERSION 5.00
Begin VB.Form frmBMI 
   BackColor       =   &H00C0E0FF&
   Caption         =   "BMI(체질량지수)"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtBMI 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "BMI 판정"
      Height          =   1815
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   3975
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "비만 3단계"
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "비만 2단계"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "비만 1단계"
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "비만 전단계"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "정상 체중"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "저 체중"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBMICalc 
      Caption         =   "BMI 계산"
      Height          =   495
      Left            =   1365
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "BMI"
      Height          =   180
      Left            =   600
      TabIndex        =   8
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "㎏"
      Height          =   180
      Left            =   3480
      TabIndex        =   5
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "체중"
      Height          =   180
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "신장"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "㎝"
      Height          =   180
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   180
   End
End
Attribute VB_Name = "frmBMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBMICalc_Click()
    Dim Height As Single
    Dim Weight As Integer
    Dim BMI As Single
    
    Height = CSng(txtHeight.Text / 100)
    Weight = CInt(txtWeight.Text)
    
    BMI = CSng(Weight / (Height * Height))
    
    txtBMI.Text = Format(BMI, "##.#")
    
    If BMI < 18.5 Then
        Option1.Value = True
    ElseIf BMI >= 18.5 And BMI <= 22.9 Then
        Option2.Value = True
    ElseIf BMI >= 23 And BMI <= 24.9 Then
        Option3.Value = True
    ElseIf BMI >= 25 And BMI <= 29.9 Then
        Option4.Value = True
    ElseIf BMI >= 30 And BMI <= 39.9 Then
        Option5.Value = True
    ElseIf BMI >= 40 Then
        Option6.Value = True
    End If
End Sub
