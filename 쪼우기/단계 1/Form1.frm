VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "쪼우기"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "패 돌리기"
      Height          =   495
      Left            =   1395
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "컴퓨터"
      Height          =   2055
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      Begin VB.PictureBox DisplayedPicture 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   2
         Left            =   240
         ScaleHeight     =   1185
         ScaleWidth      =   825
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox DisplayedPicture 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   3
         Left            =   1200
         ScaleHeight     =   1185
         ScaleWidth      =   825
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "사용자"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.PictureBox DisplayedPicture 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   1
         Left            =   1200
         ScaleHeight     =   1185
         ScaleWidth      =   825
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox DisplayedPicture 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   240
         ScaleHeight     =   1185
         ScaleWidth      =   825
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim User(1 To 2, 1 To 2) As Integer           '사용자의 화투의 값
Dim Computer(1 To 2, 1 To 2) As Integer    '컴퓨터의 화투의 값
Dim Selected(1 To 12, 1 To 2) As Boolean  '무작위로 선택된 화투


Private Sub Command1_Click()

    Command1.Enabled = False
    Dim Generated1_12 As Integer    '1~12까지의 화투 구분
    Dim Generated1_2 As Integer      '피를 제외한 1번째와 2 번째
    
    Dim UserIndex As Integer
    Dim ComputerIndex As Integer
    
    While UserIndex < 2
     Randomize                        ' 난수 발생기를 초기화하여 이전의 난수열이 반복되지 않게 합니다
    
     Generated1_12 = Int((12 * Rnd) + 1)   ' 1과 12 사이의 난수를 발생합니다.
     Generated1_2 = Int((2 * Rnd) + 1)      ' 1과 2 사이의 난수를 발생합니다.
     
        If Selected(Generated1_12, Generated1_2) <> True Then
             Selected(Generated1_12, Generated1_2) = True
             UserIndex = UserIndex + 1
             User(UserIndex, 1) = Generated1_12
             User(UserIndex, 2) = Generated1_2
        End If
    Wend
    
    
    
    While ComputerIndex < 2
     Randomize                        ' 난수 발생기를 초기화하여 이전의 난수열이 반복되지 않게 합니다
    
     Generated1_12 = Int((12 * Rnd) + 1)   ' 1과 12 사이의 난수를 발생합니다.
     Generated1_2 = Int((2 * Rnd) + 1)      ' 1과 2 사이의 난수를 발생합니다.
     
        If Selected(Generated1_12, Generated1_2) <> True Then
             Selected(Generated1_12, Generated1_2) = True
             ComputerIndex = ComputerIndex + 1
             Computer(ComputerIndex, 1) = Generated1_12
             Computer(ComputerIndex, 2) = Generated1_2
        End If
    Wend
    
    Dim strPictureFile As String
    
    Dim strPicturePath  As String
    strPicturePath = PicturePath()
    
    
    For i = 1 To 2
       strPictureFile = User(i, 1) & "-" & User(i, 2) & ".jpg"
       DisplayedPicture(i - 1).Picture = LoadPicture(strPicturePath & strPictureFile)
    Next i
    
    For i = 1 To 2
       strPictureFile = Computer(i, 1) & "-" & Computer(i, 2) & ".jpg"
       DisplayedPicture(i + 1).Picture = LoadPicture(strPicturePath & strPictureFile)
    Next i
        
    
End Sub

Private Function PicturePath() As String
    Dim strPicturePath  As String
    strPicturePath = App.Path
    
    If (Right(strPicturePath, 1) <> "\") Then
        strPicturePath = strPicturePath & "\"
    End If
    
    PicturePath = strPicturePath
End Function

Private Sub Form_Load()
    Dim strPicturePath  As String
    strPicturePath = PicturePath()
    For Index = 0 To 3
        DisplayedPicture(Index).Picture = LoadPicture(strPicturePath & "0.jpg")
    Next Index

End Sub

