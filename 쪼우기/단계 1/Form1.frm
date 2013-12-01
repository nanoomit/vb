VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ɿ��"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "�� ������"
      Height          =   495
      Left            =   1395
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ǻ��"
      Height          =   2055
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      Begin VB.PictureBox DisplayedPicture 
         Appearance      =   0  '���
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
         Appearance      =   0  '���
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
      Caption         =   "�����"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.PictureBox DisplayedPicture 
         Appearance      =   0  '���
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
         Appearance      =   0  '���
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
Dim User(1 To 2, 1 To 2) As Integer           '������� ȭ���� ��
Dim Computer(1 To 2, 1 To 2) As Integer    '��ǻ���� ȭ���� ��
Dim Selected(1 To 12, 1 To 2) As Boolean  '�������� ���õ� ȭ��


Private Sub Command1_Click()

    Command1.Enabled = False
    Dim Generated1_12 As Integer    '1~12������ ȭ�� ����
    Dim Generated1_2 As Integer      '�Ǹ� ������ 1��°�� 2 ��°
    
    Dim UserIndex As Integer
    Dim ComputerIndex As Integer
    
    While UserIndex < 2
     Randomize                        ' ���� �߻��⸦ �ʱ�ȭ�Ͽ� ������ �������� �ݺ����� �ʰ� �մϴ�
    
     Generated1_12 = Int((12 * Rnd) + 1)   ' 1�� 12 ������ ������ �߻��մϴ�.
     Generated1_2 = Int((2 * Rnd) + 1)      ' 1�� 2 ������ ������ �߻��մϴ�.
     
        If Selected(Generated1_12, Generated1_2) <> True Then
             Selected(Generated1_12, Generated1_2) = True
             UserIndex = UserIndex + 1
             User(UserIndex, 1) = Generated1_12
             User(UserIndex, 2) = Generated1_2
        End If
    Wend
    
    
    
    While ComputerIndex < 2
     Randomize                        ' ���� �߻��⸦ �ʱ�ȭ�Ͽ� ������ �������� �ݺ����� �ʰ� �մϴ�
    
     Generated1_12 = Int((12 * Rnd) + 1)   ' 1�� 12 ������ ������ �߻��մϴ�.
     Generated1_2 = Int((2 * Rnd) + 1)      ' 1�� 2 ������ ������ �߻��մϴ�.
     
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

