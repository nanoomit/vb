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
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
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
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      Filter          =   "텍스트 파일(*.txt)|*.txt|모든 파일|*.*"
      FontName        =   "굴림"
      InitDir         =   "c:\Visual Basic 테스트 룸"
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
      Align           =   2  '아래 맞춤
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
            TextSave        =   "2005-10-28"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "오후 10:47"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
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
            Object.ToolTipText     =   "새로 만들기"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "열기"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "저장"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "인쇄"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "잘라내기"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "복사"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "붙여넣기"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "글꼴"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Black"
                  Text            =   "검정"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Red"
                  Text            =   "빨강"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Green"
                  Text            =   "녹색"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Blue"
                  Text            =   "파랑"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CustomColor"
                  Text            =   "사용자 색..."
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
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
Private Function IsSaveContent() As Integer
   Dim retValue As Integer
   
   retValue = MsgBox(gFileName & "문서 내용이 변경되었습니다" _
                      & vbCrLf & "변경 내용을 저장하시겠습니까?", _
                      vbInformation Or vbDefaultButton1 Or _
                       vbYesNoCancel, gTitle)
                       
    IsSaveContent = retValue
End Function

Private Sub Form_Load()
    '폼을 화면의 한 가운데에 표시한다
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    TestBoxResize
    
    Text1.Text = ""          '텍스트 상자의 내용을 지운다
    gFileName = "제목 없음"
    DisplayTitle             '제목을 표시한다
    
    bDirty = False           '디폴트는 내용 변경 없음
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim retValue As Integer
     
   If bDirty = True Then
        retValue = IsSaveContent()
         
        Select Case retValue
            Case vbYes
                mnuSave_Click
                End     '프로그램을 종료한다
            Case vbNo
                 
            Case vbCancel
                Cancel = True
        End Select
   End If
End Sub

Private Sub Form_Resize()
    TestBoxResize
End Sub


Public Sub TestBoxResize()
    '텍스트 상자를 폼의 사용자 영역의 크기로 조정한다
    Text1.Left = Me.ScaleLeft
    Text1.Top = Me.ScaleTop + Toolbar1.Height
    Text1.Width = Me.ScaleWidth
    Text1.Height = Me.ScaleHeight - Toolbar1.Height - StatusBar1.Height
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear   '클립보드를 지운다
         
    '선택한 문자열을 클립보드에 복사합니다.
    Clipboard.SetText Text1.SelText
End Sub

Private Sub mnuCut_Click()
    Clipboard.Clear   '클립보드를 지운다
     
    '선택한 문자열을 클립보드에 복사합니다.
    Clipboard.SetText Text1.SelText
         
    '선택한 문자열을 삭제합니다.
    Text1.SelText = ""
End Sub

Private Sub mnuEditTime_Click()
    '현재 시간과 날짜를 삽입합니다.
    Text1.SelText = Format(Now, "general date")
End Sub

Private Sub mnuExit_Click()
   Dim retValue As Integer
     
   If bDirty = True Then
        retValue = IsSaveContent()
         
        Select Case retValue
            Case vbYes
                mnuSave_Click
            Case vbNo
                 
            Case vbCancel
                Exit Sub
        End Select
   End If
     
   '프로그램을 종료한다
   End
End Sub

Private Sub mnuFont_Click()
On Error GoTo Error_rtn
     
    CommonDialog1.Flags = cdlCFScreenFonts Or cdlCFEffects
     
    '글꼴 대화상자를 현재 텍스트 상자의 글꼴로 초기화한다
    CommonDialog1.FontName = Text1.Font.Name
    CommonDialog1.FontSize = Text1.Font.Size
    If Text1.FontBold Then CommonDialog1.FontBold = True
    If Text1.FontItalic Then CommonDialog1.FontItalic = True
    If Text1.FontUnderline Then CommonDialog1.FontUnderline = True
    If Text1.FontStrikethru Then CommonDialog1.FontStrikethru = True
    CommonDialog1.Color = Text1.ForeColor
     
    CommonDialog1.ShowFont
     
    '글꼴 대화상자의 설정 내용을 텍스트 상자 컨트롤에 반영한다
    Text1.Font.Name = CommonDialog1.FontName
    Text1.Font.Size = CommonDialog1.FontSize
    If CommonDialog1.FontBold Then Text1.FontBold = True
    If CommonDialog1.FontItalic Then Text1.FontItalic = True
    If CommonDialog1.FontUnderline Then Text1.FontUnderline = True
    If CommonDialog1.FontStrikethru Then Text1.FontStrikethru = True
    Text1.ForeColor = CommonDialog1.Color
     
Error_rtn:
    If Err.Number = 32755 Or Err.Number = 0 Then
        Exit Sub
    Else
        MsgBox Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub mnuNew_Click()
    Dim retValue As Integer
     
    If bDirty = True Then
        retValue = IsSaveContent()
        
        Select Case retValue
            Case vbYes
                mnuSave_Click
            Case vbNo
                 
            Case vbCancel
                Exit Sub
        End Select
    End If
         
    bDirty = False
    gFileName = "제목없음"
     
    DisplayTitle       '제목을 표시한다
    Text1.Text = ""    '텍스트의 내용을 지운다
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Error_rtn

    CommonDialog1.InitDir = "C:\Program Files"
    CommonDialog1.Filter = "텍스트파일(*.txt)|*.txt|모든 파일|*.*"
    CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.ShowOpen      '열기 대화상자 표시
    gFileName = CommonDialog1.FileName
    
    '텍스트 파일을 읽어서 텍스트 상자에 표시한다
    Dim TotalData, readData As String
    
    Open gFileName For Input As #1
        While Not EOF(1)
            Line Input #1, readData
            TotalData = TotalData & readData & vbCrLf
        Wend
    Close #1
    
    Text1.Text = TotalData
    
    bDirty = False    '내용 변경 없음
    DisplayTitle      '제목을 표시한다
    Exit Sub
Error_rtn:
    If Err.Number = 32755 Or Err.Number = 0 Then
        Exit Sub
    Else
        MsgBox Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub DisplayTitle()
    Form1.Caption = gFileName & " - " & gTitle
End Sub

Private Sub mnuPaste_Click()
    '클립보드의 문자열을 활성 컨트롤에 붙여 넣는다.
    Text1.SelText = Clipboard.GetText()
End Sub

Private Sub mnuPrint_Click()
    '텍스트 상자에 있는 내용을 기본 프린터로 출력한다
    Printer.NewPage
    Printer.Print Text1.Text
    Printer.EndDoc
End Sub

Private Sub mnuSave_Click()
On Error GoTo Error_rtn

    CommonDialog1.InitDir = "C:\Program Files"
    CommonDialog1.Filter = "텍스트파일(*.txt)|*.txt|모든 파일|*.*"
    CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.FileName = gFileName
    CommonDialog1.ShowSave      '저장 대화상자 표시
    
    gFileName = CommonDialog1.FileName
    
    '텍스트 상자의 내용을 파일로 기록한다
    Dim TotalData, readData As String
    
    Open gFileName For Output As #1
        Print #1, Text1.Text
    Close #1
   
    bDirty = False    '내용 변경 없음
    DisplayTitle      '제목을 표시한다
    Exit Sub
Error_rtn:
    If Err.Number = 32755 Or Err.Number = 0 Then
        Exit Sub
    Else
        MsgBox Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub mnuSelectAll_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_Change()
    bDirty = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuNew_Click
        Case "Open"
            mnuOpen_Click
        Case "Save"
            mnuSave_Click
        Case "Print"
            mnuPrint_Click
        Case "Cut"
            mnuCut_Click
        Case "Copy"
            mnuCopy_Click
        Case "Paste"
            mnuPaste_Click
        Case "Font"
            mnuFont_Click
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo Error_rtn
    Dim lColor As Long
    Dim strMessage As String
     
    Select Case ButtonMenu.Key
        Case "Black"
            lColor = RGB(0, 0, 0)
            strMessage = "검정"
        Case "Red"
            lColor = RGB(255, 0, 0)
            strMessage = "빨강"
        Case "Green"
            lColor = RGB(0, 255, 0)
            strMessage = "녹색"
        Case "Blue"
            lColor = RGB(0, 0, 255)
            strMessage = "파랑"
        Case "CustomColor"
            strMessage = "사용자 색상"
            CommonDialog1.ShowColor
            lColor = CommonDialog1.Color
    End Select
     
    '텍스트 색상
    Text1.ForeColor = lColor
    '메시지바의 메시지
    StatusBar1.Panels.Item(1).Text = strMessage & " 단추 메뉴가 선택되었습니다"
    Exit Sub
Error_rtn:
    If Err.Number = 32755 Or Err.Number = 0 Then
        Exit Sub
    Else
        MsgBox Err.Number & ": " & Err.Description
    End If
End Sub
