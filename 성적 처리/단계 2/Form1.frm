VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "���� ó��"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdScoreSort 
      Caption         =   "������ ǥ��"
      Height          =   375
      Left            =   3908
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdScoreInput 
      Caption         =   "���� �Է�"
      Height          =   375
      Left            =   1988
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   60
      Cols            =   7
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Text = "KKK"
End Sub

Private Sub cmdScoreInput_Click()
    Form2.Show
End Sub

Private Sub cmdScoreSort_Click()
    MSFlexGrid2.Rows = Seq + 1
    MSFlexGrid2.Cols = 9
    MSFlexGrid2.Visible = True
    
    Dim i As Integer
    Dim j As Integer
    Dim IrumTemp As String
    Dim ScoreTemp As Integer
    
    '���� ������ ����
    For i = 1 To Seq - 1
        For j = i + 1 To Seq
            If Guk(i) + Math(i) + Sahe(i) + Sci(i) < _
                Guk(j) + Math(j) + Sahe(j) + Sci(j) Then
                
                IrumTemp = Irum(i)
                Irum(i) = Irum(j)
                Irum(j) = IrumTemp
                
                ScoreTemp = Math(i)
                Math(i) = Math(j)
                Math(j) = ScoreTemp
                
                ScoreTemp = Guk(i)
                Guk(i) = Guk(j)
                Guk(j) = ScoreTemp
                
                ScoreTemp = Sahe(i)
                Sahe(i) = Sahe(j)
                Sahe(j) = ScoreTemp
                
                ScoreTemp = Sci(i)
                Sci(i) = Sci(j)
                Sci(j) = ScoreTemp
            End If
        Next j
    Next i
    
    'FlexGrid�� ���� ǥ��
    MSFlexGrid2.Row = 0
    MSFlexGrid2.Col = 1
    MSFlexGrid2.Text = "����"
    
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Text = "�̸�"

    MSFlexGrid2.Col = 3
    MSFlexGrid2.Text = "����"

    MSFlexGrid2.Col = 4
    MSFlexGrid2.Text = "����"

    MSFlexGrid2.Col = 5
    MSFlexGrid2.Text = "��ȸ"

    MSFlexGrid2.Col = 6
    MSFlexGrid2.Text = "����"
    
    MSFlexGrid2.Col = 7
    MSFlexGrid2.Text = "����"

    MSFlexGrid2.Col = 8
    MSFlexGrid2.Text = "���"
    
    For i = 1 To Seq
        MSFlexGrid2.Row = i
        
        
        MSFlexGrid2.Col = 1
        MSFlexGrid2.Text = CStr(i)
        
        MSFlexGrid2.Col = 2
        MSFlexGrid2.Text = Irum(i)
        
        MSFlexGrid2.Col = 3
        MSFlexGrid2.Text = CStr(Guk(i))
        
        MSFlexGrid2.Col = 4
        MSFlexGrid2.Text = CStr(Math(i))
        
        MSFlexGrid2.Col = 5
        MSFlexGrid2.Text = CStr(Sahe(i))
        
        MSFlexGrid2.Col = 6
        MSFlexGrid2.Text = CStr(Sci(i))
        
        MSFlexGrid2.Col = 7
        MSFlexGrid2.Text = CStr(Guk(i) + Math(i) + Sahe(i) + Sci(i))
        
        MSFlexGrid2.Col = 8
        MSFlexGrid2.Text = Format((Guk(i) + Math(i) + Sahe(i) + Sci(i)) / 4, "##0.0#")
    Next i
    
End Sub

Private Sub Form_Load()
    'FlexGrid�� ���� ǥ��
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = "����"
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Text = "�̸�"

    MSFlexGrid1.Col = 3
    MSFlexGrid1.Text = "����"

    MSFlexGrid1.Col = 4
    MSFlexGrid1.Text = "����"

    MSFlexGrid1.Col = 5
    MSFlexGrid1.Text = "��ȸ"

    MSFlexGrid1.Col = 6
    MSFlexGrid1.Text = "����"

    ' (1,1)�� ���� ��ġ�� ��Ŀ���� �ش�
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 1
End Sub

