VERSION 5.00
Begin VB.UserControl Bar 
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   PropertyPages   =   "Bar.ctx":0000
   ScaleHeight     =   855
   ScaleWidth      =   4410
   Begin VB.PictureBox picBar 
      Align           =   2  '�Ʒ� ����
      ClipControls    =   0   'False
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   4350
      TabIndex        =   1
      Top             =   360
      Width           =   4410
      Begin VB.Shape shpBar 
         BorderStyle     =   0  '����
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  '�ܻ�
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  '��� ����
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Const conBarHeight As Double = 0.5
'�⺻ �Ӽ� ��:
Const m_def_Percent = 50
'�Ӽ� ����:
Dim m_Percent As Integer

Public Event Click()
Public Event Change()





Private Sub lblCaption_Click()

    RaiseEvent Click
    
End Sub


Private Sub picBar_Click()

    RaiseEvent Click
    
End Sub

Private Sub UserControl_InitProperties()

    On Error Resume Next
    Me.Caption = Extender.Caption
    Set Me.Font = Ambient.Font
    Me.BackColor = Ambient.BackColor
    
    m_Percent = m_def_Percent
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.Caption = PropBag.ReadProperty _
                                ("Caption", lblCaption.Caption)
    
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Label1")
    lblCaption.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    shpBar.FillColor = PropBag.ReadProperty("FillColor", &HFFC0FF)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Percent = PropBag.ReadProperty("Percent", m_def_Percent)
    
    Call SetPercent
    
End Sub

Private Sub UserControl_Resize()

    '�� ��Ʈ���� �ʺ� �����Ѵ�
    '�� ��Ʈ���� ���̸� ��� ������ �����Ѵ�
    lblCaption.Move 0, 0, UserControl.ScaleWidth, _
                          UserControl.ScaleHeight * conBarHeight
                          
    picBar.Move lblCaption.Height, lblCaption.Width, _
                UserControl.ScaleHeight * (1 - conBarHeight)
                
    shpBar.Move 0, 0, shpBar.Width, shpBar.Height

End Sub
'
'Public Property Get Caption() As String
'
'    Caption = lblCaption.Caption
'
'End Property
'
'Public Property Let Caption(ByVal NewCaption As String)
'
'    lblCaption.Caption = NewCaption
'    PropertyChanged "Caption"
'
'End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    '�Ӽ� ���� ����ҿ� �����Ѵ�
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
    
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Label1")
    Call PropBag.WriteProperty("BackColor", lblCaption.BackColor, &H8000000F)
    Call PropBag.WriteProperty("FillColor", shpBar.FillColor, &HFFC0FF)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Percent", m_Percent, m_def_Percent)
End Sub
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "�Ϲ�"
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=lblCaption,lblCaption,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "��ü�� �ؽ�Ʈ�� �׷����� ǥ���ϱ� ���� ���Ǵ� ������ ��ȯ�ϰų� �����մϴ�."
    BackColor = lblCaption.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lblCaption.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=shpBar,shpBar,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "����, ��, ���ڸ� ä��� �� ���� ���� ��ȯ�ϰų� �����մϴ�."
    FillColor = shpBar.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    shpBar.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font ��ü�� ��ȯ�մϴ�."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,50
Public Property Get Percent() As Integer
Attribute Percent.VB_Description = "�ٿ� ǥ�õǾ� �ִ� %�� �����ϰų�\r\n��´�"
Attribute Percent.VB_ProcData.VB_Invoke_Property = "�Ϲ�"
    Percent = m_Percent
End Property

Public Property Let Percent(ByVal New_Percent As Integer)
    
    If New_Percent < 0 Then New_Percent = 0
    If New_Percent > 100 Then New_Percent = 100
    
    m_Percent = New_Percent
    
    Call SetPercent
    
    PropertyChanged "Percent"
    
End Property


Private Sub SetPercent()
    
    shpBar.Width = picBar.Width * Me.Percent / 100
    RaiseEvent Change
    
End Sub
