Attribute VB_Name = "Module1"
Public Seq As Integer           '����

Private Type Score
    Irum As String  '�̸�
    Guk As Integer  '���� ����
    Math As Integer '���� ����
    Sahe As Integer '��ȸ ����
    Sci As Integer  '���� ����
End Type

Public ScoreData(1 To 60) As Score   '60 ���� �л�
Public ScoreDataTemp As Score

