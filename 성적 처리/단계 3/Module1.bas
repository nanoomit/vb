Attribute VB_Name = "Module1"
Public Seq As Integer           '순번

Private Type Score
    Irum As String  '이름
    Guk As Integer  '국어 성적
    Math As Integer '수학 성적
    Sahe As Integer '사회 성적
    Sci As Integer  '과학 성적
End Type

Public ScoreData(1 To 60) As Score   '60 명의 학생
Public ScoreDataTemp As Score

