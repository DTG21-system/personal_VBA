Attribute VB_Name = "fillYellow"
'------------------------------
'�I��͈͂̕�������A�ԐF�ɕύX����
'------------------------------
Sub fillYellow()
    
    '�ϐ��F�I��͈�
    Dim range As range
    
    ' �I��͈͂��擾
    Set range = Selection
    
    ' �I��͈͂̃Z�������F�ɓh��Ԃ�
    range.Interior.Color = RGB(255, 255, 0)
    
End Sub
