Attribute VB_Name = "changeTextColorToRed"
'------------------------------
'�I��͈͂̕�������A�ԐF�ɕύX����
'------------------------------
Sub changeTextColorToRed()
    
    '�ϐ��F�I��͈�
    Dim renge As range
    
    '�I��͈͂��擾
    Set renge = Selection
    
    '�I��͈͂̕����F��ԂɕύX
    renge.Font.Color = RGB(255, 0, 0)

End Sub
