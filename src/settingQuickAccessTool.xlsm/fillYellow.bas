Attribute VB_Name = "fillYellow"
'------------------------------
'�I��͈͂̕�������A�ԐF�ɕύX����
'------------------------------
Sub fillYellow()
    
    '�I��͈͂����݂��Ȃ��ꍇ
    If Selection Is Nothing Then
        MsgBox "�Z����I�����Ă��������B"
        Exit Sub
    End If

    '�I��͈͂��Z���ł͂Ȃ��ꍇ
    If TypeName(Selection) <> "Range" Then
        MsgBox "�Z����I�����Ă��������B"
        Exit Sub
    End If
    
    '�ϐ��F�I��͈�
    Dim range As range
    
    ' �I��͈͂��擾
    Set range = Selection
    
    ' �I��͈͂̃Z�������F�ɓh��Ԃ�
    range.Interior.Color = RGB(255, 255, 0)
    
End Sub
