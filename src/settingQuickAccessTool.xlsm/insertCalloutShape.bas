Attribute VB_Name = "insertCalloutShape"
'------------------------------
'�����o���F���̐}�`��}������
'�i�}�`�̘g����ԁA�h��Ԃ��𔒁A�����F��Ԃɐݒ肷��j
'------------------------------
Sub insertCalloutShape()

    '�ϐ��F���[�N�V�[�g
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    '�ϐ��F�}�`
    Dim shape As shape
    
    '�����o���F����}������
    Set shape = ws.Shapes.AddShape(msoShapeLineCallout1, 100, 100, 200, 100)
    
    '�����o���̓h��Ԃ��𔒂ɐݒ肷��
    shape.Fill.ForeColor.RGB = RGB(255, 255, 255)

    '�����o���̘g����ԐF�ɐݒ肷��
    shape.Line.ForeColor.RGB = RGB(255, 0, 0)

    '�����o���̕����F��ԐF�ɐݒ肷��
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    
    '�e�L�X�g��ҏW���̏�Ԃɂ���
    shape.TextFrame2.TextRange.Select
    
End Sub
