Attribute VB_Name = "resetSheets"
'------------------------------
'�V�[�g�̃Z����A1�A�\���{����100%�ɕύX����
'------------------------------
Sub resetSheets()

    '�ϐ��F���[�N�V�[�g
    Dim ws As Worksheet
    
    '�V�[�g�̌������A�J��Ԃ�
    For Each ws In ActiveWorkbook.Worksheets
    
        '�Z��A1��I��
        ws.Activate
        ws.range("A1").Select
        
        '�\���{����100%�ɐݒ肷��
        ActiveWindow.Zoom = 100
    
    Next ws
    
    '�Ώۃu�b�N�̈�Ԑ擪�̃V�[�g��\������
    ThisWorkbook.Sheets(1).Select
    
End Sub
