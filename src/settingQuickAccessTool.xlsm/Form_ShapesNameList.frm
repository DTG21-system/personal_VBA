VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_ShapesNameList 
   Caption         =   "ShapesNameList"
   ClientHeight    =   4170
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4680
   OleObjectBlob   =   "Form_ShapesNameList.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Form_ShapesNameList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------
'�}�`���̈ꗗ_��������
'�ΏۃV�[�g���ɑ��݂���}�`�����[�U�[�t�H�[���Ɉꗗ�\������
'------------------------------
Private Sub UserForm_Initialize()

    '�ϐ��F���[�N�V�[�g
    Dim ws As Worksheet
    
    '�ϐ��F�}�`
    Dim shape As shape
    
    '�ϐ��F�}�`����
    Dim shapeNames As String

    '�ϐ��F�J��Ԃ����񃋁[�v����p
    Dim firstLoop As Boolean: firstLoop = True
    
    ' �A�N�e�B�u�V�[�g���擾
    Set ws = ActiveSheet
    
    ' �V�[�g���ɑ��݂���A�}�`�̐������J��Ԃ�
    For Each shape In ws.Shapes
        
        '��ԍŏ��Ɏ擾�����}�`�����e�L�X�g�{�b�N�X�ɐݒ肷��
        If firstLoop Then
            
            '�}�`���̂��A�e�L�X�g�{�b�N�X�ɐݒ肷��
            TextBox_ShapeName = shape.Name
            
            '�J��Ԃ����񃋁[�v����p�ϐ���false�ɐݒ肷��
            firstLoop = False
        
        End If
        
        '�}�`���̈ꗗ���[�U�[�t�H�[���̃��X�g�{�b�N�X�ɒǉ�����
        ListBox_Shapes.AddItem shape.Name
                
    Next shape
    
End Sub

'------------------------------
'�}�`���̈ꗗ_���X�g�{�b�N�X�I��l�؂�ւ�
'------------------------------
Private Sub ListBox_Shapes_Change()

    '�ϐ��F���[�N�V�[�g
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    '�ϐ��F�}�`
    Dim shape As shape
    
    '���X�g�{�b�N�X��NULL�̏ꍇ�A�������I������B
    If IsNull(ListBox_Shapes) Then
        Exit Sub
    End If
    
    '�ϐ��F�I��}�`����
    Dim selectedShapeName As String: selectedShapeName = ListBox_Shapes.Value
    
    '�I�������}�`�����e�L�X�g�{�b�N�X�ɐݒ�
    TextBox_ShapeName.Value = selectedShapeName
    
    '�A�N�e�B�u�V�[�g���擾
    For Each shape In ws.Shapes
    
        '���X�g�{�b�N�X�őI�������}�`���̂���v���Ă���ꍇ
        If shape.Name = selectedShapeName Then
        
            '�Ώۂ̐}�`��I����Ԃɂ���
            shape.Select
            
            '�J��Ԃ������𔲂���
            Exit For
        
        End If
    
    Next shape
    
End Sub

'------------------------------
'�}�`���̈ꗗ_���X�g�{�b�N�X�L�[�{�^������
'------------------------------
Private Sub ListBox_Shapes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    '�G���^�[�L�[�������ꂽ�ꍇ
    If KeyCode = vbKeyReturn Then
        
        '���[�U�[�t�H�[�������
        Unload Me
        
    End If

End Sub


'------------------------------
'�}�`���̈ꗗ_���l�[���{�^������
'------------------------------
Private Sub Button_Rename_Click()

    '�ϐ��F���[�N�V�[�g
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    '���X�g�{�b�N�X��NULL�̏ꍇ�A�������I������B
    If IsNull(ListBox_Shapes) Then
        Exit Sub
    End If
    
    '�ϐ��F�I��}�`��
    Dim selectedShapeName As String: selectedShapeName = ListBox_Shapes.Value
    
    '�ϐ��F���l�[���}�`����
    Dim newShapeName As String: newShapeName = TextBox_ShapeName.Value
    
    '�ϐ��F�}�`
    Dim shape As shape
    
    '�d���`�F�b�N
    If isDuplication(ws, newShapeName) Then
        
        '�d���G���[�̃��b�Z�[�W��ݒ肷��B
        MsgBox "���ꖼ����͂��Ă��邩�A���l�[����̐}�`���̂��d�����Ă��܂��B" _
               & vbCrLf & _
               "�ʂ̖��O����͂��Ă��������B", Buttons:=vbCritical
        
        '�������I������
        Exit Sub
        
    End If
    
    
    '�}�`���̂�ύX
    For Each shape In ws.Shapes
        
        '���X�g�{�b�N�X�őI�������}�`�ƈ�v���Ă���ꍇ
        If shape.Name = selectedShapeName Then
        
            '�}�`���̂�ύX����
            shape.Name = newShapeName
            
            '�����𔲂���
            Exit For
        
        End If
    
    Next shape
    
    '���l�[�������̃��b�Z�[�W���o�͂���
    MsgBox "�}�`���̍X�V���������܂����B" _
           & vbCrLf & _
           "�ύX�O�F" & selectedShapeName & " " & _
           "�ύX��F" & newShapeName
    
    
    '���X�g�{�b�N�X���N���A����
    ListBox_Shapes.Clear
    
    '���X�g�{�b�N�X���X�V����
    For Each shape In ws.Shapes
    
        '�}�`���̈ꗗ���[�U�[�t�H�[���̃��X�g�{�b�N�X�ɒǉ�����
        ListBox_Shapes.AddItem shape.Name
    
    Next shape
    
End Sub


'------------------------------
'�d���`�F�b�N
'
'�����FWorksheets �V�[�g��
'�����FnewShapeName ���l�[���㕶����
'�߂�l�Fboolean�@�d���`�F�b�N���茋��
'------------------------------
Function isDuplication(ByVal ws As Worksheet, _
                       ByVal newShapeName As String) As Boolean

    '�ϐ��F�}�`
    Dim shape As shape
    
    '�ϐ��F����}�`���J�E���^
    Dim count As Integer: count = 0
    
    '�ΏۃV�[�g�ɑ��݂���}�`�̐������J��Ԃ�
    For Each shape In ws.Shapes
        
        '���X�g�{�b�N�X�őI�������}�`�ƈ�v���Ă���ꍇ
        If shape.Name = newShapeName Then
            
            '�J�E���^���C���N�������g����
            count = count + 1
            
        End If
    
    Next shape
    
    '����}�`���J�E���^��2�ȏ�̏ꍇ
    If count >= 1 Then
    
        '�d���`�F�b�N���茋�ʂ�TRUE�ɐݒ肷��
        isDuplication = True
        
        '�������I������B
        Exit Function
    
    End If
    
    '�d���`�F�b�N���茋�ʂ�FALSE�ɐݒ肷��
    isDuplication = False
    
End Function


