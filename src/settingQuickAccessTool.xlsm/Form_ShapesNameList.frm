VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_ShapesNameList 
   Caption         =   "ShapesNameList"
   ClientHeight    =   4170
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4680
   OleObjectBlob   =   "Form_ShapesNameList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Form_ShapesNameList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------
'図形名称一覧_初期処理
'対象シート内に存在する図形をユーザーフォームに一覧表示する
'------------------------------
Private Sub UserForm_Initialize()

    '変数：ワークシート
    Dim ws As Worksheet
    
    '変数：図形
    Dim shape As shape
    
    '変数：図形名称
    Dim shapeNames As String

    '変数：繰り返し初回ループ制御用
    Dim firstLoop As Boolean: firstLoop = True
    
    ' アクティブシートを取得
    Set ws = ActiveSheet
    
    ' シート内に存在する、図形の数だけ繰り返す
    For Each shape In ws.Shapes
        
        '一番最初に取得した図形名をテキストボックスに設定する
        If firstLoop Then
            
            '図形名称を、テキストボックスに設定する
            TextBox_ShapeName = shape.Name
            
            '繰り返し初回ループ制御用変数をfalseに設定する
            firstLoop = False
        
        End If
        
        '図形名称一覧ユーザーフォームのリストボックスに追加する
        ListBox_Shapes.AddItem shape.Name
                
    Next shape
    
End Sub

'------------------------------
'図形名称一覧_リストボックス選択値切り替え
'------------------------------
Private Sub ListBox_Shapes_Change()

    '変数：ワークシート
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    '変数：図形
    Dim shape As shape
    
    'リストボックスがNULLの場合、処理を終了する。
    If IsNull(ListBox_Shapes) Then
        Exit Sub
    End If
    
    '変数：選択図形名称
    Dim selectedShapeName As String: selectedShapeName = ListBox_Shapes.Value
    
    '選択した図形名をテキストボックスに設定
    TextBox_ShapeName.Value = selectedShapeName
    
    'アクティブシートを取得
    For Each shape In ws.Shapes
    
        'リストボックスで選択した図形名称が一致している場合
        If shape.Name = selectedShapeName Then
        
            '対象の図形を選択状態にする
            shape.Select
            
            '繰り返し処理を抜ける
            Exit For
        
        End If
    
    Next shape
    
End Sub

'------------------------------
'図形名称一覧_リストボックスキーボタン押下
'------------------------------
Private Sub ListBox_Shapes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'エンターキーが押された場合
    If KeyCode = vbKeyReturn Then
        
        'ユーザーフォームを閉じる
        Unload Me
        
    End If

End Sub


'------------------------------
'図形名称一覧_リネームボタン押下
'------------------------------
Private Sub Button_Rename_Click()

    '変数：ワークシート
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    'リストボックスがNULLの場合、処理を終了する。
    If IsNull(ListBox_Shapes) Then
        Exit Sub
    End If
    
    '変数：選択図形名
    Dim selectedShapeName As String: selectedShapeName = ListBox_Shapes.Value
    
    '変数：リネーム図形名称
    Dim newShapeName As String: newShapeName = TextBox_ShapeName.Value
    
    '変数：図形
    Dim shape As shape
    
    '重複チェック
    If isDuplication(ws, newShapeName) Then
        
        '重複エラーのメッセージを設定する。
        MsgBox "同一名を入力しているか、リネーム後の図形名称が重複しています。" _
               & vbCrLf & _
               "別の名前を入力してください。", Buttons:=vbCritical
        
        '処理を終了する
        Exit Sub
        
    End If
    
    
    '図形名称を変更
    For Each shape In ws.Shapes
        
        'リストボックスで選択した図形と一致している場合
        If shape.Name = selectedShapeName Then
        
            '図形名称を変更する
            shape.Name = newShapeName
            
            '処理を抜ける
            Exit For
        
        End If
    
    Next shape
    
    'リネーム完了のメッセージを出力する
    MsgBox "図形名の更新が完了しました。" _
           & vbCrLf & _
           "変更前：" & selectedShapeName & " " & _
           "変更後：" & newShapeName
    
    
    'リストボックスをクリアする
    ListBox_Shapes.Clear
    
    'リストボックスを更新する
    For Each shape In ws.Shapes
    
        '図形名称一覧ユーザーフォームのリストボックスに追加する
        ListBox_Shapes.AddItem shape.Name
    
    Next shape
    
End Sub


'------------------------------
'重複チェック
'
'引数：Worksheets シート名
'引数：newShapeName リネーム後文字列
'戻り値：boolean　重複チェック判定結果
'------------------------------
Function isDuplication(ByVal ws As Worksheet, _
                       ByVal newShapeName As String) As Boolean

    '変数：図形
    Dim shape As shape
    
    '変数：同一図形名カウンタ
    Dim count As Integer: count = 0
    
    '対象シートに存在する図形の数だけ繰り返す
    For Each shape In ws.Shapes
        
        'リストボックスで選択した図形と一致している場合
        If shape.Name = newShapeName Then
            
            'カウンタをインクリメントする
            count = count + 1
            
        End If
    
    Next shape
    
    '同一図形名カウンタが2以上の場合
    If count >= 1 Then
    
        '重複チェック判定結果をTRUEに設定する
        isDuplication = True
        
        '処理を終了する。
        Exit Function
    
    End If
    
    '重複チェック判定結果をFALSEに設定する
    isDuplication = False
    
End Function


