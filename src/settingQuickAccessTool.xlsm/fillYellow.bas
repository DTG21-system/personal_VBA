Attribute VB_Name = "fillYellow"
'------------------------------
'選択範囲の文字列を、赤色に変更する
'------------------------------
Sub fillYellow()
    
    '選択範囲が存在しない場合
    If Selection Is Nothing Then
        MsgBox "セルを選択してください。"
        Exit Sub
    End If

    '選択範囲がセルではない場合
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルを選択してください。"
        Exit Sub
    End If
    
    '変数：選択範囲
    Dim range As range
    
    ' 選択範囲を取得
    Set range = Selection
    
    ' 選択範囲のセルを黄色に塗りつぶす
    range.Interior.Color = RGB(255, 255, 0)
    
End Sub
