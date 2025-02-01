Attribute VB_Name = "changeTextColorToRed"
'------------------------------
'選択範囲の文字列を、赤色に変更する
'------------------------------
Sub changeTextColorToRed()
    
    '変数：選択範囲
    Dim renge As range
    
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
    
    '選択範囲を取得
    Set renge = Selection
    
    '選択範囲の文字色を赤に変更
    renge.Font.Color = RGB(255, 0, 0)

End Sub
