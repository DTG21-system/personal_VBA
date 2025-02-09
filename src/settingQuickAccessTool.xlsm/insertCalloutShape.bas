Attribute VB_Name = "insertCalloutShape"
'------------------------------
'吹き出し：線の図形を挿入する
'（図形の枠線を赤、塗りつぶしを白、文字色を赤に設定する）
'------------------------------
Sub insertCalloutShape()

    '変数：ワークシート
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    '変数：図形
    Dim shape As shape
    
    '選択範囲がセルではない場合
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルを選択してください。"
        Exit Sub
    End If
    
    '選択範囲
    Dim selectedRange As range: Set selectedRange = Selection
    
    '選択されているセルを取得
    Dim selectedCellLeft As Double: selectedCellLeft = selectedRange.Left
    Dim selectedCellTop As Double: selectedCellTop = selectedRange.Top
    
    '吹き出し：線を挿入する
    Set shape = ws.Shapes.AddShape(msoShapeLineCallout1, selectedCellLeft, selectedCellTop, 200, 100)
    
    '吹き出しの塗りつぶしを白に設定する
    shape.Fill.ForeColor.RGB = RGB(255, 255, 255)

    '吹き出しの枠線を赤色に設定する
    shape.Line.ForeColor.RGB = RGB(255, 0, 0)

    '吹き出しの文字色を赤色に設定する
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
    
    'テキストを編集中の状態にする
    shape.TextFrame2.TextRange.Select
    
End Sub
