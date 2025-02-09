Attribute VB_Name = "resetSheets"
'------------------------------
'シートのセルをA1、表示倍率を100%に変更する
'------------------------------
Sub resetSheets()

    '変数：ワークシート
    Dim ws As Worksheet
    
    'シートの件数分、繰り返す
    For Each ws In ActiveWorkbook.Worksheets
    
        'セルA1を選択
        ws.Activate
        ws.range("A1").Select
        
        '表示倍率を100%に設定する
        ActiveWindow.Zoom = 100
    
    Next ws
    
    '対象ブックの一番先頭のシートを表示する
    ThisWorkbook.Sheets(1).Select
    
End Sub
