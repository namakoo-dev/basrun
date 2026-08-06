Option VBASupport 1

Sub 金額計算
    Dim lastRow As Long
    Dim i As Long

    ' A列を一番下から上へ探す（Ctrl+↑ と同じ）。途中の空白行に強い
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' 1行目は見出しなので 2 から回す
    For i = 2 To lastRow
        Cells(i, 3).Value = Cells(i, 1).Value * Cells(i, 2).Value
    Next i
End Sub
