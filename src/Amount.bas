Option VBASupport 1
Option Explicit

' 単価 x 数量 を E 列に入れ、日付を yyyy/mm/dd、金額をカンマ編集にする。
'
' ★ このファイルは平文の .bas であって、文書には格納されていない。
'    実行時にライブラリへ流し込まれる。文書は普通の .xlsx のまま。
'
' ★ LibreOffice の罠 (2026-08-04 に踏んだ):
'    Cells(Rows.Count, "A").End(xlUp).Row は列を文字で指定すると
'    静かに例外を投げて Sub ごと中断する。数値の列番号なら通る。
Sub FillAmounts(oDoc As Object)
    Dim oSheet As Object, oCell As Object
    Dim lastRow As Long, i As Long
    Dim price As Double, qty As Double

    ' ★ ThisComponent に頼らない。呼び出し側が文書を渡す
    oSheet = oDoc.Sheets.getByIndex(0)

    ' 最終行を探す (A 列を上から走査。UsedRange に頼らない)
    lastRow = 1
    Do While oSheet.getCellByPosition(0, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1   ' 0 起点の最終データ行

    For i = 1 To lastRow
        price = oSheet.getCellByPosition(2, i).getValue()
        qty = oSheet.getCellByPosition(3, i).getValue()
        oCell = oSheet.getCellByPosition(4, i)
        oCell.setValue(price * qty)
    Next i

    ' 書式: 金額をカンマ編集
    Dim oFormats As Object, sFmt As String, nFmt As Long
    Dim aLocale As New com.sun.star.lang.Locale
    oFormats = oDoc.getNumberFormats()
    sFmt = "#,##0"
    nFmt = oFormats.queryKey(sFmt, aLocale, False)
    If nFmt = -1 Then nFmt = oFormats.addNew(sFmt, aLocale)
    oSheet.getCellRangeByPosition(4, 1, 4, lastRow).NumberFormat = nFmt

    ' 呼び出し側が結果を確認できるように、処理件数を A1 の右隣に残す
    oSheet.getCellByPosition(6, 0).setString("rows=" & lastRow)
End Sub
