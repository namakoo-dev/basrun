Option VBASupport 1
Option Explicit

' 文書の全段落を走査して、版数の行を差し替える。
' ★ Calc / Impress と同じく、道具 (basrun) は何も変えていない。
'   文書の型を知っているのはこの .bas だけ。
Sub Revise(oDoc As Object)
    Dim oEnum As Object, oPar As Object
    Dim n As Integer

    oEnum = oDoc.Text.createEnumeration()
    Do While oEnum.hasMoreElements()
        oPar = oEnum.nextElement()
        If oPar.supportsService("com.sun.star.text.Paragraph") Then
            If InStr(oPar.getString(), "2026-08-06 版") > 0 Then
                oPar.setString("情報システム部　2026-08-06 版（basrun 改訂）")
                n = n + 1
            End If
        End If
    Loop

    ' 表の件数を末尾に足す。★ 表にも到達できることを示す
    Dim oText As Object, oCur As Object
    oText = oDoc.Text
    oCur = oText.createTextCursorByRange(oText.getEnd())
    oText.insertString(oCur, Chr(13) & "表 " & oDoc.TextTables.Count & " 個 / 差し替え " & n & " 箇所", False)
End Sub
