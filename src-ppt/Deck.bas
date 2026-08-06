Option VBASupport 1
Option Explicit

' プレゼンを一括で手直しする。
'
' ★ 対象は .pptx / .ppt。どちらもマクロを格納できない (マクロは .pptm 側) ので、
'   これは文書に埋め込まれず、実行時にライブラリへ流し込まれる。
'
' ★ basrun 側は Calc 用に書いたものを 1 行も変えていない。文書の型を知っているのは
'   この .bas だけで、道具は形式に依存していない。

Private Const PREFIX As String = "【改訂】"
Private Const FOOT_TAG As String = "BASRUN p."

Sub ReviseDeck(oDoc As Object)
    Dim oPages As Object, oPage As Object, oShape As Object
    Dim i As Integer, j As Integer
    Dim oFoot As Object
    Dim sFoot As String

    oPages = oDoc.getDrawPages()

    For i = 0 To oPages.getCount() - 1
        oPage = oPages.getByIndex(i)
        sFoot = FOOT_TAG & (i + 1)

        ' --- 1) 題 ---------------------------------------------------------
        If oPage.getCount() > 0 Then
            oShape = oPage.getByIndex(0)
            If Not IsNull(oShape) Then
                ' ★ 冪等にする。既に付いていれば二度付けない。
                '   2026-08-06: 処理済みの資料に当て直して
                '   「【改訂】【改訂】」になるのを実機で確認した。
                If InStr(oShape.getString(), PREFIX) <> 1 Then
                    oShape.setString(PREFIX & oShape.getString())
                End If

                ' ★ 文字を足すと枠から溢れる。実機で 2 行に折り返すのを確認した。
                '   python-pptx で読むと文字列は正しいので、**テキストの検査では
                '   出ない欠陥**。図形側に自動縮小を持たせて、枠に収める。
                On Error Resume Next
                oShape.TextFitToSize = _
                    com.sun.star.drawing.TextFitToSizeType.AUTOFIT
                On Error Goto 0
            End If
        End If

        ' --- 2) フッタ -----------------------------------------------------
        ' ★ 既にあれば作り直さず書き換える。ここが冪等性の本体。
        oFoot = Nothing
        For j = 0 To oPage.getCount() - 1
            If InStr(oPage.getByIndex(j).getString(), FOOT_TAG) = 1 Then
                oFoot = oPage.getByIndex(j)
                Exit For
            End If
        Next j

        If IsNull(oFoot) Then
            oFoot = NewFooter(oDoc, oPage)
        End If
        oFoot.setString(sFoot)
    Next i
End Sub


' フッタ用のテキスト図形を作って配置する。
'
' ★ 位置はスライドの実寸から決める。決め打ちにすると用紙比が 4:3 と 16:9 で
'   はみ出す。幅・高さも実寸に対する割合で取る。
Private Function NewFooter(oDoc As Object, oPage As Object) As Object
    Dim oBox As Object
    Dim aSize As New com.sun.star.awt.Size
    Dim aPos As New com.sun.star.awt.Point
    Dim nW As Long, nH As Long

    nW = oPage.Width
    nH = oPage.Height

    oBox = oDoc.createInstance("com.sun.star.drawing.TextShape")
    oPage.add(oBox)

    aSize.Width = nW / 4
    aSize.Height = nH / 14
    aPos.X = nW / 20
    aPos.Y = nH - (nH / 14) - (nH / 25)   ' 下端から少し上げる

    oBox.setSize(aSize)
    oBox.setPosition(aPos)
    NewFooter = oBox
End Function
