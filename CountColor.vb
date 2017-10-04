' @Tool Name      :CalcScore(得点記録表算出ツール)
' @Module Name    :CountColor(条件色のカウント処理)
' @Implementer    :R.Ishikawa
' @Version        :1.0
' @Last Updated   :2017/10/04

'@ Version History
'@ 1. Create New     R.Ishikawa      Ver.1.0

Function CountColor(計算範囲, 条件色セル)
    Application.Volatile (True)
    CountColor = 0
    Application.ScreenUpdating = False
    For y = 1 To 計算範囲.Columns.Count
        Application.ScreenUpdating = False
        For x = 1 To 計算範囲.Rows.Count
            If 計算範囲.Rows(x).Columns(y).Interior.ColorIndex = 条件色セル.Interior.ColorIndex Then
                CountColor = CountColor + 1
            End If
        Next x
        Application.ScreenUpdating = True
    Next y
    Application.ScreenUpdating = True
End Function
