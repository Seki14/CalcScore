' @Tool Name      :CalcScore(得点記録表算出ツール)
' @Module Name    :SectionChange(クリックしたセルの色を変える)
' @Implementer    :R.Ishikawa
' @Version        :1.0
' @Last Updated   :2017/10/04

'@ Version History
'@ 1. Create New     R.Ishikawa      Ver.1.0


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
 If (Target.Row >= 9 And Target.Row <= 23) And (Target.Column = 3 Or Target.Column = 5 Or Target.Column = 7) Then
     Select Case Target.Interior.ColorIndex
     Case Is = xlNone
     Target.Interior.ColorIndex = 3
     Case Else
     Target.Interior.ColorIndex = xlNone
     End Select
 ElseIf (Target.Row >= 9 And Target.Row <= 23) And (Target.Column = 4 Or Target.Column = 6 Or Target.Column = 8) Then
     Select Case Target.Interior.ColorIndex
     Case Is = xlNone
     Target.Interior.ColorIndex = 4
     Case Else
     Target.Interior.ColorIndex = xlNone
     End Select
 End If
End Sub
