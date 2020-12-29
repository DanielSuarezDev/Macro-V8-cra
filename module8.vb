Sub Macro2()
'
' Macro2 Macro
'

'
    Columns("AL:AL").Select
    Range("AL8").Activate
    Selection.NumberFormat = "@"
    Selection.TextToColumns Destination:=Range("AL1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("AM:AM").Select
    Range("AM8").Activate
    Selection.NumberFormat = "@"
    Selection.TextToColumns Destination:=Range("AM1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub



