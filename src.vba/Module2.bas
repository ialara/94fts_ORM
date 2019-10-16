Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("UpdateBtn")).Select
    Selection.Characters.Text = "Submit Status Update"
    With Selection.Characters(Start:=1, Length:=20).Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("J2").Select
End Sub