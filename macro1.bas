Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+Shift+I
'
    Set tbl = ActiveSheet.ListObjects(1)
    tbl.HeaderRowRange.Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Font
        .Name = "Aptos Narrow"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Sub
