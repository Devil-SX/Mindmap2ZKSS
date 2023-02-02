' encoding = GB 2312
Sub Main()
    For Each st In ThisWorkbook.Sheets
        st.Activate 
        SetTitle
        SetLayout
        SetStyle
        GenerateID
        AutoFit
    Next
End Sub


Private Sub SetTitle()
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "模块名称"
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "功能"

    Range("C2").Select
    ActiveCell.FormulaR1C1 = "功能说明"

    Range("D2").Select
    ActiveCell.FormulaR1C1 = "测试用例"

    Range("E2").Select
    If Not IsEmpty(Range("E2")) And IsNumeric(Range("E2").Value) Then
        ActiveCell.FormulaR1C1 = "步骤"
    Else
        ActiveCell.FormulaR1C1 = "前提"
    End If

    Range("F2").Select
    ActiveCell.FormulaR1C1 = "业务操作"

    Range("G2").Select
    ActiveCell.FormulaR1C1 = "预期结果"
End Sub


Private Sub SetLayout()
' Must After SetTitle()
    Rows("2:2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("E2:G2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("E2:G2").Select
    ActiveCell.FormulaR1C1 = "测试步骤"

    Range("H3").Select
    ActiveCell.FormulaR1C1 = "测试日期"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "所用数据/正常测试"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "所用数据/异常测试"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "执行结果"

    ' Start Merge
    MergeT "A"
    MergeT "B"
    MergeT "C"
    MergeT "D"
    MergeT "H"
    MergeT "I"
    MergeT "J"
    MergeT "K"

    Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("B3").Select
    ActiveCell.FormulaR1C1 = "角色"
    MergeT "B"
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("E3").Select
    ActiveCell.FormulaR1C1 = "测试ID"
    MergeT "E"

    Range("A1:M1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub


Private Sub SetStyle()
' Must After SetLayout()
    Range("F1:L2").Select
    With Selection.Font
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Range("F2:H2").Select
    Selection.Font.Bold = False
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub


Sub GenerateID
' TestId gep_[Module]_gn_[SubModule]_id
    Dim prefix, module, midfix, submodule As String
    prefix = "gep_"
    midfix = "_gn_"
    module = PY(Range("A1").Value)
    Dim id, maxid, i, maxrow as Integer

    i = 4
    maxrow = [A1].CurrentRegion.Rows.Count
    Do While i <= maxrow And maxrow <> 4
        maxid = Cells(i, 1).MergeArea.Cells.Count
        submodule = PY(Range("A" & CStr(i)).Value)
        For id = 1 To maxid
            Range("E" & i).Select
            ActiveCell.FormulaR1C1 = prefix & module & midfix & submodule & "_" & CStr(id)
            i = i + 1
        Next
    Loop

End Sub

Private Sub MergeT(arg As String)
    Range(arg & "2:" & arg & "3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub


Private Function PY(myStr)
' Get the first character of Chinese Character, copy from https://blog.csdn.net/w_dexu/article/details/107420366
    Dim Str$, L$, Temp$, Special$
    Str = Replace(Replace(myStr, " ", ""), " ", "")
    dict = [{"吖","a";"八","b";"擦","c";"咑","d";"鵽","e";"发","f";"伽","g";"哈","h";"丌","j";"咔","k";"垃","l";"妈","m";"拿","n";"哦","o";"妑","p";"七","q";"然","r";"仨","s";"他","t";"屲","w";"夕","x";"丫","y";"帀","z"}]
    Special = "仇Q覃Q"
    For i = 1 To Len(Str)
    L = Mid$(Str, i, 1)
        j = InStr(tmp, Mid(Str, i, 1))
        If L Like "[一-龥]" Then
            Temp = Temp & IIf(j, Mid(Special, j + 1, 1), Application.Lookup(L, dict))
        Else
            Temp = Temp & L
        End If
    Next i
    PY = Temp
End Function

Private Sub AutoFit
    [A1].CurrentRegion.Select ' CurrentRegion Auto Expand equivalent to C-A
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
End Sub