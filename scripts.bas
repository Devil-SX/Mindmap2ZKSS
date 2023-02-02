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
    ActiveCell.FormulaR1C1 = "ģ������"
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "����"

    Range("C2").Select
    ActiveCell.FormulaR1C1 = "����˵��"

    Range("D2").Select
    ActiveCell.FormulaR1C1 = "��������"

    Range("E2").Select
    If Not IsEmpty(Range("E2")) And IsNumeric(Range("E2").Value) Then
        ActiveCell.FormulaR1C1 = "����"
    Else
        ActiveCell.FormulaR1C1 = "ǰ��"
    End If

    Range("F2").Select
    ActiveCell.FormulaR1C1 = "ҵ�����"

    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Ԥ�ڽ��"
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
    ActiveCell.FormulaR1C1 = "���Բ���"

    Range("H3").Select
    ActiveCell.FormulaR1C1 = "��������"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "��������/��������"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "��������/�쳣����"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "ִ�н��"

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
    ActiveCell.FormulaR1C1 = "��ɫ"
    MergeT "B"
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("E3").Select
    ActiveCell.FormulaR1C1 = "����ID"
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
    dict = [{"߹","a";"��","b";"��","c";"��","d";"�z","e";"��","f";"٤","g";"��","h";"آ","j";"��","k";"��","l";"��","m";"��","n";"Ŷ","o";"�r","p";"��","q";"Ȼ","r";"��","s";"��","t";"��","w";"Ϧ","x";"Ѿ","y";"��","z"}]
    Special = "��Q��Q"
    For i = 1 To Len(Str)
    L = Mid$(Str, i, 1)
        j = InStr(tmp, Mid(Str, i, 1))
        If L Like "[һ-��]" Then
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