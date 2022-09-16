Sub CopytToolsTo2022()
    i = 0
    For Each cell In Worksheets("Tools").Range("sum[ñëåí]")
        If cell <> 0 Then
            Application.CutCopyMode = False
            cell.Copy
            Sheets("2022").Select
            Range("J3").Offset(i).Select
            If (Selection.MergeCells = False) Then
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            End If
        End If
    i = i + 1
    Next cell
End Sub