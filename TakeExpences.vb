Sub TakeExpences()
'    ActiveSheet.Range("Temp_Exsp[#All]").RemoveDuplicates Columns:=1, Header:= _
'        xlYes
    
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("Q3").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste

'    ActiveSheet.Range("DB_Classification[#All]").RemoveDuplicates Columns:=Array( _
'        1, 2), Header:=xlYes
    ActiveSheet.Range("DB_Classification[#All]").RemoveDuplicates Columns:=1, _
        Header:=xlYes
    ActiveWorkbook.Worksheets("Tools").ListObjects("DB_Classification").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Tools").ListObjects("DB_Classification").Sort. _
        SortFields.Add2 Key:=Range("DB_Classification[[#All],[ùí äåöàä]]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tools").ListObjects("DB_Classification").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("2022").Select
    Range("B2").Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Tools").Select
    Range("J2").Select
    ActiveSheet.Paste
End Sub