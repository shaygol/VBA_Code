for i = 1 to 10
	Sub CopytToolsTo2022()
		Sheets("Tools").Select
		Range("K2").Select
		If Selection <> 0 Then
			Application.CutCopyMode = False
			Selection.Copy
			Sheets("2022").Select
			Range("J3").Select
			Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
				:=False, Transpose:=False
    End If
Next i