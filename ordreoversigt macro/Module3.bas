Attribute VB_Name = "Module3"
Sub sja()
Attribute sja.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sja Makro
'

    ActiveWorkbook.RefreshAll
    Application.Run "'Prodordre forside.xlsm'!SortRute"

'
    
    Sheets("VareLinier").Select
     Range("L2").Select
    ActiveCell.FormulaR1C1 = "Save"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("L2").Select
    Selection.ClearContents
  
     Sheets("Forside").Select
    Range("C15").Select
End Sub
Sub SortRute()
Attribute SortRute.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortRute Makro
'

'
    Sheets("Rute sort").Select
    Cells.Select
    Selection.ClearContents
    Range("F22").Select
    Sheets("Rute").Select
    Columns("A:W").Select
    Selection.Copy
    Sheets("Rute sort").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
'    ActiveWorkbook.Worksheets("Rute sort").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("Rute sort").Sort.SortFields.Add Key:=Range( _
'        "B2:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    ActiveWorkbook.Worksheets("Rute sort").Sort.SortFields.Add Key:=Range( _
'        "P2:P1000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'     xlSortNormal
    With ActiveWorkbook.Worksheets("Rute sort").Sort
        .SetRange Range("A1:W1000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

    
    Sheets("VareLinier").Select
    Columns("O:O").Select
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Rows("8:555").EntireRow.AutoFit


End Sub
Sub close_Awb()
  Application.DisplayAlerts = False
  With ActiveWorkbook
    .Saved = False
    .Close
  Application.Quit

  End With
End Sub


