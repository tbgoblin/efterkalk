Attribute VB_Name = "Module25"
Sub sj()
Attribute sj.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sj Makro
'

'
    Sheets("LaserListe").Select
    Range("Tabel_Forespørgsel_fra_Visma4711614[[#Headers],[ProdNr.]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("Forside").Select
End Sub
