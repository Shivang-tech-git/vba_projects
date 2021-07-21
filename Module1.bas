Attribute VB_Name = "Module1"
Sub Coopergaypremium_tool()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Declare constants----------------------------------
Const colA = 1
Dim colK As Integer 'Last row in raw file
Dim colK1 As Integer
Dim column_name(1 To 3) As String
Dim alpha(1 To 11) As Integer

caption_col = 3
gp_col = 4
np_col = 5
com_col = 6
pivot_start_col = 7

inception_date = 7
expiry_date = 8

column_name(1) = "YOA"
column_name(2) = "Policy No"
column_name(3) = "Section"

'''---------------Error Handling Routine-----------------------------
'On Error Resume Next
'On Error GoTo errorHandler:
'On Error GoTo 0
''-------------Variable Declarations---------------------------------------
Dim filename As String, i As Integer, header_row As Integer, raw_range As Range, new_wb As Workbook
Dim macroSheet As Worksheet, euro_prog As Worksheet, raw_file As Workbook, raw_sheet As Worksheet
Dim pvtsht As Object, fso As Object
Dim rng As Range
Dim c As Range
Dim pc As PivotCache
Dim pt As PivotTable
''-------------Set Object References----------------------------------------
Set fso = CreateObject("Scripting.FileSystemObject")
Set macroSheet = ThisWorkbook.Worksheets("Macro")
''-----------------Dialog box-------------------------------------
With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Select Cooper Gay file"
    .Show
    filename = .SelectedItems(1)
End With
''-----------------Clerical stuff---------------------------------
deleteWS:
For ws = 1 To ThisWorkbook.Worksheets.Count
    If ThisWorkbook.Worksheets(ws).Name <> "Macro" Then
    ThisWorkbook.Worksheets(ws).Delete
    GoTo deleteWS
    End If
Next ws

Workbooks.Add
Set new_wb = ActiveWorkbook
Set euro_prog = new_wb.Worksheets(1)
euro_prog.Name = "premium data"

Workbooks.Open filename
Set raw_file = ActiveWorkbook
Set raw_sheet = ActiveWorkbook.Worksheets("Premium Paid")

''--------------------------------------------------------------------------
''======================= copy paste Premium sheet from col a to col 138 ==================================
'On Error GoTo errorHandler
raw_sheet.Activate
header_row = raw_sheet.Range("1:20").Find("Class").Row
On Error GoTo 0

last_row = raw_sheet.Cells(Rows.Count, 1).End(xlUp).Row
colK = raw_sheet.Cells(header_row, Columns.Count).End(xlToLeft).Column
Set raw_range = raw_sheet.Range(raw_sheet.Cells(header_row, colA), raw_sheet.Cells(last_row, colK))
raw_range.Copy
euro_prog.Cells(1, 1).PasteSpecial xlPasteValues
total_row = euro_prog.Range("A:A").Find("Totals").Row
euro_prog.Rows(total_row & ":" & Rows.Count).Delete
euro_prog.Rows("2:2").Delete
    If Not IsDate(euro_prog.Cells(2, inception_date).Value) Then
    euro_prog.Rows("2:2").Delete
    End If
last_row = euro_prog.Cells(Rows.Count, 1).End(xlUp).Row

euro_prog.Range("H:J").EntireColumn.Insert
euro_prog.Range("H1").Value = "YOA"

euro_prog.Range("I1").Value = "Policy No"

euro_prog.Range("J1").Value = "Section"






''====================== Apply formulas ===================================================
With euro_prog
'plus = colK + 1
'For Each col_item In column_name
'   .Cells(1, plus) = col_item
'   plus = plus + 1
'Next col_item

'On Error GoTo errorHandler

share_col = .Range("1:1").Find("share", lookat:=xlWhole).Column
gross_premium_col = .Range("1:1").Find("GROSS PREMIUM", after:=.Cells(1, share_col), lookat:=xlWhole).Column
cargo_wr_gge_col = .Range("1:1").Find("CARGO WR GGE", after:=.Cells(1, share_col), lookat:=xlWhole).Column
net_premium_col = .Range("1:1").Find("Net premium", after:=.Cells(1, share_col), lookat:=xlWhole).Column
commission_col = .Range("1:1").Find("COMMISSION", after:=.Cells(1, share_col), lookat:=xlWhole).Column

'On Error GoTo 0
On Error Resume Next
For col = 1 To 3
    For Row = 2 To last_row

           .Cells(Row, 8).Value = Year(.Cells(Row, inception_date).Value)
           .Cells(Row, 9).Value = "FRM0000002MA" & Right((.Cells(Row, expiry_date).Value), 2) & "A"
           .Cells(Row, 10).Value = "Marine"

    Next Row
Next col
On Error GoTo 0
col_last = .Cells(1, Columns.Count).End(xlToLeft).Column
Set rng = .Range(.Cells(1, 1), .Cells(last_row, col_last))
End With

'''---------------------- create pivot table-----------------------------------------------

new_wb.Worksheets.Add after:=new_wb.Worksheets(new_wb.Worksheets.Count)
new_wb.Worksheets(new_wb.Worksheets.Count).Name = "Pivot"
Set pvtsht = new_wb.Worksheets("Pivot")
'
Set pc = new_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng)
Set pt = pc.CreatePivotTable(TableDestination:=pvtsht.Cells(1, 1), TableName:="pivot table 1")
''
With pt.PivotFields("Policy No")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With
With pt.PivotFields("Section")
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With
With pt.PivotFields(cargo_wr_gge_col)
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With

pt.AddDataField pvtsht.PivotTables("pivot table 1").PivotFields(gross_premium_col), , xlSum
pt.AddDataField pvtsht.PivotTables("pivot table 1").PivotFields(net_premium_col), , xlSum
pt.AddDataField pvtsht.PivotTables("pivot table 1").PivotFields(commission_col), , xlSum

For i = 1 To pt.DataFields.Count
    pt.DataFields(i).NumberFormat = "0.00"
Next i

 pt.RowAxisLayout xlTabularRow
''------------------- marine and war --------------------

With pvtsht
war_row = .Cells(Rows.Count, 1).End(xlUp).Row + 2
pvtsht_lastrow = .Cells(Rows.Count, 1).End(xlUp).Row + 2
''----------------- war column headings --------------------------
.Cells(pvtsht_lastrow, gp_col) = "GROSS PREMIUM2"
.Cells(pvtsht_lastrow, np_col) = "NET PREMIUM"
.Cells(pvtsht_lastrow, com_col) = "COMMISSION"
.Cells(pvtsht_lastrow, caption_col) = "SECTION"

''----------- for each cargo row calculate gross premium share for war ------------------
For each_cargo = 1 To pt.RowFields(3).PivotItems.Count
    pvtsht_lastrow = pvtsht_lastrow + 1
    cargo_value = pt.RowFields(3).PivotItems(each_cargo)
    cargo_upon_gross_premium = cargo_value / pt.GetPivotData(pt.PivotFields(gross_premium_col), pt.PivotFields(cargo_wr_gge_col), cargo_value)
    
    .Cells(pvtsht_lastrow, caption_col) = "WAR"
    .Cells(pvtsht_lastrow, gp_col) = cargo_upon_gross_premium * pt.GetPivotData(pt.PivotFields(gross_premium_col), pt.PivotFields(cargo_wr_gge_col), cargo_value)
    .Cells(pvtsht_lastrow, np_col) = cargo_upon_gross_premium * pt.GetPivotData(pt.PivotFields(net_premium_col), pt.PivotFields(cargo_wr_gge_col), cargo_value)
    .Cells(pvtsht_lastrow, com_col) = cargo_upon_gross_premium * pt.GetPivotData(pt.PivotFields(commission_col), pt.PivotFields(cargo_wr_gge_col), cargo_value)
Next each_cargo
''------------ sum of each column of war ------------------------------------------------
    war_total_row = pvtsht_lastrow + 1
    .Cells(war_total_row, caption_col) = "Total"
    .Cells(war_total_row, gp_col) = WorksheetFunction.Sum(.Range(.Cells(war_row + 1, gp_col), .Cells(pvtsht_lastrow, gp_col)))
    .Cells(war_total_row, np_col) = WorksheetFunction.Sum(.Range(.Cells(war_row + 1, np_col), .Cells(pvtsht_lastrow, np_col)))
    .Cells(war_total_row, com_col) = WorksheetFunction.Sum(.Range(.Cells(war_row + 1, com_col), .Cells(pvtsht_lastrow, com_col)))
''------------ calculate marine ---------------------------------------------------------
    marine_row = war_total_row + 2
    .Cells(marine_row, caption_col) = "SECTION"
    .Cells(marine_row, gp_col) = "GROSS PREMIUM2"
    .Cells(marine_row, np_col) = "NET PREMIUM"
    .Cells(marine_row, com_col) = "COMMISSION"
    marine_row = marine_row + 1
    .Cells(marine_row, caption_col) = "MARINE"
    .Cells(marine_row, gp_col) = pt.GetPivotData(pt.PivotFields(gross_premium_col)) - .Cells(war_total_row, gp_col).Value
    .Cells(marine_row, np_col) = pt.GetPivotData(pt.PivotFields(net_premium_col)) - .Cells(war_total_row, np_col).Value
    .Cells(marine_row, com_col) = pt.GetPivotData(pt.PivotFields(commission_col)) - .Cells(war_total_row, com_col).Value
    .Cells(war_row, caption_col).CurrentRegion.Borders.LineStyle = xlContinuous
''---------------- pivot sheet for war ------------------------------
Dim war_rng As Range, pc_war As PivotCache, pt_war As PivotTable
Set war_rng = .Range(.Cells(war_row, caption_col), .Cells(war_total_row - 1, com_col))
Set pc_war = new_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=war_rng)
Set pt_war = pc_war.CreatePivotTable(TableDestination:=.Cells(war_row, pivot_start_col), TableName:="pivot table war")
pt_war.PivotFields("SECTION").Orientation = xlRowField
pt_war.PivotFields("SECTION").Position = 1
pt_war.AddDataField pvtsht.PivotTables("pivot table war").PivotFields("GROSS PREMIUM2"), , xlSum
pt_war.AddDataField pvtsht.PivotTables("pivot table war").PivotFields("NET PREMIUM"), , xlSum
pt_war.AddDataField pvtsht.PivotTables("pivot table war").PivotFields("COMMISSION"), , xlSum
pt_war.RowAxisLayout xlTabularRow
''---------------- pivot sheet for marine ------------------------------
Dim marine_rng As Range, pc_marine As PivotCache, pt_marine As PivotTable
Set marine_rng = .Range(.Cells(marine_row - 1, caption_col), .Cells(marine_row, com_col))
Set pc_marine = new_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=marine_rng)
Set pt_marine = pc_marine.CreatePivotTable(TableDestination:=.Cells(marine_row, pivot_start_col), TableName:="pivot table marine")
pt_marine.PivotFields("SECTION").Orientation = xlRowField
pt_marine.PivotFields("SECTION").Position = 1
pt_marine.AddDataField pvtsht.PivotTables("pivot table marine").PivotFields("GROSS PREMIUM2"), , xlSum
pt_marine.AddDataField pvtsht.PivotTables("pivot table marine").PivotFields("NET PREMIUM"), , xlSum
pt_marine.AddDataField pvtsht.PivotTables("pivot table marine").PivotFields("COMMISSION"), , xlSum
pt_marine.RowAxisLayout xlTabularRow
End With
''-------------The End-------------------------------------
With euro_prog
'.Range("G:G").NumberFormat = "DD/MM/YYYY"
.Range("A1:EK1").Interior.ColorIndex = 35
.Range("A:EK").EntireColumn.AutoFit
.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
End With
''----------------- save new file -------------------------
num = 1
save_path = Replace(filename, ".xls", "Final.xls")
While fso.FileExists(save_path)
save_path = Replace(save_path, ".xls", num & " .xls")
num = num + 1
Wend
new_wb.SaveAs save_path
''----------------------------------------------------------
raw_file.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "Cooper Gay Tool"
Set new_wb = Nothing
Set raw_sheet = Nothing
Set raw_file = Nothing
Set euro_prog = Nothing
Set macroSheet = Nothing
End Sub
