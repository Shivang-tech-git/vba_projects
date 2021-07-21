Attribute VB_Name = "Module1"
Public lastcolumn As Integer
Public template As Worksheet, endorsement As Worksheet, coding As Worksheet
Public pivotData As Worksheet, macroSheet As Worksheet, Query As Worksheet
Public sh As Shape, sh2 As Shape

Sub Filter()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim pt As PivotTable
Dim pvtsheet As Worksheet
Dim pvtCache  As PivotCache
Dim pvtrng As Range, ch As Chart
Dim StartPvt As String, ch2 As Chart
Dim SrcData As Range, rngCrit As Range, vCrit As Variant
Dim objField As PivotField, cell As Range
Dim ws As Worksheet

Set macroSheet = ThisWorkbook.Sheets("Macro")
Set endorsement = ThisWorkbook.Worksheets("Endorsement")
Set pivotData = ThisWorkbook.Sheets("data to create pivot")
Set Query = ThisWorkbook.Worksheets("QUERY")
Set coding = ThisWorkbook.Worksheets("Coding Sheet")
Set template = ThisWorkbook.Worksheets("PDF TEMPLATE")

Query.Range("1:1").AutoFilter
coding.Range("1:1").AutoFilter
endorsement.Range("1:1").AutoFilter
pivotData.Rows("2:" & Rows.Count).ClearContents
'Lastrow = Source.Range("O" & Rows.Count).End(xlUp).Row
'------------------------Sum for premium ------D to J--------------
'lastcolumn = Source.Cells(1, Columns.Count).End(xlToLeft).Column + 1
'Source.Cells(1, lastcolumn).Value = "Total Amount"
'For j = 2 To Lastrow
'Source.Cells(j, lastcolumn).Value = Application.WorksheetFunction.Sum(Source.Range(Source.Cells(j, 4), Source.Cells(j, 10)))
'Next j
''-----------------Filter for Coding sheet-------------------------
'Set rngCrit = macroSheet.Range(macroSheet.Cells(4, 4), macroSheet.Cells(4, 4).End(xlDown))
'vCrit = rngCrit.Value
'Source.Range("1:1").AutoFilter Field:=15, Criteria1:=Application.Transpose(vCrit), Operator:=xlFilterValues
Call codingSheet
Call endorsementSheet
''------------------Filter for Endorsement sheet---------------------
'Set rngCrit = macroSheet.Range(macroSheet.Cells(4, 3), macroSheet.Cells(4, 3).End(xlDown))
'vCrit = rngCrit.Value
'Source.Range("1:1").AutoFilter Field:=15, Criteria1:=Application.Transpose(vCrit), Operator:=xlFilterValues

On Error Resume Next
ThisWorkbook.Worksheets.Add , ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Name = "pivot tables"
Set pvtsheet = ThisWorkbook.Worksheets("pivot tables")
Set pvtrng = pivotData.UsedRange
Set pvtCache = ThisWorkbook.PivotCaches.Create(xlDatabase, pvtrng)
''----------------Create Pivot for P&C----------------------
Set pt = pvtCache.CreatePivotTable(TableDestination:=template.Cells(13, 2), TableName:="monthlytable")
With pt.PivotFields("Currency")
.Orientation = xlRowField
.Position = 1
.Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With
With pt.PivotFields("Divide")
.Orientation = xlRowField
.Position = 2
.Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With
pt.AddDataField template.PivotTables("monthlytable").PivotFields("Amount"), , xlSum
With pt.PivotFields("Product")
       .Orientation = xlPageField
       .Position = 1
End With
pt.PivotFields("Product").CurrentPage = "P&C"
pt.RowAxisLayout xlTabularRow
''--------------------Create pivot for pollution-----------------------
Set pt = pvtCache.CreatePivotTable(TableDestination:=template.Cells(13, 7), TableName:="monthlytable2")
With pt.PivotFields("Currency")
.Orientation = xlRowField
.Position = 1
.Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With
With pt.PivotFields("Divide")
.Orientation = xlRowField
.Position = 2
.Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With
pt.AddDataField template.PivotTables("monthlytable2").PivotFields("Amount"), , xlSum
With pt.PivotFields("Product")
       .Orientation = xlPageField
       .Position = 1
End With
pt.PivotFields("Product").CurrentPage = "Pollution"
pt.RowAxisLayout xlTabularRow
''----------------- Transfer data from query to CODING sheet----------------------------
lastcolumnquery = Query.Cells(1, Columns.Count).End(xlToLeft).Column + 1
lastrowquery = Query.Cells(Rows.Count, 1).End(xlUp).Row

For Each cell In Query.Range(Query.Cells(2, lastcolumnquery), Query.Cells(lastrowquery, lastcolumnquery))
cell.Value = Query.Cells(cell.Row, 11) & " - " & Query.Cells(cell.Row, 12)
Next cell

Set rngCrit = macroSheet.Range(macroSheet.Cells(4, 4), macroSheet.Cells(4, 4).End(xlDown))
vCrit = rngCrit.Value
Query.Range("1:1").AutoFilter Field:=lastcolumnquery, Criteria1:=Application.Transpose(vCrit), Operator:=xlFilterValues

lastrowCoding = coding.Cells(Rows.Count, 4).End(xlUp).Row + 1
Query.Range("K2:K" & lastrowquery).SpecialCells(xlCellTypeVisible).Copy Destination:=coding.Range("a" & lastrowCoding)
Query.Range("L2:L" & lastrowquery).SpecialCells(xlCellTypeVisible).Copy Destination:=coding.Range("b" & lastrowCoding)
lastrowcoding2 = coding.Cells(Rows.Count, 2).End(xlUp).Row
For Each cell In coding.Range(coding.Cells(lastrowCoding, 2), coding.Cells(lastrowcoding2, 2))
cell.Offset(0, 5).Value = "$0.00"
cell.Offset(0, 9).Value = "USD"
cell.Offset(0, 12).Value = "QUERY"
Next cell
''----------------- Transfer data from query to Endorsement----------------------------

Set rngCrit = macroSheet.Range(macroSheet.Cells(4, 3), macroSheet.Cells(4, 3).End(xlDown))
vCrit = rngCrit.Value
Query.Range("1:1").AutoFilter Field:=lastcolumnquery, Criteria1:=Application.Transpose(vCrit), Operator:=xlFilterValues

lastrowendorsement = endorsement.Cells(Rows.Count, 1).End(xlUp).Row + 1
Query.Range("K2:K" & lastrowquery).SpecialCells(xlCellTypeVisible).Copy Destination:=endorsement.Range("a" & lastrowendorsement)
Query.Range("L2:L" & lastrowquery).SpecialCells(xlCellTypeVisible).Copy Destination:=endorsement.Range("b" & lastrowendorsement)
lastrowendorsement2 = endorsement.Cells(Rows.Count, 2).End(xlUp).Row
For Each cell In endorsement.Range(endorsement.Cells(lastrowendorsement, 2), endorsement.Cells(lastrowendorsement2, 2))
cell.Offset(0, 5).Value = "$0.00"
cell.Offset(0, 9).Value = "USD"
cell.Offset(0, 12).Value = "QUERY"
Next cell
On Error GoTo 0
''-------------------Create pivot chart for coding-----------------------------
Set pvtrng = coding.Range(coding.Cells(1, 1), coding.Cells(lastrowcoding2, 16))
Set pvtCache = ThisWorkbook.PivotCaches.Create(xlDatabase, pvtrng)
Set pt = pvtCache.CreatePivotTable(pvtsheet.Range("A1"), "pivotforchart")

pt.AddDataField _
Field:=pt.PivotFields("Status"), _
Function:=XlConsolidationFunction.xlCount

pt.AddFields ColumnFields:="Status", RowFields:="Activity", PageFields:="Product"
pt.PivotFields("Product").CurrentPage = "P&C"
template.Activate
template.Cells(1, 1).Select
Set sh = template.Shapes.AddChart2(, XlChartType.xlColumnClustered, 20, 275, 300, 210)
sh.Chart.SetSourceData pt.TableRange1
sh.Chart.ChartTitle.Text = "P&C"
sh.Chart.HasDataTable = True
sh.Chart.DataTable.HasBorderOutline = True
''--------------Create pivot chart for endorsement----------------------------
Set pvtrng = endorsement.UsedRange
Set pvtCache = ThisWorkbook.PivotCaches.Create(xlDatabase, pvtrng)
Set pt = pvtCache.CreatePivotTable(pvtsheet.Range("G1"), "pivotforchart2")
pt.AddDataField _
Field:=pt.PivotFields("Status"), _
Function:=XlConsolidationFunction.xlCount
pt.AddFields _
RowFields:="Activity", _
PageFields:="Product", _
ColumnFields:="Status"
pt.PivotFields("Product").CurrentPage = "Pollution"
Set sh2 = template.Shapes.AddChart2(, XlChartType.xlColumnClustered, 330, 275, 300, 210)
sh2.Chart.SetSourceData pt.TableRange1
sh2.Chart.ChartTitle.Text = "Pollution"
sh2.Chart.HasDataTable = True
sh2.Chart.DataTable.HasBorderOutline = True
''--------------------Key highlights-----------------------------
pcTotal = Application.WorksheetFunction.SumIf(pivotData.Range("A:A"), "P&C", pivotData.Range("G:G"))
pollutionTotal = Application.WorksheetFunction.SumIf(pivotData.Range("A:A"), "Pollution", pivotData.Range("G:G"))
If pcTotal > pollutionTotal Then
template.Range("E4").Value = Replace(template.Range("E4").Value, "%1", "P&C")
Else: template.Range("E4").Value = Replace(template.Range("E4").Value, "%1", "Pollution")
End If
With pivotData
    .Range("1:1").AutoFilter
    Lastrow = .Range("G" & Rows.Count).End(xlUp).Row
    .Range(.Cells(1, 1), .Cells(Lastrow, 7)).Sort key1:=.Cells(1, 7), order1:=xlDescending, Header:=xlYes
    template.Range("E5").Value = Replace(template.Range("E5").Value, "%2", .Cells(2, 7).Value)
    template.Range("E5").Value = Replace(template.Range("E5").Value, "%1", .Cells(2, 1).Value & " - " & .Cells(2, 2).Value)
End With
    template.Range("E8").Value = Replace(template.Range("E8").Value, "%", Application.WorksheetFunction.SumIf( _
                                    pivotData.Range("K:K"), "USD", pivotData.Range("G:G")))
    template.Range("E9").Value = Replace(template.Range("E9").Value, "%", Application.WorksheetFunction.SumIf( _
                                    pivotData.Range("K:K"), "CAD", pivotData.Range("G:G")))
    template.Range("E10").Value = Replace(template.Range("E10").Value, "%", Application.WorksheetFunction.SumIf( _
                                    pivotData.Range("K:K"), "PCN", pivotData.Range("G:G")))
    template.Range("E2").Value = Replace(template.Range("E2").Value, "%1", Format(Date, "mmmm"))

''----------------------The End-----------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "Tool"
Set pvtsheet = Nothing
Set pvtrng = Nothing
Set pvtCache = Nothing
Set pt = Nothing
Set coding = Nothing
Set endorsement = Nothing
End Sub
Sub export()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set coding = ThisWorkbook.Worksheets("Coding Sheet")
Set template = ThisWorkbook.Worksheets("PDF TEMPLATE")
Set endorsement = ThisWorkbook.Worksheets("Endorsement")
Set pivotData = ThisWorkbook.Sheets("data to create pivot")
Set Query = ThisWorkbook.Worksheets("QUERY")
''---------------------Export PDF Template sheet as PDF---------------------------------
'If template Is Nothing Then
'MsgBox "Please click on step 1 'Data Adjustments button.'", vbCritical, "Tool"
'Exit Sub
'End If
template.Activate
template.PageSetup.Orientation = xlLandscape
ActiveSheet.Range("A1:J35").Select
template.Select
Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
ThisWorkbook.Path & "\Americas Environmental Month End Close - " & Format(Date, "mmmm") & ".pdf", Quality:=xlQualityStandard, _
IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
''---------------------Delete Pivot tables--------------------------
template.PivotTables("monthlytable").ClearTable
template.PivotTables("monthlytable2").ClearTable
template.Range("A11:J30").Delete shift:=xlToLeft
ThisWorkbook.Sheets("pivot tables").Delete
sh.Delete
sh2.Delete
template.Range("E4").Value = "Higher premium booked under %1"
template.Range("E5").Value = "Highest premium booked in %1  $(%2)"
template.Range("E6").Value = "Total %1 Codings & Renewal Certificates, %2 Endorsements"
template.Range("E8").Value = "USD $ %"
template.Range("E9").Value = "CAD $ %"
template.Range("E10").Value = "PCN $ %"
template.Range("E2").Value = "Americas Environmental Month End Close – %1"

pivotData.Rows("2:" & Rows.Count).ClearContents
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "Tool"
Set template = Nothing
Set Source = Nothing
Set sh = Nothing
Set sh2 = Nothing
End Sub

Sub codingSheet()
Dim Lastrow As Long
'------------------------For Coding Data------------------------------------------------------------
'ws.Range(ws.Cells(2, lastcolumn), ws.Cells(Lastrow, lastcolumn)).SpecialCells(xlCellTypeVisible).Copy Destination:=ws1.Range("G2")
'ws.Range("b2:b" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws1.Range("c2")
'ws.Range("c2:c" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws1.Range("d2")
'ws.Range("R2:R" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws1.Range("e2")
'ws.Range("V2:V" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws1.Range("h2")
'ws.Range("O2:O" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws1.Range("B2")
'ws.Range("T2:T" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws1.Range("M2")

With coding
For i = 2 To .Cells(Rows.Count, 4).End(xlUp).Row
.Cells(i, 1).Value = Split(.Cells(i, 2).Value, " ")(0)
.Cells(i, 5).NumberFormat = "yyyy-mm-dd"
.Cells(i, 6).Value = "WINS"
.Cells(i, 8).NumberFormat = "yyyy-mm-dd"
.Cells(i, 2).Value = Replace(.Cells(i, 2).Value, "P&C - ", "")
.Cells(i, 2).Value = Replace(.Cells(i, 2).Value, "Pollution - ", "")
.Cells(i, 2).Value = Replace(.Cells(i, 2).Value, "Pollution ", "")
.Cells(i, 15).Value = .Cells(i, 15).Value / 60

Select Case Left((Cells(i, 4)), 3)
Case "PCN"
.Cells(i, 11).Value = "PCN"
Case "CAD"
.Cells(i, 11).Value = "CAD"
Case Else
.Cells(i, 11).Value = "USD"
End Select
    
    If Len(.Cells(i, 4).Value) <> 8 Or Len(.Cells(i, 4).Value) <> 10 Then
    .Cells(i, 4).EntireRow.Interior.ColorIndex = 6
    End If
    
    If .Cells(i, 14).Value = "Pending" Then
    .Cells(i, 14).Interior.ColorIndex = 7
    End If

'.Cells(i, 14).Value = "COMPLETED"
.Cells(i, 16).Value = "Coding"

    If .Cells(i, 7).Value <> "0" Then
    .Cells(i, 9) = "Yes"
    ElseIf .Cells(i, 7).Value = "0" Then
    .Cells(i, 10) = "Yes"
    End If
    
    If Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("People").Range("A:A"), .Cells(i, 13).Value) > 0 Then
    .Cells(i, 12).Value = Application.WorksheetFunction.VLookup(.Cells(i, 13).Value, ThisWorkbook.Worksheets("People").Range("A:B"), 2, 0)
    Else: .Cells(i, 12).Value = "Check"
    End If

Next i
template.Range("E6").Value = Replace(template.Range("E6").Value, "%1", .Cells(Rows.Count, 2).End(xlUp).Row - 1)
Lastrow = coding.Cells(Rows.Count, 4).End(xlUp).Row
End With
pivotData.Range("A1").CurrentRegion.ClearContents
coding.Range("A1:P" & Lastrow).Copy
pivotData.Range("A1").PasteSpecial

End Sub



Sub endorsementSheet()
Dim Lastrow As Long
'ws.Activate
'Lastrow = ActiveSheet.Range("O" & Rows.Count).End(xlUp).Row
''------------------------For P&C Data------------------------------------------------------------
'ws.Range(ws.Cells(2, lastcolumn), ws.Cells(Lastrow, lastcolumn)).SpecialCells(xlCellTypeVisible).Copy Destination:=ws2.Range("G2")
'ws.Range("b2:b" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws2.Range("c2")
'ws.Range("c2:c" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws2.Range("d2")
'ws.Range("R2:R" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws2.Range("e2")
'ws.Range("O2:O" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws2.Range("b2")
'ws.Range("V2:V" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws2.Range("h2")
'ws.Range("T2:T" & Lastrow).SpecialCells(xlCellTypeVisible).Copy Destination:=ws2.Range("M2")

With endorsement

For i = 2 To .Cells(Rows.Count, 4).End(xlUp).Row
.Cells(i, 1).Value = Split(.Cells(i, 2).Value, " ")(0)
.Cells(i, 5).NumberFormat = "yyyy-mm-dd"
.Cells(i, 6).Value = "WINS"
.Cells(i, 8).NumberFormat = "yyyy-mm-dd"
.Cells(i, 2).Value = Replace(.Cells(i, 2).Value, "P&C - ", "")
.Cells(i, 2).Value = Replace(.Cells(i, 2).Value, "Pollution - ", "")
.Cells(i, 2).Value = Replace(.Cells(i, 2).Value, "Pollution ", "")
.Cells(i, 15).Value = .Cells(i, 15).Value / 60

Select Case Left((Cells(i, 4)), 3)
Case "PCN"
.Cells(i, 11).Value = "PCN"
Case "CAD"
.Cells(i, 11).Value = "CAD"
Case Else
.Cells(i, 11).Value = "USD"
End Select
    
    If Len(.Cells(i, 4).Value) <> 8 Or Len(.Cells(i, 4).Value) <> 10 Then
    .Cells(i, 4).EntireRow.Interior.ColorIndex = 6
    End If
    
    If .Cells(i, 14).Value = "Pending" Then
    .Cells(i, 14).Interior.ColorIndex = 7
    End If

'.Cells(i, 14).Value = "COMPLETED"
.Cells(i, 16).Value = "Endorsment"
    If .Cells(i, 7).Value <> "0" Then
    .Cells(i, 9) = "Yes"
    ElseIf .Cells(i, 7).Value = "0" Then
    .Cells(i, 10) = "Yes"
    End If
    
If Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("People").Range("A:A"), .Cells(i, 13).Value) > 0 Then
.Cells(i, 12).Value = Application.WorksheetFunction.VLookup(.Cells(i, 13).Value, ThisWorkbook.Worksheets("People").Range("A:B"), 2, 0)
Else: .Cells(i, 12).Value = "Check"
End If

Next i
template.Range("E6").Value = Replace(template.Range("E6").Value, "%2", .Cells(Rows.Count, 4).End(xlUp).Row - 1)
Lastrow = .Cells(Rows.Count, 4).End(xlUp).Row
End With

endorsement.Range("A2:P" & Lastrow).Copy
pivotData.Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).PasteSpecial

End Sub


