Attribute VB_Name = "Module1"
Sub cyber_reprt()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim gmrWB As Workbook
Dim path As String, filename As String
Dim detailExtractSheet As Worksheet
Dim dataSheet As Worksheet
Dim lr2 As Long, filterRow As Integer
Dim lc As Long, cell As Range
Dim pvtsht As Worksheet
Dim rng As Range, underwriterFilterValue As Range
Dim pc As PivotCache
Dim pt As PivotTable
Dim columnName(1 To 15) As String
Dim macroSheet As Worksheet, visibleCells As Range
Dim jmReport As Workbook, area As Range, visibleFilterCells As Range
''-----------------Column Names for pivot table-------------------
columnName(1) = "Insured"
columnName(2) = "Inception"
columnName(3) = "Policy Number"
columnName(4) = "Underwriter"
columnName(5) = "Broker"
columnName(6) = "New_Renew"
columnName(7) = "Primary/ Excess"
columnName(8) = "XL Lead Y/N"
columnName(9) = "Insured Country"
columnName(10) = "LimitValue100Pcnt_Highest"
columnName(11) = "XL Share %"
columnName(12) = "SIR Limit"
columnName(13) = "AttachmentPoint100Pcnt"
columnName(14) = "Brokerage %"
columnName(15) = "Premium Booked XL Share"
''--------------------Create New workbook--------------
Workbooks.Add
Set jmReport = ActiveWorkbook
''--------------Delete all sheets except Data and Tool------------
For Each Worksheet In ThisWorkbook.Worksheets
If Worksheet.Name <> "Tool" And Worksheet.Name <> "Data" Then
Worksheet.Delete
End If
Next Worksheet
''--------------Get path from macroSheet------------------
Set macroSheet = ThisWorkbook.Worksheets("Tool")
If macroSheet.Range("B3").Value = "" Then
MsgBox "Please enter path of 'GMR Professional Report' in cell B3 and Run again.", vbInformation, "Cyber Report Tool"
Exit Sub
Else: path = macroSheet.Range("B3").Value
End If
''-------------Get data from GMR workbook-------------------
Set dataSheet = ThisWorkbook.Sheets("Data")
On Error GoTo errorHandlerOpen
Workbooks.Open (path)
Set gmrWB = ActiveWorkbook
Set detailExtractSheet = gmrWB.Worksheets("DetailExtract")
On Error Resume Next
dataSheet.UsedRange.Delete
''-------------Loop to find header row ----------------
For y = 1 To 6
If Application.WorksheetFunction.CountA(detailExtractSheet.Range("A" & y).EntireRow) > 20 Then
filterRow = y
Exit For
End If
Next y
On Error GoTo errorHandlerFilter
detailExtractSheet.Range(filterRow & ":" & filterRow).AutoFilter
detailExtractSheet.Range(filterRow & ":" & filterRow).AutoFilter Field:=6, Criteria1:=macroSheet.Range("D3").Value, _
                                    Operator:=xlOr, Criteria2:=macroSheet.Range("E3").Value
On Error Resume Next
lastrow = detailExtractSheet.Cells.SpecialCells(xlCellTypeVisible)(Rows.Count, 1).End(xlUp).Row
detailExtractSheet.Range("A4:DD" & lastrow).Copy
dataSheet.Cells(1, 1).PasteSpecial
''---------------Create pivot table for each year-----------
If macroSheet.Range("C3").Value = "" Then
MsgBox "Please enter all inception year's in column C and Run again.", vbInformation, "Cyber Report Tool"
Else
For Each cell In macroSheet.Range(macroSheet.Cells(3, 4), macroSheet.Cells(3, 4).End(xlToRight))
Set rng = dataSheet.Range("A1").CurrentRegion
ThisWorkbook.Worksheets.Add , ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
Set pvtsht = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
pvtsht.Name = cell
Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng)
Set pt = pc.CreatePivotTable(TableDestination:=pvtsht.Cells(1, 1), TableName:=cell.Value)
''-----------------Add columns to row field--------------
For x = 1 To 15
With pt.PivotFields(columnName(x))
.Orientation = xlRowField
.Position = x
.Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With
Next x
''---------------Add columns for filters-----------------
Z = 3
For y = 1 To 4
With pt.PivotFields(macroSheet.Cells(Z, 3).Value)
       .Orientation = xlPageField
       .Position = y
End With
Z = Z + 1
Next y
With pt.PivotFields(macroSheet.Cells(9, 3).Value)
       .Orientation = xlPageField
       .Position = 5
End With
''--------------------Apply pivot filters from macro sheet----------------------
With macroSheet
pt.PivotFields(.Cells(3, 3).Value).CurrentPage = cell.Value
pt.PivotFields(.Cells(4, 3).Value).CurrentPage = .Cells(4, cell.Column).Value
pt.PivotFields(.Cells(5, 3).Value).CurrentPage = .Cells(5, cell.Column).Value
End With

With pt.PivotFields(macroSheet.Cells(6, 3).Value)
        For i = 1 To .PivotItems.Count
        If .PivotItems(.PivotItems(i).Name) = macroSheet.Cells(6, cell.Column).Value Or _
        .PivotItems(.PivotItems(i).Name) = macroSheet.Cells(7, cell.Column).Value Or _
        .PivotItems(.PivotItems(i).Name) = macroSheet.Cells(8, cell.Column).Value Then
        .PivotItems(.PivotItems(i).Name).Visible = True
        Else: .PivotItems(.PivotItems(i).Name).Visible = False
        End If
        Next i
End With

With pt.PivotFields(macroSheet.Cells(9, 3).Value)
        For i = 1 To .PivotItems.Count - 1
        .PivotItems(.PivotItems(i).Name).Visible = False
        Next i
        lastVisible = False
        For Each underwriterFilterValue In macroSheet.Range(macroSheet.Cells(9, cell.Column), macroSheet.Cells(9, cell.Column).End(xlDown))
        .PivotItems(underwriterFilterValue.Value).Visible = True
        If (.PivotItems(.PivotItems(.PivotItems.Count).Name) = underwriterFilterValue.Value) Then
        lastVisible = True
        End If
        Next underwriterFilterValue
        If lastVisible <> True Then
        .PivotItems(.PivotItems(.PivotItems.Count).Name).Visible = False
        End If
        
End With
pt.RowAxisLayout xlTabularRow
''------------Copy pivot data to JMReport sheet----------------
pvtsht.UsedRange.Copy
jmReport.Worksheets.Add after:=jmReport.Worksheets(jmReport.Worksheets.Count)

With jmReport.Worksheets(jmReport.Worksheets.Count)
.Name = cell.Value
.Range("A1").PasteSpecial xlPasteValues
''--------Apply formula and formatting the jmReport worksheet---------------
lastRowJM = .Cells(Rows.Count, 1).End(xlUp).Row
For rowNum = 4 To lastRowJM
.Cells(rowNum, 10).Value = .Cells(rowNum, 10).Value / 100
.Cells(rowNum, 13).Value = .Cells(rowNum, 13).Value / 100
Next rowNum
.Cells.Font.Name = "Arial"
.Cells.Font.Size = 10
.Cells.HorizontalAlignment = xlCenter
.Range("B:B").NumberFormat = "dd-mm-yyyy"
.Range("I:I").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
.Range("J:J").NumberFormat = "0.00%"
.Range("K:K").NumberFormat = "#,##0.00"
.Range("L:L").NumberFormat = "#,##0.00"
.Range("M:M").NumberFormat = "0.00%"
.Range("N:N").NumberFormat = "#,##0.00"
.Cells(lastRowJM, 1).EntireRow.Delete
.Range("1:6").Delete
''------------Remove last row (Grand total)-----------------------------------

''-----------Delete UK policy row--------------------------------------------
lastrow = .Cells(Rows.Count, 1).End(xlUp).Row
For rowA = 1 To lastrow
    If .Cells(rowA, 1) = "" Then
    .Cells(rowA, 1).Value = .Cells(rowA, 1).Offset(-1, 0).Value
    .Cells(rowA, 2).Value = .Cells(rowA, 2).Offset(-1, 0).Value
    End If
Next rowA
'jmReport.Worksheets("Sheet1").Range("1:1").AutoFilter
'jmReport.Worksheets("Sheet1").Rows("1:" & Rows.Count).ClearContents
'.Range("A1:A" & lastrow).Copy
'jmReport.Worksheets("Sheet1").Range("A1").PasteSpecial xlPasteValues
'jmReport.Worksheets("Sheet1").Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes

'lastrow = jmReport.Worksheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row

'jmReport.Worksheets("Sheet1").Range("B1").Formula = "=COUNTIFS('" & cell.Value & "'!A:A,Sheet1!A1,'" & cell.Value & "'!C:C," & Chr(34) & "IE*" & Chr(34) & ")"
'jmReport.Worksheets("Sheet1").Range("C1").Formula = "=COUNTIFS('" & cell.Value & "'!A:A,Sheet1!A1,'" & cell.Value & "'!C:C," & Chr(34) & "UK*" & Chr(34) & ")"
'jmReport.Worksheets("Sheet1").Activate
'ActiveSheet.Range("B1:B" & lastrow).Select
'Selection.FillDown
'ActiveSheet.Range("C1:C" & lastrow).Select
'Selection.FillDown

'jmReport.Worksheets("Sheet1").Range("1:1").AutoFilter Field:=2, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>"
'jmReport.Worksheets("Sheet1").Range("1:1").AutoFilter Field:=3, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>"

'lastrow = jmReport.Worksheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
'For Each visibleFilterCells In jmReport.Worksheets("Sheet1").Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible)
'.Range("1:1").AutoFilter Field:=1, Criteria1:=Trim(visibleFilterCells.Value), Operator:=xlFilterValues
'lastRowNew = .Cells(Rows.Count, 1).End(xlUp).Row
'moreThanOne:
'    For Each visibleCells In .Range("C2:C" & lastRowNew).SpecialCells(xlCellTypeVisible)
'        If Left(visibleCells.Value, 2) = "UK" Then
'        visibleCells.EntireRow.Delete
'        GoTo moreThanOne
'        End If
'    Next visibleCells
'Next visibleFilterCells
.Range("1:1").AutoFilter
''--------------------------------------------------------------------------
lastrow = .Cells(Rows.Count, 1).End(xlUp).Row
For q = 2 To lastrow
.Range("A" & q).Value = WorksheetFunction.Trim(.Range("A" & q).Value)
.Range("H" & q).Value = WorksheetFunction.Trim(.Range("H" & q).Value)
Next q
.Range("A1:N" & lastrow).Sort .Range("B2:B" & lastrow), xlAscending, , , , , , xlYes, , False, xlSortColumns, xlPinYin, xlSortNormal
.Range("O1").Value = "Additional Comments"
.Columns("A:O").AutoFit
.UsedRange.Borders.LineStyle = xlContinuous
.Range("1:1").Interior.ColorIndex = 40
.Range("1:1").AutoFilter
End With
Next cell
''-------------save jmReport sheet-------------------
jmReport.Worksheets("Sheet1").Delete
filename = "\JM Report " & Format(Date, "dd.mm.yyyy") & ".xlsx"
jmReport.SaveAs ThisWorkbook.path & filename
End If
''---------------The end----------------------
gmrWB.Close
Set pt = Nothing
Set pc = Nothing
Set rng = Nothing
Set pvtsht = Nothing
Set dataSheet = Nothing
Set gmrWB = Nothing
Set detailExtractSheet = Nothing
Set macroSheet = Nothing
Set jmReport = Nothing

Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "Cyber Report Tool"

Exit Sub
errorHandlerOpen:
MsgBox Err.Description, vbCritical, "Cyber Report Tool"
Exit Sub
errorHandlerFilter:
MsgBox "Inception Year not found in 'DetailExtract' worksheet.", vbCritical, "Cyber Report Tool"
End Sub

