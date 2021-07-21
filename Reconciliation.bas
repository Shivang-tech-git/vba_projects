Attribute VB_Name = "Reconciliation"
Sub reconPart2()
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''-------------Variable Declarations---------------------------------------
Dim macroSheet As Worksheet, usmSheet As Worksheet, bdxSheet As Worksheet, reconSheet As Worksheet, paidSheet As Worksheet
Dim paidOrWrittenWB As Workbook, paidOrWrittenWS As Worksheet
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set usmSheet = ThisWorkbook.Worksheets("USM")
Set bdxSheet = ThisWorkbook.Worksheets("BDX")
Set reconSheet = ThisWorkbook.Worksheets("Reconciliation")
Set paidSheet = ThisWorkbook.Worksheets("Paid not Written")
''--------------Clear contents-----------------------------------------------------
paidSheet.Cells.ClearContents
reconSheet.Range("1:1").AutoFilter
reconSheet.Rows("2:" & Rows.Count).ClearContents
usmSheet.Range("1:1").AutoFilter
bdxSheet.Range("1:1").AutoFilter
''------------Check if (Usm = Certificate Ref)--------------------------------
lastRowUSM = usmSheet.Cells(Rows.Count, 10).End(xlUp).Row
lastRowbdx = bdxSheet.Cells(Rows.Count, 5).End(xlUp).Row
lastRowMacro = macroSheet.Cells(Rows.Count, 5).End(xlUp).Row
With reconSheet
''--------------copy concat column from usmSheet and bdx column to reconSheet and remove duplicates---------------
usmSheet.Range(("L2"), ("L") & lastRowUSM).Copy
.Range("A2").PasteSpecial xlPasteValues
bdxSheet.Range(("T2"), ("T") & lastRowbdx).Copy
.Range("A" & .Cells(Rows.Count, 1).End(xlUp).Row + 1).PasteSpecial xlPasteValues
.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes

For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
''-----------------VLookUp for concat column from reconSheet to (bdx sheet) and (usm sheet)---------------------------------
.Cells(i, 2) = Application.WorksheetFunction.VLookup(.Cells(i, 1), usmSheet.Range("L:L"), 1, 0)
.Cells(i, 3) = Application.WorksheetFunction.VLookup(.Cells(i, 1), bdxSheet.Range("T:T"), 1, 0)

    If .Cells(i, 2) = .Cells(i, 3) Then
    .Cells(i, 4) = "True"
    Else
    .Cells(i, 4) = "False"
    End If
''-----------------Index and match for (currency) column from usm sheet------------------------------
.Cells(i, 5) = Application.WorksheetFunction.Index(usmSheet.Range(("G2"), ("G") & lastRowUSM), _
               Application.WorksheetFunction.Match(.Cells(i, 1), usmSheet.Range(("L2"), ("L") & lastRowUSM), 0))
''-----------------sum if for (calculated original net amount) column from usm sheet----------------
.Cells(i, 6) = Application.WorksheetFunction.SumIf(usmSheet.Range("L:L"), .Cells(i, 1), usmSheet.Range("H:H"))
.Cells(i, 6) = Format(.Cells(i, 6), "0.00")
''-----------------Index and match for (Signing Date) column from usm sheet--------------------------
.Cells(i, 7) = Application.WorksheetFunction.Index(usmSheet.Range(("C2"), ("C") & lastRowUSM), _
               Application.WorksheetFunction.Match(.Cells(i, 1), usmSheet.Range(("L2"), ("L") & lastRowUSM), 0))
.Cells(i, 7) = Format(.Cells(i, 7), "mm/dd/yyyy")
''----------------Index and match for (Original Currency) column from bdx sheet-----------------
.Cells(i, 8) = Application.WorksheetFunction.Index(bdxSheet.Range(("K2"), ("K") & lastRowbdx), _
               Application.WorksheetFunction.Match(.Cells(i, 1), bdxSheet.Range(("T2"), ("T") & lastRowbdx), 0))
''-----------------sum if for (BDX Premium) column from bdx sheet--------------
.Cells(i, 9) = Application.WorksheetFunction.SumIf(bdxSheet.Range("T:T"), .Cells(i, 1), bdxSheet.Range("S:S"))
.Cells(i, 9) = Format(.Cells(i, 9), "0.00")
''----------------Index and match for (Year of Account) column from bdx sheet-----------------
.Cells(i, 10) = Application.WorksheetFunction.Index(bdxSheet.Range(("D2"), ("D") & lastRowbdx), _
               Application.WorksheetFunction.Match(.Cells(i, 1), bdxSheet.Range(("T2"), ("T") & lastRowbdx), 0))
''----------------------------In case of match------------------------------------------------
        If .Cells(i, 5) = .Cells(i, 8) Then
        .Cells(i, 11) = "True"
        .Cells(i, 12) = .Cells(i, 9) - .Cells(i, 6)
''------------------------------Get values in USD--------------------------------------------
                If .Cells(i, 8).Value <> "USD" Then
                    On Error GoTo exRateHandler
                    exRate = Application.WorksheetFunction.Index(macroSheet.Range("F4:F" & lastRowMacro), _
                             Application.WorksheetFunction.Match(reconSheet.Cells(i, 8), macroSheet.Range("E4:E" & lastRowMacro), 0))
                    On Error Resume Next
                    .Cells(i, 13) = .Cells(i, 12) * exRate
                Else
                    .Cells(i, 13) = .Cells(i, 12)
                End If
        Else
        .Cells(i, 11) = "False"
''-----------------------------In case of mismatch and no value in bdx----------------------
            If .Cells(i, 9) = "" Or .Cells(i, 9) = 0 Then
            .Cells(i, 12) = .Cells(i, 6)
                If .Cells(i, 5).Value <> "USD" Then
                    On Error GoTo exRateHandler
                    exRate = Application.WorksheetFunction.Index(macroSheet.Range("F4:F" & lastRowMacro), _
                             Application.WorksheetFunction.Match(reconSheet.Cells(i, 5), macroSheet.Range("E4:E" & lastRowMacro), 0))
                    On Error Resume Next
                    .Cells(i, 13) = .Cells(i, 12) * exRate
                Else
                .Cells(i, 13) = .Cells(i, 12)
                End If
''------------------------------In case of mismatch and no value in usm----------------------
            Else
            .Cells(i, 12) = .Cells(i, 9)
            If .Cells(i, 8).Value <> "USD" Then
                    On Error GoTo exRateHandler
                    exRate = Application.WorksheetFunction.Index(macroSheet.Range("F4:F" & lastRowMacro), _
                             Application.WorksheetFunction.Match(reconSheet.Cells(i, 8), macroSheet.Range("E4:E" & lastRowMacro), 0))
                    On Error Resume Next
                    .Cells(i, 13) = .Cells(i, 12) * exRate
                Else
                    .Cells(i, 13) = .Cells(i, 12)
                End If
            End If
        End If
.Cells(i, 12) = Format(.Cells(i, 12), "0.00")
.Cells(i, 13) = Format(.Cells(i, 13).Value, "0.00")
.Cells(i, 1) = Split(.Cells(i, 1).Value, " ")(0)
''---------------------Evaluate Balance column-----------------------------------------
If .Cells(i, 13) <> "" Then
    If .Cells(i, 13) <= 5000 Then
    .Cells(i, 14) = "Nominal balance"
    GoTo exitIF
    ElseIf .Cells(i, 13) <= 50000 Then
    .Cells(i, 14) = "Small balance"
    ElseIf .Cells(i, 13) > 50000 Then
    .Cells(i, 14) = "Top balance"
exitIF:
    End If
End If

Next i
End With
''-------------The End------------------------------------------------------
MsgBox "Done !", vbInformation, "ACT RECONCILIATION TOOL"
Set usmSheet = Nothing
Set bdxSheet = Nothing
Set reconSheet = Nothing
Set macroSheet = Nothing
Set paidSheet = Nothing
Set paidOrWrittenWB = Nothing
Set paidOrWrittenWS = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Exit Sub
exRateHandler:
exRate = 0
Resume Next
End Sub

Sub paidOrWritten()
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim reconSheet As Worksheet, paidSheet As Worksheet, macroSheet As Worksheet
Dim paidOrWrittenWB As Workbook, paidOrWrittenWS As Worksheet

Set reconSheet = ThisWorkbook.Worksheets("Reconciliation")
Set paidSheet = ThisWorkbook.Worksheets("Paid not Written")
Set macroSheet = ThisWorkbook.Worksheets("Macro")
''-------------------------Paid Sheet------------------------
wbPath = Dir(ThisWorkbook.Path & "\Files\" & macroSheet.Cells(5, 8).Value & "*")
On Error GoTo errorHandler:
Workbooks.Open ThisWorkbook.Path & "\Files\" & wbPath
Set paidOrWrittenWB = ActiveWorkbook
Set paidOrWrittenWS = paidOrWrittenWB.Worksheets(macroSheet.Cells(6, 8).Value)
On Error Resume Next

''--------------Clear contents-----------------------------------------------------
With paidSheet
.Cells.ClearContents
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Interior.Color = xlNone
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Border.LineStyle = xlNone
End With
''----------------------------------------------------------------------------------
reconSheet.Range("1:1").AutoFilter Field:=3, Criteria1:="<>""", Operator:=xlFilterValues
reconSheet.Range("1:1").AutoFilter Field:=2, Criteria1:="", Operator:=xlFilterValues
reconSheet.Cells.SpecialCells(xlCellTypeVisible).Copy
paidSheet.Cells(1, 1).PasteSpecial xlPasteValues
paidSheet.Columns("A:A").Insert Shift:=xlToRight
paidSheet.Cells(1, 1) = "Paid or Written"

For Z = 2 To paidSheet.Cells(Rows.Count, 2).End(xlUp).Row
paidSheet.Cells(Z, 1) = Application.WorksheetFunction.Index(paidOrWrittenWS.Range("C:C"), _
                        Application.WorksheetFunction.Match(paidSheet.Cells(Z, 2), paidOrWrittenWS.Range("A:A"), 0))
Next Z
With paidSheet
.Columns("A:O").AutoFit
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Interior.ColorIndex = 38
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Border.LineStyle = xlContinuous
End With
MsgBox "Done !", vbInformation, "ACT RECONCILIATION TOOL"
''-------------The End------------------------------------------------------
Set macroSheet = Nothing
Set reconSheet = Nothing
Set paidSheet = Nothing
Set paidOrWrittenWB = Nothing
Set paidOrWrittenWS = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical + vbOKCancel, "ACT RECONCILIATION TOOL"
End Sub

Sub lineslipPolicy()
'On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim reconSheet As Worksheet, lineslip As Worksheet, macroSheet As Worksheet, masterpolicySheet As Worksheet
Set reconSheet = ThisWorkbook.Worksheets("Reconciliation")
Set lineslip = ThisWorkbook.Worksheets("Lineslip policy")
Set macroSheet = ThisWorkbook.Worksheets("Macro")
lastRowMacro = macroSheet.Cells(Rows.Count, 5).End(xlUp).Row
''-----------open lineslip workbook----------------------
wbPath = Dir(ThisWorkbook.Path & "\Files\" & macroSheet.Cells(7, 8).Value & "*")
On Error GoTo errorHandler:
Workbooks.Open ThisWorkbook.Path & "\Files\" & wbPath
Set masterpolicySheet = ActiveWorkbook.Worksheets(macroSheet.Cells(8, 8).Value)
On Error Resume Next

''--------------Clear contents-----------------------------------------------------
With lineslip
.Cells.ClearContents
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Interior.Color = xlNone
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Border.LineStyle = xlNone
''----------------------------------------------------------------------------------
reconSheet.Range("1:1").AutoFilter Field:=3, Criteria1:="<>""", Operator:=xlFilterValues
reconSheet.Range("1:1").AutoFilter Field:=2, Criteria1:="", Operator:=xlFilterValues
reconSheet.Cells.SpecialCells(xlCellTypeVisible).Copy
.Cells(1, 1).PasteSpecial xlPasteValues
reconSheet.Range("1:1").AutoFilter
''-----------------------------------------------------------------------------------
.Cells(1, 15).Value = "Master Policy number"
For Z = 2 To lineslip.Cells(Rows.Count, 1).End(xlUp).Row
.Cells(Z, 15) = Application.WorksheetFunction.Index(masterpolicySheet.Range("B:B"), _
                        Application.WorksheetFunction.Match(.Cells(Z, 1), masterpolicySheet.Range("A:A"), 0))
If .Cells(Z, 15) <> "" Then
.Cells(Z, 15) = "B1526" & .Cells(Z, 15)
.Cells(Z, 2) = Application.WorksheetFunction.Index(reconSheet.Range("B:B"), _
                        Application.WorksheetFunction.Match(.Cells(Z, 15), reconSheet.Range("A:A"), 0))
.Cells(Z, 5) = Application.WorksheetFunction.Index(reconSheet.Range("E:E"), _
                        Application.WorksheetFunction.Match(.Cells(Z, 15), reconSheet.Range("A:A"), 0))
.Cells(Z, 6) = Application.WorksheetFunction.Index(reconSheet.Range("F:F"), _
                        Application.WorksheetFunction.Match(.Cells(Z, 15), reconSheet.Range("A:A"), 0))
.Cells(Z, 7) = Format(Application.WorksheetFunction.Index(reconSheet.Range("G:G"), _
                        Application.WorksheetFunction.Match(.Cells(Z, 15), reconSheet.Range("A:A"), 0)), "mm/dd/yyyy")
If .Cells(Z, 5) = .Cells(Z, 8) Then
        .Cells(Z, 11) = "True"
        .Cells(Z, 12) = .Cells(Z, 9) - .Cells(Z, 6)
        .Cells(Z, 12) = Format(.Cells(Z, 12), "0.00")
''------------------------------Get values in USD--------------------------------------------
                If .Cells(Z, 8).Value <> "USD" Then
                On Error GoTo exRateHandler
                    exRate = Application.WorksheetFunction.Index(macroSheet.Range("F4:F" & lastRowMacro), _
                             Application.WorksheetFunction.Match(.Cells(Z, 8), macroSheet.Range("E4:E" & lastRowMacro), 0))
                On Error Resume Next
                    .Cells(Z, 13) = .Cells(Z, 12) * exRate
                Else
                    .Cells(Z, 13) = .Cells(Z, 12)
                End If
                .Cells(Z, 13) = Format(.Cells(Z, 13), "0.00")


''---------------------Evaluate Balance column-----------------------------------------
                
                    If .Cells(Z, 13) <= 5000 Then
                    .Cells(Z, 14) = "Nominal balance"
                    GoTo exitIF
                    ElseIf .Cells(Z, 13) <= 50000 Then
                    .Cells(Z, 14) = "Small balance"
                    ElseIf .Cells(Z, 13) > 50000 Then
                    .Cells(Z, 14) = "Top balance"
exitIF:
                    End If
Else
                .Cells(Z, 11) = "False"
End If
End If
Next Z
.Columns("A:0").AutoFit
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Interior.ColorIndex = 15
.Range(.Cells(1, 1), .Cells(1, .Cells(1, Columns.Count).End(xlToLeft).Column)).Border.LineStyle = xlContinuous
End With
MsgBox "Done !", vbInformation, "ACT RECONCILIATION TOOL"
''-------------The End------------------------------------------------------
Set reconSheet = Nothing
Set lineslip = Nothing
Set macroSheet = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical + vbOKCancel, "ACT RECONCILIATION TOOL"
Exit Sub
exRateHandler:
exRate = 0
Resume Next
End Sub
