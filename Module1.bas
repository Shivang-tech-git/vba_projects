Attribute VB_Name = "Module1"
Sub pdfExtract()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Error Handling Routine-----------------------------
''on error resume next
''On Error GoTo errorHandler:
''on error goto 0
''-------------Variable Declarations---------------------------------------
Dim fileName As String, I As Integer
Dim macroSheet As Worksheet, datasheet As Worksheet, statement As Worksheet, sheet1 As Worksheet
Dim fso As Object, pagestart As Range, pageend As Range, receipt As Range, cedent As String
Dim rowNum As Integer, totalReceipts As Integer, lastword As Integer, soaid As String, treaty As String
Dim cell As Range, firstLineItem As Range, lastLineItem As Range, lineItems As Range
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set datasheet = ThisWorkbook.Worksheets("DATA")
Set statement = ThisWorkbook.Worksheets("Statement")
Set fso = CreateObject("scripting.filesystemobject")
''----------------Select PDF from file dialog-----------
With Application.FileDialog(msoFileDialogFilePicker)
    If .Show <> 0 Then
    .Title = "Select Euler Hermes PDF."
    fileName = .SelectedItems(1)
    End If
End With
pdfName = Replace(fso.getfile(fileName).Name, ".pdf", "")
''-----------------Clear contents---------------------------------
With datasheet
statement.Range("1:1").AutoFilter
statement.Rows("2:" & Rows.Count).ClearContents
datasheet.Rows("2:" & Rows.Count).ClearContents

rowNum = 2
''---------------------Open PDF and wait 10 seconds----------------------------
On Error GoTo errorHandlerPath
ThisWorkbook.FollowHyperlink fileName
On Error GoTo 0
MsgBox "Please press OK once document's processing status completes.", vbOK + vbExclamation, "EULER HERMES TOOL"
AppActivate (pdfName)
Application.Wait (Now + TimeValue("00:00:05"))
''------------------view --> page display --> enable scrolling-----
SendKeys ("%v")
SendKeys ("p")
SendKeys ("c")
'-------------------transfer data to excel-----------------------
SendKeys ("^a")
SendKeys ("^c")
Application.Wait (Now + TimeValue("00:00:05"))
lastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
.Range("A" & lastRow).PasteSpecial xlPasteAll
''----------Define column numbers---------------
lastrowData = .Cells(Rows.Count, 1).End(xlUp).Row + 1
.Cells(lastrowData, 1).Value = "Run Number"

treatyCol = 1
soaidCol = 2
item = 3
periodYear = 4
basis100 = 5
shareCal100 = 6
cededAmount = 7
cededBasis = 8
sharePer = 9
cededValue = 10
curr = 11
Cntry = 12

rowNum = 2
''----- set receipt to a statement of account ---------
Set pagestart = .UsedRange.Find("Soaid")
Set pageend = .UsedRange.Find("Run Number", pagestart)
Set receipt = .Range(pagestart.Address, pageend.Address)
    While Not pagestart Is Nothing
''----------Set lineItems to all wanted line items in one page in current statement-----------
    Set firstLineItem = receipt.Find("Item Period Basis 100% Share Calc Ceded Amount").Offset(2, 0)
    Set lastLineItem = receipt.Find("per Section/CoB", firstLineItem).Offset(-1, 0)
    Set lineItems = .Range(firstLineItem.Address, lastLineItem.Address)
        While Not firstLineItem Is Nothing
        
            soaid = Split(receipt.Find("soaid").Value, " ")(1)
            treaty = Left(Split(receipt.Find("Treaty Partner").Offset(1, 0).Value, "(")(1), "1") & "Y"
            If Split(receipt.Find("Cedent").Value, " ")(2) = "USX001" Then
            cedent = "CA"
            ElseIf Split(receipt.Find("Cedent").Value, " ")(2) = "US0025" Then
            cedent = "US"
            End If
            
            For Each cell In lineItems
                lastword = UBound(Split(cell.Value, " "))
                If lastword = 3 Then: GoTo increment
                statement.Cells(rowNum, treatyCol) = treaty
                statement.Cells(rowNum, soaidCol) = soaid
                statement.Cells(rowNum, item) = Mid(cell.Value, 1, InStr(1, cell.Value, Split(cell.Value, " ")(lastword - 4), vbTextCompare) - 2)
                statement.Cells(rowNum, periodYear) = Split(cell.Value, " ")(lastword - 4)
                statement.Cells(rowNum, basis100) = Split(cell.Value, " ")(lastword - 3)
                statement.Cells(rowNum, shareCal100) = Split(cell.Value, " ")(lastword - 2)
                statement.Cells(rowNum, cededAmount) = Split(cell.Value, " ")(lastword - 1) & " " & Split(cell.Value, " ")(lastword)
                statement.Cells(rowNum, curr) = Split(cell.Value, " ")(lastword)
                statement.Cells(rowNum, Cntry) = cedent
                rowNum = rowNum + 1
increment:
            Next cell
        Set firstLineItem = receipt.Find("Item Period Basis 100% Share Calc Ceded Amount", lastLineItem).Offset(2, 0)
        If firstLineItem.Row < lastLineItem.Row Then
        GoTo endOfLineItems
        End If
        Set lastLineItem = receipt.Find("per Section/CoB", firstLineItem).Offset(-1, 0)
        Set lineItems = .Range(firstLineItem.Address, lastLineItem.Address)
        Wend
endOfLineItems:
    Set pagestart = .UsedRange.Find("Soaid", pageend)
        If pagestart.Row < pageend.Row Then
        GoTo endOfPDF
        End If
    Set pageend = .UsedRange.Find("Run Number", pagestart)
    Set receipt = .Range(pagestart.Address, pageend.Address)
    soaid = ""
    treaty = ""
    cedent = ""
    Wend
endOfPDF:
End With
''--------------------Apply formulae and refresh Pivot table--------------
With statement
lastrowstatement = .Cells(Rows.Count, 1).End(xlUp).Row
For x = 2 To lastrowstatement
''-----------------Change amount from us format to indian format---------------
        .Cells(x, basis100).Value = Replace(.Cells(x, basis100).Value, ".", "")
        .Cells(x, basis100).Value = Replace(.Cells(x, basis100).Value, ",", ".")
''-------------------If amount is negative, shift the negative sign.-----------------
                If Right(.Cells(x, basis100).Value, 1) = "-" Then
                .Cells(x, cededBasis) = "-" & Left(.Cells(x, basis100).Value, Len(.Cells(x, basis100).Value) - 1)
                Else
                .Cells(x, cededBasis) = .Cells(x, basis100).Value
                End If
''-------------Get share percentage and ceded value------------------------------
        .Cells(x, sharePer) = .Cells(x, shareCal100).Value / 100
        .Cells(x, cededValue) = .Cells(x, cededBasis).Value * .Cells(x, sharePer).Value
Next x
End With
ThisWorkbook.RefreshAll
''---------------DATA sheet formatting----------------------
statement.Columns("A:AI").AutoFit
''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Complete.", vbInformation, "EULER HERMES TOOL"
Set macroSheet = Nothing
Set datasheet = Nothing
Set statement = Nothing
Set fso = Nothing
Exit Sub
errorHandlerPath:
MsgBox "Please select Euler Hermes PDF. " & Err.Description, vbCritical, "EULER HERMES TOOL"
End Sub
















