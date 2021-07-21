Attribute VB_Name = "ActiveX"
Sub recon()
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''-------------Variable Declarations---------------------------------------
Dim macroSheet As Worksheet, usmSheet As Worksheet, bdxSheet As Worksheet
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set usmSheet = ThisWorkbook.Worksheets("USM")
Set bdxSheet = ThisWorkbook.Worksheets("BDX")
Set exchangeRatesSheet = ThisWorkbook.Worksheets("Exchange Rates")
''------------Remove filters---------------------------------------
usmSheet.Range("1:1").AutoFilter
bdxSheet.Range("1:1").AutoFilter
''---------Replace B1966 with B1526-----------------------------------
With usmSheet
lastRowUSM = .Cells(Rows.Count, 10).End(xlUp).Row
For Z = 2 To lastRowUSM
If UCase(Left(.Cells(Z, 10).Value, 5)) = "B1966" Then
.Cells(Z, 10) = Replace(.Cells(Z, 10), Left(.Cells(Z, 10), 5), "B1526")
End If
''---------------umr and (orig curr) concat--------------------------
.Cells(Z, 12) = .Cells(Z, 10) & " " & .Cells(Z, 7)
Next Z
End With

With bdxSheet
lastRowbdx = .Cells(Rows.Count, 12).End(xlUp).Row
For i = 2 To lastRowbdx
''--------------calculate our share-----------------------------
    YOA = 0.25
    Select Case .Cells(i, 4).Value
    Case "2019"
    YOA = 0.3425
    Case "2020"
    YOA = 0.255
    End Select
'----------LIC commission and BDX premium calculation----------------
    If UCase(.Cells(i, 1).Value) = "B1526CBSPS1900007" Or UCase(.Cells(i, 1).Value) = "B1526CBSPS2000007" Then
    .Cells(i, 18) = .Cells(i, 12) * 0.0275
    .Cells(i, 19) = (.Cells(i, 13) - .Cells(i, 18) - .Cells(i, 14) + .Cells(i, 15)) * YOA
    Else
    .Cells(i, 19) = (.Cells(i, 13) - .Cells(i, 14) + .Cells(i, 15)) * YOA
    End If
''----------------(cert ref) and Original Currency concat--------------
.Cells(i, 20) = .Cells(i, 5) & " " & .Cells(i, 11)
Next i
.Range("R1:S" & lastRowbdx).NumberFormat = "0.00"
End With
MsgBox "Complete.", vbInformation, "ACT RECONCILIATION TOOL"
Set usmSheet = Nothing
Set bdxSheet = Nothing
Set reconSheet = Nothing
Set macroSheet = Nothing
Set exchangeRatesSheet = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub ActiveXDataObject()
On Error Resume Next
Application.DisplayAlerts = False
Application.ScreenUpdating = False
''-------------Variable Declarations---------------------------------------
Dim oconn As ADODB.Connection
Dim oRS As ADODB.Recordset, oRS2 As ADODB.Recordset, lineslip As Worksheet, paidOrWritten As Worksheet
Dim sheetName As String, RowNum As Integer, ColNum As Integer, headerName As String
Dim macroSheet As Worksheet, usmSheet As Worksheet, bdxSheet As Worksheet, reconSheet As Worksheet
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set usmSheet = ThisWorkbook.Worksheets("USM")
Set bdxSheet = ThisWorkbook.Worksheets("BDX")
Set reconSheet = ThisWorkbook.Worksheets("Reconciliation")
Set lineslip = ThisWorkbook.Worksheets("Lineslip Policy")
Set paidOrWritten = ThisWorkbook.Worksheets("Paid not Written")
Set oconn = New ADODB.Connection
Set oRS = New ADODB.Recordset
Set oRS2 = New ADODB.Recordset
''--------------Clear contents-----------------------------------------------------
bdxSheet.Range("1:1").AutoFilter
bdxSheet.Rows("2:" & Rows.Count).ClearContents
usmSheet.Range("1:1").AutoFilter
usmSheet.Rows("2:" & Rows.Count).ClearContents
reconSheet.Range("1:1").AutoFilter
reconSheet.Rows("2:" & Rows.Count).ClearContents
lineslip.Range("1:1").AutoFilter
lineslip.Rows("2:" & Rows.Count).ClearContents
paidOrWritten.Range("1:1").AutoFilter
paidOrWritten.Rows("2:" & Rows.Count).ClearContents
''---------------Calculate number of total files in bdx folder----------------------
cntr = 0
ChDir ThisWorkbook.Path & "\BDX\"
wbName = Dir("*.xls*")
Do While wbName <> ""
cntr = cntr + 1
wbName = Dir
Loop
''---------------Loop through every file in BDX folder-----------------------------
cntr2 = 0
ChDir ThisWorkbook.Path & "\BDX\"
wbName = Dir("*.xls*")
Do While wbName <> ""
cntr2 = cntr2 + 1
''--------------Create ADO connection to source sheet -----------------------------
    With oconn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties").Value = "Excel 12.0"
        .Open wbName
        sheetName = .OpenSchema(adSchemaTables).Fields("table_name") & "A2:HQ"
        sheetName = Replace(sheetName, "'", "")
    End With
    ''------Run query on each BDX column of source sheet and get data to BDX sheet ----------
    LastRow = bdxSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
    ColNum = 1
    For RowNum = 4 To macroSheet.Cells(Rows.Count, 4).End(xlUp).Row
    headerName = macroSheet.Range("D" & RowNum).Value
    ''---------Special treatment for row 16------------------------
    If RowNum = 16 Then
     Sql = "SELECT * FROM [" & sheetName & "]"
     oRS2.Open Sql, oconn
    For i = 1 To oRS2.Fields.Count
         If InStr(1, oRS2.Fields(i).Name, headerName, vbTextCompare) > 0 Then
         headerName = oRS2.Fields(i).Name
         Exit For
         End If
     Next i
     End If
     ''----------------------------------------------------------------
        Sql = "SELECT [" & headerName & "] FROM [" & sheetName & "]"
        oRS.Open Sql, oconn
        bdxSheet.Cells(LastRow, ColNum).CopyFromRecordset oRS

        oRS.Close
        ColNum = ColNum + 1
    Next RowNum
    oconn.Close
''----------------Progress bar---------------------------------
Progress.Show (vbModeless)
Progress.Text.Caption = "BDX Files Processed " & cntr2 & "  of " & cntr
Progress.Bar.Width = (cntr2 / cntr) * 100
Progress.Repaint
wbName = Dir
Loop
''---------------Calculate number of total files in usm folder----------------------
usmcntr = 0
ChDir ThisWorkbook.Path & "\USM\"
wbName = Dir("*.xls*")
Do While wbName <> ""
usmcntr = usmcntr + 1
wbName = Dir
Loop
''-----------Do same for USM sheet-----------------------------------------
usmcntr2 = 0
ChDir ThisWorkbook.Path & "\USM\"
wbName = Dir("*.xls*")
Do While wbName <> ""
usmcntr2 = usmcntr2 + 1
''--------------Create ADO connection to source sheet -----------------------------
    With oconn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties").Value = "Excel 12.0"
        .Open wbName
        sheetName = .OpenSchema(adSchemaTables).Fields("table_name")
    End With
    ''------Run query on each USM column of source sheet and get data to USM sheet ----------
    LastRow = usmSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
    ColNum = 1
    For RowNum = 4 To macroSheet.Cells(Rows.Count, 3).End(xlUp).Row
        Sql = "SELECT [" & macroSheet.Range("C" & RowNum).Value & "] FROM [" & sheetName & "]"
        oRS.Open Sql, oconn
        usmSheet.Cells(LastRow, ColNum).CopyFromRecordset oRS
        oRS.Close
        ColNum = ColNum + 1
    Next RowNum
    oconn.Close
''----------------Progress bar---------------------------------
Progress.Text.Caption = "USM Files Processed " & usmcntr2 & "  of " & usmcntr
Progress.Bar.Width = (usmcntr2 / usmcntr) * 100
Progress.Repaint
wbName = Dir
Loop
''-------------The End------------------------------------------------------
Progress.Hide
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "ACT RECONCILIATION TOOL"
Set oconn = Nothing
Set oRS = Nothing
Set reconSheet = Nothing
Set macroSheet = Nothing
Set usmSheet = Nothing
Set bdxSheet = Nothing
Set oRS2 = Nothing
End Sub

'Sub find()
'Dim bdx As Workbook, rowcount As Double
'Dim sheet As Worksheet
'rowcount = 0
'filename = Dir("C:\Users\X134391\Desktop\ACT Recon Tool\BDX\*.xlsx")
'Do While filename <> ""
'Workbooks.Open "C:\Users\X134391\Desktop\ACT Recon Tool\BDX\" & filename
'Set bdx = ActiveWorkbook
'Set sheet = bdx.Worksheets(bdx.Worksheets.Count)
'rowcount = rowcount + sheet.Cells(Rows.Count, 1).End(xlUp).Row
'
''Set celladdress = sheet.UsedRange.find("B0823AE1656011")
''
''If celladdress Is Nothing Then
'bdx.Close
''End If
'filename = Dir()
'Loop
'MsgBox rowcount
'End Sub
''109068

