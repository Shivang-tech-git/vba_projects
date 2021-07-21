Attribute VB_Name = "Reconciliation"
Sub ELDvsELTO()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Error Handling Routine-----------------------------
''on error resume next
''On Error GoTo errorHandler:
''on error goto 0
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer, cell As Range
Dim macroSheet As Worksheet, lowLevelComparison As Workbook, TempWB As Workbook
Dim originalWS As String, currentWS As String, fso As Object
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
''''---------------Move downloaded files to thisworkbook location-------------------
A = Format(DateAdd("M", -1, Now), "MMMM YYYY")
Set fso = CreateObject("scripting.filesystemobject")
ChDir ("C:\Users\" & VBA.Environ("Username") & "\Downloads")
filePath = Dir(A & "*(0030)*")
    If fso.fileexists(filePath) Then
    fso.copyFile filePath, ThisWorkbook.Path & "\" & filePath
    Else
    MsgBox "0030 file not found in Downloads folder.", vbCritical, "ELTO Tool"
    Exit Sub
    End If

filePath = Dir(A & "*(0056)*")
    If fso.fileexists(filePath) Then
    fso.copyFile filePath, ThisWorkbook.Path & "\" & filePath
    Else
    MsgBox "0056 file not found in Downloads folder.", vbCritical, "ELTO Tool"
    Exit Sub
    End If
''------------------Create and save Low Level Comparison workbook-----------------
Workbooks.Add
Set lowLevelComparison = ActiveWorkbook
lowLevelComparison.Worksheets.Add , , 3
lowLevelComparison.Worksheets("Sheet4").Name = "Genius XLICSE data"
lowLevelComparison.Worksheets("Sheet3").Name = "Genius XLCICL data"
lowLevelComparison.Worksheets("Sheet2").Name = "Filtered ELTO 0030 Data"
lowLevelComparison.Worksheets("Sheet1").Name = "Filtered ELTO 0056 Data"
lowLevelComparison.SaveAs ThisWorkbook.Path & "\Low Level Comparison - " & Format(DateAdd("M", -1, Date), "MMMM YYYY") ''#### CHANGE PATH HERE ####
''---------------------------Copy and paste ELD and ELTO data to low level comparison WB----------------

ChDir ThisWorkbook.Path ''#### CHANGE PATH HERE ####
Workbooks.Open ThisWorkbook.Path & "\" & Dir("*0030*") ''#### CHANGE PATH HERE ####
Set TempWB = ActiveWorkbook
TempWB.Worksheets(TempWB.Worksheets.Count).Copy lowLevelComparison.Worksheets("Filtered ELTO 0030 Data")
lowLevelComparison.Worksheets(TempWB.Worksheets(TempWB.Worksheets.Count).Name).Name = "Original ELTO 0030 Data"
TempWB.Close
Set TempWB = Nothing

Workbooks.Open ThisWorkbook.Path & "\" & Dir("*0056*")  ''#### CHANGE PATH HERE ####
Set TempWB = ActiveWorkbook
TempWB.Worksheets(TempWB.Worksheets.Count).Copy lowLevelComparison.Worksheets("Filtered ELTO 0056 Data")
lowLevelComparison.Worksheets(TempWB.Worksheets(TempWB.Worksheets.Count).Name).Name = "Original ELTO 0056 Data"
TempWB.Close
Set TempWB = Nothing

Workbooks.Open ThisWorkbook.Path & "\" & Dir("*EL Section*")  ''#### CHANGE PATH HERE ####
Set TempWB = ActiveWorkbook
TempWB.Worksheets(1).Copy lowLevelComparison.Worksheets("Genius XLICSE data")
lowLevelComparison.Worksheets(TempWB.Worksheets(1).Name).Name = "Original Genius Report"
TempWB.Close
Set TempWB = Nothing

''''--------------------------------------COLUMN REFERENCE-----------------------------------------
With lowLevelComparison

'    With .Worksheets("Original ELTO 0030 Data").Range("1:1")
'    masterPolicyNumber0030 = .Find("Master Policy Number").Column
'    End With
'
'    With .Worksheets("Original ELTO 0056 Data").Range("1:1")
'    masterPolicyNumber0056 = .Find("Master Policy Number").Column
'    End With
'
'    With .Worksheets("Genius XLICSE data").Range("1:1")
'
'    End With
''----------------------------------0030 and 0056 Data formatting and applying formulas-------------------------
currentWS = "Filtered ELTO 0030 Data"
originalWS = "Original ELTO 0030 Data"
For x = 1 To 2
.Worksheets(originalWS).Range("1:1").AutoFilter
.Worksheets(originalWS).UsedRange.SpecialCells(xlCellTypeVisible).Copy
''-----------------Remove master policy numbers starting from or containing "PC".------------
    With .Worksheets(currentWS)
    .Range("C1").PasteSpecial xlPasteAll
    .Range("A1").Value = "Is Policy on Genius Data?"
    .Range("B1").Value = "Comments"
    .Range("1:1").AutoFilter field:=10, Criteria1:="PC*", Operator:=xlOr, Criteria2:="*PC*"
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
''--------------------Remove cover start date before 1 April 2011 for 0030------
    Select Case currentWS
    Case "Filtered ELTO 0030 Data"
    .Range("1:1").AutoFilter field:=13, Criteria1:="<=4/01/2011"
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("A2:A" & .Cells(Rows.Count, 10).End(xlUp).Row).Formula = "=IF(MATCH(J:J,'Genius XLICSE data'!D:D,0),""Yes"")"
''----------------------Remove master policy number not starting from UK for 0056----
''----------------------Remove cover start date before 1 January 2019 for 0056-------
    Case "Filtered ELTO 0056 Data"
    .Range("1:1").AutoFilter field:=10, Criteria1:="<>UK*", Operator:=xlFilterValues
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("1:1").AutoFilter field:=13, Criteria1:="<=1/01/2019"
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("A2:A" & .Cells(Rows.Count, 10).End(xlUp).Row).Formula = "=IF(MATCH(J:J,'Genius XLCICL data'!D:D,0),""Yes"")"
    End Select
''---------------------Trim master policy number column----------------
    lastrow = .Cells(Rows.Count, 10).End(xlUp).Row
    For Each cell In .Range(.Cells(2, 10), .Cells(lastrow, 10))
    cell.Value = Application.WorksheetFunction.Trim(cell.Value)
    Next cell
    
''--------------------Remove duplicate master policy number rows---------
    .Range(.Cells(2, 1), .Cells(lastrow, .Cells(1, Columns.Count).End(xlToLeft).Column)).RemoveDuplicates Columns:=10, Header:=xlYes
''-----------------------Comments column---------------------------------------
.Range("1:1").AutoFilter field:=1, Criteria1:="Yes", Operator:=xlFilterValues
lastrow = .Cells(Rows.Count, 10).End(xlUp).Row
.Range(.Cells(2, 2), .Cells(lastrow, 2)).SpecialCells(xlCellTypeVisible).Value = "Policy is on the Genius"
.Range("1:1").AutoFilter
lastrow = .Cells(Rows.Count, 10).End(xlUp).Row
For Each cell In .Range(.Cells(2, 10), .Cells(lastrow, 10)).SpecialCells(xlCellTypeVisible)
        Select Case Trim(.Cells(cell.Row, 32).Value)
        Case "123/BE12345", "123/AB12345", "N/A - EXEMPT"
        .Cells(cell.Row, 2) = "Binder Policy"
        End Select
Next cell
'.Range("1:1").AutoFilter field:=32, Criteria1:=Array("123/BE12345", "123/AB12345", "N/A - EXEMPT"), Operator:=xlFilterValues
'.Range(.Cells(2, 2), .Cells(lastrow, 2)).SpecialCells(xlCellTypeVisible).Value = "Binder Policy"
'.Range("1:1").AutoFilter
.Range("1:1").AutoFilter field:=1, Criteria1:="<>Yes", Operator:=xlFilterValues
''--------------------Hide columns---------------------------------------
.Range("C:I,K:L,O:S,V:AN").EntireColumn.Hidden = True
.Cells.SpecialCells(xlCellTypeVisible).EntireColumn.AutoFit
    End With
originalWS = "Original ELTO 0056 Data"
currentWS = "Filtered ELTO 0056 Data"
Next x
''==========================================================================
''-----------------------Formatting the genius report-----------------------
''==========================================================================
currentWS = "Genius XLICSE data"
originalWS = "Original Genius Report"
For x = 1 To 2
.Worksheets(originalWS).Range("1:1").AutoFilter
.Worksheets(originalWS).UsedRange.SpecialCells(xlCellTypeVisible).Copy
''-----------------Remove master policy numbers starting from or containing "PC".------------
    With .Worksheets(currentWS)
    .Range("C1").PasteSpecial xlPasteAll
    .Range("A1").Value = "Is Policy on ELTO Data?"
    .Range("B1").Value = "Comments"
''--------------------Remove inception date before 1 April 2011 and after 1 Jan 2019 for XLICSE------
    Select Case currentWS
    Case "Genius XLICSE data"
    .Range("1:1").AutoFilter field:=7, Criteria1:="<=4/01/2011"
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("1:1").AutoFilter field:=7, Criteria1:=">=1/01/2019"
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("A2:A" & .Cells(Rows.Count, 4).End(xlUp).Row).Formula = "=IF(MATCH(D:D,'Filtered ELTO 0030 Data'!J:J,0),""Yes"")"
''-------------------Remove PKPolNbr not starting from UK for XLCICL--------------
''-------------------Remove Inception Date before 1 January 2019 for XLCICL-------
''-------------------Remove FKCompanyName not equal to XLCICL-UK------------------
    Case "Genius XLCICL data"
    .Range("1:1").AutoFilter field:=4, Criteria1:="<>UK*", Operator:=xlFilterValues
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("1:1").AutoFilter field:=7, Criteria1:="<=1/01/2019"
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("1:1").AutoFilter field:=25, Criteria1:="<>XLCICL-UK", Operator:=xlFilterValues
    .Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    .Range("1:1").AutoFilter
    .Range("A2:A" & .Cells(Rows.Count, 4).End(xlUp).Row).Formula = "=IF(MATCH(D:D,'Filtered ELTO 0056 Data'!J:J,0),""Yes"")"
    End Select
''--------------------Remove duplicate PKPolNbr rows---------
    lastrow = .Cells(Rows.Count, 4).End(xlUp).Row
    .Range(.Cells(2, 1), .Cells(lastrow, .Cells(1, Columns.Count).End(xlToLeft).Column)).RemoveDuplicates Columns:=4, Header:=xlYes
''----------------------Comments column------------------------------
    lastrow = .Cells(Rows.Count, 4).End(xlUp).Row
    .Range("1:1").AutoFilter field:=1, Criteria1:="Yes", Operator:=xlFilterValues
    .Range(.Cells(2, 2), .Cells(lastrow, 2)).SpecialCells(xlCellTypeVisible).Value = "Policy is on the ELD"
    .Range("1:1").AutoFilter field:=1, Criteria1:="<>Yes", Operator:=xlFilterValues
    
    For Each cell In .Range(.Cells(2, 4), .Cells(lastrow, 4)).SpecialCells(xlCellTypeVisible)
        Select Case True
        Case InStr(1, .Cells(cell.Row, 6).Value, "XOL", vbTextCompare)
        .Cells(cell.Row, 2) = "XOL"
        
        Case Right(Trim(.Cells(cell.Row, 6).Value), 2) = "IE"
        .Cells(cell.Row, 2) = "Irish policies"
        
        Case InStr(1, .Cells(cell.Row, 4).Value, "MM", vbTextCompare) > 0
        .Cells(cell.Row, 2) = "Dummy regional numbers"
        
        Case Trim(.Cells(cell.Row, 9).Value) = 0
        .Cells(cell.Row, 2) = "One day policies"
        
        Case Trim(.Cells(cell.Row, 5).Value) = "Private Client"
        .Cells(cell.Row, 2) = "Private client"
        End Select
    Next cell
    .Columns.AutoFit
    End With
currentWS = "Genius XLCICL data"
Next x
End With
''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
lowLevelComparison.Activate
MsgBox "Done !", vbInformation, "ELTO LOW COMPARISON TOOL"
Set macroSheet = Nothing
'Exit Sub
'errorHandler:
'Resume Next
End Sub























