Attribute VB_Name = "Module1"
Sub subName()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Error Handling Routine-----------------------------
''on error resume next
''On Error GoTo errorHandler:
''on error goto 0
''-------------column number as constants-------------------------
Const ProgramNumber As Integer = 14
Const PolicyType As Integer = 16
Const LEName As Integer = 5
Const ContractDescription As Integer = 23
Const FacilityFlag As Integer = 19
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer
Dim macroSheet As Worksheet, frameSheet As Worksheet
Dim Facilities As Integer
Dim Declaration As Integer
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set frameSheet = ThisWorkbook.Worksheets("UK Frame business")
''--------------------------------------------------------------------------
With frameSheet
Facilities = .Cells(1, Columns.Count).End(xlToLeft).Column + 1
Declaration = .Cells(1, Columns.Count).End(xlToLeft).Column + 2
.Range("1:1").AutoFilter
.Range(.Cells(2, Facilities), .Cells(Rows.Count, Declaration).End(xlUp)).ClearContents
.Cells(1, Facilities) = "Facilities"
.Cells(1, Declaration) = "Declaration"
''==================================================================================================
''----------------------------Declarations = yes----------------------------------------------------
''==================================================================================================
''-------------------Exclude 0 and /0 from (program number) column----------------------------------
.Range("1:1").AutoFilter field:=ProgramNumber, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>*/0"
lastRow = .Cells(Rows.Count, ProgramNumber).End(xlUp).Row
.Range(.Cells(2, Declaration), .Cells(lastRow, Declaration)).SpecialCells(xlCellTypeVisible) = "Yes"
''------------PolicyType = Delegated authority, lename = Synd 5345----------------------------------
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter field:=PolicyType, Criteria1:="Delegated Authority", Operator:=xlFilterValues
.Range("1:1").AutoFilter field:=LEName, Criteria1:="Synd 5345*", Operator:=xlFilterValues
lastRow = .Cells(Rows.Count, LEName).End(xlUp).Row
.Range(.Cells(2, Declaration), .Cells(lastRow, Declaration)).SpecialCells(xlCellTypeVisible) = "Yes"
''==================================================================================================
''----------------------------Declarations = NO-----------------------------------------------------
''==================================================================================================
''--------------If Contract Description = blank or binding authority then declaration = NO-----------
.Range("1:1").AutoFilter field:=ContractDescription, Criteria1:="Lineslip", Operator:=xlOr, Criteria2:="Consortium"
lastRow = .Cells(Rows.Count, ContractDescription).End(xlUp).Row
.Range(.Cells(2, Declaration), .Cells(lastRow, Declaration)).SpecialCells(xlCellTypeVisible) = "No"
.Range("1:1").AutoFilter field:=LEName
.Range("1:1").AutoFilter field:=ContractDescription, Criteria1:=Array("", "Binding Authority", "-"), Operator:=xlFilterValues
lastRow = .Cells(Rows.Count, PolicyType).End(xlUp).Row
.Range(.Cells(2, Declaration), .Cells(lastRow, Declaration)).SpecialCells(xlCellTypeVisible) = "No"
''---------------------All remaining Declarations are No.--------------------------------------------
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter field:=Declaration, Criteria1:="", Operator:=xlFilterValues
lastRow = .Cells(Rows.Count, PolicyType).End(xlUp).Row
.Range(.Cells(2, Declaration), .Cells(lastRow, Declaration)).SpecialCells(xlCellTypeVisible) = "No"
''==================================================================================================
''----------------------------FACILITIES = YES------------------------------------------------------
''==================================================================================================
''--------------policy type and contract description filter for yes category-----------------------
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter field:=PolicyType, Criteria1:="Delegated Authority", Operator:=xlFilterValues
.Range("1:1").AutoFilter field:=ContractDescription, _
 Criteria1:=Array("Binding Authority", "Consortium", "Lineslip", "Lineslip Treaty"), Operator:=xlFilterValues
lastRow = .Cells(Rows.Count, ContractDescription).End(xlUp).Row
.Range(.Cells(2, Facilities), .Cells(lastRow, Facilities)).SpecialCells(xlCellTypeVisible) = "Y"
''--------------------------------------------------------------------------------------------------
.Range("1:1").AutoFilter field:=ContractDescription, Criteria1:="Treaty", Operator:=xlFilterValues
.Range("1:1").AutoFilter field:=FacilityFlag, Criteria1:="Y", Operator:=xlFilterValues
lastRow = .Cells(Rows.Count, FacilityFlag).End(xlUp).Row
.Range(.Cells(2, Facilities), .Cells(lastRow, Facilities)).SpecialCells(xlCellTypeVisible) = "Y"
''---------------------------------------------------------------------------------------------------
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter field:=PolicyType, Criteria1:="Direct", Operator:=xlFilterValues
.Range("1:1").AutoFilter field:=ContractDescription, Criteria1:="Consortium", Operator:=xlFilterValues
lastRow = .Cells(Rows.Count, ContractDescription).End(xlUp).Row
.Range(.Cells(2, Facilities), .Cells(lastRow, Facilities)).SpecialCells(xlCellTypeVisible) = "Y"
''==================================================================================================
''----------------------------FACILITIES = NO-------------------------------------------------------
''==================================================================================================
''--------------------------All remaining rows are N-------------------------------------------------
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter field:=Facilities, Criteria1:="", Operator:=xlFilterValues
lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
.Range(.Cells(2, Facilities), .Cells(lastRow, Facilities)).SpecialCells(xlCellTypeVisible) = "N"
.Range("1:1").AutoFilter
End With
''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "Facilities Mapping Tool"
Set macroSheet = Nothing
Set frameSheet = Nothing
'Exit Sub
'errorHandler:
'Resume Next
End Sub
