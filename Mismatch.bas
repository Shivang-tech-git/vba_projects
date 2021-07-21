Attribute VB_Name = "Mismatch"
Sub Mismatch()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next
Dim detailExtract As Worksheet, macroSheet As Worksheet
Dim vCrit As Variant, rngcrit As Range, titles(1 To 2) As Variant
Dim titleSheetName(1 To 4) As String
Dim titleColumn As Integer
If MsgBox("Run tool.", vbOKCancel + vbQuestion, "Proceed?") = vbOK Then
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set detailExtract = ThisWorkbook.Worksheets("Detail Extract")
With detailExtract
''---------------------Delete All Sheets-----------------------------
For Each Worksheet In ThisWorkbook.Worksheets
    If Worksheet.Name <> "Macro" And Worksheet.Name <> "Detail Extract" Then
    Worksheet.Delete
    End If
Next Worksheet
''--------------------------Calculate last row ----------------------
lastcolumnextract = .Cells(1, Columns.Count).End(xlToLeft).Column + 1
.Cells(1, lastcolumnextract).Value = "LastColumn"
.Range("1:1").AutoFilter
lastRowExtract = detailExtract.Cells(Rows.Count, 1).End(xlUp).Row
'' Unhide all columns and rows to avoid mismatch of fields.
.Columns.EntireColumn.Hidden = False
.Rows.EntireRow.Hidden = False

'filter out blank rows from Quote_Nbr column
.Range("1:1").AutoFilter Field:=Quote_Nbr, Criteria1:="<>"

'filter for array from FKPolStatusDesc column
.Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:=Array("New Awaiting information", "New Complete", _
"RNW Awaiting Information", "RNW Complete", "New Program Master Policy", "Policy Handover to Local Cntry", _
"New Fronting Partner instructd", "Lapsed"), Operator:=xlFilterValues

Call myLoop(Quote_LeadYn, Policy_LeadYn, "Lead Flag inconsistent", "str")
Call myLoop(Quote_Title, PolTitle, "Title", "str")
Call myLoop(Quote_Insured, Policy_OrigInsured, "Insured code mismatch", "str")
Call myLoop(Quote_LocalUnderwriter, Policy_LocalUnderwriterName, "LUW", "str")
Call myLoop(Quote_PAMName, Policy_PAMName, "PAM", "str")
Call myLoop(Quote_InceptionDate, Policy_InceptionDate, "Inception period inconsistent", "dt")
Call myLoop(Quote_ExpiryDate, Policy_ExpiryDate, "Expiry period inconsistent", "dt")
Call myLoop(Quote_BrokerOrAgent, Policy_BrokerOrAgent, "Broker code mismatch", "str")
Call myLoop(Quote_MANName, Policy_MANName, "MAN", "str")
Call myLoop(Quote_TerritoryScopeDesc, Policy_TerritoryScopeDesc, "Territory Scope inconsistent", "str")
Call myLoop(Quote_CountryOfSettlementDesc, Policy_CountryOfSettlementDesc, "Country of settlement mismatch", "str")
Call myLoop(Quote_Portfolio_Prot, Policy_Portfolio_Prot, "Portfolio Prot", "str")
''===========================================================================================
''----------------------------Quote and Policy Title Prefix and sufix------------------------
''===========================================================================================
titles(1) = Array("ACCH", "HLTH", "TRA", "EQUI", "LMOR", "LMKT", "LSPE", "LCCC", "LMTC")
titles(2) = Array("GB", "UK", "IE", "US", "CA", "SG", "AU", "MY", "IN", "AT", "IT", "CH", "NL", "ES", "SE", "DE", "BE", "HK", "SO")
titleSheetName(1) = "Quote Title Prefix"
titleSheetName(2) = "Policy Title Prefix"
titleColumn = Quote_Title
''Outer loop for prefix and sufix
For x = 1 To 2
''Inner loop for Quote and policy column
    For y = 1 To 2
        For I = 2 To lastRowExtract
        If x = 1 Then
        .Cells(I, lastcolumnextract).Value = Left(.Cells(I, titleColumn).Value, 4)
        Else: .Cells(I, lastcolumnextract).Value = Right(.Cells(I, titleColumn).Value, 2)
        End If
        Next I
        
        ''Get unique error values for filter criteria
        .Columns(lastcolumnextract).Copy
        macroSheet.Range("J1").PasteSpecial xlPasteValues
        macroSheet.Range("J:J").RemoveDuplicates 1, xlYes
        
        Set rngTitle = macroSheet.Range("I:I")
        rngTitle.Value = Application.Transpose(titles(x))
        
        For Z = 1 To macroSheet.Cells(Rows.Count, 10).End(xlUp).Row
        If Application.WorksheetFunction.CountIf(macroSheet.Range("I:I"), macroSheet.Cells(Z, 10).Value) > 0 Then
        macroSheet.Cells(Z, 10).Delete
        Z = Z - 1
        End If
        Next Z
        
        lastrowMacro = macroSheet.Cells(Rows.Count, 10).End(xlUp).Row
        Set rngcrit = macroSheet.Range("J1:J" & lastrowMacro).SpecialCells(xlCellTypeVisible)
        vCrit = rngcrit.Value
        
        ''Apply filter for error values of title
        .Range("1:1").AutoFilter Field:=lastcolumnextract, Criteria1:=Application.Transpose(vCrit), Operator:=xlFilterValues
    
    Call myLoop(lastcolumnextract, , titleSheetName(y))
    .Range("1:1").AutoFilter
    .Columns(lastcolumnextract).ClearContents
    .Cells(1, lastcolumnextract).Value = "LastColumn"
    'filter out blank rows from Quote_Nbr column
    .Range("1:1").AutoFilter Field:=Quote_Nbr, Criteria1:="<>"
    
    'filter for array from FKPolStatusDesc column
    .Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:=Array("New Awaiting information", "New Complete", _
    "RNW Awaiting Information", "RNW Complete", "New Program Master Policy", "Policy Handover to Local Cntry", _
    "New Fronting Partner instructd", "Lapsed"), Operator:=xlFilterValues

    
    macroSheet.Range("J:J").ClearContents
    titleColumn = PolTitle
    Next y
titleColumn = Quote_Title
titleSheetName(1) = "Quote Title Suffix"
titleSheetName(2) = "Policy Title Suffix"
Next x
macroSheet.Range("I:J").ClearContents
''===========================================================================================
''UW YOA and inception date
For j = 2 To lastRowExtract
If .Cells(j, Policy_UnderwritingYear).Value <> Year(.Cells(j, Policy_InceptionDate).Value) Then
    .Cells(j, lastcolumnextract).Value = "Error"
End If
Next j
.Range("1:1").AutoFilter Field:=lastcolumnextract, Criteria1:="Error", Operator:=xlFilterValues
Call myLoop(lastcolumnextract, , "UW YOA and inception date")
''===========================================================================================
''Frequency, No.of Installments, Processing Type should be in sync
.Range("1:1").AutoFilter
.Columns(lastcolumnextract).ClearContents
.Cells(1, lastcolumnextract).Value = "LastColumn"
For k = 2 To lastRowExtract
Select Case .Cells(k, FrequencyDesc).Value
Case "ANN annually", "SGL Single"
    If .Cells(k, NbrOfInstallments).Value <> "1" Then
    .Cells(k, lastcolumnextract).Value = "Error"
    End If
Case "MTH Monthly"
    If .Cells(k, NbrOfInstallments).Value = "1" Then
    .Cells(k, lastcolumnextract).Value = "Error"
    End If
Case "HLF Half - yearly"
    If .Cells(k, NbrOfInstallments).Value <> "2" Then
    .Cells(k, lastcolumnextract).Value = "Error"
    End If
Case "QTR Quarterly"
    If .Cells(k, NbrOfInstallments).Value <> "4" Then
    .Cells(k, lastcolumnextract).Value = "Error"
    End If
Case "OTH Other"
    If .Cells(k, NbrOfInstallments).Value = "1" Or .Cells(k, NbrOfInstallments).Value = "2" Or .Cells(k, NbrOfInstallments).Value = "4" Then
    .Cells(k, lastcolumnextract).Value = "Error"
    End If
End Select
Next k
.Range("1:1").AutoFilter Field:=lastcolumnextract, Criteria1:="Error", Operator:=xlFilterValues
Call myLoop(lastcolumnextract, , "No.of Installments Error")
''==================================================================================
''Policy authorize but NI status
.Range("1:1").AutoFilter
.Columns(lastcolumnextract).ClearContents
.Cells(1, lastcolumnextract).Value = "LastColumn"
.Range("1:1").AutoFilter Field:=Policy_Authorization_Date, Criteria1:="<>", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:="New Awaiting information", Operator:=xlFilterValues
Call myLoop(FKPolStatusDesc, , "Policy authorize but NI status")
''===================================================================================
''Quote BD with No Policy
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter Field:=Quote_Status, Criteria1:="Bound", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=PKPolNbr, Criteria1:="", Operator:=xlFilterValues
Call myLoop(Quote_Status, , "Quote BD with No Policy")
''===================================================================================
''Lapsed, but not declined
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:="Lapsed", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Quote_Status, Criteria1:="<>Declined", Operator:=xlFilterValues
Call myLoop(Quote_Status, , "Lapsed, but not declined")
''====================================================================================
''Declined, but not lapsed
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter Field:=Quote_Status, Criteria1:="Declined", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:="<>Lapsed", Operator:=xlFilterValues
Call myLoop(Quote_Status, , "Declined, but not lapsed")
''====================================================================================
''Signed and written not equal
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:=Array("New Awaiting information", "New Complete", _
"RNW Awaiting Information", "RNW Complete"), Operator:=xlFilterValues
Call myLoop(SignedLinePcnt, WrittenLinePcnt, "Signed and written not equal", "str")
''====================================================================================
''Inc Date=Exp Date, not cancel
.Range("1:1").AutoFilter
For D = 2 To lastRowExtract
        If DateDiff("d", .Cells(D, Policy_InceptionDate).Value, .Cells(D, Policy_ExpiryDate).Value) = 0 Then
            If Split(.Cells(D, FKPolStatusDesc).Value, " ")(0) <> "Cncl" And Split(.Cells(D, FKPolStatusDesc).Value, " ")(0) <> "Cancel" Then
            .Cells(D, lastcolumnextract) = "Error"
            End If
        End If
Next D
.Range("1:1").AutoFilter Field:=lastcolumnextract, Criteria1:="Error", Operator:=xlFilterValues
Call myLoop(lastcolumnextract, , "Inc Date=Exp Date, not cancel")
.Range("1:1").AutoFilter
.Columns(lastcolumnextract).ClearContents
''======================================================================================
''Equine Hunters Commission
.Range("1:1").AutoFilter Field:=FKMainLineDesc, Criteria1:="Equine", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_BrokerOrAgent, Criteria1:="HUNT0232", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_DeductionPercent, Criteria1:="<>22.5", Operator:=xlFilterValues
Call myLoop(Policy_DeductionPercent, , "Equine Hunters Commission")

''----------------LTA - Y with No LTA dates--------------------------
.Range("1:1").AutoFilter
.Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:=Array("New Awaiting information", "New Complete", _
"RNW Awaiting Information", "RNW Complete"), Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_LTAYN, Criteria1:="Y", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_LTAInceptionDate, Criteria1:="", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_LTAExpiryDate, Criteria1:="", Operator:=xlFilterValues
Call myLoop(Policy_LTAYN, , "LTA - Y with No LTA dates")

''---------------LTA - N with LTA dates-------------------------
.Range("1:1").AutoFilter Field:=Policy_LTAYN, Criteria1:="N", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_LTAInceptionDate, Criteria1:="<>", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_LTAExpiryDate, Criteria1:="<>", Operator:=xlFilterValues
Call myLoop(Policy_LTAYN, , "LTA - N with LTA dates")

''-------------Policy Period >366 with LTA – N--------------------
''---------------Remove all filters-------------------------------
.Range("1:1").AutoFilter
.Cells(1, lastcolumnextract).Value = "Difference"
.Range("1:1").AutoFilter Field:=FKPolStatusDesc, Criteria1:=Array("New Awaiting information", "New Complete", _
"RNW Awaiting Information", "RNW Complete"), Operator:=xlFilterValues
''---------------Subtract dates-----------------------------------
For I = 2 To lastRowExtract
        If DateDiff("d", .Cells(I, Policy_LTAInceptionDate).Value, .Cells(I, Policy_LTAExpiryDate).Value) > 366 Then
        .Cells(I, lastcolumnextract) = "Greater than 366"
        End If
Next I
.Range("1:1").AutoFilter Field:=lastcolumnextract, Criteria1:="Greater than 366", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_LTAYN, Criteria1:="N", Operator:=xlOr, Criteria2:=""

Call myLoop(lastcolumnextract, , "Policy Period >366 with LTA – N")
''-------------Policy Period <366 with LTA - Y--------------------
.Range("1:1").AutoFilter Field:=lastcolumnextract, Criteria1:="", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=Policy_LTAYN, Criteria1:="Y", Operator:=xlFilterValues
Call myLoop(Policy_LTAYN, , "Policy Period <366 with LTA - Y")

'''-------------Invalid Attachment Point PrimaryLayer E--------------------
'.Range("1:1").AutoFilter Field:=lastcolumnextract
'.Range("1:1").AutoFilter Field:=FKLongTermAgreement_Ind
'.Range("1:1").AutoFilter Field:=Policy_PrimaryLayer, Criteria1:="E", Operator:=xlFilterValues
'.Range("1:1").AutoFilter Field:=AttachmentPoint100PcntOrigCcy, Criteria1:="0", Operator:=xlOr, Criteria2:=""
'Call myLoop(Policy_PrimaryLayer, , "Attachment Point PrimaryLayer E")
'''-------------Invalid Attachment Point PrimaryLayer N--------------------
'.Range("1:1").AutoFilter Field:=Policy_PrimaryLayer, Criteria1:="N", Operator:=xlFilterValues
'.Range("1:1").AutoFilter Field:=AttachmentPoint100PcntOrigCcy, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>"
'Call myLoop(Policy_PrimaryLayer, , "Attachment Point PrimaryLayer N")
'''---------------Invalid Master Status--------------
'.Range("1:1").AutoFilter
'.Columns(lastcolumnextract).EntireColumn.Delete
'.Range("1:1").AutoFilter Field:=PolicyStatus, Criteria1:="NC", Operator:=xlOr, Criteria2:="RC"
'.Range("1:1").AutoFilter Field:=PolicyAuthoriseStatus, Criteria1:="N", Operator:=xlFilterValues
'Call myLoop(PolicyStatus, , "Invalid Master Status")
''------------------------------------------------------------------
.Range("1:1").AutoFilter
.Columns(lastcolumnextract).EntireColumn.Delete
End With
ThisWorkbook.Worksheets("Macro").Activate
MsgBox "Validations completed!", vbInformation, "Information."
Else: MsgBox "Process cancelled!", vbInformation, "Information."

End If
Set macroSheet = Nothing
Set detailExtract = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub myLoop(Optional quoteColumnToCompare, Optional policyColumnToCompare, Optional validation, Optional compareType As String)
On Error Resume Next
Dim wb As Worksheet, macroSheet As Worksheet
Dim lastrow As Long
Dim I As Long
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set wb = ThisWorkbook.Worksheets("Detail Extract")
With wb
'Paste visible rows to new sheet.
ThisWorkbook.Sheets.Add.Name = "Temporary"
.UsedRange.Copy
ThisWorkbook.Worksheets("Temporary").Range("A1").PasteSpecial xlPasteAll
End With
Set wb = ThisWorkbook.Worksheets("Temporary")
With wb
lastrow = .Cells(Rows.Count, quoteColumnToCompare).End(xlUp).Row
lastcolumn = .Cells(1, Columns.Count).End(xlToLeft).Column + 1
''In case of date comparison, convert validation columns to date format.
If compareType = "dt" Then
.Range(.Cells(2, quoteColumnToCompare), .Cells(lastrow, quoteColumnToCompare)).NumberFormat = "m/d/yyyy"
.Range(.Cells(2, policyColumnToCompare), .Cells(lastrow, policyColumnToCompare)).NumberFormat = "m/d/yyyy"
End If
'------Compare quoteColumnToCompare & policyColumnToCompare-----------
If compareType = "str" Then
    For I = 2 To lastrow
    If StrComp(.Cells(I, quoteColumnToCompare).Value, .Cells(I, policyColumnToCompare).Value, vbTextCompare) <> 0 Then
        .Cells(I, lastcolumn) = "Error"
        End If
    Next I
    GoTo removeNoError
ElseIf compareType = "dt" Then
    For I = 2 To lastrow
        If DateDiff("d", .Cells(I, quoteColumnToCompare).Value, .Cells(I, policyColumnToCompare).Value) <> 0 Then
        .Cells(I, lastcolumn) = "Error"
        End If
    Next I
    GoTo removeNoError
End If
resumeProcedure:
lastrow = .Cells(Rows.Count, quoteColumnToCompare).End(xlUp).Row
'Check if error is there(rename sheet) or not(delete sheet).
If lastrow > 1 Then
wb.Name = validation
wb.Columns("A:CZ").AutoFit
wb.Range("A1:CZ1").Interior.ColorIndex = 40
wb.UsedRange.Borders.LineStyle = xlContinuous
wb.Columns(Policy_Inception_Date).EntireColumn.NumberFormat = "mm/dd/yyyy"
wb.Columns(Policy_Expiry_Date).EntireColumn.NumberFormat = "mm/dd/yyyy"
wb.Columns(Policy_LTA_Inception_Date).EntireColumn.NumberFormat = "mm/dd/yyyy"
wb.Columns(Policy_LTA_Expiry_Date).EntireColumn.NumberFormat = "mm/dd/yyyy"
Else
wb.Delete
End If
End With
''------------Count number of mismatches-----------------------
With macroSheet
lastrow2 = .Cells(Rows.Count, 3).End(xlUp).Row
    For j = 4 To lastrow2
        If .Range("C" & j).Value = validation Then
            If lastrow = 1 Then
            .Range("F" & j).Value = "No Mismatch found"
            Else
            .Range("F" & j).Value = lastrow - 1
            Exit For
            End If
        End If
    Next j
End With
Set macroSheet = Nothing
Set wb = Nothing
Exit Sub
removeNoError:
ThisWorkbook.Worksheets("Temporary").Range("1:1").AutoFilter Field:=lastcolumn, Criteria1:="<>Error", Operator:=xlFilterValues
ThisWorkbook.Worksheets("Temporary").Rows("2:" & Rows.Count).SpecialCells(xlCellTypeVisible).Select
Selection.Delete
ThisWorkbook.Worksheets("Temporary").Range("1:1").AutoFilter
GoTo resumeProcedure
End Sub

