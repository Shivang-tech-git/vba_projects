Attribute VB_Name = "report360"
Dim final_report_360 As Workbook, report_360_final As Workbook, report_wb As Workbook, report_ws As Worksheet
Dim user_id As String, report_last_col As Long, report_last_row As Long, final_ws As Worksheet
Sub report360()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer
Dim macroSheet As Worksheet, fso As FileSystemObject
''-------------Set Object References----------------------------------------
Set fso = New FileSystemObject
Set macroSheet = ThisWorkbook.Worksheets("Macro")
''----------------- check if final report 360 file already exist ------
If fso.FileExists(ThisWorkbook.Path & "\Final Report 360.xlsx") Then
MsgBox "File ""Final Report 360.xlsx"" already exist.", vbCritical, "Broker Reminder Tool"
Exit Sub
End If
''----------------- open 360 report ---------------------------------
With Application.FileDialog(msoFileDialogFilePicker)
.Title = "Please select 360 Report."
.Show
filename = .SelectedItems(1)
End With
Set report_wb = Workbooks.Open(filename)
Set report_ws = report_wb.Worksheets(1)
report_last_row = report_ws.Cells(Rows.count, 1).End(xlUp).Row
report_last_col = report_ws.Cells(1, Columns.count).End(xlToLeft).Column + 1
report_ws.Range("A1:A1").AutoFilter
''--------------------- copy to final_report_360 ------
Set report_360_final = Workbooks.Add
Set final_ws = report_360_final.Worksheets(1)
''-------------------------- remove data type other than OS ------
DATA_TYPE_col = report_ws.Range("1:1").Find("DATA_TYPE", lookat:=xlWhole).Column
report_ws.Range("A1:A1").AutoFilter Field:=DATA_TYPE_col, Criteria1:="OS", Operator:=xlFilterValues
report_ws_lr = report_ws.Cells(Rows.count, 1).End(xlUp).Row
''------------ copy to new workbook ---------
report_ws.Range(report_ws.Cells(1, 1), report_ws.Cells(report_ws_lr, report_last_col)).SpecialCells(xlCellTypeVisible).Copy
final_ws.Range("A1").PasteSpecial xlPasteAll
Call france_exception

report_wb.Close False
report_360_final.SaveAs (ThisWorkbook.Path & "\Final Report 360.xlsx")

''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "Broker Reminder Tool"
Set fso = Nothing
Set report_wb = Nothing
Set report_ws = Nothing
Set macroSheet = Nothing
Set report_360_final = Nothing
Exit Sub
errorHandler:
MsgBox "Macro couldn't finish because one of the columns required for filters is either misspelled or missing.", vbCritical, "BROKER REMINDER TOOL"
End Sub

Sub france_exception()
With final_ws
.Activate
DATA_TYPE_col = .Range("1:1").Find("DATA_TYPE", lookat:=xlWhole).Column
minor_acc_col = .Range("1:1").Find("MINOR_ACCOUNT_TYPE", lookat:=xlWhole).Column
AC_ANALYSIS_4_FLAG_col = .Range("1:1").Find("AC_ANALYSIS_4_FLAG", lookat:=xlWhole).Column
ACCOUNT_NAME_col = .Range("1:1").Find("ACCOUNT_NAME", lookat:=xlWhole).Column
due_date_col = .Range("1:1").Find("DUE_DATE", lookat:=xlWhole).Column
entry_date_col = .Range("1:1").Find("ENTRY_DATE", lookat:=xlWhole).Column

.Cells(1, report_last_col) = "NPCC FLAG"
.Cells(1, report_last_col + 1) = "REMINDER TO BE SENT"
.Cells(1, report_last_col + 2) = "AGEING DAYS"
''--------------------------NPCC and France FILTER --------------------------

.Range("A1:A1").AutoFilter Field:=AC_ANALYSIS_4_FLAG_col, Criteria1:=Array("ANP", "ANPA", "ENP", "ENPA", "ENPG", "LNP", "LNPA"), Operator:=xlFilterValues
report_ws_lr = .Cells(Rows.count, AC_ANALYSIS_4_FLAG_col).End(xlUp).Row
If report_ws_lr > 1 Then
.Range(.Cells(2, report_last_col), .Cells(report_ws_lr, report_last_col)).SpecialCells(xlCellTypeVisible).Value = "NPCC"
.Range(.Cells(2, report_last_col + 1), .Cells(report_ws_lr, report_last_col + 1)).SpecialCells(xlCellTypeVisible).Value = "To be excluded"
End If
.Range("A1:A1").AutoFilter Field:=AC_ANALYSIS_4_FLAG_col
.Range("A1:A1").AutoFilter Field:=report_last_col, Criteria1:="", Operator:=xlFilterValues
report_ws_lr = .Cells(Rows.count, 1).End(xlUp).Row
.Range(.Cells(2, report_last_col), .Cells(report_ws_lr, report_last_col)).SpecialCells(xlCellTypeVisible).Value = "FRANCE"
''--------------------------- filter for future due date ------------------------------------------------------------------------
.Range("A1:A1").AutoFilter
latest_due_date = DateAdd("d", -30, Date)
.Range("A1:A1").AutoFilter Field:=due_date_col, Criteria1:=">" & latest_due_date
report_ws_lr = .Cells(Rows.count, 1).End(xlUp).Row
.Range(.Cells(2, report_last_col + 1), .Cells(report_ws_lr, report_last_col + 1)).SpecialCells(xlCellTypeVisible).Value = "To be excluded"
''--------------------------- FILTER FOR MINOR ACCOUNT TYPE, FIN, DUE DATE - ENTRY DATE AND BEFORE 2019 YEAR ---------------------
.Range("A1:A1").AutoFilter
.Range("A1:A1").AutoFilter Field:=minor_acc_col, Criteria1:=Array("BKR", "ARI", "DIR", "SPC", "LDR"), Operator:=xlFilterValues
.Range("A1:A1").AutoFilter Field:=ACCOUNT_NAME_col, Criteria1:="<>FIN *", Operator:=xlFilterValues
.Range("A1:A1").AutoFilter Field:=report_last_col + 1, Criteria1:="", Operator:=xlFilterValues
''---------------------------------------------------------------------------------------------------------------------------------
report_ws_lr = .Cells(Rows.count, 1).End(xlUp).Row
Set visible_rng = .Range(.Cells(2, due_date_col), .Cells(report_ws_lr, due_date_col))

For Each cell In visible_rng.SpecialCells(xlCellTypeVisible)

'.Cells(cell.Row, due_date_col) = .Cells(cell.Row, due_date_col)
'.Cells(cell.Row, entry_date_col) = .Cells(cell.Row, entry_date_col)

    If DateDiff("d", .Cells(cell.Row, entry_date_col), cell.Value) <= 29 Then
        .Cells(cell.Row, report_last_col + 1) = "Due date error"
    ElseIf Year(.Cells(cell.Row, entry_date_col)) <= 2019 And Year(cell.Value) <= 2019 Then
         .Cells(cell.Row, report_last_col + 1) = "Backlog"
    Else
        .Cells(cell.Row, report_last_col + 1) = "To be reminded"
        .Cells(cell.Row, report_last_col + 2) = DateDiff("d", cell.Value, Date)
    End If
    
Next cell

.Range("A1:A1").AutoFilter
.Range(.Cells(2, report_last_col + 2), .Cells(report_ws_lr, report_last_col + 2)).NumberFormat = "General"
.Range("1:1").Interior.ColorIndex = 6
.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
.Columns("A:FA").AutoFit
.Columns(report_last_col).Interior.ColorIndex = 19
.Columns(report_last_col + 1).Interior.ColorIndex = 19
.Columns(report_last_col + 1).Interior.ColorIndex = 19
End With
End Sub

Sub broker_wise_template()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim dict As New Scripting.Dictionary
Dim data(), r As Long, acc_codes_per_email() As String, count As Integer
Set dict = CreateObject("Scripting.Dictionary")
Dim fso As FileSystemObject, example_template_eng As Workbook, broker_visible_range As Range
Dim broker_list As Workbook, example_template_frn As Workbook, broker_ws As Worksheet
Dim all_codes_per_email As String
''------------------ create folder and open report 360 file --------------------
Set fso = New FileSystemObject
If fso.FolderExists(ThisWorkbook.Path & "\All Broker statements") Then
MsgBox "Folder ""All Broker statements"" already exist.", vbCritical, "Broker Reminder Tool"
Exit Sub
Else
Set statement_folder = fso.CreateFolder(ThisWorkbook.Path & "\All Broker statements")
fso.CreateFolder (statement_folder.Path & "\BKR")
fso.CreateFolder (statement_folder.Path & "\ARI")
fso.CreateFolder (statement_folder.Path & "\DIR")
fso.CreateFolder (statement_folder.Path & "\SPC")
fso.CreateFolder (statement_folder.Path & "\LDR")
End If
sFound = Dir(ThisWorkbook.Path & "\*Final Report 360*")    'the first one found
If sFound <> "" Then
Set final_report_360 = Workbooks.Open(ThisWorkbook.Path & "\" & sFound)
End If
''------------------- open broker contact list ----------------------------------
Set broker_list = Workbooks.Open(ThisWorkbook.Path & "\Broker Contact List.xlsb")
Set broker_ws = broker_list.Worksheets(1)
broker_ws.AutoFilterMode = False
email_col = 5
lang_col = 4
code_col = 1

''--------------- get unique Emails -----------------
data = broker_ws.Columns(email_col).Value
For r = 1 To UBound(data)
    dict(data(r, 1)) = Empty
Next
data = WorksheetFunction.Transpose(dict.Keys())
''--------------------- example statement -----------------------
Set example_template_eng = Workbooks.Open(ThisWorkbook.Path & "\Example statement - English.xlsx")
Set example_template_frn = Workbooks.Open(ThisWorkbook.Path & "\Example statement - French.xlsx")
''------------------ open final report 360 workbook -------------------------------
With final_report_360.Worksheets(1)
.AutoFilterMode = False
last_col_360 = .Cells(1, Columns.count).End(xlToLeft).Column

On Error GoTo errorHandler:
acc_code_360 = .Range("1:1").Find("ACCOUNT_CODE", lookat:=xlWhole).Column
acc_name = .Range("1:1").Find("ACCOUNT_NAME", lookat:=xlWhole).Column
lob = .Range("1:1").Find("LINE_OF_BUSINESS", lookat:=xlWhole).Column
inc_date = .Range("1:1").Find("INCEPTION_DATE", lookat:=xlWhole).Column
due_date = .Range("1:1").Find("DUE_DATE", lookat:=xlWhole).Column
entry_date = .Range("1:1").Find("ENTRY_DATE", lookat:=xlWhole).Column
trans_type = .Range("1:1").Find("TRANS_TYPE", lookat:=xlWhole).Column
inst_nbr = .Range("1:1").Find("INST_NBR", lookat:=xlWhole).Column
ins_name = .Range("1:1").Find("INSURED_NAME", lookat:=xlWhole).Column
their_ref = .Range("1:1").Find("THEIR_REF", lookat:=xlWhole).Column
lead_flag = .Range("1:1").Find("LEAD_FLAG", lookat:=xlWhole).Column
policy = .Range("1:1").Find("POLICY", lookat:=xlWhole).Column
'policy_title = .Range("1:1").Find("POLICY_TITLE", lookat:=xlWhole).Column
'audit_no = .Range("1:1").Find("AUDIT_NO", lookat:=xlWhole).Column
ccy_360 = .Range("1:1").Find("ORIG_CCY", lookat:=xlWhole).Column
premium = .Range("1:1").Find("GRS_PRM_ORG", lookat:=xlWhole).Column
tax = .Range("1:1").Find("GRS_TAX_ORG", lookat:=xlWhole).Column
brokerage = .Range("1:1").Find("GRS_COM_ORG", lookat:=xlWhole).Column
amt_rem = .Range("1:1").Find("AMOUNT_REMAINING_ORIG", lookat:=xlWhole).Column
sum_all = .Range("1:1").Find("SUM_ALLOCATED_ORIG", lookat:=xlWhole).Column
'acc_curr = .Range("1:1").Find("ACCOUNT_CURRENCY", lookat:=xlWhole).Column
amt_rem_base = .Range("1:1").Find("AMOUNT_REMAINING_BASE", lookat:=xlWhole).Column
narrative_col = .Range("1:1").Find("NARRATIVE", lookat:=xlWhole).Column
reminder_status_col = .Range("1:1").Find("REMINDER TO BE SENT", lookat:=xlWhole).Column
On Error GoTo 0
''------------ filter final report 360 for to be "reminded" --------------
.Range("1:1").AutoFilter Field:=reminder_status_col, Criteria1:="To be reminded", Operator:=xlFilterValues
''------------ columns for broker list ----------
acc_code_list = 1
broker_type_col = 2
On Error Resume Next
brokerTypes = Array("SPC", "BKR", "ARI", "DIR", "LDR")
sheetNames = Array("FIRST REMINDER", "SECOND REMINDER", "THIRD REMINDER")
'' ----------------------- first loop for minor account type and second loop for unique email address -----------

For Each broker_type In brokerTypes

broker_ws.Range("1:1").AutoFilter Field:=broker_type_col, Criteria1:=broker_type, Operator:=xlFilterValues

For Each email_Add In data

If Not email_Add = Empty And Not email_Add = "Adresse mail" Then
broker_ws.Range("1:1").AutoFilter Field:=email_col, Criteria1:=email_Add, Operator:=xlFilterValues

broker_type_last_row = broker_ws.Cells(Rows.count, 1).End(xlUp).Row
If broker_type_last_row > 1 Then
Set broker_visible_range = broker_ws.Range(broker_ws.Cells(2, code_col), broker_ws.Cells(broker_type_last_row, code_col))
temp_lang = broker_ws.Cells(broker_type_last_row, lang_col).Value
count = 0
Erase acc_codes_per_email
all_codes_per_email = ""
''------------------- get all visible account codes in an array -----------------
For Each bkr_cell In broker_visible_range.SpecialCells(xlCellTypeVisible)
    ReDim Preserve acc_codes_per_email(count)
    acc_codes_per_email(count) = bkr_cell
    all_codes_per_email = all_codes_per_email & " " & bkr_cell
    count = count + 1
Next bkr_cell

eur_1_row = 15
usd_1_row = 25
other_1_row = 35
eur_2_row = 15
usd_2_row = 25
other_2_row = 35
eur_3_row = 15
usd_3_row = 25
other_3_row = 35

''------------- filter 360 sheet for account code from broker list --------------
    .Range("1:1").AutoFilter Field:=acc_code_360, Criteria1:=acc_codes_per_email, Operator:=xlFilterValues
    last_row_360 = .Cells(Rows.count, 1).End(xlUp).Row
    If last_row_360 > 1 Then
    ''---------------- open example statement workbook --------------
    If temp_lang = "French" Then
        Set example_template = example_template_frn
    ElseIf temp_lang = "English" Then
        Set example_template = example_template_eng
    End If
    example_template.SaveCopyAs (statement_folder.Path & "\" & broker_type & "\" & Trim(all_codes_per_email) & ".xlsx")
    Set example_template = Workbooks.Open(statement_folder.Path & "\" & broker_type & "\" & Trim(all_codes_per_email) & ".xlsx")
    ''----------------- update month and year in template -------------
    
    Dim sheet As Worksheet
    Dim idx As Integer
    idx = 0
    For Each sheet In example_template.Worksheets
    sheet.Range("A2").Value = Replace(Replace(sheet.Range("A2").Value, "MONTH", Format(Date, "mmmm")), "YEAR", Year(Date))
    sheet.Name = sheetNames(idx)
    idx = idx + 1
    Next sheet
        ''-------------- for each row for filtered account code in broker list -------------
        Set visible_range = .Range(.Cells(2, acc_code_360), .Cells(last_row_360, acc_code_360))
        For Each cell In visible_range.SpecialCells(xlCellTypeVisible)
        
        ''------------ check due days and currency ---------------
        narrative = .Cells(cell.Row, narrative_col).Value
        orig_curr = .Cells(cell.Row, ccy_360).Value
        ''----------- copy cell values from final 360 workbook ---------------
        row_per_account = Array(.Cells(cell.Row, acc_code_360), _
                                .Cells(cell.Row, acc_name), _
                                .Cells(cell.Row, lob), _
                                .Cells(cell.Row, inc_date), _
                                .Cells(cell.Row, due_date), _
                                .Cells(cell.Row, entry_date), _
                                .Cells(cell.Row, trans_type), _
                                .Cells(cell.Row, inst_nbr), _
                                .Cells(cell.Row, ins_name), _
                                .Cells(cell.Row, their_ref), _
                                .Cells(cell.Row, lead_flag), _
                                .Cells(cell.Row, policy), _
                                .Cells(cell.Row, ccy_360), _
                                .Cells(cell.Row, premium) + .Cells(cell.Row, tax), _
                                .Cells(cell.Row, brokerage), _
                                .Cells(cell.Row, premium) + .Cells(cell.Row, tax) + .Cells(cell.Row, brokerage), _
                                .Cells(cell.Row, sum_all), _
                                .Cells(cell.Row, premium) + .Cells(cell.Row, tax) + .Cells(cell.Row, brokerage) - .Cells(cell.Row, sum_all))
                                
''---- insert row in first reminder worksheet ------
    If example_template.Worksheets("FIRST REMINDER").Cells(eur_1_row + 2, 1) <> "" Then
        example_template.Worksheets("FIRST REMINDER").Rows(eur_1_row & ":" & eur_1_row + 1).Insert
        usd_1_row = usd_1_row + 2
        other_1_row = other_1_row + 2
    End If
    If example_template.Worksheets("FIRST REMINDER").Cells(usd_1_row + 2, 1) <> "" Then
        example_template.Worksheets("FIRST REMINDER").Rows(usd_1_row & ":" & usd_1_row + 1).Insert
        other_1_row = other_1_row + 2
    End If
    If example_template.Worksheets("FIRST REMINDER").Cells(other_1_row + 2, 1) <> "" Then
        example_template.Worksheets("FIRST REMINDER").Rows(other_1_row & ":" & other_1_row + 1).Insert
    End If
    
''---- insert row in second reminder worksheet ------

    If example_template.Worksheets("SECOND REMINDER").Cells(eur_2_row + 2, 1) <> "" Then
        example_template.Worksheets("SECOND REMINDER").Rows(eur_2_row & ":" & eur_2_row + 1).Insert
        usd_2_row = usd_2_row + 2
        other_2_row = other_2_row + 2
    End If
    If example_template.Worksheets("SECOND REMINDER").Cells(usd_2_row + 2, 1) <> "" Then
        example_template.Worksheets("SECOND REMINDER").Rows(usd_2_row & ":" & usd_2_row + 1).Insert
        other_2_row = other_2_row + 2
    End If
    If example_template.Worksheets("SECOND REMINDER").Cells(other_2_row + 2, 1) <> "" Then
        example_template.Worksheets("SECOND REMINDER").Rows(other_2_row & ":" & other_2_row + 1).Insert
    End If
    
''---- insert row in third reminder worksheet ------

    If example_template.Worksheets("THIRD REMINDER").Cells(eur_3_row + 2, 1) <> "" Then
        example_template.Worksheets("THIRD REMINDER").Rows(eur_3_row & ":" & eur_3_row + 1).Insert
        usd_3_row = usd_3_row + 2
        other_3_row = other_3_row + 2
    End If
    If example_template.Worksheets("THIRD REMINDER").Cells(usd_3_row + 2, 1) <> "" Then
        example_template.Worksheets("THIRD REMINDER").Rows(usd_3_row & ":" & usd_3_row + 1).Insert
        other_3_row = other_3_row + 2
    End If
    If example_template.Worksheets("THIRD REMINDER").Cells(other_3_row + 2, 1) <> "" Then
        example_template.Worksheets("THIRD REMINDER").Rows(other_3_row & ":" & other_3_row + 1).Insert
    End If

''==============================================================
''------------ paste to example statement workbook -------------
''==============================================================

''-------------- SECOND REMINDER -------------------
        If InStr(1, narrative, "1st reminder", vbTextCompare) > 0 Then
        
            If InStr(1, orig_curr, "EUR", vbTextCompare) Then
            
                example_template.Worksheets("SECOND REMINDER").Range("A" & eur_2_row & ":R" & eur_2_row) = row_per_account
                eur_2_row = eur_2_row + 1
                
            ElseIf InStr(1, orig_curr, "USD", vbTextCompare) Then
            
                example_template.Worksheets("SECOND REMINDER").Range("A" & usd_2_row & ":R" & usd_2_row) = row_per_account
                usd_2_row = usd_2_row + 1
                
            Else
            
                example_template.Worksheets("SECOND REMINDER").Range("A" & other_2_row & ":R" & other_2_row) = row_per_account
                example_template.Worksheets("SECOND REMINDER").Cells(other_2_row, 19) = .Cells(cell.Row, amt_rem_base)
                other_2_row = other_2_row + 1
                
            End If
''-------------- THIRD REMINDER -------------------
        ElseIf InStr(1, narrative, "2nd reminder", vbTextCompare) > 0 Then
        
            If InStr(1, orig_curr, "EUR", vbTextCompare) Then
            
                example_template.Worksheets("THIRD REMINDER").Range("A" & eur_3_row & ":R" & eur_3_row) = row_per_account
                eur_3_row = eur_3_row + 1
                
            ElseIf InStr(1, orig_curr, "USD", vbTextCompare) Then
            
                example_template.Worksheets("THIRD REMINDER").Range("A" & usd_3_row & ":R" & usd_3_row) = row_per_account
                usd_3_row = usd_3_row + 1
                
            Else
            
                example_template.Worksheets("THIRD REMINDER").Range("A" & other_3_row & ":R" & other_3_row) = row_per_account
                example_template.Worksheets("THIRD REMINDER").Cells(other_3_row, 19) = .Cells(cell.Row, amt_rem_base)
                other_3_row = other_3_row + 1
                
            End If
           
''-------------- FIRST REMINDER -------------------
        Else
            If InStr(1, orig_curr, "EUR", vbTextCompare) Then
        
                example_template.Worksheets("FIRST REMINDER").Range("A" & eur_1_row & ":R" & eur_1_row) = row_per_account
                eur_1_row = eur_1_row + 1
                
            ElseIf InStr(1, orig_curr, "USD", vbTextCompare) Then
            
                example_template.Worksheets("FIRST REMINDER").Range("A" & usd_1_row & ":R" & usd_1_row) = row_per_account
                usd_1_row = usd_1_row + 1
                
            Else
            
                example_template.Worksheets("FIRST REMINDER").Range("A" & other_1_row & ":R" & other_1_row) = row_per_account
                example_template.Worksheets("FIRST REMINDER").Cells(other_1_row, 19) = .Cells(cell.Row, amt_rem_base)
                other_1_row = other_1_row + 1
                
            End If
        End If
 ''----------------- go to next row -------------
        row_per_account = ""
        Reminder = ""
        narrative = ""
        orig_curr = ""
        Next cell

''---------------------- Formatting ----------------------

    example_template.Worksheets("FIRST REMINDER").Range("A" & eur_1_row).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("FIRST REMINDER").Range("A" & usd_1_row).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("FIRST REMINDER").Range("A" & other_1_row - 1).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("SECOND REMINDER").Range("A" & eur_2_row).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("SECOND REMINDER").Range("A" & usd_2_row).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("SECOND REMINDER").Range("A" & other_2_row - 1).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("THIRD REMINDER").Range("A" & eur_3_row).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("THIRD REMINDER").Range("A" & usd_3_row).CurrentRegion.Borders.LineStyle = xlContinuous
    example_template.Worksheets("THIRD REMINDER").Range("A" & other_3_row - 1).CurrentRegion.Borders.LineStyle = xlContinuous
    On Error GoTo 0
    
    row_arr = Array(eur_1_row, usd_1_row, other_1_row, eur_2_row, usd_2_row, other_2_row, eur_3_row, usd_3_row, other_3_row)
    inc = 0
    
    For Each sheet In example_template.Worksheets
    
    ''------------------------------ if premium is zero then highlight row in orange ----------------------
    last_row_prem = sheet.Cells(Rows.count, 14).End(xlUp).Row
'        For prem_row = 15 To last_row_prem
'            If sheet.Cells(prem_row, 14).Value <> "" And sheet.Cells(prem_row, 14).Value < 0.5 And sheet.Cells(prem_row, 14).Value > -0.5 Then
'            sheet.Range("A" & prem_row & ":R" & prem_row).Interior.ColorIndex = 45
'            End If
'        Next prem_row
    ''------------------------------- sum total for columns N to R ----------------------------------------
        If temp_lang = "English" Then
            sheet.Cells(row_arr(0 + inc), 13) = "Total Outstanding Premium(s)"
            sheet.Cells(row_arr(1 + inc), 13) = "Total Outstanding Premium(s)"
            sheet.Cells(row_arr(2 + inc), 13) = "Total Outstanding Premium(s)"
        ElseIf temp_lang = "French" Then
            sheet.Cells(row_arr(0 + inc), 13) = "Total Prime(s) impayé(s)"
            sheet.Cells(row_arr(1 + inc), 13) = "Total Prime(s) impayé(s)"
            sheet.Cells(row_arr(2 + inc), 13) = "Total Prime(s) impayé(s)"
        End If
        row1 = sheet.Cells(row_arr(0 + inc), 13 + col).End(xlDown).Row
        row2 = sheet.Cells(row_arr(1 + inc), 13 + col).End(xlDown).Row
        For col = 1 To 6
            If row_arr(0 + inc) >= 16 Then
               sheet.Cells(row_arr(0 + inc), 13 + col) = WorksheetFunction.Sum(sheet.Range(sheet.Cells(15, 13 + col), sheet.Cells(row_arr(0 + inc) - 1, 13 + col)))
            End If
            
            If row_arr(1 + inc) >= 26 Then
               sheet.Cells(row_arr(1 + inc), 13 + col) = WorksheetFunction.Sum(sheet.Range(sheet.Cells(row1, 13 + col), sheet.Cells(row_arr(1 + inc) - 1, 13 + col)))
            End If
            
            If row_arr(2 + inc) >= 36 Then
               sheet.Cells(row_arr(2 + inc), 13 + col) = WorksheetFunction.Sum(sheet.Range(sheet.Cells(row2, 13 + col), sheet.Cells(row_arr(2 + inc) - 1, 13 + col)))
            End If
        Next col
    inc = inc + 3
    Next sheet
    If temp_lang = "French" Then
    example_template.Worksheets("FIRST REMINDER").Name = "1er Relance"
    example_template.Worksheets("SECOND REMINDER").Name = "2eme Relance"
    example_template.Worksheets("THIRD REMINDER").Name = "3eme Relance"
    End If
    
''-------------- save broker reminder and go to next broker ---------
        example_template.Save
        example_template.Close
        temp_lang = ""
    Set example_template = Nothing
    Set visible_range = Nothing

    End If

End If '''' if no rows found after minor type and email_add filter
End If '''' if email_add is empty or address mail
Next email_Add
Next broker_type '''' minor account type
End With
''-------------The End------------------------------------------------------
Workbooks("Example statement - English.xlsx").Close False
Workbooks("Example statement - French.xlsx").Close False
final_report_360.Close False
'third_rem_BKR.Range("1:1").Interior.ColorIndex = 19
'third_rem_BKR.Range("D:F").NumberFormat = "DD/MM/YYYY"
'third_rem_BKR.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
'third_rem_BKR.Columns("A:T").AutoFit
'third_rem_ari.Range("1:1").Interior.ColorIndex = 19
'third_rem_ari.Range("D:F").NumberFormat = "DD/MM/YYYY"
'third_rem_ari.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
'third_rem_ari.Columns("A:T").AutoFit
broker_list.Close False
On Error GoTo 0
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Set fso = Nothing
Set example_template_eng = Nothing
Set example_template_frn = Nothing
Set statement_folder = Nothing
Set final_report_360 = Nothing
Set example_template = Nothing
Set broker_list = Nothing
MsgBox "Done !", vbInformation, "Broker Reminder Tool"
Exit Sub
errorHandler:
MsgBox "Macro couldn't finish because one of the columns required for filters is either misspelled or missing.", vbCritical, "BROKER REMINDER TOOL"
End Sub

Sub Outlook()
Dim OutApp As Object, all_broker_folder As Folder, file As file
Dim OutMail As Object, fso As FileSystemObject, broker_list As Workbook
Dim dict As New Scripting.Dictionary
Dim data(), r As Long
Set dict = CreateObject("Scripting.Dictionary")
Set OutApp = CreateObject("Outlook.Application")
Set fso = New FileSystemObject

''----------------------- all columns -------------------------------
brokerTypes = Array("SPC", "BKR", "ARI", "DIR", "LDR")
account_code_col = 1
acc_name_col = 3
lang_col = 4
receiver_col = 5
broker_type_col = 2
all_codes_per_email = ""

If fso.FolderExists(ThisWorkbook.Path & "\All Broker statements") Then
''------------------- open broker contact list ----------------------------------
Set broker_list = Workbooks.Open(ThisWorkbook.Path & "\Broker Contact List.xlsb")
Set broker_ws = broker_list.Worksheets(1)
broker_ws.AutoFilterMode = False
With broker_ws
''--------------- get unique Emails -----------------
data = .Columns(receiver_col).Value
For r = 1 To UBound(data)
    dict(data(r, 1)) = Empty
Next
data = WorksheetFunction.Transpose(dict.Keys())
''---------------------------------------------------
For Each broker_type_str In brokerTypes

.Range("1:1").AutoFilter Field:=broker_type_col, Criteria1:=broker_type_str, Operator:=xlFilterValues

For Each email_Add In data

If Not email_Add = Empty And Not email_Add = "Adresse mail" Then

.Range("1:1").AutoFilter Field:=receiver_col, Criteria1:=email_Add, Operator:=xlFilterValues
last_row = .Cells(Rows.count, receiver_col).End(xlUp).Row
If last_row > 1 Then
''----------------------------------------------------
    Set list_range = .Range(.Cells(2, receiver_col), .Cells(last_row, receiver_col))
''-------------------------- get all details --------------------
    name_of_acc = .Cells(last_row, acc_name_col)
    email_lang = .Cells(last_row, lang_col)
''-------------- check attachments -------------
    For Each cell In list_range.SpecialCells(xlCellTypeVisible)
    all_codes_per_email = all_codes_per_email & " " & .Cells(cell.Row, account_code_col)
    Next cell
    
    If fso.FileExists(ThisWorkbook.Path & "\All Broker statements\" & broker_type_str & "\" & Trim(all_codes_per_email) & ".xlsx") Then
            att_exists = True
    End If
        
''------------- English language subject and body ---------------------
    If att_exists = True Then
    
    If InStr(1, email_lang, "English", vbTextCompare) Then
        Set OutMail = OutApp.CreateItemFromTemplate(ThisWorkbook.Path & "\English.oft")
        OutMail.Subject = "AXA XL outstanding premium overview - " & Format(Date, "mmmm yyyy") & " - " & name_of_acc
''------------- Spanish Language subject and body --------------------
    ElseIf InStr(1, email_lang, "French", vbTextCompare) Then
        Set OutMail = OutApp.CreateItemFromTemplate(ThisWorkbook.Path & "\French.oft")
        OutMail.Subject = "AXA XL Prime Impayées - " & Format(Date, "mmmm yyyy") & " - " & name_of_acc
    End If
''-------------- add attachments -------------
    OutMail.Attachments.Add ThisWorkbook.Path & "\All Broker statements\" & broker_type_str & "\" & Trim(all_codes_per_email) & ".xlsx"

    ''OutMail.SentOnBehalfOfName = "XLIberiaCreditControl@xlcatlin.com"
    ''OutMail.SendUsingAccount = Session.Accounts.Item(ThisWorkbook.Worksheets("Macro").Range("E6"))
    OutMail.To = email_Add
    OutMail.CC = "RM-FranceCreditControl@axaxl.com;BALGDFAPDFSIN@axaxl.com"
    OutMail.Close (olSave)
    End If
    End If
''---------- reset variables --------
all_codes_per_email = ""
last_row = 0
email_lang = ""
att_exists = False
Set OutMail = Nothing
''-----------------------------------
End If
Next email_Add
Next broker_type_str

End With
broker_list.Close False
MsgBox "Done !", vbInformation, "Broker Reminder Tool"

'---------------------------------
Else
MsgBox "All broker statements folder does not exist !", vbCritical, "Broker Reminder Tool"
End If
Set final_report_360 = Nothing
Set OutApp = Nothing
Set fso = Nothing
End Sub

