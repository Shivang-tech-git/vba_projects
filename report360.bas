Attribute VB_Name = "report360"
Dim final_report_360 As Workbook, report_360_final As Workbook, report_wb As Workbook, report_ws As Worksheet
Dim user_id As String
Sub report360()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Error Handling Routine-----------------------------
''on error resume next
''On Error GoTo errorHandler:
''on error goto 0
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
''--------------------- filter for data type and copy to final_report_360 ------
Set report_360_final = Workbooks.Add
report_last_row = report_ws.Cells(Rows.Count, 1).End(xlUp).Row
report_last_col = report_ws.Cells(1, Columns.Count).End(xlToLeft).Column
''-------------- find columns data type and due date --------------
On Error GoTo errorHandler
data_type_col = report_ws.Range("1:1").Find("DATA_TYPE", lookat:=xlWhole).Column
due_date_col = report_ws.Range("1:1").Find("DUE_DATE", lookat:=xlWhole).Column
On Error GoTo 0
''------------ filter 360 report for OS data type and copy to new workbook ---------
report_ws.Range("1:1").AutoFilter Field:=data_type_col, Criteria1:="OS", Operator:=xlFilterValues
Call spain_exception
report_ws.Range(report_ws.Cells(1, 1), report_ws.Cells(report_last_row, report_last_col)).SpecialCells(xlCellTypeVisible).Copy
report_360_final.Worksheets(1).Range("A1").PasteSpecial xlPasteAll
report_wb.Close
''--------------------- apply filter for due date ----------------
With report_360_final.Worksheets(1)
report_last_row = .Cells(Rows.Count, 1).End(xlUp).Row
.Range(.Cells(1, due_date_col + 1), .Cells(1, due_date_col + 2)).EntireColumn.Insert
.Cells(1, due_date_col + 1).Value = "Run Date"
.Cells(1, due_date_col + 2).Value = "Ageing"
.Range(.Cells(2, due_date_col + 1), .Cells(report_last_row, due_date_col + 1)).Formula = Format(Date, "dd/mm/yyyy")

    For Each cell In .Range(.Cells(2, due_date_col + 2), .Cells(report_last_row, due_date_col + 2))
    .Cells(cell.Row, due_date_col) = Format(.Cells(cell.Row, due_date_col), "dd/mm/yyyy")
    cell.Value = DateDiff("d", Format(.Cells(cell.Row, due_date_col), "dd/mm/yyyy"), Date)
    Next cell
    
.Range(.Cells(2, due_date_col + 2), .Cells(report_last_row, due_date_col + 2)).NumberFormat = "General"
End With

report_360_final.SaveAs (ThisWorkbook.Path & "\Final Report 360.xlsx")
report_360_final.Close
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

Sub spain_exception()
With report_ws

minor_acc_col = .Range("1:1").Find("MINOR_ACCOUNT_TYPE", lookat:=xlWhole).Column
policy_col = .Range("1:1").Find("POLICY", lookat:=xlWhole).Column
sett_col = .Range("1:1").Find("COUNTRY_OF_SETTLEMENT", lookat:=xlWhole).Column

.Range("1:1").AutoFilter Field:=minor_acc_col, Criteria1:=Array("BKR", "ARI", "DIR"), Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=policy_col, Criteria1:="<>ESA*", Operator:=xlFilterValues
.Range("1:1").AutoFilter Field:=sett_col, Criteria1:="ES", Operator:=xlFilterValues

End With
End Sub

Sub broker_wise_template()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim fso As FileSystemObject, example_template As Workbook
Dim broker_list As Workbook, third_rem_wb As Workbook, third_rem_ws As Worksheet
''------------------ create folder and open required files --------------------
Set fso = New FileSystemObject
If fso.FolderExists(ThisWorkbook.Path & "\All Broker statements") Then
MsgBox "Folder ""All Broker statements"" already exist.", vbCritical, "Broker Reminder Tool"
Exit Sub
Else
Set statement_folder = fso.CreateFolder(ThisWorkbook.Path & "\All Broker statements")
fso.CreateFolder (statement_folder.Path & "\BKR")
fso.CreateFolder (statement_folder.Path & "\ARI")
End If
sFound = Dir(ThisWorkbook.Path & "\*Final Report 360*")    'the first one found
If sFound <> "" Then
Set final_report_360 = Workbooks.Open(ThisWorkbook.Path & "\" & sFound)
End If

Set broker_list = Workbooks.Open(ThisWorkbook.Path & "\Updated Broker List.xlsx")

broker_count_bkr = broker_list.Worksheets("BKR").Cells(Rows.Count, 1).End(xlUp).Row
broker_count_ari = broker_list.Worksheets("ARI").Cells(Rows.Count, 1).End(xlUp).Row
broker_list.Worksheets("BKR").AutoFilterMode = False
broker_list.Worksheets("ARI").AutoFilterMode = False
''------------------- create workbook for all third reminders ----------------------
Set example_template = Workbooks.Open(ThisWorkbook.Path & "\Example statement.xlsx")
Set third_rem_wb = Workbooks.Add
third_rem_wb.Worksheets("Sheet1").Name = "BKR"
Set third_rem_BKR = third_rem_wb.Worksheets("BKR")
third_rem_wb.Worksheets.Add after:=third_rem_wb.Worksheets(third_rem_wb.Worksheets.Count)
third_rem_wb.Worksheets(third_rem_wb.Worksheets.Count).Name = "ARI"
Set third_rem_ari = third_rem_wb.Worksheets("ARI")
example_template.Worksheets("FIRST REMINDER").Range("A14:U14").Copy
third_rem_BKR.Range("A1:U1").PasteSpecial xlPasteAll
third_rem_BKR.Range("V1") = "PAM_NAME"
third_rem_ari.Range("A1:U1").PasteSpecial xlPasteAll
third_rem_ari.Range("V1") = "PAM_NAME"
''------------------ open final report 360 workbook -------------------------------
With final_report_360.Worksheets(1)
.AutoFilterMode = False
acc_code_list = 3

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
policy = .Range("1:1").Find("POLICY", lookat:=xlWhole).Column
policy_title = .Range("1:1").Find("POLICY_TITLE", lookat:=xlWhole).Column
audit_no = .Range("1:1").Find("AUDIT_NO", lookat:=xlWhole).Column
ccy_360 = .Range("1:1").Find("ORIG_CCY", lookat:=xlWhole).Column
premium = .Range("1:1").Find("GRS_PRM_ORG", lookat:=xlWhole).Column
brokerage = .Range("1:1").Find("GRS_COM_ORG", lookat:=xlWhole).Column
amt_rem = .Range("1:1").Find("AMOUNT_REMAINING_ORIG", lookat:=xlWhole).Column
acc_curr = .Range("1:1").Find("ACCOUNT_CURRENCY", lookat:=xlWhole).Column
amt_rem_acc = .Range("1:1").Find("AMOUNT_REMAINING_ACCOUNTING", lookat:=xlWhole).Column
ageing_360 = .Range("1:1").Find("Ageing", lookat:=xlWhole).Column
narrative_col = .Range("1:1").Find("NARRATIVE", lookat:=xlWhole).Column
uw_name_col = .Range("1:1").Find("PAM_NAME", lookat:=xlWhole).Column
minor_acc_col = .Range("1:1").Find("MINOR_ACCOUNT_TYPE", lookat:=xlWhole).Column
leader_ced_col = .Range("1:1").Find("LEADER_CEDANT_EXTREF", lookat:=xlWhole).Column
On Error GoTo 0
''------------ columns for broker list ----------
yORn = 11
uw_name = 12

last_col_360 = .Cells(1, Columns.Count).End(xlToLeft).Column
broker_type = "BKR"
broker_type_last_row = broker_count_bkr
''---------------- create templates for each broker ----------------------------
On Error Resume Next
For broker_sheet = 1 To 2
For Broker = 2 To broker_type_last_row

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
    .Range("1:1").AutoFilter Field:=acc_code_360, Criteria1:=broker_list.Worksheets(broker_type).Cells(Broker, acc_code_list).Value, Operator:=xlFilterValues
    last_row_360 = .Cells(Rows.Count, 1).End(xlUp).Row
    If last_row_360 > 1 Then
    ''---------------- open example statement workbook --------------
    Set example_template = Workbooks.Open(ThisWorkbook.Path & "\Example statement.xlsx")
    accountCode = broker_list.Worksheets(broker_type).Cells(Broker, acc_code_list).Value
    example_template.SaveCopyAs (statement_folder.Path & "\" & broker_type & "\" & accountCode & ".xlsx")
    Set example_template = Workbooks.Open(statement_folder.Path & "\" & broker_type & "\" & accountCode & ".xlsx")
    ''----------------- update month and year in template -------------
    Dim sheet As Worksheet
    For Each sheet In example_template.Worksheets
    sheet.Range("A2").Value = Replace(Replace(sheet.Range("A2").Value, "MONTH", Format(Date, "mmmm")), "YEAR", Year(Date))
    Next sheet
        ''-------------- for each row for filtered account code in broker list -------------
        Set visible_range = .Range(.Cells(2, acc_code_360), .Cells(last_row_360, acc_code_360)).SpecialCells(xlCellTypeVisible)
        For Each cell In visible_range
        ''------------ check due days and currency ---------------
        narrative = .Cells(cell.Row, narrative_col).Value
        orig_curr = .Cells(cell.Row, ccy_360).Value
        pam_name = .Cells(cell.Row, uw_name_col).Value
        ''----------- copy cell values from final 360 workbook ---------------
        row_per_account = Array(.Cells(cell.Row, acc_code_360), _
                                .Cells(cell.Row, acc_name), _
                                .Cells(cell.Row, minor_acc_col), _
                                .Cells(cell.Row, lob), _
                                .Cells(cell.Row, inc_date), _
                                .Cells(cell.Row, due_date), _
                                .Cells(cell.Row, entry_date), _
                                .Cells(cell.Row, trans_type), _
                                .Cells(cell.Row, inst_nbr), _
                                .Cells(cell.Row, ins_name), _
                                .Cells(cell.Row, their_ref), _
                                .Cells(cell.Row, leader_ced_col), _
                                .Cells(cell.Row, policy), _
                                .Cells(cell.Row, policy_title), _
                                .Cells(cell.Row, audit_no), _
                                .Cells(cell.Row, ccy_360), _
                                .Cells(cell.Row, premium), _
                                .Cells(cell.Row, brokerage), _
                                .Cells(cell.Row, amt_rem), _
                                .Cells(cell.Row, acc_curr), _
                                .Cells(cell.Row, amt_rem_acc))

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
            
                example_template.Worksheets("SECOND REMINDER").Range("A" & eur_2_row & ":U" & eur_2_row) = row_per_account
                eur_2_row = eur_2_row + 1
                
            ElseIf InStr(1, orig_curr, "USD", vbTextCompare) Then
            
                example_template.Worksheets("SECOND REMINDER").Range("A" & usd_2_row & ":U" & usd_2_row) = row_per_account
                usd_2_row = usd_2_row + 1
                
            Else
            
                example_template.Worksheets("SECOND REMINDER").Range("A" & other_2_row & ":U" & other_2_row) = row_per_account
                other_2_row = other_2_row + 1
                
            End If
''-------------- THIRD REMINDER -------------------
        ElseIf InStr(1, narrative, "2nd reminder", vbTextCompare) > 0 Then
        
            If InStr(1, orig_curr, "EUR", vbTextCompare) Then
            
                example_template.Worksheets("THIRD REMINDER").Range("A" & eur_3_row & ":U" & eur_3_row) = row_per_account
                eur_3_row = eur_3_row + 1
                
            ElseIf InStr(1, orig_curr, "USD", vbTextCompare) Then
            
                example_template.Worksheets("THIRD REMINDER").Range("A" & usd_3_row & ":U" & usd_3_row) = row_per_account
                usd_3_row = usd_3_row + 1
                
            Else
            
                example_template.Worksheets("THIRD REMINDER").Range("A" & other_3_row & ":U" & other_3_row) = row_per_account
                other_3_row = other_3_row + 1
                
            End If
            ''--------------- paste to all third reminders workbook ---------------------
            If broker_type = "BKR" Then
                lr_3_rem_bkr = third_rem_BKR.Cells(Rows.Count, 1).End(xlUp).Row + 1
                third_rem_BKR.Range("A" & lr_3_rem_bkr & ":U" & lr_3_rem_bkr) = row_per_account
                third_rem_BKR.Range("V" & lr_3_rem_bkr) = pam_name
            ElseIf broker_type = "ARI" Then
                lr_3_rem_ari = third_rem_ari.Cells(Rows.Count, 1).End(xlUp).Row + 1
                third_rem_ari.Range("A" & lr_3_rem_ari & ":U" & lr_3_rem_ari) = row_per_account
                third_rem_ari.Range("V" & lr_3_rem_ari) = pam_name
            End If
            
        Else
        
''-------------- FIRST REMINDER -------------------
            If InStr(1, orig_curr, "EUR", vbTextCompare) Then
        
                example_template.Worksheets("FIRST REMINDER").Range("A" & eur_1_row & ":U" & eur_1_row) = row_per_account
                eur_1_row = eur_1_row + 1
                
            ElseIf InStr(1, orig_curr, "USD", vbTextCompare) Then
            
                example_template.Worksheets("FIRST REMINDER").Range("A" & usd_1_row & ":U" & usd_1_row) = row_per_account
                usd_1_row = usd_1_row + 1
                
            Else
            
                example_template.Worksheets("FIRST REMINDER").Range("A" & other_1_row & ":U" & other_1_row) = row_per_account
                other_1_row = other_1_row + 1
                
            End If
        End If
 ''----------------- go to next row -------------
        row_per_account = ""
        Reminder = ""
        narrative = ""
        orig_curr = ""
        pam_name = ""
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
    For Each sheet In example_template.Worksheets
    last_row_prem = sheet.Cells(Rows.Count, 17).End(xlUp).Row
        For prem_row = 15 To last_row_prem
            If sheet.Cells(prem_row, 17).Value <> "" And sheet.Cells(prem_row, 17).Value < 0.5 And sheet.Cells(prem_row, 17).Value > -0.5 Then
            sheet.Range("A" & prem_row & ":U" & prem_row).Interior.ColorIndex = 45
            End If
        Next prem_row
    Next sheet
    
    
''-------------- save broker reminder and go to next broker ---------
        example_template.Save
        example_template.Close
    Set example_template = Nothing
    Set visible_range = Nothing

    End If

Next Broker
broker_type = "ARI"
broker_type_last_row = broker_count_ari
Next broker_sheet
End With
''-------------The End------------------------------------------------------
Workbooks("Example statement.xlsx").Close
final_report_360.Close
third_rem_BKR.Range("1:1").Interior.ColorIndex = 19
third_rem_BKR.Range("E:G").NumberFormat = "DD/MM/YYYY"
third_rem_BKR.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
third_rem_BKR.Columns("A:V").AutoFit
third_rem_ari.Range("1:1").Interior.ColorIndex = 19
third_rem_ari.Range("E:G").NumberFormat = "DD/MM/YYYY"
third_rem_ari.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
third_rem_ari.Columns("A:V").AutoFit
broker_list.Close
third_rem_wb.SaveAs statement_folder.Path & "\All Third Reminders.xlsx"
third_rem_wb.Close
On Error GoTo 0
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Set fso = Nothing
Set statement_folder = Nothing
Set final_report_360 = Nothing
Set example_template = Nothing
Set broker_list = Nothing
Set third_rem_wb = Nothing
Set third_rem_BKR = Nothing
Set third_rem_ari = Nothing
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
If fso.FolderExists(ThisWorkbook.Path & "\All Broker statements") Then
Set all_broker_folder = fso.GetFolder(ThisWorkbook.Path & "\All Broker statements\BKR")
Set broker_list = Workbooks.Open(ThisWorkbook.Path & "\Updated Broker List.xlsx")

axaImage = ThisWorkbook.Path & "\axa.bmp"
account_code_col = 3
acc_name_col = 4
lang_col = 9
receiver_col = 10
'english_body = "<p>Dear Sir or Madam,</p>" & _
'                "<p>Attached you will find the monthly overview with open / outstanding premiums.</p>" & _
'                "<p>Looking forward to receive your comments and confirmation of the settlement. Our bank details are included in the sheet.</p>" & _
'                "<p>Please do not hesitate to contact me for any further information.</p>" & _
'                "<p>Thank you and regards,</p>" & _
'                "<img src ='cid:axa.bmp" & "'" _
'                 & "width='50' height='50'><br>" & _
'                "<p>Credit Control<br>" & _
'                "Middle Office<br>" & _
'                "AXA XL, a division of AXA<br>" & _
'                "Plaza de la Lealtad, 4,  2ª Planta, 28014 Madrid - Spain<br>" & _
'                "Email:XLIberiaCreditControl@ xlcatlin.com</p>" & _
'                "<p>axaxl.com<br>" & _
'                "XL Insurance Company SE | Sede: 8 St. Stephen's Green, Dublin 2, Irlanda | Inscrita en el registro de sociedades de Irlanda (Companies Registration Office) | Sociedad nº 641686</p>" & _
'                "<p>Compañía de seguros regulada por la Central Bank of Ireland (www.centralbank.ie)<br>" & _
'                "Sucursal en España (Madrid): Plaza de la Lealtad, 4 - 28014 Madrid | Inscrita en el Registro Mercantil de Madrid, Tomo: 28325, Libro: 0, Folio: 217, Sección: 8, Hoja: M 321046, Inscripción 23 C.I.F. W-0065403-H | Inscrita con la<br>" & _
'                "Dirección General de Seguros y de Fondos de Pensiones bajo la clave E0134</p>"
'
'spanish_body = "<p>Estimados,</p>" & _
'                "<p>Adjunto podrán encontrar el fichero mensual de primas pendientes en nuestro sistema.</p>" & _
'                "<p>Agradecemos vuestros comentarios así como confirmación de pago. Adjunto encontraran también los detalles bancarios.</p>" & _
'                "<p>No duden en contactarnos para mayor información.</p>" & _
'                "<p>Gracias y saludos,</p>" & _
'                "<img src ='cid:axa.bmp" & "'" _
'                 & "width='50' height='50'><br>" & _
'                "<p>Credit Control<br>" & _
'                "Middle Office<br>" & _
'                "AXA XL, a division of AXA<br>" & _
'                "Plaza de la Lealtad, 4,  2ª Planta, 28014 Madrid - Spain<br>" & _
'                "Email: XLIberiaCreditControl@ xlcatlin.com</p>" & _
'                "<p>axaxl.com<br>" & _
'                "XL Insurance Company SE | Sede: 8 St. Stephen's Green, Dublin 2, Irlanda | Inscrita en el registro de sociedades de Irlanda (Companies Registration Office) | Sociedad nº 641686</p>" & _
'                "<p>Compañía de seguros regulada por la Central Bank of Ireland (www.centralbank.ie)<br>" & _
'                "Sucursal en España (Madrid): Plaza de la Lealtad, 4 - 28014 Madrid | Inscrita en el Registro Mercantil de Madrid, Tomo: 28325, Libro: 0, Folio: 217, Sección: 8, Hoja: M 321046, Inscripción 23 C.I.F. W-0065403-H | Inscrita con la <br>" & _
'                "Dirección General de Seguros y de Fondos de Pensiones bajo la clave E0134</p>"


account_code_col = 3
acc_name_col = 4
lang_col = 9
receiver_col = 10
broker_type_str = "BKR"
''--------------- send each broker file --------------
For broker_type = 1 To 2
With broker_list.Worksheets(broker_type_str)
.AutoFilterMode = False
''--------------- get unique Emails -----------------
data = .Columns(receiver_col).Value
For r = 1 To UBound(data)
    dict(data(r, 1)) = Empty
Next
data = WorksheetFunction.Transpose(dict.Keys())
att_exists = False
att_count = 0
''------------- for each email id -------------------
For Each email_id In data
Debug.Print email_id
If Not email_id = Empty And Not email_id = "Emails" Then
''------------- create new email item ---------------
    
    'Set OutMail = OutApp.CreateItem(0)
''------------- filter for broker list --------------

    .Range("1:1").AutoFilter Field:=receiver_col, Criteria1:=email_id, Operator:=xlFilterValues
    last_row = .Cells(Rows.Count, receiver_col).End(xlUp).Row
    If last_row > 1 Then
''----------------------------------------------------
    Set list_range = .Range(.Cells(1, receiver_col), .Cells(last_row, receiver_col))
''-------------------------- get all details --------------------
    name_of_acc = .Cells(last_row, acc_name_col)
    email_lang = .Cells(last_row, lang_col)
''-------------- check attachments -------------
    For Each cell In list_range.SpecialCells(xlCellTypeVisible)
        If fso.FileExists(ThisWorkbook.Path & "\All Broker statements\" & broker_type_str & "\" & .Cells(cell.Row, account_code_col) & ".xlsx") Then
            att_exists = True
            att_count = att_count + 1
            accountCode = .Cells(cell.Row, account_code_col)
        End If
    Next cell
''------------- English language subject and body ---------------------
    If att_exists = True Then
    
    If InStr(1, email_lang, "English", vbTextCompare) Then
    Set OutMail = OutApp.CreateItemFromTemplate(ThisWorkbook.Path & "\English.oft")
        If att_count = 1 Then
        OutMail.Subject = "AXA XL outstanding premium overview - " & Format(Date, "mmmm - yyyy") & " - " & accountCode & " " & name_of_acc
        ElseIf att_count > 1 Then
        OutMail.Subject = "AXA XL outstanding premium overview - " & Format(Date, "mmmm - yyyy") & " " & name_of_acc
        End If
''------------- Spanish Language subject and body --------------------
    ElseIf InStr(1, email_lang, "Spanish", vbTextCompare) Then
    Set OutMail = OutApp.CreateItemFromTemplate(ThisWorkbook.Path & "\Spanish.oft")
        If att_count = 1 Then
        OutMail.Subject = "AXA XL Primas Pendiente - " & Format(Date, "mmmm - yyyy") & " - " & accountCode & " " & name_of_acc
        ElseIf att_count > 1 Then
        OutMail.Subject = "AXA XL Primas Pendiente - " & Format(Date, "mmmm - yyyy") & " " & name_of_acc
        End If
    End If
''-------------- add attachments -------------
    For Each cell In list_range.SpecialCells(xlCellTypeVisible)
            If fso.FileExists(ThisWorkbook.Path & "\All Broker statements\" & broker_type_str & "\" & .Cells(cell.Row, account_code_col) & ".xlsx") Then
                OutMail.Attachments.Add ThisWorkbook.Path & "\All Broker statements\" & broker_type_str & "\" & .Cells(cell.Row, account_code_col) & ".xlsx"
            End If
    Next cell
    ''OutMail.SendUsingAccount = Session.Accounts.Item(ThisWorkbook.Worksheets("Macro").Range("E6"))
    OutMail.SentOnBehalfOfName = "XLIberiaCreditControl@xlcatlin.com"
    OutMail.To = email_id
    OutMail.CC = "XLIberiaCreditControl@xlcatlin.com"
    OutMail.Close (olSave)
    End If
    End If
''---------- reset variables --------
last_row = 0
att_count = 0
email_lang = ""
att_exists = False
Set OutMail = Nothing
''-----------------------------------
End If
Next email_id
End With
Set all_broker_folder = fso.GetFolder(ThisWorkbook.Path & "\All Broker statements\ARI")
broker_type_str = "ARI"
Next broker_type
broker_list.Close
MsgBox "Done !", vbInformation, "Broker Reminder Tool"

'---------------------------------
Else
MsgBox "All broker statements folder does not exist !", vbCritical, "Broker Reminder Tool"
End If
Set final_report_360 = Nothing
Set OutApp = Nothing
Set fso = Nothing
End Sub

Sub ResolveDisplayNameToSMTP(user_name)
    
    Dim oRecip As Outlook.Recipient
    Dim oEU As Outlook.ExchangeUser
    Dim oEDL As Outlook.ExchangeDistributionList
    Dim osess As Outlook.Application
    
    Set osess = New Outlook.Application
    Set oRecip = osess.Session.CreateRecipient(user_name)
    
    user_id = ""
    oRecip.Resolve
    If oRecip.Resolved Then
    Select Case oRecip.AddressEntry.AddressEntryUserType
    Case OlAddressEntryUserType.olExchangeUserAddressEntry
    Set oEU = oRecip.AddressEntry.GetExchangeUser
    If Not (oEU Is Nothing) Then
    user_id = oEU.PrimarySmtpAddress
    End If
    Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry
    Set oEDL = oRecip.AddressEntry.GetExchangeDistributionList
    If Not (oEDL Is Nothing) Then
    user_id = oEDL.PrimarySmtpAddress
    End If
    End Select
    End If
    
End Sub
'''--------------- send each broker file --------------
'broker_type_str = "BKR"
'For broker_type = 1 To 2
'With broker_list.Worksheets(broker_type_str)
'For Each file In all_broker_folder.Files
'''------------- create new email item ---------------
'    Set OutMail = OutApp.CreateItem(0)
'''------------- get underwriter emails from final report 360 ------------
'    final_report_360.Worksheets(1).Range("1:1").AutoFilter Field:=acc_code_360, Criteria1:=Replace(file.Name, ".xlsx", ""), Operator:=xlFilterValues
'    final_report_360.Worksheets(1).Range("1:1").AutoFilter Field:=narrative_col, Criteria1:="*2nd reminder*", Operator:=xlFilterValues
'    last_row_360 = final_report_360.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
'    If last_row_360 > 1 Then
'    Set visible_range = final_report_360.Worksheets(1).Range(final_report_360.Worksheets(1).Cells(2, uw_name_col), final_report_360.Worksheets(1).Cells(last_row_360, uw_name_col)).SpecialCells(xlCellTypeVisible)
'        For Each cell In visible_range
'            ResolveDisplayNameToSMTP (cell.Value)
'            If user_id <> "" And InStr(1, all_outmail_cc, user_id, vbTextCompare) = 0 Then
'                all_outmail_cc = all_outmail_cc & ";" & user_id
'            End If
'        Next cell
'    End If
'''------------- filter for broker list --------------
'    .Range("1:1").AutoFilter Field:=account_code_col, Criteria1:=Replace(file.Name, ".xlsx", ""), Operator:=xlFilterValues
'    last_row = .Cells(Rows.Count, 1).End(xlUp).Row
'''------------- English language ---------------------
'    If InStr(1, .Cells(last_row, lang_col), "English", vbTextCompare) Then
'        OutMail.To = .Cells(last_row, receiver_col)
'        OutMail.Subject = "AXA XL outstanding premium overview - " & Format(Date, "mmmm - yyyy") & " - " & .Cells(last_row, account_code_col) & " " & .Cells(last_row, acc_name_col)
'        OutMail.HTMLBody = english_body
'        OutMail.Attachments.Add file.Path
'            If all_outmail_cc <> "" Then
'            OutMail.CC = all_outmail_cc
'            End If
'        OutMail.Close (olSave)
'''------------- Spanish Language --------------------
''Debug.Print InStr(1, .Cells(last_row, lang_col).Value, "Spanish", vbTextCompare)
'    ElseIf InStr(1, .Cells(last_row, lang_col).Value, "Spanish", vbTextCompare) Then
'        OutMail.To = .Cells(last_row, receiver_col)
'        OutMail.Subject = "AXA XL Primas Pendiente - " & Format(Date, "mmmm - yyyy") & " - " & .Cells(last_row, account_code_col) & " " & .Cells(last_row, acc_name_col)
'        OutMail.HTMLBody = spanish_body
'        OutMail.Attachments.Add file.Path
'            If all_outmail_cc <> "" Then
'            OutMail.CC = all_outmail_cc
'            End If
'        OutMail.Close (olSave)
'    End If
'last_row = 0
'all_outmail_cc = ""
'Set OutMail = Nothing
'Next file
'End With
'Set all_broker_folder = fso.GetFolder(ThisWorkbook.Path & "\All Broker statements\ARI")
'broker_type_str = "ARI"
'Next broker_type
'
''
