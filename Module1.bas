Attribute VB_Name = "Module1"
Sub Signature()
Attribute Signature.VB_ProcData.VB_Invoke_Func = "q\n14"
''On Error GoTo errorHandler:
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer
Dim macroSheet As Worksheet, emailSheet As Worksheet, obuSheet As Worksheet
Dim taxonomySheet As Worksheet, bar_wb As Workbook, bar_ws As Worksheet, policy_level_ws As Worksheet
Dim olApp As Outlook.Application, olNs As Outlook.Namespace
Dim olFolder As Outlook.MAPIFolder, olMail As Outlook.MailItem
Dim doctype(1 To 3) As String
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set emailSheet = ThisWorkbook.Worksheets("Emails")
Set obuSheet = ThisWorkbook.Worksheets("OBU")
Set taxonomySheet = ThisWorkbook.Worksheets("Taxonomy")
bar_wb_name = Dir(ThisWorkbook.Path & "\Business Activity Report*.xlsx")
Set bar_wb = Workbooks.Open(ThisWorkbook.Path & "\" & bar_wb_name)
Set bar_ws = bar_wb.Worksheets("Business Activity Report")
Set policy_level_ws = bar_wb.Worksheets("Policy_level_Data_Orig")
Set olApp = New Outlook.Application
Set olNs = olApp.GetNamespace("MAPI")
Set olFolder = olNs.Folders(macroSheet.Cells(5, 4).Value)
Set olFolder = olFolder.Folders(macroSheet.Cells(6, 4).Value)
If macroSheet.Cells(7, 4).Value <> "" Then
Set olFolder = olFolder.Folders(macroSheet.Cells(7, 4).Value)
End If

''---------------- find required columns in all worksheets -----------------
Const subject_col = 1
Const document_name_col = 2
Const policy_or_quote_col = 3
Const category_col = 4
Const comments_col = 5
Const fkproduct_col = 6
Const lob_col = 7
Const office_col = 8
Const obu_col = 9
Const document_class_col = 10
Const document_category_col = 11
Const document_type_col = 12

bar_policy_col = bar_ws.Range("1:1").Find("Policy_number").Column
bar_quote_col = bar_ws.Range("1:1").Find("Quote_Number").Column
bar_fkproduct_col = bar_ws.Range("1:1").Find("FKProduct").Column
bar_office_col = bar_ws.Range("1:1").Find("FKProducingOffice").Column
bar_lob_col = bar_ws.Range("1:1").Find("Quote_Policy_Title").Column
policy_level_policy_col = policy_level_ws.Range("1:1").Find("Policy_No").Column
policy_level_quote_col = policy_level_ws.Range("1:1").Find("Quote_No").Column
policy_level_fkproduct_col = policy_level_ws.Range("1:1").Find("FKProduct").Column
policy_level_lob_col = policy_level_ws.Range("1:1").Find("Line of Business").Column
policy_level_office_col = policy_level_ws.Range("1:1").Find("FKProducingOffice").Column

'' --------------------- delete old data -----------------------
emailSheet.Rows("2:" & Rows.Count).ClearContents
emailSheet.Cells.Borders.LineStyle = xlNone

start_row = 2
start_col = 2
RowCount = start_row
colcount = start_col
'' ---------------------- extract emails -------------------------
For I = 1 To olFolder.Items.Count Step 1
    If TypeOf olFolder.Items(I) Is MailItem Then
        Set olMail = olFolder.Items(I)
       '' If InStr(1, olMail.Subject, "FW: test email for IDOC BOT", vbTextCompare) Then
        ''------------------------extract table from email body ----------------------------------
        Set documents_table = olMail.GetInspector.WordEditor.Tables.Item(1)
        total_rows = documents_table.Rows.Count
        total_columns = documents_table.Rows.Item(1).Cells.Count
            For table_row = 2 To total_rows
                If InStr(1, documents_table.Rows.Item(table_row).Cells.Item(3).Range.Text, "Choose an item", vbTextCompare) = 0 Then
                    For table_cell = 1 To total_columns
                        emailSheet.Cells(RowCount, colcount) = documents_table.Rows.Item(table_row).Cells.Item(table_cell).Range.Text
                        colcount = colcount + 1
                    Next table_cell
                    emailSheet.Cells(RowCount, subject_col) = olMail.Subject
                    colcount = start_col
                    RowCount = RowCount + 1
                End If
            Next table_row
       '' End If
    End If
Next I

''---------------- find last rows for all worksheets ---------------------
email_last_row = emailSheet.Cells(Rows.Count, subject_col).End(xlUp).Row
obu_last_row = obuSheet.Cells(Rows.Count, 9).End(xlUp).Row
taxonomy_last_row = taxonomySheet.Cells(Rows.Count, 1).End(xlUp).Row
policy_level_last_row = policy_level_ws.Cells(Rows.Count, 1).End(xlUp).Row
bar_last_row = bar_ws.Cells(Rows.Count, 1).End(xlUp).Row

''------------------ text to column --------------------------------------

ttc_array = Array(document_name_col, policy_or_quote_col, category_col, comments_col)

emailSheet.Columns(category_col).TextToColumns Destination:=emailSheet.Columns(category_col), _
                                            DataType:=xlDelimited, _
                                            ConsecutiveDelimiter:=True, _
                                            Tab:=False, _
                                            Semicolon:=False, _
                                            Comma:=False, _
                                            Space:=False, _
                                            Other:=True, OtherChar:=Chr(7)
                                            
For Each col In ttc_array

    emailSheet.Columns(col).TextToColumns Destination:=emailSheet.Columns(col), _
                                            DataType:=xlDelimited, _
                                            ConsecutiveDelimiter:=True, _
                                            Tab:=True, _
                                            Semicolon:=False, _
                                            Comma:=False, _
                                            Space:=False, _
                                            Other:=False
   
Next col
                                            
'' ---------------------- FIND FKProduct and Line of business ------------------------------
For email_row = 2 To email_last_row

    policy_row = ""
    fk_product = ""
    line_of_business = ""
    producing_office = ""
    obu = ""
    document_class = ""
    document_category = ""
    document_type = ""
    policy_number = False
'' -------------------------- C = policy, q = quote  -----------

    If InStr(1, Left(emailSheet.Cells(email_row, 3), 1), "C", vbTextCompare) > 0 Then
        policy_number = True
        bar_nbr = bar_policy_col
        pldo_nbr = policy_level_policy_col
        obu_prod_lob_concat = 1
        obu_prod_office_concat = 3
    ElseIf InStr(1, Left(emailSheet.Cells(email_row, 3), 1), "Q", vbTextCompare) > 0 Then
        policy_number = False
        bar_nbr = bar_quote_col
        pldo_nbr = policy_level_quote_col
        obu_prod_lob_concat = 2
        obu_prod_office_concat = 4
    End If
    
    ''-----------------Index and match from business activity report------------------------------
                          
    If policy_number = False And Application.WorksheetFunction.CountIf(bar_ws.Range(bar_ws.Cells(1, bar_nbr), bar_ws.Cells(bar_last_row, bar_nbr)), _
                                                emailSheet.Cells(email_row, policy_or_quote_col)) > 0 Then
    
    policy_row = Application.WorksheetFunction.Match(emailSheet.Cells(email_row, policy_or_quote_col), _
                                bar_ws.Range(bar_ws.Cells(1, bar_nbr), bar_ws.Cells(bar_last_row, bar_nbr)), 0)

    fk_product = Application.WorksheetFunction.Index(bar_ws.Range(("A1"), ("AH") & bar_last_row), policy_row, bar_fkproduct_col)
                                    
    line_of_business = Split(Application.WorksheetFunction.Index(bar_ws.Range(("A1"), ("AH") & bar_last_row), policy_row, bar_lob_col))(0)
                                    
    producing_office = Application.WorksheetFunction.Index(bar_ws.Range(("A1"), ("AH") & bar_last_row), policy_row, bar_office_col)
                                    
''-----------------Index and match from policy level sheet------------------------------

    ElseIf Application.WorksheetFunction.CountIf(policy_level_ws.Range(policy_level_ws.Cells(1, pldo_nbr), _
                                            policy_level_ws.Cells(policy_level_last_row, pldo_nbr)), emailSheet.Cells(email_row, policy_or_quote_col)) > 0 Then
    
    policy_row = Application.WorksheetFunction.Match(emailSheet.Cells(email_row, policy_or_quote_col), _
                                    policy_level_ws.Range(policy_level_ws.Cells(1, pldo_nbr), policy_level_ws.Cells(policy_level_last_row, pldo_nbr)), 0)

    fk_product = Application.WorksheetFunction.Index(policy_level_ws.Range(("A1"), ("AU") & policy_level_last_row), policy_row, policy_level_fkproduct_col)
                                    
    line_of_business = Application.WorksheetFunction.Index(policy_level_ws.Range(("A1"), ("AU") & policy_level_last_row), policy_row, policy_level_lob_col)

    producing_office = Application.WorksheetFunction.Index(policy_level_ws.Range(("A1"), ("AU") & policy_level_last_row), policy_row, policy_level_office_col)
                                    
    End If
    
    emailSheet.Cells(email_row, fkproduct_col) = fk_product
    emailSheet.Cells(email_row, lob_col) = line_of_business
    emailSheet.Cells(email_row, office_col) = producing_office
    
    ''---------------------------- Find OBU --------------------------------------------

    If emailSheet.Cells(email_row, office_col) = "AEB" Then
        
        obu_row = Application.WorksheetFunction.Match(emailSheet.Cells(email_row, fkproduct_col) & emailSheet.Cells(email_row, office_col), _
                                obuSheet.Range(obuSheet.Cells(1, obu_prod_office_concat), obuSheet.Cells(obu_last_row, obu_prod_office_concat)), 0)
        
        obu = Application.WorksheetFunction.Index(obuSheet.Range(("A1"), ("I") & obu_last_row), obu_row, 9)
    
    
    ElseIf Application.WorksheetFunction.CountIf(obuSheet.Range(obuSheet.Cells(1, obu_prod_lob_concat), obuSheet.Cells(obu_last_row, obu_prod_lob_concat)), _
                                                emailSheet.Cells(email_row, fkproduct_col) & emailSheet.Cells(email_row, lob_col)) > 0 Then
                                                
        obu_row = Application.WorksheetFunction.Match(emailSheet.Cells(email_row, fkproduct_col) & emailSheet.Cells(email_row, lob_col), _
                                obuSheet.Range(obuSheet.Cells(1, obu_prod_lob_concat), obuSheet.Cells(obu_last_row, obu_prod_lob_concat)), 0)
        
        obu = Application.WorksheetFunction.Index(obuSheet.Range(("A1"), ("I") & obu_last_row), obu_row, 9)
    
    
    End If
    ''------------------- Document class, category and type ----------------------------

    If Application.WorksheetFunction.CountIf(taxonomySheet.Range(taxonomySheet.Cells(1, 1), taxonomySheet.Cells(taxonomy_last_row, 1)), _
                                                Trim(emailSheet.Cells(email_row, category_col))) > 0 Then
    
    category_row = Application.WorksheetFunction.Match(Trim(emailSheet.Cells(email_row, category_col)), _
                            taxonomySheet.Range(taxonomySheet.Cells(1, 1), taxonomySheet.Cells(taxonomy_last_row, 1)), 0)

    document_class = Application.WorksheetFunction.Index(taxonomySheet.Range(("A1"), ("D") & taxonomy_last_row), category_row, 2)
    document_category = Application.WorksheetFunction.Index(taxonomySheet.Range(("A1"), ("D") & taxonomy_last_row), category_row, 3)
    document_type = Application.WorksheetFunction.Index(taxonomySheet.Range(("A1"), ("D") & taxonomy_last_row), category_row, 4)
    
    End If
    
    emailSheet.Cells(email_row, obu_col) = obu
    emailSheet.Cells(email_row, document_class_col) = document_class
    emailSheet.Cells(email_row, document_category_col) = document_category
    emailSheet.Cells(email_row, document_type_col) = document_type
    
Next email_row

''----------------------------------------------------------------------------------
With emailSheet.UsedRange.Borders
    .Weight = xlThin
    .LineStyle = xlContinuous
End With
emailSheet.UsedRange.HorizontalAlignment = xlLeft
ThisWorkbook.Save

''-------------The End------------------------------------------------------

Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Complete.", vbInformation, "IDOC Indexing tool"

Set macroSheet = Nothing
Set emailSheet = Nothing
Set obuSheet = Nothing
Set taxonomySheet = Nothing
Set bar_wb = Nothing
Set bar_ws = Nothing
Set policy_level_ws = Nothing
Set olApp = Nothing
Set olNs = Nothing
Set olFolder = Nothing

End Sub






























