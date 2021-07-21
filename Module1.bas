Attribute VB_Name = "Module1"
Dim regex As RegExp
Dim inv_details_ws As Worksheet

Sub InvoiceDetails()
Attribute InvoiceDetails.VB_ProcData.VB_Invoke_Func = "q\n14"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim bar As Object, rCount1 As Integer
    Dim olApp As Outlook.Application, olNs As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder, olMail As Outlook.MailItem
    Dim eFolder As Outlook.Folder
    Dim I As Long
    Dim payments_tool As Workbook
    
    Set regex = New RegExp
    Set payments_tool = ThisWorkbook
    Set inv_details_ws = payments_tool.Sheets("Invoice details")
    Set macro_ws = payments_tool.Sheets("Tool")
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.Folders(macro_ws.Cells(5, 4).Value)
    Set olFolder = olFolder.Folders("Inbox")
    Set olFolder = olFolder.Folders(macro_ws.Cells(6, 4).Value)
    
    inv_details_ws.Rows("2:" & Rows.Count).ClearContents
    inv_details_ws.UsedRange.Borders.LineStyle = xlNone
    
    rCount1 = 2
    Const sub_col = 2
    Const claim_col = 3
    Const vendor_col = 4
    Const total_col = 5
    Const invoice_col = 6
    Const BT_col = 7
    
    last_working_date = WorksheetFunction.WorkDay(Date, -2)
    Set recent_mails = olFolder.Items.Restrict("[ReceivedTime] > '" & Format(last_working_date, "MM/DD/YYYY") & "'")
    For Each olMail In recent_mails
        ''----------- s. no. and subject -----
        inv_details_ws.Cells(rCount1, 1).Value = rCount1 - 1
        inv_details_ws.Cells(rCount1, sub_col).Value = olMail.Subject
        ''------------- claim number ------------
        regular_expression_engine "(Claim\s#: )(\d+)", olMail.body, rCount1, claim_col
        ''------------- vendor name ------------
        regular_expression_engine "(Vendor Name: )(.*)", olMail.body, rCount1, vendor_col
        ''------------- total to pay ------------
        regular_expression_engine "(Total To Pay: )(.*)", olMail.body, rCount1, total_col
        ''------------- invoice number ------------
        regular_expression_engine "(Invoice #: )(.*)", olMail.body, rCount1, invoice_col
        ''------------- BT Total ------------
        regular_expression_engine "(BT Total: )(.*)", olMail.body, rCount1, BT_col
        rCount1 = rCount1 + 1
    Next olMail
    
    With inv_details_ws.UsedRange
        .Borders.Weight = xlThin
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlLeft
    End With
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
Set olFolder = Nothing
Set inv_details_ws = Nothing
Set regex = Nothing

 MsgBox "Complete.", vbInformation, "Bottomline Payments Tool"
 
 End Sub
 
Sub regular_expression_engine(ByRef pattern As String, ByRef source_string As String, ByRef row As Integer, col)
regex.pattern = pattern
Set extracted_string = regex.Execute(source_string)
If extracted_string.Count > 0 Then
    If extracted_string(0).SubMatches.Count > 1 Then
        inv_details_ws.Cells(row, col).Value = extracted_string(0).SubMatches.Item(1)
    End If
End If
End Sub
    
    

