Attribute VB_Name = "Z_GetMails"
    Sub ExtractEmailData()
    
    'On Error GoTo catch
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim olApp As Outlook.Application, olNs As Outlook.Namespace
    Dim olFolder As Outlook.MAPIFolder, olMail As Outlook.MailItem
    Dim eFolder As Outlook.Folder
    Dim I As Long
    Dim x As Date, ws As Worksheet
    Dim objSender
    Dim exUser
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim m As Workbook
    Dim olRecip, ShareInbox
    Dim x1 As Variant, RI As Integer, CI As Integer, proceed As Boolean
    Dim rCount1 As Integer
    Dim counter(1 To 2) As Integer
    'Set ws = ActiveSheet
    Set olApp = New Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    x = Date
    
    '-------open Final Data sheet and move data to compiled file---------

    Set m = ThisWorkbook
    Set ws1 = m.Sheets("Final Data")
    Workbooks.Open Filename:=ThisWorkbook.Path & "\Compiled.xlsx"
    Set ws2 = Workbooks("Compiled.xlsx").Worksheets("Sheet1")
    
    LR = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    lr2 = ws2.Cells(Rows.Count, 2).End(xlUp).Row + 2
    
    '-----------------Select rows that do not contain any empty cell------------------
    counter(1) = 0
    proceed = False
    For RI = 16 To LR Step 1
        For CI = 1 To 15 Step 1
            If ws1.Cells(RI, CI) = "" Then
            proceed = False
            Exit For
            Else
            proceed = True
            End If
        Next CI
        If proceed = True Then
        ws1.Cells(RI, CI).EntireRow.Copy
        lr2 = lr2 + 1
        ws2.Cells(lr2, 1).PasteSpecial xlPasteValues
        ws1.Cells(RI, CI).EntireRow.Delete
        counter(1) = counter(1) + 1
        RI = RI - 1
        proceed = False
        End If
    Next RI
    
    With ws2.UsedRange.Borders
    .Weight = xlThin
    .LineStyle = xlContinuous
    ws2.UsedRange.HorizontalAlignment = xlLeft
    End With
    On Error Resume Next
    ws2.Range("X2", "X" & lr2).NumberFormat = "yyyy-mm-dd"
    ws2.Range("AA2", "AA" & lr2).NumberFormat = "yyyy-mm-dd"
    
    Workbooks("Compiled.xlsx").Save
    Workbooks("Compiled.xlsx").Close
    
    '---------------------------------------------------------------
   
    x1 = ws1.Range("C16").Value
    counter(2) = 0
    For mailBoxCount = 3 To 5
    
    If m.Sheets("Defaults").Cells(4, mailBoxCount) = "" Then
    Exit For
    End If
    
    
    Set olFolder = olNs.Folders(m.Sheets("Defaults").Cells(4, mailBoxCount).Value)
    ''Set olFolder = olFolder.Folders("Inbox")
    Set olFolder = olFolder.Folders(m.Sheets("Defaults").Cells(7, mailBoxCount).Value)

    
    
    rCount1 = ws1.Range("C" & Rows.Count).End(-4162).Row + 1
    
    For I = 1 To olFolder.Items.Count Step 1
    If TypeOf olFolder.Items(I) Is MailItem Then
    Set olMail = olFolder.Items(I)
    
    '''''''''''''''''''''''''''''''''''
    'START Getting sender email address
    '''''''''''''''''''''''''''''''''''
    
    If olMail.SenderEmailType = "SMTP" Then
    SenderEmail = olMail.SenderEmailAddress
    Else
    
    SenderEmail = olMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
    If Len(SenderEmail) = 0 Then
    Set objSender = olMail.Sender
    If Not (objSender Is Nothing) Then
    'read PR_SMTP_ADDRESS_W
    SenderEmail = objSender.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
    If Len(SenderEmail) = 0 Then
    'last resort
    Set exUser = objSender.GetExchangeUser
    If Not (exUser Is Nothing) Then
      SenderEmail = exUser.PrimarySmtpAddress
    End If
    End If
    End If
    End If
    End If
    
    '''''''''''''''''''''''''''''''''''
    'END Getting sender email address
    '''''''''''''''''''''''''''''''''''

    ws1.Cells(rCount1, 3).Value = SenderEmail
    ws1.Cells(rCount1, 5).Value = olMail.Subject
    ws1.Cells(rCount1, 2).Value = olMail.SenderName
    ws1.Cells(rCount1, 4).Value = olMail.categories
    ws1.Cells(rCount1, 10).Value = Format(Date, "m-d-yyyy")
    rCount1 = rCount1 + 1
    counter(2) = counter(2) + 1
    End If
    Next I
    
    With ws1.Range(ws1.Cells(15, 1), ws1.Cells(rCount1, 22))
    .Borders.Weight = xlThin
    .Borders.LineStyle = xlContinuous
    .HorizontalAlignment = xlLeft
    End With
Next mailBoxCount
    
    
    Set olFolder = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
    '
    'catch:
    '
    '    If Err.Number <> 0 Then
    '        MsgBox ("Failed " & Err.Description)
    '    End If
    MsgBox " 1.  " & counter(1) & " Rows transferred to Compiled.xlsx file." & vbCrLf & " 2.  " & counter(2) & " New Emails appended.", vbInformation, "Information."
    End Sub
    
'What Get Mails button will do ?
'1. Move previous day completed allocation to compiled sheet. If any cell within column A:O is empty, that row will not get transferred to compiled sheet.
'2. For new mails, fill S. No. column automatically in ascending order. S. No. for previous mails will remain same. S. No. is important for Vlookup from End of Day Allocation status.
'3. Get all the mails from specified folder and get Sender Email, subject, sender name, categories and received date in their respective columns.
'4. VlookUp from Definations sheet for Processing and Review time.
'What Allocate button will do?
'1. VlookUP (from Defaults sheet) for user email ID to whom mail will be forwarded .
'2. Emails with today's allocation date will get allocated. Separate sheet will be created for review and process allocation for a processor.
'3. File name will be saved with processor's name and today's date in Output folder.
'What forward Mail button will do?
'1. VlookUP (from Defaults sheet) for user email ID to whom mail will be forwarded .
'2. Forward all mails starting from 16th row to processor. (Reviewer will not get the mails.)
'What allocation status button will do?
'1. VLookup from S.No. for Process Sheet and review sheet for column O to column U.

