Attribute VB_Name = "Module1"
Sub Signature()
Attribute Signature.VB_ProcData.VB_Invoke_Func = "q\n14"
''On Error GoTo errorHandler:
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer
Dim macroSheet As Worksheet, emailSheet As Worksheet
Dim olApp As Outlook.Application, olNs As Outlook.Namespace
Dim olFolder As Outlook.MAPIFolder, olMail As Outlook.MailItem
Dim sent_folder As Outlook.Folder, recipient_email_address As String
Dim dict As Scripting.Dictionary, subject_filter As String, fwd_mail As MailItem

''-------------Set Object References----------------------------------------
Set dict = New Scripting.Dictionary
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set emailSheet = ThisWorkbook.Worksheets("Emails")
Set olApp = New Outlook.Application
Set olNs = olApp.GetNamespace("MAPI")
Set olFolder = olNs.Folders(macroSheet.Cells(5, 4).Value)
Set sent_folder = olFolder.Folders("Sent Items")
Set olFolder = olFolder.Folders(macroSheet.Cells(6, 4).Value)
If macroSheet.Cells(7, 4).Value <> "" Then
Set olFolder = olFolder.Folders(macroSheet.Cells(7, 4).Value)
End If

''---------------- constants -----------------
Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
Const subject_col = 1
Const received_time_col = 2
Const reminder_col = 3
Const recipient_col = 4
start_row = 2
dict.Add 30, "REMINDER #5"
dict.Add 25, "REMINDER #4"
dict.Add 20, "REMINDER #3"
dict.Add 15, "REMINDER #2"
dict.Add 10, "REMINDER #1"
'' --------------------- delete old data -----------------------
emailSheet.Rows("2:" & Rows.Count).ClearContents
emailSheet.Cells.Borders.LineStyle = xlNone

''----------------------- filter outlook to include emails older than 10 working days -------------
last_working_date = WorksheetFunction.WorkDay(Date, -10)
Set recent_mails = olFolder.Items.Restrict("[ReceivedTime] < '" & Format(last_working_date, "MM/DD/YYYY") & "'")
'' ---------------------- Send Reminders -------------------------
For I = 1 To olFolder.Items.Count
    If TypeOf olFolder.Items(I) Is MailItem Then
        Set olMail = olFolder.Items(I)
'' --------- Check if recipient email address if from axaxl.com --------------
        recipient_email_address = olMail.Recipients(1).PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        If recipient_email_address Like "*axaxl.com" Then
            subject_filter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " LIKE '%" & olMail.Subject & "%'"
            reminder_filter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " LIKE '%" & dict(Key) & "%'"
'' --------- check if email is (30 or 25 or 20 or 15 or 10) days old and it's reminder is not sent already ---------
            For Each Key In dict.Keys
            Debug.Print DateDiff("d", olMail.ReceivedTime, Format(WorksheetFunction.WorkDay(Date, -(Key)), "MM/DD/YYYY"))
                If dict(Key) <> Empty And _
                DateDiff("d", olMail.ReceivedTime, Format(WorksheetFunction.WorkDay(Date, -(Key)), "MM/DD/YYYY")) >= 0 _
                And sent_folder.Items.Restrict(subject_filter).Restrict(reminder_filter).Count = 0 Then
''----------------------------------- forward email -----------------------------------
                Set fwd_mail = olMail.Forward
                fwd_mail.Subject = olMail.Subject & " " & dict(Key)
                fwd_mail.Recipients.Add recipient_email_address
                If Key >= 25 And Not macroSheet.Range("C10:C14").Find(recipient_email_address) Is Nothing Then
                fwd_mail.CC = macroSheet.Range("C10:C14").Find(recipient_email_address).Offset(0, 1).Value
                End If
'                fwd_mail.Display
'                fwd_mail.Send
''------------- save details of forwarded email in Emails worksheet ---------------
                emailSheet.Cells(start_row, subject_col) = fwd_mail.Subject
                emailSheet.Cells(start_row, received_time_col) = Format(olMail.ReceivedTime, "MM/DD/YYYY")
                emailSheet.Cells(start_row, reminder_col) = dict(Key)
                emailSheet.Cells(start_row, recipient_col) = recipient_email_address
                start_row = start_row + 1
                Exit For
                End If
            Next Key
        End If
    End If
Next I

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
MsgBox "Complete, Please check Emails sheet for sent reminder details.", vbInformation, "Outlook reminders sendout tool"

Set fwd_mail = Nothing
Set dict = Nothing
Set macroSheet = Nothing
Set emailSheet = Nothing
Set olApp = Nothing
Set olNs = Nothing
Set olFolder = Nothing

End Sub






























