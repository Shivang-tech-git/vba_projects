Attribute VB_Name = "A1_Fetchdata"
Sub SaveXLAttachment()
 Dim xlApp As Object
 Dim xlWB As Object
 Dim xlSheet As Object
 Dim rCount As Long
 Dim olkMsg As Object
 Dim olApp As Outlook.Application
 Dim olFolder As Outlook.MAPIFolder
 Dim olFolder1 As Outlook.MAPIFolder
 Dim olNs As Outlook.Namespace
 Dim ws1 As Worksheet
 Dim ws2 As Worksheet
 Dim m As Workbook
 Dim olItem As MailItem
 Dim olAttach As Attachment
 Dim Filename As String
 Dim strname As String
 Dim strpath As String
 Dim fso As FileSystemObject
Path = ThisWorkbook.Path
strpath = ThisWorkbook.Path & "\Word Doc\Temp Folder\" 'the path where the attachments are to be saved
srtname = "Underwriter Referral Template*"
'--------------- Setting up workbook---------------------------------------------------
Set m = ThisWorkbook
Set fso = New FileSystemObject
'Set ws1 = m.Sheets("MailBody")
'Set ws2 = m.Sheets("Defaults")

'----------------Defining Sub folder---------------------------------------------------
Set objOL = Outlook.Application
Set olNs = objOL.GetNamespace("MAPI")
Set objFolder = olNs.GetDefaultFolder(olFolderInbox).Folders("NDBI")

'---------------Downloading attachment-------------------------------------------------
For Each olItem In objFolder.Items
Subject = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(olItem.Subject, ":", ""), "/", ""), """", ""), "\", ""), "?", ""), "*", ""), "<", ""), ">", ""), "|", "")
fso.CreateFolder strpath & Subject & " " & Format(olItem.ReceivedTime, "HH-MM-SS")
Set email_item_folder = fso.GetFolder(strpath & Subject & " " & Format(olItem.ReceivedTime, "HH-MM-SS"))
    For Each olAttach In olItem.Attachments
    '--------------- save underwriter referral template.docx -------------
        If olAttach.Filename <= srtname And _
        Right(LCase(olAttach.Filename), 4) = "docx" Then
            Filename = email_item_folder.Path & "\" & olAttach.Filename
            olAttach.SaveAsFile Filename
    '--------------- save pdf or docx -----------------------------
        ElseIf Right(LCase(olAttach.Filename), 3) = "pdf" Or _
        Right(LCase(olAttach.Filename), 4) = "docx" Then
            Filename = email_item_folder.Path & "\" & olAttach.Filename
            olAttach.SaveAsFile Filename
        End If
    Next olAttach
Set email_item_folder = Nothing
Next olItem
'-----------------------------------------------------------------------------------------------



MsgBox "Done! Email attachments downloaded to ""Temp Folder."""
   
End Sub




