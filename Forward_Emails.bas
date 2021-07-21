Attribute VB_Name = "Forward_Emails"
Dim category As String

Sub ForwardEmails()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
'On Error GoTo catch
    
    If MsgBox("Emails will be forwarded. Proceed?", vbOKCancel + vbQuestion) = vbOK Then
        
        Dim olApp As Outlook.Application
        Dim olNs As Outlook.Namespace
        Dim Inbox As Outlook.MAPIFolder
        Dim Item As Object
        Dim MsgFwd As MailItem
        Dim Items As Outlook.Items
        Dim Recip As Recipient
        Dim EmailTO As String
        Dim EmailFrom As String
        Dim EmailCC As String
        Dim EmailSignature As String
        Dim RequesterEmail As String
        Dim ExchangeAccountNo As Integer
        Dim ItemSubject As String
        Dim lngCount As Long
        Dim I As Long
        Dim EmailSentCounter As Integer
        Dim olFolder As Outlook.MAPIFolder
        Dim ws1 As Worksheet
        Dim ws2 As Worksheet
        Dim m As Workbook
        Dim confidence As Boolean
        Set m = ThisWorkbook
    
    Set ws1 = m.Sheets("Final Data")
        Set olApp = CreateObject("Outlook.Application")
        Set olNs = olApp.GetNamespace("MAPI")
        
'        Set Inbox = olNs.GetDefaultFolder(olFolderInbox)
'        Set TargetFolder = olNs.GetDefaultFolder(olFolderInbox).Folders("outlook_forms")
'        Set Items = TargetFolder.Items
        
        
    
        
 '--------------User Email ID VLookUP-------------------
 LastRow = m.Sheets("Final Data").Cells(Rows.Count, 12).End(xlUp).Row
 With ws1
  RowInc = 16
    For RowInc = 16 To LastRow
          If .Range("V" & RowInc) = "" Then
          On Error Resume Next
          .Range("V" & RowInc) = Application.WorksheetFunction.VLookup(.Range("L" & RowInc), m.Sheets("Defaults").Range("B:C"), 2, 0)
          End If
    Next RowInc
 End With
 
 ''-----------create unique id's for forwarding an email only once -----------------
 
 RowInc = 16
 With ws1
    For RowInc = 16 To LastRow
    m.Worksheets("Help").Cells(RowInc - 15, 1) = Application.WorksheetFunction.Concat(.Cells(RowInc, 5).Value, .Cells(RowInc, 12).Value)
    Next RowInc
 End With
 m.Sheets("Help").Range("$A:$A").RemoveDuplicates Columns:=1
 
  '-----------Get Sender Details------------------------
  
        With m.Sheets("Defaults")
            EmailFrom = .Cells(4, 3)
            EmailCC = .Cells(5, 3)
            EmailSignature = .Cells(6, 3)
        End With
        
        EmailSentCounter = 0
        'I = StartRow 'StartRow is defined from when get emails button is clicked
        'MsgBox (ActiveSheet.UsedRange.Columns(4).Column)
        With m.Worksheets("Final Data") ' Sheet Name
        
        ws1.Activate
        LastRow = ws1.Cells(Rows.Count, 12).End(xlUp).Row
    
        For K = 16 To LastRow
        
            If (IsEmpty(Cells(K, 22).Value) = False) Then
            
                If Cells(K, 10) = Date Then 'dont forward mail if it's not allocated today.
                
                   confidence = False
                   For rownum = 1 To m.Worksheets("Help").Cells(Rows.Count, 1).End(xlUp).Row
                   If InStr(1, m.Worksheets("Help").Cells(rownum, 1).Value, Application.WorksheetFunction.Concat(Cells(K, 5).Value, Cells(K, 12).Value), vbBinaryCompare) <> 0 Then
                   m.Worksheets("Help").Cells(rownum, 1).EntireRow.Delete
                   confidence = True
                   End If
                   Next rownum
                   
                   
                   
                    RequesterEmail = .Cells(K, 3).Value & ";"  'Appened ; so that this can be directly used in CC
                    ItemSubject = .Cells(K, 5).Value
                    EmailTO = .Cells(K, 22).Value
                    'ExchangeAccountNo = GetExchangeAccountNo(EmailFrom) 'Get the exchange account number like for default it is 1 and for other account 2 and so on.
                  If confidence = True Then
                    
                        Call categories(UCase(Cells(K, 4).Value))   '''Check wether email category falls within prohibited category or not
                        If category = "" Then             '''if it does not falls with in prohibited category then forward mail
                            If (EmailTO = "") Then
                                
                                MsgBox ("No email was specified for " & K & "Row, eMail will not be forwarded for this row")
                                
                            Else
                            
                For mailBoxCount = 3 To 5
                
                If m.Sheets("Defaults").Cells(4, mailBoxCount) = "" Then
                Exit For
                End If
                                Set olFolder = olNs.Folders(m.Sheets("Defaults").Cells(4, mailBoxCount).Value)
                                  ''Set olFolder = olFolder.Folders("Inbox")
                                Set olFolder = olFolder.Folders(m.Sheets("Defaults").Cells(7, mailBoxCount).Value)
                                Set Items = olFolder.Items
                                
                                '// Loop through Inbox Items backwards
                                For lngCount = Items.Count To 1 Step -1
                                    Set Item = Items.Item(lngCount)
                                    
                                    If Item.Subject = ItemSubject Then ' if Subject found then
                                        
                                         Set MsgFwd = Item.Forward
                                         MsgFwd.HTMLBody = "Hi, <br><br>Please process.<br><br>" & EmailSignature & "<br><br>" & Item.Forward.HTMLBody
                                         Set Recip = MsgFwd.Recipients.Add(EmailTO) ' add Recipient
                                         Recip.Type = olTo
                                         MsgFwd.CC = RequesterEmail & EmailCC
                                         'MsgFwd.Display
                                        MsgFwd.SendUsingAccount = olApp.Session.Accounts.Item(m.Sheets("Defaults").Cells(3, mailBoxCount).Value)
                                     ''    MsgFwd.Display
                                         EmailSentCounter = EmailSentCounter + 1
                                       MsgFwd.Send
                                    
                                    End If
                                Next ' exit loop
                                
                Next mailBoxCount
                            End If
                        'k = k + 1 '  = Row 2 + 1 = Row 3
                    End If
                  
                  End If
                
                End If
                End If 'IsEmpty(Cells(k, 4).Value) = False
            
        Next K
        
            'Do Until (.Cells(I, 4) = "")
                
                
            'Loop
        End With
            
            Set olApp = Nothing
            Set olNs = Nothing
            Set Inbox = Nothing
            Set Item = Nothing
            Set MsgFwd = Nothing
            Set Items = Nothing
            MsgBox "Total emails sent / forwarded: " & EmailSentCounter, vbInformation, "Information."
    Else
        
        MsgBox ("Process canceled, No emails sent!")
        
        Exit Sub
        
    End If
    m.Sheets("Help").Columns("A:A").ClearContents
    Set ws1 = Nothing
    
'catch:
'
'    If Err.Number <> 0 Then
'        MsgBox ("Failed " & Err.Description)
'    End If
    
End Sub

'Gets the Exchange Account no. of the specified email id.
''1 = Default email id 2 for other and so on...
'Private Function GetExchangeAccountNo(ByVal FromEmail As String)
''Don't forget to set a reference to Outlook in the VBA editor
'    Dim OutApp As Outlook.Ap plication
'    Dim I As Long
'
'    Set OutApp = CreateObject("Outlook.Application")
'
'    For I = 1 To OutApp.Session.Accounts.Count
'
'        If (LCase(OutApp.Session.Accounts.Item(I)) = LCase(FromEmail)) Then
'            GetExchangeAccountNo = I
'        End If
'    Next I
'
'End Function

Sub categories(emailCategory As String)

category = ""

Dim categoryType(0 To 22) As String
    
    categoryType(0) = UCase("Final Auto Adjustment")
    categoryType(1) = UCase("P&C - Issuance - AUT")
    categoryType(2) = UCase("P&C - Issuance - GL")
    categoryType(3) = UCase("P&C - Issuance - UMB")
    categoryType(4) = UCase("Pollution - Issuance - Others")
    categoryType(5) = UCase("Pollution - Issuance - PCP / CPL")
    categoryType(6) = UCase("Pollution - Issuance - EVPCP")
    categoryType(7) = UCase("Pollution - Issuance - PRL")
    categoryType(8) = UCase("Pollution - Issuance - STAG")
    categoryType(9) = UCase("ANI")
    categoryType(10) = UCase("Correction Post Bind")
    categoryType(11) = UCase("Rework Post Bind")
    categoryType(12) = UCase("Triaging Post Bind")
    categoryType(13) = UCase("Forms/MANUS")
    categoryType(14) = UCase("Class Codes")
    categoryType(15) = UCase("Script")
    categoryType(16) = UCase("Location Coding – Excel")
    categoryType(17) = UCase("GL Coding Locations")
    categoryType(18) = UCase("State")
    categoryType(19) = UCase("FAC")
    categoryType(20) = UCase("Vehicles")
    categoryType(21) = UCase("MCS 90")
    categoryType(22) = UCase("Location Coding")

For Z = 0 To 22
If categoryType(Z) = emailCategory Then
category = categoryType(Z)
End If
Next Z


End Sub
