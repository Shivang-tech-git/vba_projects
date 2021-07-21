Attribute VB_Name = "Development"
Public total As Integer
Sub appendToDatabase()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Error Handling Routine-----------------------------
''on error resume next
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer, getOut As Boolean
Dim ufSheet As Worksheet, ole As OLEObjects
Dim OutlookApp As Outlook.Application
Dim OutlookMail As Outlook.MailItem
''-------------Set Object References----------------------------------------
Set ufSheet = ThisWorkbook.Worksheets("Userform")
Set ole = ufSheet.OLEObjects
Set OutlookApp = New Outlook.Application
Set OutlookMail = OutlookApp.CreateItem(olMailItem)
''----------------Checking Mandatory fields------------------------
getOut = False


'If ole.Item(12).Object.Value = 1 Then
'Call groupnomination ''''(***********GOING HERE )
'Else

If ufSheet.OLEObjects("Checkbox1").Object.Value = True Then
Call GroupNomination_macro
Else


If ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "COMMITMENT/ATTITUDE" Then
'''''------------------HERE SHOULLD HAVE BEEN ADDIOTINAL CHECK, IF THE PERSON NOMINATED 4 TIMES OR IF TOTAL AMMOUNT OF GIVEN POINTS ITS 40 !
        If ole.Item(11).Object.text = "" Then
        MsgBox "Please provide Nominee's Name.", vbInformation, "Rewards and Recognition Tool"
        getOut = True
'        ElseIf ole.Item(5).Object.text = "" Then
'        MsgBox "Please provide Nominee's Position.", vbInformation, "Rewards and Recognition Tool"
'        getOut = True
'        ElseIf ole.Item(3).Object.text = "" Then
'        MsgBox "Please provide Nominee's Band.", vbInformation, "Rewards and Recognition Tool"
'        getOut = True
        ElseIf ole.Item(9).Object.text = "" Then
        MsgBox "Please don't leave 'What do you nominate for?' field empty!", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        'ElseIf ole.Item(9).Object.text = "" Then
        'MsgBox "Please don't leave 'Data and any necessary comments on Efficiency.' field empty!", vbInformation, "Rewards and Recognition Tool"
        'getOut = True
        'ElseIf ole.Item(8).Object.text = "" Then
        'MsgBox "Effect/impact on the service provided", vbInformation, "Rewards and Recognition Tool"
        'getOut = True
'        ElseIf ole.Item(4).Object.text = "" Then
'        MsgBox "Please provide Business Relation.", vbInformation, "Rewards and Recognition Tool"
'        getOut = True

        
          ElseIf ufSheet.Shapes("Drop Down 6").ControlFormat.Value = 0 Then
           MsgBox "Please select prize", vbInformation, "Rewards and Recognition Tool"
            getOut = True
            End If
        If getOut = True Then Exit Sub
Else

'MsgBox UFSheet.Shapes("Drop Down 6").ControlFormat.Value
'MsgBox ole.Item(5).Object.text
'MsgBox ole.Item(3).Object.text
'MsgBox ole.Item(2).Object.text
'MsgBox ole.Item(9).Object.text
'MsgBox ole.Item(8).Object.text

        If ole.Item(11).Object.text = "" Then
        MsgBox "Please provide Nominee's Name.", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        ElseIf ole.Item(2).Object.text = "" Then
        MsgBox "Please provide Nominee's Position.", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        ElseIf ole.Item(1).Object.text = "" Then
        MsgBox "Please provide Nominee's Band.", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        ElseIf ole.Item(9).Object.text = "" Then
        MsgBox "Please don't leave 'What do you nominate for?' field empty!", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        ElseIf ole.Item(4).Object.text = "" Then
        MsgBox "Please don't leave 'Data and any necessary comments on Efficiency.' field empty!", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        ElseIf ole.Item(6).Object.text = "" Then
        MsgBox "Effect/impact on the service provided", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        ElseIf ufSheet.Shapes("Drop Down 6").ControlFormat.Value = 0 Then
        MsgBox "Please select prize", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        ElseIf ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "CI" And ole.Item(10).Object.text = "" Then
        MsgBox "Please provide SCCP number", vbInformation, "Rewards and Recognition Tool"
        getOut = True
        
        
        End If
        

End If
If ufSheet.OLEObjects("Checkbox1").Object.Value = False Then
If getOut = True Then Exit Sub
''=================Lukasz, this is for you=========================
''--------------Check Points limit for nomination Eligibility -----
If ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "COMMITMENT/ATTITUDE" Then
Prize = Replace(ufSheet.Shapes("Drop Down 6").ControlFormat.List(ufSheet.Shapes("Drop Down 6").ControlFormat.Value), " points", "")
Call checkEligibility
If Prize > (40 - total) Then
MsgBox "Please review points awarded. This year you have already awarded " & total & " Points. Maximum yearly limit is 40 Points.", vbCritical, "REWARD AND RECOGNITION TOOL"
total = 0
Exit Sub
End If
End If
End If

'-----------------------Append Query-------------------------------
Call structuredQueryLanguage("Table1", "appendQuery")
With RecSet
.AddNew
RecSet!Date = Date
    If ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "COMMITMENT/ATTITUDE" Then
RecSet![Approved] = "Approved"
End If
RecSet![Group_nomination] = 0
RecSet![Nominated By] = Application.UserName
RecSet![Nominee's Name] = ole.Item(11).Object.text
RecSet!Position = ole.Item(2).Object.text
RecSet!Band = ole.Item(1).Object.text
RecSet!category = ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value)
RecSet!Prize = ufSheet.Shapes("Drop Down 6").ControlFormat.List(ufSheet.Shapes("Drop Down 6").ControlFormat.Value)
RecSet![What do you nominate for?] = ole.Item(9).Object.text
RecSet![Comments on Efficiency(Performance) or Saving(CI)] = ole.Item(4).Object.text
RecSet![Additional assignments] = ole.Item(7).Object.text
RecSet![Feedback] = ole.Item(5).Object.text
RecSet![Effect/impact] = ole.Item(6).Object.text
RecSet![Email] = ole.Item(3).Object.text
RecSet![SCCP_number] = ole.Item(10).Object.text
'RecSet![Business Relation] = ole.Item(4).Object.text
On Error GoTo errorHandler:
.update '-------------------- stores the new record---------------------------
On Error GoTo 0
End With
RecSet.Close
Set RecSet = Nothing
cn.Close
Set cn = Nothing
End If

''---------------------------Send Email to Nominee, only in case of (commitment/attitude) category.------------------------

If ufSheet.OLEObjects("Checkbox1").Object.Value = False Then
If ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "COMMITMENT/ATTITUDE" Then
 With OutlookMail
    ' .CC = "lukasz.sosulski@contractor.axaxl.com"
    .CC = Application.UserName
    .To = ole.Item(3).Object.text
    .Subject = "Nomination in Commitment/Attitude Category."
    .BodyFormat = olFormatHTML
    .Body = "Dear " & ole.Item(11).Object.text & "," _
    & vbNewLine & vbNewLine & "Happy to announce that you have been awarded by your colleague " & Application.UserName & " in category COMMITMENT/ATTITUDE. You received " _
    & ufSheet.Shapes("Drop Down 6").ControlFormat.List(ufSheet.Shapes("Drop Down 6").ControlFormat.Value) & "." & _
   vbNewLine & "You were appreciated for: " & ole.Item(9).Object.text _
   & vbNewLine & "Please note that total amount of your points and prize catalog can be checked in the R&R Tool." _
   & vbNewLine & vbNewLine & vbNewLine & "Congratulations!" & vbNewLine & vbNewLine & _
  "Best regards," & vbNewLine & _
 "Rewards and Recognition team"
 ''  Application.UserName
 
   ' & " from my team for " & UFSheet.Shapes("Drop Down 5").ControlFormat.List(UFSheet.Shapes("Drop Down 5").ControlFormat.Value) & vbNewLine & _
vbNewLine & "Thanks & Regards" & vbNewLine
  .Display
 End With
 ElseIf ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "CI" Or ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "WOW Effect" Then
 
 ''---------------------------Send Emai to HR ppl, in CI/WOW category.------------------------
  With OutlookMail
     .CC = ActiveWorkbook.Worksheets("Sheet2").Range("b7")
     .To = Application.UserName
    ''.CC = ole.Item(4).Object.text
    .Subject = "Nomination to be considered. Please approve."
    .BodyFormat = olFormatHTML
    .Body = "Dear " & Application.UserName & "," & vbNewLine & vbNewLine & _
    "Please be informed that your nomination for " & ole.Item(11).Object.text & " in the category " & ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) & " has been submitted." & vbNewLine & _
    "All nominations will be discussed by local management team and If the prize is granted, you will be notified in a separate email." _
    & vbNewLine & vbNewLine & _
     "Best regards," & vbNewLine & _
    "Rewards and Recognition team"
    '''''& vbNewLine & "* or not, we need to discuss much more about this." _

''    .Body = "Hi Team" & vbNewLine & vbNewLine & "I would like to nominate " & ole.Item(9).Object.text _
''    & " for " & UFSheet.Shapes("Drop Down 5").ControlFormat.List(UFSheet.Shapes("Drop Down 5").ControlFormat.Value) & " category." & vbNewLine & _
''    vbNewLine & "Thanks & Regards" & vbNewLine _
''    & Application.UserName
    
    If ole.Item(8).Object.text <> "" Then
    On Error GoTo attachmentError
    .Attachments.Add ole.Item(8).Object.text
    On Error GoTo 0
    End If
    
  .Display
 End With
 
End If
End If



''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "Tool"
Set ole = Nothing
Set ufSheet = Nothing
Set OutlookApp = Nothing
Set OutlookMail = Nothing
Exit Sub
attachmentError:
MsgBox "Invalid attachment path.", vbCritical, "REWARD AND RECOGNITION TOOL"
Resume Next
errorHandler:
MsgBox "Cannot update database!", vbCritical, "REWARD AND RECOGNITION TOOL"
Resume Next
End Sub

Sub viewDatabase() ''My nominations
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim ufSheet As Worksheet
Dim query As String
Dim userSheet As Worksheet
''-------------Set Object References----------------------------------------
Set ufSheet = ThisWorkbook.Worksheets("Userform")
Set userSheet = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
On Error GoTo Handler:
''userSheet.Name = Application.UserName
userSheet.Name = "MyNominations"
On Error GoTo 0
selectQueryMyNominations = "SELECT *" & _
                            "FROM Table1 " & _
                            "WHERE [Nominated By] = '" & Application.UserName & "' " & _
                            "ORDER BY Date"
Call structuredQueryLanguage("Table1", "selectQuery", selectQueryMyNominations)
colNum = 1
For q = 0 To RecSet.Fields.Count - 1
userSheet.Cells(1, colNum) = RecSet.Fields(q).Name
colNum = colNum + 1
Next q
''----------------Get database records ------------------------------
userSheet.Range("A2").CopyFromRecordset RecSet
''-------------sheet formatting------------------------------
''userSheet.Range("B:B").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Approved,Reject,Pending"
userSheet.Range("D:D").NumberFormat = "MM/DD/YYYY"
userSheet.Range("A1:Q1").Interior.ColorIndex = 40
userSheet.Columns("A:Q").AutoFit
userSheet.Columns("A:Q").Locked = True
userSheet.Protect Password:="IloveLukasz", UserInterFaceOnly:=True

userSheet.UsedRange.Borders.LineStyle = xlContinuous
RecSet.Close
cn.Close
Set RecSet = Nothing
Set cn = Nothing
Set objcmd = Nothing
''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
'MsgBox "Done !", vbInformation, "REWARD AND RECOGNITION TOOL" ''' On Natalia Kolodziej mail 2020.06.18
Set ufSheet = Nothing
Set userSheet = Nothing
Exit Sub
ErrHandler:
Call errorHandler
Exit Sub
Handler:
userSheet.Delete
MsgBox "Worksheet (" & Application.UserName & ") already exist.", vbInformation, "Rewards and Recognition Tool."
End Sub

Sub allNominations()
Dim query As String, managerWiseData As managerWiseData
Set managerWiseData = New managerWiseData
If Not IsError(Application.Match(Environ("Username"), Sheets("Sheet2").Range("I:I"), 0)) Then
distinctAllNominations = "select distinct [Nominated By] from Table1 where Approved = 'Pending' or Approved is null" '''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call structuredQueryLanguage("Table1", "selectQuery", distinctAllNominations)
Do While Not RecSet.EOF
managerWiseData.ComboBox1.AddItem RecSet(0)
RecSet.MoveNext
Loop
managerWiseData.Show
Else
MsgBox "You are not allowed to use this function"
End If
End Sub

Sub allCommitments()
    Dim query As String, managerWiseData As managerWiseData
    'Set managerWiseData = New managerWiseData
    'distinctAllNominations = "select distinct [Nominated By] from Table1"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
 
    Dim userSheet As Worksheet
    
    Set userSheet = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    On Error GoTo Handler:
    userSheet.Name = "Commitments"
    On Error GoTo 0
    selectQueryAllNominations = "select [Unique Key], [Date], [Nominated By], [Nominee's Name], [category], [Prize], [What do you nominate for?],[Feedback]  from Table1 where [Category] ='COMMITMENT/ATTITUDE'"
    'selectQueryAllNominations = "select * from Table1 where [Category] ='COMMITMENT/ATTITUDE'"
    Call structuredQueryLanguage("Table1", "selectQuery", selectQueryAllNominations)
    
    ''-----------------Get database fields name in row 1-------------------
    colNum = 1
    For q = 0 To RecSet.Fields.Count - 1
    userSheet.Cells(1, colNum) = RecSet.Fields(q).Name
    colNum = colNum + 1
    Next q
    ''----------------Get database records ------------------------------
    userSheet.Range("A2").CopyFromRecordset RecSet
    
    ''-------------sheet formatting------------------------------
    ''userSheet.Range("B:B").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Approve,Reject,Pending"
    userSheet.Range("B:B").NumberFormat = "MM/DD/YYYY"
    userSheet.Range("A1:H1").Interior.ColorIndex = 40
    userSheet.Columns("A:H").AutoFit
    userSheet.UsedRange.Borders.LineStyle = xlContinuous
    ''MsgBox "Done!", vbInformation, "Rewards and Recognition Tool" ''' On Natalia Kolodziej mail 2020.06.18
    
    ''-------------The End------------------------
    Set userSheet = Nothing
    Set RecSet = Nothing
    Set objcmd = Nothing
    Set cn = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
Handler:
    userSheet.Delete
    'MsgBox "Worksheet (" & Me.ComboBox1.Value & ") already exist.", vbInformation, "Rewards and Recognition Tool."
End Sub

Sub update()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

''i added here that non everybody can change this.
''if Environ("Username")
If Not IsError(Application.Match(Environ("Username"), Sheets("Sheet2").Range("I:I"), 0)) Then
   'The value present in that range
  '' MsgBox Environ("Username")

''----------variable declaration----------
Dim getSheetName As getSheetName
''---------set object reference-----------
Set getSheetName = New getSheetName
getSheetName.Show

''---------------The End--------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Else
MsgBox "You are not allowed to use this function"
End If
End Sub


Sub ViewPoints()
    Dim mp_bottom As Integer
    Dim rew_bottom As Integer
    
    Dim query As String, managerWiseData As managerWiseData
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim userSheet As Worksheet
    
    Set userSheet = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    On Error GoTo Handler:
    userSheet.Name = "MyPoints"
    On Error GoTo 0
    selectQueryAllNominations = "select * from Table1 where [Prize] like '%points%' and [Nominee's Name]='" & Application.UserName & "' Order by Date"
    
    Call structuredQueryLanguage("Table1", "selectQuery", selectQueryAllNominations)
    
    ''-----------------Get database fields name in row 1-------------------
    colNum = 1
    For q = 0 To RecSet.Fields.Count - 1
    userSheet.Cells(1, colNum) = RecSet.Fields(q).Name
    colNum = colNum + 1
    Next q
    ''----------------Get database records ------------------------------
    userSheet.Range("A2").CopyFromRecordset RecSet
    
    ''-------------sheet formatting------------------------------
    ''userSheet.Range("B:B").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Approve,Reject,Pending"
    userSheet.Range("D:D").NumberFormat = "MM/DD/YYYY"
    userSheet.Range("A1:Q1").Interior.ColorIndex = 40
  ''xxx (!)
    userSheet.UsedRange.Borders.LineStyle = xlContinuous
     ''' adding rewards table
    mp_bottom = Cells(Rows.Count, 1).End(xlUp).Row
    
    ThisWorkbook.Worksheets("Rewards").Activate
    rew_bottom = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(rew_bottom, 2)).Copy ThisWorkbook.Worksheets("MyPoints").Cells(mp_bottom + 3, 1)

    ThisWorkbook.Worksheets("MyPoints").Activate
    userSheet.Columns("A:Q").AutoFit ''' it was stolen from xxx (!)
    userSheet.Range(Cells(mp_bottom + 3, 1), Cells(mp_bottom + 3, 2)).Interior.ColorIndex = 40
    '''MsgBox "Done!", vbInformation, "Rewards and Recognition Tool" ''' On Natalia Kolodziej mail 2020.06.18
    
    ''-------------The End------------------------
    Set userSheet = Nothing
    Set RecSet = Nothing
    Set objcmd = Nothing
    Set cn = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
Handler:
    userSheet.Delete
    'MsgBox "Worksheet (" & Me.ComboBox1.Value & ") already exist.", vbInformation, "Rewards and Recognition Tool."
End Sub


Sub attachment()
Dim ole As OLEObjects
Set ole = ThisWorkbook.Worksheets("Userform").OLEObjects
With Application.FileDialog(msoFileDialogFilePicker)
.Title = "Select file to be attached to EMail."
    If .Show <> 0 Then
    ole.Item(8).Object.text = .SelectedItems(1)
    End If
End With
Set ole = Nothing
End Sub

''=================Lukasz, this is for you=========================
Sub checkEligibility()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim ufSheet As Worksheet
Dim point As Variant
Set ufSheet = ThisWorkbook.Worksheets("Userform")
Set ole = ufSheet.OLEObjects

selectQueryAwardedPoints = "select Prize from Table1 where year(Date) = '" & Year(Date) & "' and [Category] = 'COMMITMENT/ATTITUDE' and [Nominated By]='" & Application.UserName & "' Order by Date"
    
Call structuredQueryLanguage("Table1", "selectQuery", selectQueryAwardedPoints)

Do While Not RecSet.EOF
point = Replace(RecSet(0), " points", "")
total = total + point
RecSet.MoveNext
Loop

Set ole = Nothing
Set userSheet = Nothing
Set RecSet = Nothing
Set objcmd = Nothing
Set cn = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub



