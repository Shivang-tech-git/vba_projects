Attribute VB_Name = "GroupNomination"

Sub GroupNomination_macro() ''''(***********GOING HERE )

Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Error Handling Routine-----------------------------
''on error resume next
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer, getOut As Boolean
Dim ufSheet As Worksheet, ole As OLEObjects
Dim OutlookApp As Outlook.Application
Dim OutlookMail As Outlook.MailItem
Dim Amount_of_names As Integer
'Dim Names(1 To 100)


''-------------Set Object References----------------------------------------
Set ufSheet = ThisWorkbook.Worksheets("Userform")
Set ole = ufSheet.OLEObjects
Set OutlookApp = New Outlook.Application
Set OutlookMail = OutlookApp.CreateItem(olMailItem)
''----------------Checking Mandatory fields------------------------
getOut = False


'If ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) = "COMMITMENT/ATTITUDE" Then
' If ufSheet.Shapes("Drop Down 43").ControlFormat.ListCount = 0 Then
'        MsgBox "Please provide Nominee's Name.", vbInformation, "Rewards and Recognition Tool"
'        getOut = True
'      ElseIf ole.Item(9).Object.text = "" Then
'        MsgBox "Please don't leave 'What do you nominate for?' field empty!", vbInformation, "Rewards and Recognition Tool"
'        getOut = True
'        ElseIf ufSheet.Shapes("Drop Down 6").ControlFormat.Value = 0 Then
'           MsgBox "Please select prize", vbInformation, "Rewards and Recognition Tool"
'            getOut = True
'            End If
'        If getOut = True Then Exit Sub
'
'
'      Else
        If ufSheet.Shapes("Drop Down 43").ControlFormat.ListCount = 0 Then
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
'   End If
   
   If getOut = True Then Exit Sub
   
'''''''''''''-

  Amount_of_names = ufSheet.Shapes("Drop Down 43").ControlFormat.ListCount
all_names = ""
For Z = 1 To Amount_of_names
If all_names = "" Then
all_names = ufSheet.Shapes("Drop Down 43").ControlFormat.List(Z)
Else
all_names = all_names & "; " & ufSheet.Shapes("Drop Down 43").ControlFormat.List(Z)
End If
  Next Z
  
 ' MsgBox Namess(I)
 ' MsgBox Emailss(I)
 '----- look for the max value from access (!) Group_Nomination

selectQueryMaxGroupNomination = "select max(Group_nomination) as MaxNomination from Table1"

Call structuredQueryLanguage("Table1", "selectQuery", selectQueryMaxGroupNomination)
    maxvalue = RecSet(0)
    insertValue = maxvalue + 1
'-----------------------Append Query-------------------------------
 For I = 1 To Amount_of_names
Call structuredQueryLanguage("Table1", "appendQuery")

 With RecSet
.AddNew

   Namess = ufSheet.Shapes("Drop Down 43").ControlFormat.List
   Emailss = ufSheet.Shapes("Drop Down 44").ControlFormat.List
name_to_database = Namess(I)
email_to_database = Emailss(I)
'prizee = ufSheet.Shapes("Drop Down 6").ControlFormat.List(ufSheet.Shapes("Drop Down 6").ControlFormat.Value) / 2
If ufSheet.Shapes("Drop Down 6").ControlFormat.List(ufSheet.Shapes("Drop Down 6").ControlFormat.Value) = "1000 pln" Then
prizee = Round(1000 / Amount_of_names, 2) & " pln"
ElseIf ufSheet.Shapes("Drop Down 6").ControlFormat.List(ufSheet.Shapes("Drop Down 6").ControlFormat.Value) = "300 points" Then
prizee = Round(300 / Amount_of_names, 2) & " points"
End If


RecSet!Group_nomination = insertValue
RecSet!Date = Date
RecSet![Nominated By] = Application.UserName
RecSet![Nominee's Name] = name_to_database
RecSet!Position = ole.Item(2).Object.text

RecSet!Band = ole.Item(1).Object.text
RecSet!category = ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value)
RecSet!Prize = prizee
RecSet![What do you nominate for?] = ole.Item(9).Object.text
RecSet![Comments on Efficiency(Performance) or Saving(CI)] = ole.Item(4).Object.text
RecSet![Additional assignments] = ole.Item(7).Object.text
RecSet![Feedback] = ole.Item(5).Object.text
RecSet![Effect/impact] = ole.Item(6).Object.text
RecSet![Email] = email_to_database
RecSet![SCCP_number] = ole.Item(10).Object.text
'RecSet![Business Relation] = ole.Item(4).Object.text
'On Error GoTo errorHandler:
.update '-------------------- stores the new record---------------------------
    

On Error GoTo 0
End With
RecSet.Close
Set RecSet = Nothing
cn.Close
Set cn = Nothing
Next I


With OutlookMail

     .CC = "tool_administrator"
     .To = Application.UserName
    .Subject = "Nomination to be considered. Please approve."
    .BodyFormat = olFormatHTML
    .Body = "Dear " & Application.UserName & "," & vbNewLine & vbNewLine & _
    "Please be informed that your nomination for: " & all_names & " in the category " & ufSheet.Shapes("Drop Down 5").ControlFormat.List(ufSheet.Shapes("Drop Down 5").ControlFormat.Value) & " has been submitted." & vbNewLine & _
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
    'On Error GoTo attachmentError
    .Attachments.Add ole.Item(8).Object.text
    On Error GoTo 0
 End If
    
  .Display
 End With


End Sub
'Me.ComboBox1.ListCount
