VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} getSheetName 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   10
   ClientTop       =   10
   ClientWidth     =   3480
   OleObjectBlob   =   "getSheetName.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "getSheetName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
Dim userSheet As Worksheet
Dim dol As Integer
Dim OutlookApp As Outlook.Application
Dim OutlookMail As Outlook.MailItem


Set OutlookApp = New Outlook.Application
Set OutlookMail = OutlookApp.CreateItem(olMailItem)
If Me.ComboBox1.Value = "" Then
MsgBox "Please select worksheet name for which you want to update database.", vbInformation, "Rewards and Recognition Tool"
Else
Set userSheet = ThisWorkbook.Worksheets(Me.ComboBox1.Value)
''---------------------get recordset----------------------
Call structuredQueryLanguage("Table1", "appendQuery")
''--------------Loop through each unique key to update database--------------------
rowInc = 1
Do While Not userSheet.Cells(1, 1).Offset(rowInc, 0).Value = ""
RecSet.Filter = "[Unique Key] = '" & userSheet.Cells(1, 1).Offset(rowInc, 0).Value & "'"
If RecSet.EOF Then
MsgBox "Unique Key in row " & rowInc + 1 & " does not exist in database.", vbCritical, "Rewards and Recognition Tool"
Else
RecSet!Approved = userSheet.Cells(1, 1).Offset(rowInc, 1).Value
End If
RecSet.update
rowInc = rowInc + 1
Loop
RecSet.Close
Set RecSet = Nothing
cn.Close
Set cn = Nothing
''----------------Delete sheet after updating---------------------------

Me.Hide

ThisWorkbook.Worksheets(Me.ComboBox1.Value).Activate
dol = Cells(Rows.Count, 1).End(xlUp).Row
''' loop for sending emails
If dol <> 1 Then
        For I = 2 To dol
            cc_value = Cells(I, 4).Value
            who_value = Cells(I, 5).Value
            cat_value = Cells(I, 8).Value
            prize_value = Cells(I, 9).Value
            
            Set OutlookMail = Outlook.CreateItem(olMailItem)
            
            If Cells(I, 2).Value = "Approve" Then
            
      
             With OutlookMail
                     .To = cc_value
                    .Subject = "Your nomination has been approved."
                    .BodyFormat = olFormatHTML
                    .Body = "Dear " & Cells(I, 4) & "," & vbNewLine & vbNewLine & _
                    "Please be informed that your nomination for " & who_value & " in the category " & cat_value & " has been accepted and the award has been granted." & vbNewLine & _
                    "Prize is: " & prize_value & ". Please note that total number of points and prize catalog can be checked in the R&R Tool." _
                     & vbNewLine & vbNewLine & _
                     "Best regards," & vbNewLine & _
                    "Rewards and Recognition team"
                    
                  .Display
                  .Save
            End With
            
        ElseIf Cells(I, 2).Value = "Reject" Then
         With OutlookMail
                     .To = cc_value
                    .Subject = "Your nomination has been rejected."
                    .BodyFormat = olFormatHTML
                    .Body = "Dear " & Cells(I, 4) & "," & vbNewLine & vbNewLine & _
                    "Please be informed that your nomination for " & who_value & " in the category " & cat_value & " has been rejected." & vbNewLine & _
                    "More details can be provided by your line manager." _
                    & vbNewLine & vbNewLine & _
                     "Best regards," & vbNewLine & _
                    "Rewards and Recognition team"
                    
                  .Display
                  .Save
            End With
        
        End If
        Set OutlookMail = Nothing
    Next I
End If

userSheet.Delete
MsgBox "Done!", vbInformation, "Rewards and Recognition Tool"
End If
End Sub


Sub UserForm_Initialize()
Me.Caption = "Update pending approvals:"
Me.BackColor = RGB(255, 204, 153)

Me.CommandButton1.Caption = "UPDATE"
Me.CommandButton1.BackColor = RGB(204, 153, 255)

Me.Label1.Caption = "Select worksheet name:"
Me.Label1.BackColor = RGB(255, 204, 153)

For I = 1 To ThisWorkbook.Worksheets.Count
If ThisWorkbook.Worksheets(I).Name <> "Userform" _
    And ThisWorkbook.Worksheets(I).Name <> "Sheet2" _
     And ThisWorkbook.Worksheets(I).Name <> "Commitments" _
     And ThisWorkbook.Worksheets(I).Name <> "MyNominations" _
    And ThisWorkbook.Worksheets(I).Name <> "Rewards" Then
Me.ComboBox1.AddItem ThisWorkbook.Worksheets(I).Name
End If
Next I

Me.StartUpPosition = 1
''--------------------keep code running after displaying userform---------
End Sub



'Private Sub CommandButton1_Click()
'Application.ScreenUpdating = False
'Application.DisplayAlerts = False
'''---------------Error Handling Routine-----------------------------
'''on error resume next
'''-------------Variable Declarations---------------------------------------
'Dim UFSheet As Worksheet, ole As OLEObjects, query As String
'Dim sheet2 As Worksheet
'''-------------Set Object References----------------------------------------
'Set UFSheet = ThisWorkbook.Worksheets("Userform")
'Set sheet2 = ThisWorkbook.Worksheets("Sheet2")
'Set ole = UFSheet.OLEObjects
''-----------------------Clear userform-----------------------
'Call clearUserForm
''--------------------- connect to the Access database------------------
'query = "SELECT * " & _
'        "FROM Table1 " & _
'        "WHERE [Unique Key] = " & Me.TextBox2.text & ";"
'
'Call structuredQueryLanguage("Table1", "selectQuery", query)
'' ------------------Fill up the userform from access---------------------
''' category feild
'For x = 1 To UFSheet.Shapes("drop down 5").ControlFormat.ListCount
'If UFSheet.Shapes("Drop Down 5").ControlFormat.List(x) = RecSet("Category").Value Then
'UFSheet.Shapes("Drop Down 5").ControlFormat.ListIndex = x
'End If
'Next x
''' Prize feild
'For x = 1 To UFSheet.Shapes("drop down 6").ControlFormat.ListCount
'If UFSheet.Shapes("Drop Down 6").ControlFormat.List(x) = RecSet("Prize").Value Then
'UFSheet.Shapes("Drop Down 6").ControlFormat.ListIndex = x
'End If
'Next x
'ole.Item(1).Object.text = RecSet("Nominee's Name").Value
'ole.Item(5).Object.text = RecSet("Position")
'ole.Item(3).Object.text = RecSet("Band")
'ole.Item(2).Object.text = RecSet("What do you nominate for?")
'ole.Item(9).Object.text = RecSet("Data and any necessary comments on Efficiency")
'ole.Item(6).Object.text = RecSet("Additional assignments")
'ole.Item(7).Object.text = RecSet("Stakeholder's feedback")
'ole.Item(8).Object.text = RecSet("Effect/impact")
'ole.Item(4).Object.text = RecSet("Business Relation")
'RecSet.Close
''-----------------------------------------------------------------
'cn.Close
'Set RecSet = Nothing
'Set cn = Nothing
'Set objcmd = Nothing
'Exit Sub
'ErrHandler:
'    'clean up
'    If RecSet.State = adStateOpen Then
'        RecSet.Close
'    End If
'
'    If cn.State = adStateOpen Then
'        cn.Close
'    End If
'
'    Set RecSet = Nothing
'    Set cn = Nothing
'    Set objcmd = Nothing
'
'    If Err <> 0 Then
'        MsgBox Err.Source & "-->" & Err.Description, , "Error"
'    End If
''EndBasicCmd
'''-------------The End------------------------------------------------------
'Application.DisplayAlerts = True
'Application.ScreenUpdating = True
'Set ole = Nothing
'Set UFSheet = Nothing
'Set sheet2 = Nothing
'End Sub


