VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} managerWiseData 
   Caption         =   "UserForm1"
   ClientHeight    =   3510
   ClientLeft      =   -40
   ClientTop       =   -100
   ClientWidth     =   2760
   OleObjectBlob   =   "managerWiseData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "managerWiseData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click() ''All Nominations
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim userSheet As Worksheet
If Me.ComboBox1.Value <> "" Then
Set userSheet = ThisWorkbook.Worksheets.Add(, ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
On Error GoTo Handler:
userSheet.Name = Me.ComboBox1.Value
On Error GoTo 0
selectQueryAllNominations = "select * from Table1 where [Nominated By] = '" & Me.ComboBox1.Value & "' and (Approved = 'Pending' or Approved is null) and (Category<>'COMMITMENT/ATTITUDE')"

Call structuredQueryLanguage("Table1", "selectQuery", selectQueryAllNominations)

''-----------------Get database fields name in row 1-------------------
colNum = 1
For q = 0 To RecSet.Fields.Count - 1
userSheet.Cells(1, colNum) = RecSet.Fields(q).Name
colNum = colNum + 1
Next q
''----------------Get database records ------------------------------
userSheet.Range("A2").CopyFromRecordset RecSet
Me.Hide
''-------------sheet formatting------------------------------
userSheet.Range("C:C").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Approved,Rejected,Pending"
userSheet.Range("J:J").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="300 points,1000 pln"
userSheet.Range("D:D").NumberFormat = "MM/DD/YYYY"
userSheet.Range("A1:Q1").Interior.ColorIndex = 40
userSheet.Columns("A:Q").AutoFit
userSheet.UsedRange.Borders.LineStyle = xlContinuous
userSheet.Cells.Locked = False
userSheet.Range("A:B").Locked = True
userSheet.Range("D:I").Locked = True
userSheet.Range("K:Q").Locked = True
userSheet.Protect Password:="IloveLukasz", UserInterFaceOnly:=True

'''MsgBox "Done!", vbInformation, "Rewards and Recognition Tool" ''' On Natalia Kolodziej mail 2020.06.18
Else
MsgBox "Please select a name from Dropdown list.", vbInformation, "Rewards and Recognition Tool."
End If
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
MsgBox "Worksheet (" & Me.ComboBox1.Value & ") already exist.", vbInformation, "Rewards and Recognition Tool."
End Sub


Private Sub Label2_Click()

End Sub

Private Sub UserForm_Initialize()
Me.Caption = "All Nominations:"
Me.BackColor = RGB(255, 204, 153)

Me.Label1.Caption = "Select [Nominated By]:"
Me.Label1.BackColor = RGB(255, 204, 153)

Me.Label2.BackColor = RGB(255, 204, 153)
Me.Label2.Caption = "SHOW PENDING NOMINATIONS"
Me.Label2.TextAlign = fmTextAlignCenter
Me.CommandButton1.BackColor = RGB(204, 255, 204)
Me.CommandButton1.Caption = "SHOW"
End Sub
