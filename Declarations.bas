Attribute VB_Name = "Declarations"

''-------------Public Variable declaration----------------------------
Public selectQueryMyNominations As String
Public getRecordFromMyNominations As String
Public distinctAllNominations As String
Public selectQueryAllNominations As String
Public selectQueryAwardedPoints As String
Public selectQueryMaxGroupNomination As String


Public RecSet As ADODB.Recordset
Public cn As ADODB.Connection
Public objcmd As ADODB.Command

''------------This procedure connects to access database---------------

Public Sub structuredQueryLanguage(tableName, queryType As String, Optional sql As String)
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0; " & _
"Data Source=" & ThisWorkbook.Worksheets("Sheet2").Cells(2, 2).Value & ";MS Access;PWD=Gfjatusheq"
'Gfjatusheq
Select Case queryType
Case "appendQuery"
''------------------open record set----------------------
Set RecSet = New ADODB.Recordset
RecSet.Open tableName, cn, adOpenKeyset, adLockOptimistic, adCmdTable
''-------------------run select query------------------------
Case "selectQuery"
Set objcmd = New ADODB.Command
objcmd.CommandText = sql
objcmd.CommandType = adcmdtext
objcmd.ActiveConnection = cn
Set RecSet = objcmd.Execute
End Select
End Sub

''------------------Common error handler for all procedures-----------------

Public Sub errorHandler()
If RecSet.State = adStateOpen Then
RecSet.Close
End If
If cn.State = adStateOpen Then
cn.Close
End If
Set RecSet = Nothing
Set cn = Nothing
Set objcmd = Nothing
If Err <> 0 Then
MsgBox Err.Source & "-->" & Err.Description, , "Error"
End If
End Sub

''----------This procedure clears all textboxes in userform-----------

Public Sub clearUserForm()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim ufSheet As Worksheet, ole As OLEObjects
Dim query As String
''-------------Set Object References----------------------------------------
Set ufSheet = ThisWorkbook.Worksheets("Userform")
Set ole = ufSheet.OLEObjects
''-------------------------------------------------------------------------
For I = 1 To 12
If I = 12 Then
        ole.Item(12).Object.Value = False
    Else
        ole.Item(I).Object.text = ""
End If
Next I

For I = 1 To 12
If I = 12 Then
        ole.Item(12).Object.Value = False
    Else
            ole.Item(I).Object.text = ""

    End If
Next I
''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Set ole = Nothing
Set ufSheet = Nothing
End Sub








'''''''''''''''''''''''''''''
'For I = 1 To ole.Count
'    If I = 12 Then
'        ole.Item(12).Object.Value = False
'    Else
'        ole.Item(I).Object.text = ""
'    End If
'Next I
'
'For I = 1 To ole.Count
'    If ole.Item(I) = 12 Then
'            ole.Item(12) = False
'    Else
'            ole.Item(I).Object.text = ""
'    End If
'
'Next I
