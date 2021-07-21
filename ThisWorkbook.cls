VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
Dim category As Shape

Set ws = ActiveWorkbook.Worksheets("Userform")
Set category = ws.Shapes("DropDown5")
ws.Activate
ActiveSheet.Protect Password:="iloveshivang", UserInterFaceOnly:=True

If IsError(Application.Match(Environ("Username"), Sheets("Sheet2").Range("K:K"), 0)) Then  '''--- check who can open workbook
MsgBox "You are nor allowed to use this Tool. For more information contact Natalia Kolodziej."
ActiveWorkbook.Close savechanges:=False

End If

If Not IsError(Application.Match(Environ("Username"), Sheets("Sheet2").Range("J:J"), 0)) Then
With category.ControlFormat
    .RemoveAllItems
        .AddItem "COMMITMENT/ATTITUDE"
        .AddItem "CI"
        .AddItem "WOW Effect"
        End With
        
 ThisWorkbook.Worksheets("Userform").OLEObjects("Checkbox1").Visible = True
Else
'' only commitment
With category.ControlFormat
        .RemoveAllItems
        .AddItem "COMMITMENT/ATTITUDE"
End With
   ThisWorkbook.Worksheets("Userform").OLEObjects("Checkbox1").Visible = False
End If
End Sub

