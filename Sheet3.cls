VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''===============================================================
''Lukasz my friend, below code is written by shivang, only for you :)
''===============================================================
''-----------CLICK ON GROUP NOMINATION CHECK BOX----------------
Private Sub checkbox1_Click()
Dim category As Shape
Set ws = ActiveWorkbook.Worksheets("Userform")
Set category = ws.Shapes("DropDown5")
'ActiveSheet.Protect Password:="iloveshivang", UserInterFaceOnly:=True

If Me.CheckBox1.Value = True Then
Me.CommandButton1.Visible = True
Me.Shapes("Drop Down 43").Visible = True
Me.Shapes("Drop Down 44").Visible = True
Me.CommandButton2.Visible = True
Me.Shapes("Rectangle: Rounded Corners 1").Visible = msoTrue
With category.ControlFormat
        .RemoveAllItems
        .AddItem "CI"
        .AddItem "WOW Effect"
        End With
Else
Me.CommandButton1.Visible = False
Me.Shapes("Drop Down 43").Visible = False
Me.Shapes("Drop Down 44").Visible = False
Me.CommandButton2.Visible = False
Me.Shapes("Rectangle: Rounded Corners 1").Visible = msoFalse
Me.Shapes("Drop Down 43").ControlFormat.RemoveAllItems
Me.Shapes("Drop Down 44").ControlFormat.RemoveAllItems
With category.ControlFormat
    .RemoveAllItems
        .AddItem "COMMITMENT/ATTITUDE"
        .AddItem "CI"
        .AddItem "WOW Effect"
        End With

End If


End Sub

''------------CLICK ON ADD NEW BUTTON---------------
Private Sub CommandButton1_Click()

For dropDownValue = 1 To Me.Shapes("Drop Down 43").ControlFormat.ListCount
If Me.Shapes("Drop Down 43").ControlFormat.List(dropDownValue) = Me.ComboBox2.Value Then
MsgBox Me.ComboBox2.Value & " already exist.", vbInformation, "RR Tool"
Exit Sub
End If
Next dropDownValue

Me.Shapes("Drop Down 43").ControlFormat.AddItem Me.ComboBox2.Value
Me.Shapes("Drop Down 44").ControlFormat.AddItem Me.TextBox6.Value
Me.ComboBox2.Value = ""
Me.TextBox6.Value = ""
End Sub

''------------CLICK ON REMOVE BUTTON----------------
Private Sub CommandButton2_Click()
On Error Resume Next
Me.Shapes("Drop Down 43").ControlFormat.RemoveItem Me.Shapes("Drop Down 43").ControlFormat.Value
Me.Shapes("Drop Down 44").ControlFormat.RemoveItem Me.Shapes("Drop Down 44").ControlFormat.Value
End Sub
''-------------Automatically select Email according to selected name-------------------
Sub selectRespectiveEmail()
a = ThisWorkbook.Worksheets("Userform").Shapes("Drop Down 43").ControlFormat.Value
ThisWorkbook.Worksheets("Userform").Shapes("Drop Down 44").ControlFormat.ListIndex = a
End Sub

'
'Private Sub ComboBox1_Change()
'Dim myrange As Range
'Dim wb As String
'wb = ActiveWorkbook.Name
'
'Set myrange = Workbooks(wb).Worksheets("Sheet2").Range("M:N")
'On Error Resume Next
'If ComboBox1 = "" Then
'TextBox6.Value = ""
'ComboBox1.DropDown
'
'Else
'TextBox6.Value = WorksheetFunction.VLookup(ComboBox1, myrange, 2, False)
'
'    If Err.Number <> 0 Then
'    TextBox3.Value = "not found"
'    End If
'
'End If
'End Sub
'

'Private Sub ComboBox1_Change()
'On Error Resume Next
'If ComboBox1 = "" Then
'
'ComboBox1.DropDown
'
'End If
'End Sub



Private Sub ComboBox2_GotFocus()
''ActiveSheet.Protect UserInterFaceOnly:=True
    ComboBox2.ListFillRange = "DropDownList"
    Me.ComboBox2.DropDown

End Sub


Private Sub ComboBox2_Change()
Dim myrange As Range
Dim wb As String
wb = ThisWorkbook.Name
''ActiveSheet.Protect UserInterFaceOnly:=True
Set myrange = Workbooks(wb).Worksheets("Sheet2").Range("M:N")
On Error Resume Next
If ComboBox2 = "" Then
TextBox6.Value = ""

Else
TextBox6.Value = WorksheetFunction.VLookup(ComboBox2, myrange, 2, False)

    If Err.Number <> 0 Then
    TextBox6.Value = "not found"
    End If

End If
End Sub


''ActiveSheet.Protect UserInterFaceOnly:=True


