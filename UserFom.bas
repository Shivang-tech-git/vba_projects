Attribute VB_Name = "UserFom"
Sub Button5()
Dim ws As Worksheet
Dim category As Object
Dim Prize As Shape
Dim label(1 To 6) As String
Dim text(1 To 6) As String
'On Error Resume Next
Set ws = ActiveSheet
Set category = ws.DropDowns("DropDown5")
Set Prize = ws.Shapes("Drop Down 6")

label(1) = "TextBox 3" ''position
label(2) = "TextBox 27" ''band
label(3) = "TextBox 71"  ''comments
label(4) = "TextBox 63" ''additional assigments
label(5) = "TextBox 68" ''effect impact
text(1) = "TextBox5 ''position"
text(2) = "TextBox4" ''band
text(3) = "TextBox7" ''comments
text(4) = "TextBox10" ''additional assigments
text(5) = "TextBox9" ''effect impact
Select Case category.List(category.ListIndex)

Case "COMMITMENT/ATTITUDE"

  ''ws.OLEObjects.Item(text(2)).Select

 
    With Prize.ControlFormat
        .RemoveAllItems
        .AddItem "10 points"
        .AddItem "20 points"
        .AddItem "30 points"
        .AddItem "40 points"
   End With
 
  
   '-------------Hide all -------------------
   For I = 1 To 5
   ws.OLEObjects.Item(text(I)).Visible = False
   ws.TextBoxes(label(I)).Visible = False
   Next I
   
   ''''---------------- SSCC number
    ws.TextBoxes("TextBox 41").Visible = False
    ws.OLEObjects.Item(10).Visible = False
    '''--- prize !
    ws.TextBoxes("TextBox 21").Visible = False
       ws.TextBoxes("TextBox 44").Visible = True
       
       
''   ws.TextBoxes("TextBox 58").Visible = True
''   ws.OLEObjects.Item(text(1)).Visible = True
''   ws.TextBoxes("TextBox 65").Visible = False
''   ws.TextBoxes("TextBox 38").Visible = True
''    ws.OLEObjects.Item(7).Visible = False
''      ws.OLEObjects.Item(4).Visible = True
''
''
''     '----Business relation----------
''  ' ws.TextBoxes("TextBox 23").Visible = False
'' '  ws.OLEObjects.Item(4).Visible = False
''    '----Band----------
''   ws.TextBoxes("TextBox 27").Visible = False
''   ws.OLEObjects.Item(2).Visible = False
''   '----Position----------
''   ws.TextBoxes("TextBox 3").Visible = False
''   ws.OLEObjects.Item(3).Visible = False
''
   
   
   
   Case "CI"

    With Prize.ControlFormat
        .RemoveAllItems
        .AddItem "300 points"
        .AddItem "1000 pln"
    End With

   For I = 1 To 5
   ws.OLEObjects.Item(text(I)).Visible = True
   ws.TextBoxes(label(I)).Visible = True
   Next I

       
       
    ''''---------------- SSCC number
    ws.TextBoxes("TextBox 41").Visible = True
    ws.OLEObjects.Item(10).Visible = True
       
       ''------ prize
     ws.TextBoxes("TextBox 21").Visible = True
      ws.TextBoxes("TextBox 44").Visible = False
          
    '-------Data and any necessary comments------------
   ws.TextBoxes("TextBox 71").Caption = "Data and any necessary comments on saving (time saving in FTE, quality improvement, customer satisfaction increase etc) 1 FTE = 115 hours"

       
Case "WOW Effect"
    
    With Prize.ControlFormat
        .RemoveAllItems
        .AddItem "300 points"
        .AddItem "1000 pln"
    End With
   
   For I = 1 To 5
   ws.OLEObjects.Item(text(I)).Visible = True
   ws.TextBoxes(label(I)).Visible = True
   Next I
     ''''---------------- SSCC number
    ws.TextBoxes("TextBox 41").Visible = False
    ws.OLEObjects.Item(10).Visible = False
    ''------ prize
     ws.TextBoxes("TextBox 21").Visible = True
      ws.TextBoxes("TextBox 44").Visible = False
    '-------Data and any necessary comments------------
   ws.TextBoxes("TextBox 71").Caption = "Data and any necessary comments on Efficiency (Volumes, Delivery Time) & Quality (in relation to team's performance, please state period of presented data"
   
   
End Select
End Sub

