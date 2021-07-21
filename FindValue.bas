Attribute VB_Name = "FindValue"
Sub extractCoding()
Attribute extractCoding.VB_ProcData.VB_Invoke_Func = "q\n14"
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'On Error Resume Next
''-------------Variable Declarations---------------------------------------
Dim filename As String, I As Integer
Dim macroSheet As Worksheet, extSheet As Worksheet, codingSheet As Worksheet, uaUwSheet As Worksheet
Dim colLetter As String, cellRef As String
Dim doColor As Boolean
doColor = True
''---------------column Names-------------------------------------------
ID = 2
Issued_By = 3
Trans_Type = 4
Status = 5
Presto_Status = 6
UW = 7
UA = 8
Assigned_To = 9
Genius_Policy = 10
Insured = 11
Eff_Date = 12
Exp_Date = 13
Bound_Date = 14
Invoice_Date = 15
Broker_Statement = 16
Date_to_India = 17
Date_to_Poland = 18
Authorized_Date = 19
Policy_Issue_Date = 20
Days_old = 21
Reason_for_Delay = 22
Broker_Name = 23
Broker_Code = 24
TRIA = 25
Commision = 26
Gross_Premium = 27
Surcharges_Taxes = 28
Coverage = 29
Sub_Line = 30
FAC = 31
Treaty = 32
Adjustable = 33
UA_Notes = 34
Poland_Notes = 35
Poland_Status = 36
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set extSheet = ThisWorkbook.Worksheets("Extraction")
Set uaUwSheet = ThisWorkbook.Worksheets("UA & UW Names")
''----------Open coding sheet ----------------------------------------
'filename = macroSheet.Range("D3").Value
With Application.FileDialog(msoFileDialogFilePicker)
    If .Show <> 0 Then
    filename = .SelectedItems(1)
        End If
End With
On Error GoTo errorHandlerPath
Workbooks.Open filename
On Error GoTo 0
Set codingSheet = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
With extSheet
extRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
''-------------no border and no color------------------------
.Cells.Borders.LineStyle = xlNone
.Rows(extRow & ":" & .Rows.Count).Interior.Color = xlNone
.Cells(extRow, 1) = "Policy Details"
.Cells(extRow, 1).Interior.Color = RGB(153, 204, 0)
End With
''------------Find columns in coding sheet-------------------
On Error Resume Next
With codingSheet
extSheet.Cells(extRow, Trans_Type) = .Cells.Find("Account Status").Offset(0, 2).Value
extSheet.Cells(extRow, Status) = "In Progress"
extSheet.Cells(extRow, Presto_Status) = "Not Added"
extSheet.Cells(extRow, UW) = .Range("AV:AV").Find("UW").Offset(0, 2).Value
extSheet.Cells(extRow, UA) = .Range("AV:AV").Find("UA").Offset(0, 2).Value
extSheet.Cells(extRow, Genius_Policy) = .Range("X:X").Find("Policy").Offset(1, 0).Value
extSheet.Cells(extRow, Insured) = .Range("F:F").Find("Insured Name").Offset(0, 2).Value
extSheet.Cells(extRow, Eff_Date) = .Range("F:F").Find("Inception Date").Offset(0, 2).Value
extSheet.Cells(extRow, Exp_Date) = .Range("AC:AC").Find("Expiry Date").Offset(0, 2).Value
extSheet.Cells(extRow, Broker_Statement) = .Cells.Find("Notes").Offset(1, 0).Value
extSheet.Cells(extRow, Broker_Name) = .Range("AA:AA").Find("Producer Name").Offset(0, 2).Value
extSheet.Cells(extRow, Broker_Code) = .Range("AA:AA").Find("Prod #").Offset(0, 2).Value
extSheet.Cells(extRow, TRIA) = .Range("BL" & codingSheet.Cells.Find("Dereg/NYFTZ").Offset(-3, 0).Row).Value
extSheet.Cells(extRow, Commision) = .Range("T:X").Find("commission").Offset(1, 0).Value
extSheet.Cells(extRow, Gross_Premium) = .Cells.Find("Tech Premium").Offset(0, 2).Value
extSheet.Cells(extRow, Surcharges_Taxes) = .Range("M:M").Find("Surcharge Premium").Offset(1, 0).Value
extSheet.Cells(extRow, FAC) = .Range("AC:AC").Find("FAC").Offset(0, 2).Value
End With
On Error GoTo 0
With extSheet
''-----------Trans type---------------------------
        If StrComp(UCase(.Cells(extRow, Trans_Type)), "NEW", vbTextCompare) = 0 Then
        .Cells(extRow, Trans_Type) = "New Business"
        End If
'' ----------Name check for UW column-------------------------
        For j = 2 To uaUwSheet.Cells(Rows.Count, 1).End(xlUp).Row
        
            If StrComp(.Cells(extRow, UW).Value, uaUwSheet.Cells(j, 1).Value, vbTextCompare) = 0 Then
            doColor = False

            ElseIf InStr(1, .Cells(extRow, UW).Value, "Morgan", vbTextCompare) > 0 _
                    And InStr(1, .Cells(extRow, UW).Value, "Thomas", vbTextCompare) > 0 Then
            .Cells(extRow, UW).Value = "Morgan, Tom"
            doColor = False

            ElseIf InStr(1, uaUwSheet.Cells(j, 1).Value, Split(.Cells(extRow, UW).Value, " ")(0), vbTextCompare) > 0 And _
                    InStr(1, uaUwSheet.Cells(j, 1).Value, Split(.Cells(extRow, UW).Value, " ")(1), vbTextCompare) > 0 Then
            .Cells(extRow, UW).Value = uaUwSheet.Cells(j, 1).Value
            doColor = False
            End If
        Next j
        If doColor = True Then
        .Cells(extRow, UW).Interior.ColorIndex = 46
        End If
        doColor = True
'' --------------------name check for UA column--------------------------
        For K = 2 To uaUwSheet.Cells(Rows.Count, 2).End(xlUp).Row
            If StrComp(.Cells(extRow, UA), uaUwSheet.Cells(K, 2).Value, vbTextCompare) = 0 Then
            doColor = False
            ElseIf InStr(1, uaUwSheet.Cells(K, 2).Value, Split(.Cells(extRow, UA), " ")(0), vbTextCompare) > 0 And _
                   InStr(1, uaUwSheet.Cells(K, 2).Value, Split(.Cells(extRow, UA), " ")(1), vbTextCompare) > 0 Then
            .Cells(extRow, UA) = uaUwSheet.Cells(K, 2).Value
            doColor = False
            End If
        Next K
        If doColor = True Then
        .Cells(extRow, UA).Interior.ColorIndex = 46
        End If
''--------------------Broker Statement column---------------------
        If InStr(1, UCase(.Cells(extRow, Broker_Statement)), "STATEMENT", vbTextCompare) > 0 Then
        .Cells(extRow, Broker_Statement) = "Yes"
        ElseIf InStr(1, UCase(.Cells(extRow, Broker_Statement)), "Invoice", vbTextCompare) > 0 Then
        .Cells(extRow, Broker_Statement) = "No"
        End If
'' ------------------TRIA column--------------------
        If .Cells(extRow, TRIA) = 0 Then
        .Cells(extRow, TRIA) = "No"
        ElseIf .Cells(extRow, TRIA) = "" Then
        .Cells(extRow, TRIA) = "No"
        Else
        .Cells(extRow, TRIA) = "Yes"
        End If
''-------------Cell format----------------
.Range("L:M").NumberFormat = "mm/dd/yyyy"
.Range("Z:Z").NumberFormat = "0.00%"
.Range("AA:AA").NumberFormat = "$#,##0.00;[Red]$#,##0.00"
End With
''-------Border and alignment------------------------------------------------
With extSheet.UsedRange.Borders
    .Weight = xlThin
    .LineStyle = xlContinuous
End With
extSheet.UsedRange.HorizontalAlignment = xlLeft
extSheet.Activate
''-------------------Check if Macro found every column, if not, then show (not found) column name in msgbox----------------

If extSheet.Cells(extRow, Trans_Type) = "" Then: msgboxtext = msgboxtext & " Trans Type" & vbNewLine

If extSheet.Cells(extRow, UW) = "" Then: msgboxtext = msgboxtext & " UW" & vbNewLine

If extSheet.Cells(extRow, UA) = "" Then: msgboxtext = msgboxtext & " UA" & vbNewLine

If extSheet.Cells(extRow, Genius_Policy) = "" Then: msgboxtext = msgboxtext & " Genius Policy" & vbNewLine

If extSheet.Cells(extRow, Insured) = "" Then: msgboxtext = msgboxtext & " Insured" & vbNewLine

If extSheet.Cells(extRow, Eff_Date) = "" Then: msgboxtext = msgboxtext & " Eff Date" & vbNewLine

If extSheet.Cells(extRow, Exp_Date) = "" Then: msgboxtext = msgboxtext & " Exp Date" & vbNewLine

If extSheet.Cells(extRow, Broker_Statement) = "" Then: msgboxtext = msgboxtext & " Broker Statement" & vbNewLine

If extSheet.Cells(extRow, Broker_Name) = "" Then: msgboxtext = msgboxtext & " Broker Name" & vbNewLine

If extSheet.Cells(extRow, Broker_Code) = "" Then: msgboxtext = msgboxtext & " Broker Code" & vbNewLine

If extSheet.Cells(extRow, TRIA) = "" Then: msgboxtext = msgboxtext & " TRIA" & vbNewLine

If extSheet.Cells(extRow, Commision) = "" Then: msgboxtext = msgboxtext & " Commission" & vbNewLine

If extSheet.Cells(extRow, Gross_Premium) = "" Then: msgboxtext = msgboxtext & " Gross Premium" & vbNewLine

If extSheet.Cells(extRow, Surcharges_Taxes) = "" Then: msgboxtext = msgboxtext & " Surcharges Taxes" & vbNewLine

If extSheet.Cells(extRow, FAC) = "" Then: msgboxtext = msgboxtext & " FAC" & vbNewLine

If msgboxtext <> "" Then
msgboxtext = "Either below Columns are empty in coding sheet or macro couldn't find them." & vbNewLine & msgboxtext
End If
''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Data extracted to row number " & extRow & "." & vbNewLine & msgboxtext, vbInformation, "CODING SHEET EXTRACTION TOOL"
Set macroSheet = Nothing
Set extSheet = Nothing
Exit Sub
errorHandlerPath:
MsgBox "Please select Coding Sheet excel file. " & Err.Description, vbCritical, "CODING SHEET EXTRACTION TOOL"
End Sub























