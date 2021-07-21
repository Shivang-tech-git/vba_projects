Attribute VB_Name = "potToTemplate"
Dim oword As Word.Application, mydoc As Word.document, mypara As Long, mypara2 As Long
Dim myrange As Word.Range, lastColumn As Integer, col As Integer, y As Integer, cov_last_col As Integer

Sub potToTemplate()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim covArr(1 To 3) As Integer
Set macro = ThisWorkbook.Worksheets("Macro")
Set details = ThisWorkbook.Worksheets("Details")
Set coverages = ThisWorkbook.Worksheets("Coverages")
Set contacts = ThisWorkbook.Worksheets("Contacts")
'Set oword = New Word.Application
Set mydoc = Word.Documents.Open(macro.Cells(4, 3).Value)
'Set mydoc = Word.Documents(Word.Documents.Count)
Set myrange = mydoc.Range
searchAndExtract "Job Number: ", "localPolicy"
''------------------Clear Details worksheet-------------------------
With details
.Range(.Cells(4, Master_Details_Col), .Cells(Rows.Count, Master_Details_Col)).ClearContents
.Range(.Cells(4, Exclusions_Col), .Cells(Rows.Count, Exclusions_Col)).ClearContents
.Range(.Cells(4, Claims_Col), .Cells(Rows.Count, Claims_Col)).ClearContents
.Range(.Cells(1, Field_Name_Col + 1), .Cells(Rows.Count, Columns.Count)).ClearContents
.UsedRange.Borders.LineStyle = xlNone
.Columns("I:ZZ").Interior.ColorIndex = 2
''------------------clear coverages worksheet ---------------------------
coverages.Range(coverages.Cells(1, coverage_field_col + 1), coverages.Cells(Rows.Count, Columns.Count)).ClearContents
coverages.UsedRange.Borders.LineStyle = xlNone
coverages.Columns("B:ZZ").Interior.ColorIndex = 2
''------------------clear contacts worksheet ---------------------------
contacts.Range(contacts.Cells(1, contacts_field_col + 1), contacts.Cells(Rows.Count, Columns.Count)).ClearContents
contacts.UsedRange.Borders.LineStyle = xlNone
contacts.Columns("B:ZZ").Interior.ColorIndex = 2
''---------------Genius master policy number--------------------------------
extractFromWord "Genius Master Policy No."
.Cells(Genius_master_policy_number, Master_Details_Col) = mydoc.Paragraphs(mypara + 1).Range.text
''--------------All Countries------------------------------------------------
colNum = 1
While myrange.Find.Execute("Country:", False, True) = True
      myrange.Select
      x = mydoc.Range(0, Word.Selection.Paragraphs(1).Range.End).Paragraphs.Count
        ''---------------------- Details ----------------------------
        .Cells(page_Number_Row, Field_Name_Col + colNum) = Word.Selection.Information(wdActiveEndPageNumber)
        .Cells(country_row, Field_Name_Col + colNum) = mydoc.Paragraphs(x + 1).Range.text
        ''---------------------- Coverages --------------------------
        coverages.Cells(page_Number_Row, coverage_field_col + colNum) = Word.Selection.Information(wdActiveEndPageNumber)
        coverages.Cells(country_row, coverage_field_col + colNum) = mydoc.Paragraphs(x + 1).Range.text
        ''---------------------- contacts ---------------------------
        contacts.Cells(page_Number_Row, contacts_field_col + colNum) = Word.Selection.Information(wdActiveEndPageNumber)
        contacts.Cells(country_row, contacts_field_col + colNum) = mydoc.Paragraphs(x + 1).Range.text
        colNum = colNum + 1
Wend
Set myrange = Nothing
Set myrange = mydoc.Range

lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
cov_last_col = coverages.Cells(1, Columns.Count).End(xlToLeft).Column
con_last_col = contacts.Cells(1, Columns.Count).End(xlToLeft).Column
''------------ Local Policy Reference ---------------------------------------

searchAndExtract "Policy Ref:", "localPolicy"
searchAndExtract2 "Policy Ref:", "localPolicy2"
''------------------- Local Brokerage and percentage -------------------------
searchAndExtract "Local Brokerage:", "brokerageAndPercentage"
''------------------- Policy Trigger -----------------------------------------
searchAndExtract2 "Policy trigger", "policyTrigger"
''------------------- Local Broker Contact ------------------------------------
searchAndExtract2 "Local broker contact", "localBrokerContact"
If contacts.Cells(Broker_Name_Row, contacts_field_col + 1) = "" Then
    searchAndExtract2 "Contact person", "localBrokerContact"
End If
If contacts.Cells(Broker_Name_Row, contacts_field_col + 1) = "" Then
    searchAndExtract2 "local contact", "localBrokerContact"
End If
If contacts.Cells(Broker_Name_Row, contacts_field_col + 1) = "" Then
    searchAndExtract2 "Broker contact", "localBrokerContact"
End If
''------------------- Local insured contact -----------------------------------

''------------------ Master Details ---------------------------------
extractFromWord "Territorial Scope"
.Cells(Territory_Scope_Row, Master_Details_Col) = Replace(mydoc.Paragraphs(mypara + 1).Range.text, "Territorial Scope", "")
extractFromWord "Trade of Business"
.Cells(Business_Activity_Row, Master_Details_Col) = Replace(mydoc.Paragraphs(mypara + 1).Range.text, "Trade of Business", "")
''------------- Exclusions --------------------------------------------
extractFromWord "Master Policy Exclusions:", "Local Policy Exclusions:"
rowNum = exclusions_start_row
For a = mypara To mypara2
    If Len(mydoc.Paragraphs(a).Range.text) > 2 Then
        .Cells(rowNum, Exclusions_Col) = mydoc.Paragraphs(a).Range.text
        rowNum = rowNum + 1
    End If
Next a
''----------------------Remove next line character from extracted text-------------------------------
On Error Resume Next
''----------------------Details----------------------------
.Range(.Cells(4, Master_Details_Col), .Cells(Rows.Count, Master_Details_Col)).TextToColumns , xlDelimited
.Range(.Cells(4, Exclusions_Col), .Cells(Rows.Count, Exclusions_Col)).TextToColumns , xlDelimited, , , True
lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    For col = Field_Name_Col + 1 To lastColumn
    .Columns(col).TextToColumns , xlDelimited
    Next col
''---------------------Coverages---------------------------
    For col = coverage_field_col + 1 To cov_last_col
    coverages.Columns(col).TextToColumns , xlDelimited, , , True
    Next col
''--------------------- Contacts ---------------------------
    For col = contacts_field_col + 1 To con_last_col
    contacts.Columns(col).TextToColumns , xlDelimited, , , True
    Next col
On Error GoTo 0
''------------ Additional insureds -----------------------
searchAndExtract "Additional insured", "additionalInsured"
''------------------- coverage and Limit -----------------------------------------

searchAndExtract2 "Limit", "limit"
searchAndExtract2 "LPPC", "lppc"
searchAndExtract2 "DPPC", "dppc"
searchAndExtract2 "DPPP", "dppc"
''------------------- Deductible -----------------------------------------
If coverages.Cells(Deductible_Row, coverage_field_col + 1) = "" Then
searchAndExtract2 "Deductible", "deductible"
End If
'''------------------- Flat premium -----------------------------------------
'searchAndExtract2 "flat", "flatPremium"
''-------------------The End-------------------------------
.Columns.AutoFit
.UsedRange.Borders.LineStyle = xlContinuous
.Range(.Cells(Local_Currency_Row, Field_Name_Col + 1), .Cells(ROE_Date_Row, lastColumn)).Interior.ColorIndex = 35

coverages.Columns.AutoFit
coverages.UsedRange.Borders.LineStyle = xlContinuous
coverages.Range(coverages.Cells(SIR_Row, coverage_field_col + 1), coverages.Cells(turnover, cov_last_col)).Interior.ColorIndex = 35

contacts.Columns.AutoFit
contacts.UsedRange.Borders.LineStyle = xlContinuous

''-------------- Increment policy number for each country --------------------------
''-------------- Details ------------------
On Error Resume Next
For Each cell In .Range(.Cells(local_Policy_row, Field_Name_Col + 1), .Cells(local_Policy_row, lastColumn))
cell.Value = Left(cell.Value, Len(cell.Value) - 3) & _
            Mid(cell.Value, Len(cell.Value) - 2, 2) + 1 & Right(cell.Value, 1)
Next cell
''--------------- Coverages ------------------
For Each cell In coverages.Range(coverages.Cells(local_Policy_row, coverage_field_col + 1), coverages.Cells(local_Policy_row, cov_last_col))
cell.Value = Left(cell.Value, Len(cell.Value) - 3) & _
            Mid(cell.Value, Len(cell.Value) - 2, 2) + 1 & Right(cell.Value, 1)
Next cell
''--------------- contacts ------------------
For Each cell In contacts.Range(contacts.Cells(local_Policy_row, contacts_field_col + 1), contacts.Cells(local_Policy_row, con_last_col))
cell.Value = Left(cell.Value, Len(cell.Value) - 3) & _
            Mid(cell.Value, Len(cell.Value) - 2, 2) + 1 & Right(cell.Value, 1)
Next cell

masterPolNum = .Cells(Genius_master_policy_number, Master_Details_Col).Value
.Cells(Genius_master_policy_number, Master_Details_Col) = Left(masterPolNum, Len(masterPolNum) - 3) & _
            Mid(masterPolNum, Len(masterPolNum) - 2, 2) + 1 & Right(masterPolNum, 1)
On Error GoTo 0
End With

''-----------------Change the number format for all amounts in coverage worksheet------------------

covArr(1) = Deductible_Row
covArr(2) = adjustable
covArr(3) = Limit_Row
For covItem = 1 To 3
    For Each cell In coverages.Range(coverages.Cells(covArr(covItem), coverage_field_col + 1), coverages.Cells(covArr(covItem), cov_last_col))
        If Right(cell.Value, 3) = ".00" Or Right(cell.Value, 3) = ",00" Then
            cell.Value = Left(cell.Value, Len(cell.Value) - 3)
        End If
    Next cell
coverages.Range(covArr(covItem) & ":" & covArr(covItem)).Replace ",", ""
coverages.Range(covArr(covItem) & ":" & covArr(covItem)).Replace ".", ""
Next covItem


''------------Amount formatting in coverages------------------

Set myrange = Nothing
Set orange2 = Nothing
Set orange = Nothing
Set mydoc = Nothing
Set oword = Nothing

Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Complete", vbInformation, "GISMO PREFILL TOOL"
End Sub

''========== GLOBAL PROCEDURE TO EXTRACT DATA FROM WORD =============

Sub extractFromWord(search1 As String, Optional search2 As String)
Dim orange As Word.Range, orange2 As Word.Range
Set orange = mydoc.Range
Set orange2 = mydoc.Range
mypara = 0
mypara2 = 0
    While orange.Find.Execute(search1, False, True) = True
        orange.Select
        mypara = mydoc.Range(0, Word.Selection.Paragraphs(1).Range.End).Paragraphs.Count
    Wend
    If search2 <> "" Then
        While orange2.Find.Execute(search2, False, True) = True
            orange2.Select
            mypara2 = mydoc.Range(0, Word.Selection.Paragraphs(1).Range.End).Paragraphs.Count
        Wend
    End If
Set orange = Nothing
Set orange2 = Nothing
End Sub

''================= SUB FOR DETAILS WORKSHEET ==================

Sub searchAndExtract(keyword As String, procedure As String)
Row = 1
        While myrange.Find.Execute(keyword, False, True) = True
              myrange.Select
              y = mydoc.Range(0, Word.Selection.Paragraphs(1).Range.End).Paragraphs.Count
                    ThisWorkbook.Worksheets("Sheet1").Cells(Row, 1) = mydoc.Paragraphs(y).Range.text
                    Row = Row + 1
        Wend
    Set myrange = Nothing
    Set myrange = mydoc.Range
End Sub

''================== SUB FOR COVERAGES WORKSHEET ===================

Sub searchAndExtract2(keyword As String, procedure As String)
        While myrange.Find.Execute(keyword, False, True) = True
              myrange.Select
              y = mydoc.Range(0, Word.Selection.Paragraphs(1).Range.End).Paragraphs.Count
            For col = coverage_field_col + 1 To cov_last_col
                If Word.Selection.Information(wdActiveEndPageNumber) = coverages.Cells(page_Number_Row, col).Value Then
                    Application.Run procedure
                    Exit For
                End If
            Next col
        Wend
    Set myrange = Nothing
    Set myrange = mydoc.Range
End Sub
''=================== Details - Local policy================================
Sub localPolicy()
    details.Cells(local_Policy_row, col) = mydoc.Paragraphs(y + 1).Range.text
End Sub
''=================== Coverages - Local policy==============================
Sub localPolicy2()
 ''   coverages.Cells(local_Policy_row, col) = Left(locPol2, Len(locPol2) - 3) & _
                                        Mid(locPol2, Len(locPol2) - 2, 2) + 1 & Right(locPol2, 1)
    coverages.Cells(local_Policy_row, col) = mydoc.Paragraphs(y + 1).Range.text
    contacts.Cells(local_Policy_row, col) = mydoc.Paragraphs(y + 1).Range.text
End Sub

Sub brokerageAndPercentage()
    If InStr(1, mydoc.Paragraphs(y + 1).Range.text, "%", vbTextCompare) > 0 Then
        details.Cells(Local_Brokerage_Row, col) = "Y"
        details.Cells(Percentage_Row, col) = Replace(mydoc.Paragraphs(y + 1).Range.text, "%", "")
    ElseIf IsNumeric(Replace(Replace(mydoc.Paragraphs(y + 1).Range.text, ",", ""), ".", "")) Then
        details.Cells(Local_Brokerage_Row, col) = "Y"
        details.Cells(Flat_Amount_Row, col) = mydoc.Paragraphs(y + 1).Range.text
    Else:
        details.Cells(Local_Brokerage_Row, col) = "Y"
        details.Cells(Percentage_Row, col) = 0
    End If
End Sub

Sub policyTrigger()
    If InStr(1, mydoc.Paragraphs(y).Range.text, "occurrence", vbTextCompare) > 0 Or _
        InStr(1, mydoc.Paragraphs(y).Range.text, "occurence", vbTextCompare) > 0 Then
        coverages.Cells(Policy_Trigger_Row, col) = "Occurence"
    ElseIf InStr(1, mydoc.Paragraphs(y).Range.text, "claims", vbTextCompare) > 0 Then
        coverages.Cells(Policy_Trigger_Row, col) = "Claims Made"
    End If
End Sub

Sub localBrokerContact()
    For inc = 1 To 6
    brokerText = mydoc.Paragraphs(y + inc).Range.text
        If InStr(1, brokerText, "+", vbTextCompare) > 0 Or InStr(1, brokerText, "Tel", vbTextCompare) > 0 Then
            contacts.Cells(Broker_Phone_Row, col) = brokerText
        ElseIf InStr(1, brokerText, "@", vbTextCompare) > 0 Then
            contacts.Cells(Broker_Name_Row, col) = Replace(Replace(Split(brokerText, "@")(0), "Email:", ""), ".", " ")
            contacts.Cells(Broker_Email_Row, col) = Replace(brokerText, "Email:", "")
        End If
    Next inc
End Sub
Sub limit()
    totalwords = UBound(Split(mydoc.Paragraphs(y).Range.text, " "))
    For Words = 0 To totalwords
        limitAmount = Split(mydoc.Paragraphs(y).Range.text, " ")(Words)
        If IsNumeric(Replace(Replace(limitAmount, ",", ""), ".", "")) Then
            coverages.Cells(Limit_Row, col).NumberFormat = "@"
            coverages.Cells(Limit_Row, col) = limitAmount
            Exit For
        ElseIf InStr(1, limitAmount, "mil", vbTextCompare) > 0 Then
            coverages.Cells(Limit_Row, col) = Trim(Replace(limitAmount, "mil", "")) & "000000"
        ElseIf InStr(1, limitAmount, "mn", vbTextCompare) > 0 Then
            coverages.Cells(Limit_Row, col) = Trim(Replace(limitAmount, "mn", "")) & "000000"
        End If
    Next Words
    
    For Z = 0 To 2
    
        If InStr(1, mydoc.Paragraphs(y + Z).Range.text, "public", vbTextCompare) > 0 And _
            InStr(1, mydoc.Paragraphs(y + Z).Range.text, "product", vbTextCompare) > 0 Then
            coverages.Cells(coverage_row, col) = "Public & Product Liability Combined"
        ElseIf InStr(1, mydoc.Paragraphs(y + Z).Range.text, "public liability", vbTextCompare) > 0 Then
            coverages.Cells(coverage_row, col) = "Public Liability"
        End If
        
    Next Z
    
End Sub
Sub deductible()
    totalwords = UBound(Split(mydoc.Paragraphs(y).Range.text, " "))
    For Words = 0 To totalwords
        If IsNumeric(Replace(Replace(Split(mydoc.Paragraphs(y).Range.text, " ")(Words), ",", ""), ".", "")) Then
             coverages.Cells(Deductible_Row, col).NumberFormat = "@"
             coverages.Cells(Deductible_Row, col) = Split(mydoc.Paragraphs(y).Range.text, " ")(Words)
        End If
    Next Words
End Sub
'Sub flatPremium()
'    totalwords = UBound(Split(mydoc.Paragraphs(y).Range.text, " "))
'    For Words = 0 To totalwords
'        If IsNumeric(Replace(Replace(Split(mydoc.Paragraphs(y).Range.text, " ")(Words), ",", ""), ".", "")) Then
'             coverages.Cells(Flat_Amount_Row, col) = "@"
'             coverages.Cells(Flat_Amount_Row, col) = Split(mydoc.Paragraphs(y).Range.text, " ")(Words)
'        End If
'    Next Words
'End Sub
Sub additionalInsured()
    j = 0
    i = 0
    While Len(mydoc.Paragraphs(y + i).Range.text) > 3
    instext = Trim(Replace(mydoc.Paragraphs(y + i).Range.text, "Additional insured", ""))
    
  ''------------For first line in additional insured-------------------
        If i = 0 Then
            If InStr(1, instext, "None", vbTextCompare) > 0 Then
                i = i + 1
                GoTo nextName
            ElseIf InStr(1, instext, "To be entered manually", vbTextCompare) > 0 Then
                i = i + 1
                GoTo nextName
            ElseIf Len(instext) <= 3 Then
                i = i + 1
                GoTo nextName
            End If
        End If
  ''---------------For more than one insured having separating address with comma in line --------------
        If InStr(1, instext, ",", vbTextCompare) > 0 Then
            details.Cells(additional_insured_name + j, col) = Trim(Left(instext, WorksheetFunction.Find(",", instext) - 1))
            details.Cells(additional_insured_add + j, col) = Trim(Mid(instext, WorksheetFunction.Find(",", instext) + 1))
                If details.Cells(additional_insured_add + j, col) = "" Then
                    details.Cells(additional_insured_add + j, col) = "Please provide address"
                    details.Cells(additional_insured_add + j, col).Interior.ColorIndex = 35
                End If
        ElseIf InStr(1, instext, "none", vbTextCompare) > 0 Then
  ''--------------for not having comma in address in line-----------------------------
        Else
            details.Cells(additional_insured_name + j, col) = instext
            details.Cells(additional_insured_add + j, col) = "Please provide address"
            details.Cells(additional_insured_add + j, col).Interior.ColorIndex = 35
        End If
        i = i + 1
        j = j + 2
nextName:
    Wend
End Sub

Sub lppc()
coverages.Cells(coverage_row, col) = "Public & Product Liability Combined"
For td = 1 To 10
    limitText = Trim(Left(mydoc.Paragraphs(y + td).Range.text, Len(mydoc.Paragraphs(y + td).Range.text) - 1))
        If IsNumeric(Replace(Replace(Left(limitText, Len(limitText) - 1), ",", ""), ".", "")) Then
            coverages.Cells(Limit_Row, col) = limitText
            coverages.Cells(Policy_Trigger_Row, col) = Trim(Left(mydoc.Paragraphs(y + td + 1).Range.text, Len(mydoc.Paragraphs(y + td + 1).Range.text) - 1))
            coverages.Cells(SIR_Row, col) = Trim(Left(mydoc.Paragraphs(y + td + 2).Range.text, Len(mydoc.Paragraphs(y + td + 2).Range.text) - 1))
        End If
Next td

End Sub

Sub dppc()
For td = 12 To 18
    dedText = Trim(Left(mydoc.Paragraphs(y + td).Range.text, Len(mydoc.Paragraphs(y + td).Range.text) - 1))
        If IsNumeric(Replace(Replace(Left(dedText, Len(dedText) - 1), ",", ""), ".", "")) Then
            coverages.Cells(Deductible_Row, col) = dedText
            Exit For
        End If
Next td
End Sub

Sub run_gismo_spider_home()
ChDir ThisWorkbook.Path
ThisWorkbook.Save
Call Shell(ThisWorkbook.Path & "\gismoSpiderHome.exe")
End Sub

'.Cells(Internal_Claim_Handler_Name_Row, Claims_Col) = y
'.Cells(Internal_Claim_Handler_Email_Row, Claims_Col) = y
'.Cells(Limit_Row, Field_Name_Col) = y
'.Cells(Policy_Trigger_Row, Field_Name_Col) = y
'.Cells(Deductible_Row, Field_Name_Col) = y
'.Cells(Type_of_Deductibe_Row, Field_Name_Col) = y
'.Cells(Coverage_Premium_Row, Field_Name_Col) = y
'.Cells(Flat_YN_Row, Field_Name_Col) = y
'.Cells(Broker_Name_Row, Field_Name_Col) = y
'.Cells(Broker_Email_Row, Field_Name_Col) = y
'.Cells(Broker_Phone_Row, Field_Name_Col) = y
'.Cells(Insured_Name_Row, Field_Name_Col) = y
'.Cells(Insured_Email_Row, Field_Name_Col) = y
'.Cells(Insured_Phone_Row, Field_Name_Col) = y
'.Cells(Local_Currency_Row, Field_Name_Col) = y
'.Cells(Rate_of_Exchange_Row, Field_Name_Col) = y
'.Cells(ROE_Date_Row, Field_Name_Col) = y
'.Cells(Local_Brokerage_Row, Field_Name_Col) = y
'.Cells(Percentage_Row, Field_Name_Col) = y
'.Cells(Flat_Amount_Row, Field_Name_Col) = y

'For Each Table In mydoc.Tables
'    For Each ro In Table.Rows
''        If InStr(1, ro.Range.Text, "Pharmaceuticals", vbTextCompare) Then
'        Debug.Print ro.Range.Text
''        End If
'    Next ro
'Next Table
