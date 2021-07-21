Attribute VB_Name = "A_Assign_Task"
Sub Assigntask()

   Application.ScreenUpdating = False
   Application.DisplayAlerts = False
    
    Dim m As Workbook
    Dim I As Workbook
    Dim x As Integer
    Dim RowInc As Integer
    '-----------path change required---------
    
   '' Workbooks.Open Filename:=ThisWorkbook.Path & "\Main.xlsx"
    Set m = ThisWorkbook

'-------------Get unique names------------

    With m.Sheets("Final Data")
    
    .Range("15:15").AutoFilter field:=15, Criteria1:="", Operator:=xlFilterValues
   .Range(.Range("L15"), .Range("L" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).Copy
    m.Sheets("Help").Range("A1").PasteSpecial xlPasteValues
    
    LastRow = m.Sheets("Help").Range("A" & Rows.Count).End(xlUp).Row + 1
    
   .Range(.Range("M16"), .Range("M" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).Copy
    m.Sheets("Help").Range("A" & LastRow).PasteSpecial xlPasteValues
    
    m.Sheets("Help").Range("$A:$A").RemoveDuplicates Columns:=1, Header:=xlYes
    
    m.Sheets("Help").Range("$A:$A").SpecialCells(xlCellTypeBlanks).Rows.Delete
    
End With
 ''----------remove "pending" and "NA" from unique names------------
 With m.Sheets("Help")
 
    LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
    
    For xy = 2 To LastRow
        capital = UCase(.Cells(xy, 1).Value)
        If capital = "PENDING" Or capital = "NA" Then
        .Cells(xy, 1).EntireRow.Delete
        xy = xy - 1
        End If
    Next xy
    
    LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
End With

'--------------User Email ID VLookUP-------------------
With m.Sheets("Final Data")
  LastRow = .Cells(Rows.Count, 12).End(xlUp).Row
  RowInc = 16
    For RowInc = 16 To LastRow
          If .Range("V" & RowInc) = "" Then
          On Error Resume Next
          .Range("V" & RowInc) = Application.WorksheetFunction.VLookup(.Range("L" & RowInc), m.Sheets("Defaults").Range("B:C"), 2, 0)
          End If
    Next RowInc
 End With

''-----------Unique serial numbers for each email--------------

For sno = 16 To LastRow
    m.Sheets("Final Data").Range("A" & sno) = sno - 15
Next sno

'''---------VlookUp from Definations sheet for Processing and Review time ---------------
'    With m.Sheets("Final Data")
'    RowInc = 16
'    For RowInc = 16 To LastRow
'          If .Range("H" & RowInc) = "" Then
'          On Error Resume Next
'          .Range("H" & RowInc) = Application.WorksheetFunction.VLookup(.Range("D" & RowInc), m.Sheets("Definations").Range("A:D"), 3, 0) * .Range("G" & RowInc)
'          End If
'
'
'
'          If .Range("I" & RowInc) = "" Then
'          .Range("I" & RowInc) = Application.WorksheetFunction.VLookup(.Range("D" & RowInc), m.Sheets("Definations").Range("A:D"), 4, 0) * .Range("G" & RowInc)
'          End If
'    Next RowInc
'    End With

'--------------Assign Task--------------

LastRow = m.Sheets("Help").Cells(Rows.Count, 1).End(xlUp).Row
    For x = 2 To LastRow
    
    Workbooks.Open Filename:=ThisWorkbook.Path & "\Sample.xlsb"
    
    'Set i = Workbooks.Add
    'i.SaveAs Filename:=ThisWorkbook.Path & "\Output\" & m.Sheets("Help").Cells(x, 1) & " - " & Format(Date, "dd-mmm-yy") & ".xlsx"
     
    Set I = Workbooks("Sample.xlsb")
    
    '-------------filter by name and Copy data-----------
    m.Sheets("Final Data").Range("15:15").AutoFilter field:=15, Criteria1:="", Operator:=xlFilterValues
    m.Sheets("Final Data").Range("15:15").AutoFilter field:=12, Criteria1:=m.Sheets("Help").Cells(x, 1)
    m.Sheets("Final Data").Activate
    Cells.Select
    Selection.Copy
    With I.Sheets("Process")
        .Range("A1").PasteSpecial xlPasteValues
        .Range("A1:A14").EntireRow.Delete
        .Range("J:J").NumberFormat = "m/d/yyyy"
        .Range("K:K").NumberFormat = "m/d/yyyy"
        .Range("N:N").NumberFormat = "m/d/yyyy"
        .Range("X:X").NumberFormat = "m/d/yyyy"
        .Range("AA:AA").NumberFormat = "m/d/yyyy"
    End With
    'new
    m.Sheets("Final Data").Range("15:15").AutoFilter
    
    '-------------filter by name and Copy data - Review Tab -----------
    m.Sheets("Final Data").Range("15:15").AutoFilter field:=15, Criteria1:="", Operator:=xlFilterValues
    m.Sheets("Final Data").Range("15:15").AutoFilter field:=13, Criteria1:=m.Sheets("Help").Cells(x, 1)
    LastReviewRow = m.Sheets("Final Data").Cells(Rows.Count, 1).End(xlUp).Row
    If LastReviewRow >= 2 Then
    I.Sheets.Add
    I.Sheets("Sheet1").Name = "Review"
    m.Sheets("Final Data").Activate
    Cells.Select
    Selection.Copy
    
    With I.Sheets("Review")
    .Range("A1").PasteSpecial xlPasteValues
    .Range("A1:A14").EntireRow.Delete
    .Range("J:J").NumberFormat = "m/d/yyyy"
    .Range("K:K").NumberFormat = "m/d/yyyy"
    .Range("N:N").NumberFormat = "m/d/yyyy"
    .Range("X:X").NumberFormat = "m/d/yyyy"
    .Range("AA:AA").NumberFormat = "m/d/yyyy"
    End With
    End If
    
    '----remove filter-----
    m.Sheets("Final Data").Range("15:15").AutoFilter
 

    
    '--------path change required--------
    
    I.SaveAs Filename:=ThisWorkbook.Path & "\Output\" & m.Sheets("Help").Cells(x, 1) & " - " & Format(Date, "dd-mmm-yy") & ".xlsb"
    I.Close
    
    Next x
    
    m.Sheets("Help").Columns("A:A").ClearContents
    m.Sheets("Final Data").Range("15:15").AutoFilter
    
    MsgBox "Done...!!"

    Set I = Nothing
    Set m = Nothing
End Sub
