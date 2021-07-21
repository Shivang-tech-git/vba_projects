Attribute VB_Name = "A_CopyData"
Sub loopDocxsbooking()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

StartTime = Timer

Dim ole As OLEObject
Dim wApp As word.Application
Dim wDoc As word.Document
Dim file As Scripting.file
Dim mySource As Object
Dim wraw As Worksheet
Dim wproc As Worksheet
Dim whelp As Worksheet
Dim rng As Range
Dim totalFiles As Integer, counter As Integer
Set wraw = ThisWorkbook.Sheets("Raw")
Set wproc = ThisWorkbook.Sheets("Processed")
Set whelp = ThisWorkbook.Sheets("Help")

wproc.Visible = xlSheetVisible
whelp.Visible = xlSheetVisible

'------delete previous data--------
wraw.Range("A:R").Delete
'wres.Range("A2:BJ1000").ClearContents
wproc.Cells.Delete
Dim obj As FileSystemObject
Set obj = New Scripting.FileSystemObject

'----file path---
Path = ThisWorkbook.Path & "\Word Doc\Temp Folder\"
Set mySource = obj.GetFolder(Path)

''-----------------------TotalFiles-------------------------
totalFiles = 0
For Each email_folder In mySource.SubFolders
    For Each file In email_folder.Files
        If InStr(1, file.Name, "Underwriter Referral Template", vbTextCompare) > 0 And Right(file.Name, 5) = ".docx" Then
        totalFiles = totalFiles + 1
        End If
    Next file
Next email_folder

For Each email_folder In mySource.SubFolders
    For Each file In email_folder.Files

If InStr(1, file.Name, "Underwriter Referral Template", vbTextCompare) > 0 And Right(file.Name, 5) = ".docx" Then
wraw.Range("A:R").Delete
wraw.Cells.UnMerge
Set wApp = CreateObject("Word.Application")

'------------------------------- Select all and copy from word to wraw -----------------------------------

Set wDoc = wApp.Documents.Open(email_folder.Path & "\" & file.Name, , ReadOnly)
wDoc.Range.Select
wDoc.Range.Copy
Application.Wait (Now() + TimeValue("00:00:02"))
'---paste--
wraw.Range("A1").PasteSpecial xlPasteValues
wraw.Cells.UnMerge

'' ---------------------- now copy from wraw to whelp - delete blank rows and match data --------------------
'----first row----

firstrow = wraw.Cells(1, 5).End(xlDown).Row
If firstrow <> Rows.Count Then
lastrow1 = firstrow
Else:
lastrow1 = wraw.Cells(Rows.Count, 1).End(xlUp).Row
End If

'clear help
whelp.Cells.ClearContents

'------Transpose and paste 1------
wraw.Range("A1:B" & lastrow1).Copy

'------     remove blanks   --------
ThisWorkbook.Sheets("Help").Activate
Range("A2").PasteSpecial xlPasteValues

lastrowhelp = Cells(Rows.Count, 1).End(xlUp).Row
On Error Resume Next
Range("B1:B" & lastrowhelp).Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
lastrowhelp = Cells(Rows.Count, 1).End(xlUp).Row

'format data
For t = 1 To lastrowhelp
If Cells(t, 1) = "" Then
Cells(t, 2).Copy
Cells(t - 1, 3).PasteSpecial xlPasteValues
End If
Next

'delete unwanted data

Range("A1:A" & lastrowhelp).Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

'' ---------------------- now copy the ready data from whelp and append at wproc -----------------

lastrowhelp = Cells(Rows.Count, 1).End(xlUp).Row
Range("A1:C" & lastrowhelp).Select
Selection.Copy

'------paste-----
lastrow2 = wproc.Cells(Rows.Count, 3).End(xlUp).Row

'---change here-----
wproc.Cells(lastrow2 + 2, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
wproc.Cells(lastrow2 + 2, 20).Value = Left(file.Name, WorksheetFunction.Find(".", file.Name) - 1)

'clear help
whelp.Cells.Delete
whelp.Visible = xlSheetHidden
wproc.Visible = xlSheetHidden

wApp.Quit
Set wApp = Nothing
End If
Next file
Next email_folder
'--Delete--
wraw.Cells.Delete
'wproc.Range("A:A").Delete
Set wApp = Nothing
'copy data
ThisWorkbook.Sheets("Processed").Activate
lastrowprocc = wproc.Cells(Rows.Count, 1).End(xlUp).Row

'remove blanks
If lastrowprocc > 1 Then
Range("A1:A" & lastrowprocc).Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End If
''----------if new wb exists then open otherwise add new workbook --------------
new_wb_name = Dir(ThisWorkbook.Path & "\Underwriter Referral -*")
If new_wb_name <> vbNullString Then
Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & new_wb_name)
Set wb = wb.Sheets("Sheet1")
lastrow_new_wb = wb.Cells(Rows.Count, 1).End(xlUp).Row
If lastrow_new_wb = 1 Then
'copy headers
wproc.Range("A2:O2").Select
Selection.Copy
wb.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
End If
Else
'add new workbook
Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\Underwriter Referral - " & Format(Date, "dd.mmm.yyyy") & ".xlsx"
Set wb = Workbooks("Underwriter Referral - " & Format(Date, "dd.mmm.yyyy") & ".xlsx").Sheets("Sheet1")
'copy headers
wproc.Activate
wproc.Range("A2:O2").Select
Selection.Copy
wb.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
lastrow_new_wb = wb.Cells(Rows.Count, 1).End(xlUp).Row
End If

'delete numbers and multiple headers
For O = lastrowprocc To 1 Step -1
If wproc.Cells(O, 1).Value = "1" Or wproc.Cells(O, 1).Value = "Underwriter Name " Or wproc.Cells(O, 1).Value = "Underwriter Name" Then
wproc.Cells(O, 1).EntireRow.Delete
End If
Next

'recalculate lastrow
lastrowprocc = wproc.Cells(Rows.Count, 1).End(xlUp).Row

'copy data
wproc.Range("A1:O" & lastrowprocc).Copy
wb.Cells(lastrow_new_wb + 1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

''-------------insert pdf as object to new wb --------------------
Dim ol As OLEObject
Const pdf_icon = "C:\WINDOWS\Installer\{AC76BA86-7AD7-1033-7B44-AC0F074E4100}\PDFFile_8.ico"
Const word_icon = "C:\WINDOWS\Installer\{90160000-000F-0000-0000-0000000FF1CE}\wordicon.exe"
''---------------- specify start row -------------------------------------
If lastrow_new_wb = 1 Then
start_row = 2
Else: start_row = lastrow_new_wb + 1
End If
''--------------for each file in subfolders in word doc folder -----------
For Each email_folder In mySource.SubFolders
start_col = 16
    For Each file In email_folder.Files
''----------------if file is a pdf or docx and not underwriter refferal template then insert as object -----------
        If Right(file.Name, 4) = ".pdf" Or Right(file.Name, 5) = ".docx" Then
        If InStr(1, file.Name, "Underwriter Referral Template", vbTextCompare) = 0 Then

''------------- hyperlink ---------------
wb.Hyperlinks.Add wb.Cells(start_row, start_col), Address:=Replace(file.Path, "\Temp Folder", "")

        start_col = start_col + 1
        End If
        End If
    Next file
start_row = start_row + 1
Next email_folder
'formatting new wb
wb.Columns("D:D").NumberFormat = "m/d/yyyy"
wb.Rows("1:1").Interior.ColorIndex = 19
wb.Rows("1:" & start_row).RowHeight = 50
wb.Columns("P:S").ColumnWidth = 25
wb.Columns("A:N").EntireColumn.AutoFit
wb.Cells.WrapText = True
wb.Cells.Borders.LineStyle = xlContinuous
wb.Cells.HorizontalAlignment = xlLeft
wb.Parent.Save
''------------ move all folders from temp folder t word doc folder ---------------
For Each email_folder In mySource.SubFolders
obj.MoveFolder email_folder.Path, ThisWorkbook.Path & "\Word Doc\"
Next email_folder
''---------------------------------------------------------------------------------
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Done! Task Completed in " & MinutesElapsed & " minutes/seconds"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
    
End Sub

