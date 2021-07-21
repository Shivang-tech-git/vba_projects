Attribute VB_Name = "C_Compile_Output"


Sub Appendfile()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wkbDest As Workbook
    Dim wkbSource As Workbook
    Set wkbDest = ThisWorkbook
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim Process_Count As Integer
    Dim Review_Count As Integer
    Dim strPath As String
    Dim counter As Integer
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.Filesystemobject")
    counter = 0
    If MsgBox("Get status and move files to (Archieve - Input) folder?", vbOKCancel + vbQuestion, "Proceed?") = vbOK Then
    
    strPath = ThisWorkbook.Path & "\Input\"
    ChDir strPath
    strextension = Dir("*.xls*")
    Do While strextension <> ""
    Set wkbSource = Workbooks.Open(strPath & strextension)
    wkbSource.Sheets("Process").Activate
    lr1 = wkbSource.Sheets("Process").Cells(Rows.Count, 1).End(xlUp).Row
    lr2 = wkbDest.Sheets("Final Data").Cells(Rows.Count, 1).End(xlUp).Row
    
    '---------------------VLookup from S.No. for Process Sheet--------------
    
    With wkbDest.Sheets("Final Data")
    RowIncrement = 16
    columnIncrement = 15
    For RowIncrement = 16 To lr2
        For columnIncrement = 15 To 21
            On Error Resume Next
            .Cells(RowIncrement, columnIncrement) = Application.WorksheetFunction.VLookup(.Cells(RowIncrement, 1), wkbSource.Sheets("Process").Range("A2:U" & lr1), Application.WorksheetFunction.Match(.Cells(15, columnIncrement), wkbSource.Sheets("Process").Range("A1:U1"), 0), 0)
        Next columnIncrement
        On Error Resume Next
    .Cells(RowIncrement, 6) = Application.WorksheetFunction.VLookup(.Cells(RowIncrement, 1), wkbSource.Sheets("Process").Range("A2:U" & lr1), Application.WorksheetFunction.Match(.Cells(15, 6), wkbSource.Sheets("Process").Range("A1:U1"), 0), 0)
    
    Next RowIncrement
    End With
    
    


    wkbSource.Close
    
    ''-------------Move files from input folder to archieve input folder--------------
    FSO.MoveFile source:=strPath & strextension, Destination:=ThisWorkbook.Path & "\Archieve - Input\" & strextension
    
    ''----------------------------------------------------------------------------------
    strextension = Dir
    counter = counter + 1
    Loop
    MsgBox counter & " files transferred.", vbInformation, "Archieve - Input folder"
    Else
    MsgBox "Tool didn't run and Input files are not transferred.", vbInformation, "Information"
    End If
    
    With wkbDest.Sheets("Final Data")
    .Range("O:O").NumberFormat = "m/d/yyyy"
    .Range("p:p").NumberFormat = "m/d/yyyy"
    End With
    
    ''---------Move files from Output folder to Archieve Output folder.----------
    counter = 0
    If MsgBox("Click Ok to move files to (Archieve - Output) folder.", vbOKCancel + vbQuestion, "Move Output files?") = vbOK Then
    Set FSO = CreateObject("Scripting.Filesystemobject")
    strPath = ThisWorkbook.Path & "\Output\"
    ChDir strPath
    strextension = Dir("*.xls*")
    Do While strextension <> ""
    FSO.MoveFile source:=strPath & strextension, Destination:=ThisWorkbook.Path & "\Archieve - Output\" & strextension
    strextension = Dir
    counter = counter + 1
    Loop
    MsgBox counter & " files transferred.", vbInformation, "Archieve - output folder"
    Else
    MsgBox "Output files are not transferred.", vbInformation, "Information"
    End If
    ''----------------------------------------------------------------------------
    Set wkbDest = Nothing
    Set wkbSource = Nothing
    
    '----------------VlookUp from S.No. for Review Sheet--------------------
    
'    For Each ws In wkbSource.Worksheets
'
'    If ws.Name = "Review" Then
'    wkbSource.Sheets("Review").Activate
'    lr1 = wkbSource.Sheets("Review").Cells(Rows.Count, 1).End(xlUp).Row
'
'    With wkbDest.Sheets("Final Data")
'    RowIncrement = 16
'    columnIncrement = 15
'    For RowIncrement = 16 To lr2
'        For columnIncrement = 15 To 21
'            On Error Resume Next
'            .Cells(RowIncrement, columnIncrement) = Application.WorksheetFunction.VLookup(.Cells(RowIncrement, 1), wkbSource.Sheets("Review").Range("A2:U" & lr1), Application.WorksheetFunction.Match(.Cells(15, columnIncrement), wkbSource.Sheets("Review").Range("A1:U1"), 0), 0)
'        Next columnIncrement
'    Next RowIncrement
'    End With
'
'    End If
'
'    Next ws
End Sub
