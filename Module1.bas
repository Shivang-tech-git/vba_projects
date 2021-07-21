Attribute VB_Name = "Module1"
Function GetSignature(ByVal fPath As String) As String
    Dim fso As Object
    Dim TSet As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set TSet = fso.GetFile(fPath).OpenAsTextStream(1, -2)
    GetSignature = TSet.Readall
    TSet.Close
End Function


Sub deductible()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''---------------Error Handling Routine-----------------------------
''on error resume next
''On Error GoTo errorHandler:
''on error goto 0
''-------------Variable Declarations---------------------------------------
Dim signature As String, I As Integer
Dim macroSheet As Worksheet
Dim OutlookApp As Outlook.Application
Dim OutlookMail As Outlook.MailItem
Dim vInspector As Object
Dim wEditor As Object
    
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set OutlookApp = CreateObject("Outlook.Application")
Set OutlookMail = OutlookApp.CreateItem(0)
Set vInspector = OutlookMail.GetInspector
Set wEditor = vInspector.WordEditor
''-----------------------Get Signature-----------------------------------
sPath = Dir(Environ("appdata") & "\Microsoft\Signatures\*.txt")
    If sPath <> "" Then
        strSignature = GetSignature(Environ("appdata") & "\Microsoft\Signatures\" & sPath)
    Else
        strSignature = ""
    End If
''----------------------Send EMail----------------------------------------

With macroSheet
OutlookMail.To = .Cells(12, 3)
OutlookMail.Subject = "Selbstbehaltanforderung " & .Cells(5, 3) & " " & .Cells(6, 3)
'OutlookMail.htmlBody = "Sehr " & .Cells(10, 2) & " " & .Cells(10, 3) & " <br> <br> in vorbezeichneter Angelegenheit hatten wir Aufwendungen von <b>" _
'                    & .Cells(7, 3) & "</b> (siehe Anlage). <br> <br>" _
'                    & "Unter Berücksichtigung des vertraglich vereinbarten Selbstbehaltes von <b>" _
'                    & .Cells(8, 3) & "</b> bitten wir Sie den Betrag von " & .Cells(11, 3) _
'                    & "<i> - unter Angabe unserer Schadennummer " & .Cells(5, 3) & "</i> - auf folgendes Konto zu überweisen: <br> <br>"

wEditor.Paragraphs(1).Range = "Sehr " & .Cells(10, 2) & " " & .Cells(10, 3)
wEditor.Paragraphs(2).Range = "in vorbezeichneter Angelegenheit hatten wir Aufwendungen von " _
                            & .Cells(7, 3) & " (siehe Anlage)."
wEditor.Paragraphs(3).Range = "Unter Berücksichtigung des vertraglich vereinbarten Selbstbehaltes von " _
                    & .Cells(8, 3) & " bitten wir Sie den Betrag von " & .Cells(11, 3) _
                    & " - unter Angabe unserer Schadennummer " & .Cells(5, 3) & " - auf folgendes Konto zu überweisen:" _
                    & vbNewLine & vbNewLine & .Cells(9, 3)

wEditor.Range(wEditor.Paragraphs(2).Range.Characters(62).Start, wEditor.Paragraphs(2).Range.Characters(68).End).Font.Bold = True
wEditor.Range(wEditor.Paragraphs(3).Range.Characters(72).Start, wEditor.Paragraphs(3).Range.Characters(78).End).Font.Bold = True
wEditor.Range(wEditor.Paragraphs(3).Range.Characters(117).Start, wEditor.Paragraphs(3).Range.Characters(159).End).Font.Italic = True
wEditor.Paragraphs(3).Range.Font.Color = vbBlack

OutlookMail.Display
End With
''-------------The End------------------------------------------------------
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "Done !", vbInformation, "DEDUCTIBLE TOOL"
Set OutlookApp = Nothing
Set OutlookMail = Nothing
Set macroSheet = Nothing
'Exit Sub
'errorHandler:
'Resume Next
End Sub



















