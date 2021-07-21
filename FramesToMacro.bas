Attribute VB_Name = "FramesToMacro"
Dim HTMLDoc As MSHTML.HTMLDocument
Dim HTMLInput As MSHTML.IHTMLElement, tabElement As IHTMLElementCollection
Dim HTMLAs As MSHTML.IHTMLElementCollection
Dim HTMLA As MSHTML.IHTMLElement
Dim loopCount As Integer
Dim IE As SHDocVw.InternetExplorerMedium

Sub framesHTML()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
''-------------Variable Declarations---------------------------------------
Dim i As Integer, accountCodeVar As String, journalSIDVar As String
Dim cell As Range
Dim divElement As MSHTML.IHTMLElement
''-------------Set Object References----------------------------------------
Set macroSheet = ThisWorkbook.Worksheets("Macro")
Set Accounts = ThisWorkbook.Worksheets("Accounts")
''------------------Open Internet Explorer--------------------------
Set IE = New InternetExplorerMedium
IE.Visible = True
IE.navigate "https://frameuatlogin.catlin.com/cas/login?spversion=20.08&spname=theFrame&envname=New%20Pit%20London&service=https%3A%2F%2Fnewpitlondon.catlin.com%3A443%2Fj_spring_cas_security_check%3Bjsessionid_frm%3DF1062E6F8D2FC8F1AC1A1C0B3E471734"
Do While IE.readyState <> READYSTATE_COMPLETE
Loop
Set HTMLDoc = IE.document
''---------------Run procedure for below account code-----------
accountCodeVar = "BKR 954764"
journalSIDVar = "49634340/1"
ledgerCodeVar = "Catlin Syndicate 2003"
settlementAmountVar = "-7,968.00"
''------------click on financials----------------
findElementByClass "mainmenu"
For Each HTMLA In HTMLAs
    If HTMLA.innerText = "Financials" Then
    HTMLA.Click
    End If
Next HTMLA
''------------click on Accounts----------------
findElementByClass "dropdown"
For Each HTMLA In HTMLAs
    If HTMLA.innerText = "Accounts" Then
    HTMLA.Click
    End If
Next HTMLA
''---------------click on search----------------
HTMLElementFinder "div", "id", "financials.accounts.search", False
HTMLInput.FirstChild.Click
''----------Select Ledger-----------------------
On Error GoTo permissionDenied
selectLedger:
HTMLElementFinder "select", "name", "ledgerCode", False
    For Each divElement In HTMLInput.Children
        If divElement.innerText = ledgerCodeVar Then
        Set HTMLInput = divElement
        HTMLInput.Selected = True
        Exit For
        End If
    Next divElement
On Error GoTo 0
''-------------------Insert Account Code-----------
HTMLElementFinder "input", "name", "accountCode", False, "insertAdjacentText", accountCodeVar
''-------------------Click search------------------
HTMLElementFinder "input", "value", "Search", False, "click"
''------------------Click on Account Code----------
On Error GoTo permissionDenied2
clickOnAccountCode:
Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByTagName("td")
Loop
Do While HTMLInput Is Nothing
For Each HTMLA In HTMLAs
    If HTMLA.classname = "tabledata" Then
        For Each divElement In HTMLA.getElementsByTagName("a")
            If divElement.innerText = accountCodeVar Then
            Set HTMLInput = divElement
            HTMLInput.Click
On Error GoTo 0
            GoTo clickOnAll
            End If
        Next divElement
    End If
Next HTMLA
Loop
 ''--------------------Click on All----------------
clickOnAll:
On Error GoTo permissionDenied3
findElementByClass "tabunselected"
Do While HTMLInput Is Nothing
    For Each HTMLA In HTMLAs
        If HTMLA.innerText = "All " Then
        Set HTMLInput = HTMLA
        HTMLInput.Click
        Exit For
        End If
    Next HTMLA
Loop
On Error GoTo 0
''---------------Find Journal ID and click on allocate-----------------
On Error GoTo PermissionDenied4
clickOnAllocate:
Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByTagName("td")
Loop
Do While HTMLInput Is Nothing
For Each HTMLA In HTMLAs
    If HTMLA.classname = "tabledata" Then
        For Each divElement In HTMLA.getElementsByTagName("a")
            If InStr(1, divElement.innerText, journalSIDVar, vbTextCompare) Then
            Set HTMLInput = HTMLA.parentElement.LastChild
            HTMLInput.FirstChild.Click
On Error GoTo 0
            GoTo clickOnDropDown
            End If
        Next divElement
    End If
Next HTMLA
Loop
''----------------------Click on Drop Down-------------------------
clickOnDropDown:
findIEwindow ("New cash allocation: Allocations")
Debug.Print HTMLDoc.title
Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByTagName("td")
Loop
Do While HTMLInput Is Nothing
For Each HTMLA In HTMLAs
    If HTMLA.classname = "tabledata" Then
        For Each divElement In HTMLA.getElementsByTagName("a")
            If InStr(1, divElement.innerText, journalSIDVar, vbTextCompare) > 0 Then
                If HTMLA.parentElement.Children(4).innerText = settlementAmountVar Then
                    Set HTMLInput = HTMLA.parentElement.FirstChild
                    HTMLInput.FirstChild.Click
                    GoTo clickOnComplete
                End If
            End If
        Next divElement
    End If
Next HTMLA
Loop
''---------------------Click on complete----------------------
clickOnComplete:
HTMLElementFinder "input", "value", "Complete", False, "click"
'IE.quit
Set HTMLDoc = Nothing
Set IE = Nothing
Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
Set macroSheet = Nothing
Application.DisplayAlerts = True
Application.ScreenUpdating = True

Exit Sub
''---------------Error Handlers-------------------------------------
permissionDenied:
If Err.Number = 70 Then
    Resume selectLedger
    Else: MsgBox Err.Description, "Frames Allocations Tool", vbCritical
End If
Exit Sub
permissionDenied2:
If Err.Number = 70 Then
    Resume clickOnAccountCode
    Else: MsgBox Err.Description, "Frames Allocations Tool", vbCritical
End If
Exit Sub
permissionDenied3:
If Err.Number = 70 Then
    Resume clickOnAll
    Else: MsgBox Err.Description, "Frames Allocations Tool", vbCritical
End If
Exit Sub
PermissionDenied4:
If Err.Number = 70 Then
    Resume clickOnAllocate
    Else: MsgBox Err.Description, "Frames Allocations Tool", vbCritical
End If
End Sub
''=========================== Activate internet explorer window ==============
Sub findIEwindow(ieTitle)
Dim SWs As New SHDocVw.ShellWindows
Dim Doc
    For Each IE In SWs
        Set Doc = IE.document
        If TypeOf Doc Is HTMLDocument Then
            Select Case Doc.title
            Case ieTitle
            Set HTMLDoc = IE.document
            End Select
        End If
    Next
    
End Sub

''==============GLOBAL PROCEDURE TO SCRAPE FRAMES========================================

Sub HTMLElementFinder(tagName As String, attributeField As String, attributeValue As String, quit As Boolean, _
                        Optional action As String, Optional text As Variant)
Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
''---------------------------------------------------------
'Application.Wait (Now + TimeValue("00:00:02"))
''-------------Search for the tag name---------------------
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByTagName(tagName)
Loop
loopCount = 0
''-------------set reference to desired element and perform action----------
Do While HTMLInput Is Nothing
    For Each HTMLA In HTMLAs
'    Debug.Print HTMLA.innerText
            If HTMLA.getAttribute(attributeField) = attributeValue Then
            Set HTMLInput = HTMLA
                Select Case action
                Case "click"
                HTMLInput.Click
                Case "insertAdjacentText"
                HTMLInput.insertAdjacentText "afterbegin", text
                End Select
                    Exit For
            End If
''-----------Exit sub if couldn't find element after 3 attempts-------------
            If quit = True Then
                If loopCount = 3 Then
                Exit Sub
                End If
            End If
    Next HTMLA
loopCount = loopCount + 1
Loop

End Sub

Sub findElementByClass(classname As String)
Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByClassName(classname)
Loop
End Sub

Sub closeMessageBox()
Dim wshshell As Object
Set wshshell = CreateObject("wscript.Shell")

Do
ret = wshshell.AppActivate("Message from webpage")
Loop Until ret = True

If ret = True Then
ret = wshshell.AppActivate("Message from webpage")
Application.Wait (Now + TimeValue("00:00:02"))
wshshell.SendKeys "{enter}"
End If

End Sub



