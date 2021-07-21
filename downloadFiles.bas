Attribute VB_Name = "downloadFiles"

Dim HTMLDoc As MSHTML.HTMLDocument
Dim HTMLInput As MSHTML.IHTMLElement
Dim HTMLAs As MSHTML.IHTMLElementCollection
Dim HTMLA As MSHTML.IHTMLElement
Dim IE As SHDocVw.InternetExplorerMedium

Private Declare PtrSafe Function apiGetClassName Lib "user32" Alias _
                "GetClassNameA" (ByVal Hwnd As Long, _
                ByVal lpClassname As String, _
                ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function apiGetDesktopWindow Lib "user32" Alias _
                "GetDesktopWindow" () As Long
Private Declare PtrSafe Function apiGetWindow Lib "user32" Alias _
                "GetWindow" (ByVal Hwnd As Long, _
                ByVal wCmd As Long) As Long
Private Declare PtrSafe Function apiGetWindowLong Lib "user32" Alias _
                "GetWindowLongA" (ByVal Hwnd As Long, ByVal _
                nIndex As Long) As Long
Private Declare PtrSafe Function apiGetWindowText Lib "user32" Alias _
                "GetWindowTextA" (ByVal Hwnd As Long, ByVal _
                lpString As String, ByVal aint As Long) As Long
Private Const mcGWCHILD = 5
Private Const mcGWHWNDNEXT = 2
Private Const mcGWLSTYLE = (-16)
Private Const mcWSVISIBLE = &H10000000
Private Const mconMAXLEN = 255

Private Function fGetCaption(Hwnd As Long) As String
    Dim strBuffer As String
    Dim intCount As Integer

    strBuffer = String$(mconMAXLEN - 1, 0)
    intCount = apiGetWindowText(Hwnd, strBuffer, mconMAXLEN)
    If intCount > 0 Then
        fGetCaption = Left$(strBuffer, intCount)
    End If
End Function

Function fEnumWindows()
Dim lngx As Long, lngLen As Long
Dim lngStyle As Long, strCaption As String

    lngx = apiGetDesktopWindow()
    'Return the first child to Desktop
    lngx = apiGetWindow(lngx, mcGWCHILD)

    Do While Not lngx = 0
        strCaption = fGetCaption(lngx)
        If Len(strCaption) > 0 Then
            lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
            'enum visible windows only
            If lngStyle And mcWSVISIBLE Then
                 If InStr(1, fGetCaption(lngx), "Internet Explorer", vbTextCompare) > 0 Then
                 AppActivate (fGetCaption(lngx))
                 End If
            End If
        End If
        lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
    Loop
End Function

Sub Elto_comparisontool()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim A As Variant, login As Worksheet
Set login = ThisWorkbook.Worksheets("Login")
Set IE = New InternetExplorer
IE.Visible = True
IE.navigate "https://reporting.mib.org.uk/EFTClient/Account/Login.htm"
Do While IE.readyState <> READYSTATE_COMPLETE
Loop
Set HTMLDoc = IE.document
A = Format(DateAdd("M", -1, Now), "MMMM YYYY")
''--------------Download 0030 file-----------------------------
ELTOElementFinder "input", "id", "username", "insertAdjacentText", login.Cells(3, 2)
ELTOElementFinder "input", "id", "password", "insertAdjacentText", login.Cells(3, 3)
ELTOElementFinder "input", "id", "loginSubmit", "click"
    Do Until Not IE.Busy And IE.readyState = 4
        DoEvents
    Loop
ELTOElementFinder "input", "ng-model", "search.name", "insertAdjacentText", A

Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByTagName("a")
Loop
Do While HTMLInput Is Nothing
    For Each HTMLA In HTMLAs
        If InStr(1, HTMLA.getAttribute("href"), A & " - Reconciliation", vbTextCompare) > 0 Then
                Set HTMLInput = HTMLA
                HTMLInput.Click
                Application.Wait (Now + TimeValue("00:00:02"))
                fEnumWindows
                Application.Wait (Now + TimeValue("00:00:10"))
                Application.SendKeys ("%s")
                Application.SendKeys ("{ENTER}")
                Exit For
        End If
Next HTMLA
Loop
    If HTMLA Is Nothing Then
    MsgBox "0030 File doesn't exist !", vbCritical, "ELTO Tool"
    End If
''-------------------Download 0056 file---------------------------------
IE.navigate "https://reporting.mib.org.uk/EFTClient/Account/Login.htm"
Do While IE.readyState <> READYSTATE_COMPLETE
Loop
Set HTMLDoc = IE.document
ELTOElementFinder "input", "id", "username", "insertAdjacentText", login.Cells(4, 2)
ELTOElementFinder "input", "id", "password", "insertAdjacentText", login.Cells(4, 3)
ELTOElementFinder "input", "id", "loginSubmit", "click"
    Do Until Not IE.Busy And IE.readyState = 4
        DoEvents
    Loop
ELTOElementFinder "input", "ng-model", "search.name", "insertAdjacentText", A

Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByTagName("a")
Loop
Do While HTMLInput Is Nothing
    For Each HTMLA In HTMLAs
        If InStr(1, HTMLA.getAttribute("href"), A & " - Reconciliation", vbTextCompare) > 0 Then
                Set HTMLInput = HTMLA
                HTMLInput.Click
                ''---------------------------------------------------------------------------------
                Application.Wait (Now + TimeValue("00:00:02"))
                fEnumWindows
                Application.Wait (Now + TimeValue("00:00:10"))
                Application.SendKeys ("%s")
                Application.SendKeys ("{ENTER}")
                Exit For
        End If
Next HTMLA
Loop
''--------------------------------------------------------------------------------
    If HTMLA Is Nothing Then
    MsgBox "0056 File doesn't exist !", vbCritical, "ELTO Tool"
    End If

MsgBox "Done", vbInformation, "ELTO Tool"
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub ELTOElementFinder(tagName As String, attributeField As String, attributeValue As String, _
                        Optional action As String, Optional text As Variant)
Set HTMLAs = Nothing
Set HTMLA = Nothing
Set HTMLInput = Nothing
''-------------Search for the tag name---------------------
Do While HTMLAs Is Nothing
Set HTMLAs = HTMLDoc.getElementsByTagName(tagName)
Loop
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
    Next HTMLA
Loop
End Sub
