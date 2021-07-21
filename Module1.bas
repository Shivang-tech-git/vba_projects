Attribute VB_Name = "Module1"
Sub LoadCustRibbon()

Dim hFile As Long
Dim path As String, fileName As String, ribbonXML As String, user As String

hFile = FreeFile
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
fileName = "olkexplorer.officeUI"

ribbonXML = "<mso:customUI      xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "      <mso:tab id='Run Macro' label='Run Macro' insertBeforeQ='mso:TabFormat'>" & vbNewLine
ribbonXML = ribbonXML + "        <mso:group id='Forward' label='Action' autoScale='true'>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='ForwardButton' label='Forward' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='ScreenNavigatorForwardMenu'      onAction='Project1.Forward_RMEq_noaction'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='ReminderButton' label='Reminder' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='ReminderGallery'      onAction='Project1.Reminder'/>" & vbNewLine
ribbonXML = ribbonXML + "          <mso:button id='RenameButton' label='Rename' " & vbNewLine
ribbonXML = ribbonXML + "imageMso='DatasheetColumnRename'      onAction='Project1.Rename'/>" & vbNewLine
ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
ribbonXML = ribbonXML + "</mso:customUI>"

ribbonXML = Replace(ribbonXML, """", "")

Open path & fileName For Output Access Write As hFile
Print #hFile, ribbonXML
Close hFile
Call pasteProject

End Sub

Sub pasteProject()
Dim fso As Object
Dim user As String, path As String, fileName As String
Set fso = CreateObject("Scripting.FileSystemObject")
user = Environ("Username")
path = "C:\Users\" & user & "\AppData\Roaming\Microsoft\Outlook\"
fileName = "VbaProject.OTM"
FolderName = ThisWorkbook.path & "\" & fileName
On Error GoTo box
fso.copyfile FolderName, path
On Error GoTo 0
Set fso = Nothing
MsgBox "Done. Please close and reopen outlook to see changes.", vbInformation, "Rewards and recognition tool."
Exit Sub
box:
MsgBox Err.Description & ", Please close Outlook and try again.", vbCritical, "Rewards and recognition tool."
End Sub
