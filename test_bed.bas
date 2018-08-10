Attribute VB_Name = "test_bed"
Sub regexpTest()
Dim re
Set re = CreateObject("VBScript.Regexp")
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = ".?.?\s?\d\d\d\d\d\d\d\d\s"
Dim strToMatch As String

strToMatch = "20161018 IT INITIATION"
Debug.Print strToMatch & " - must be true -  " & re.Test(strToMatch)

strToMatch = "RU DEM0022988 MOW VRN 100Mbps p2p link"
Debug.Print strToMatch & " - must be true - " & re.Test(strToMatch)

strToMatch = "20150818OM20136908"
Debug.Print strToMatch & " - must be FALSE - " & re.Test(strToMatch)

strToMatch = "2015-08 18 OM20136908"
Debug.Print strToMatch & " - must be FALSE - " & re.Test(strToMatch)

strToMatch = "Access to premises"
Debug.Print strToMatch & " - must be FALSE - " & re.Test(strToMatch)

End Sub

Sub LoopThruMailboxes()
Dim olApp As Outlook.Application
Dim olNS As Outlook.NameSpace
Dim folder As Outlook.MAPIFolder
Dim subfolder As Outlook.MAPIFolder
Dim secSubFolder As Outlook.MAPIFolder
Dim olInbox As Outlook.MAPIFolder
Dim strInboxEntry
' get local namespace
Set olApp = Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")
mailboxCount = olNS.Folders.Count

Set olInbox = olNS.GetDefaultFolder(olFolderInbox)
strInboxEntry = olInbox.EntryID


For Each folder In olNS.Folders
    If (InStr(folder.Name, "nokia.com") And InStr(folder.Name, "Archive") < 1) Then
    
    Debug.Print folder.Name
    Debug.Print folder.FolderPath
    Debug.Print folder.Name & " subfolders: "
        
        For Each subfolder In folder.Folders
            If subfolder.EntryID = strInboxEntry Then
                 Debug.Print "!!!!!!!!!!!!This is Inbox!!!!!!!!!!!!!!!"
            End If

        Debug.Print subfolder.FolderPath
            For Each secSubFolder In subfolder.Folders
                Debug.Print secSubFolder.FolderPath
            Next secSubFolder
        Next subfolder
    End If
    Next folder
'Destruct====================
Set olApp = Nothing
Set olInbox = Nothing
Set olNS = Nothing
End Sub

Sub subMessageDetails()
Dim olApp As New Outlook.Application
Dim olExp As Outlook.Explorer
Dim olSel As Outlook.Selection
Dim olEmail As Object
Set olExp = olApp.ActiveExplorer
Set olSel = olExp.Selection
Dim oAppt As AppointmentItem
    Debug.Print olSel.Item(1).MessageClass
Dim olMeeting As MeetingItem
    If olSel.Item(1).MessageClass = "IPM.Schedule.Meeting.Request" Or olSel.Item(1).MessageClass = "IPM.Appointment" Then
        Set olMeeting = olSel.Item(1)
Set oAppt = olMeeting.GetAssociatedAppointment(True)
'Debug.Print oAppt.Delete
Dim olResponse
'to accept automatically
'Set oResponse = oAppt.Respond(olMeetingAccepted, True)
Set olResponse = oAppt.Respond(olMeetingDeclined, True)
'to send a response
'oResponse.Display '.Send
' to decline without sending a response

'olResponse.Close (olSave)
olMeeting.Delete
Else
    Exit Sub
End If
End Sub

Sub callerRenameTask()
    Tools.subRenameTask "TEST FOLDER", "RU TEST FOLDER"
End Sub

'carefully!
Private Sub getSubFolders()
    Dim objFS As Object
    Dim fsFolder As Object
    Dim fsSubFolder As Object
    Dim strSubfolder As String
    subSetFolders
    Set objFS = CreateObject("Scripting.FilesystemObject")
    Set fsFolder = objFS.getfolder(WorkingFolder)
    
        For Each fsSubFolder In fsFolder.subfolders
            strSubfolder = Replace(fsSubFolder, WorkingFolder, "")
            Debug.Print strSubfolder
        Next fsSubFolder
End Sub

Sub subSmallTask()
Dim objExplorer As Explorer
 Dim objMail As MailItem
 Set objExplorer = Application.ActiveExplorer
 
 Dim strPaste  As Variant
 
 strPaste = objMail.Se
'init MS forms data object
Dim objClipboard As MSForms.DataObject
'assign the object
Set objClipboard = New MSForms.DataObject
'get clipboard data
DataObj.GetFromClipboard
'paste clipboard data
strPaste = DataObj.GetText(1)
'Get the needed text
DataObj.SetText oMail.Body
'put the needed text to clipboard
DataObj.PutInClipboard

If strPaste = False Then Exit Sub
If strPaste = "" Then Exit Sub

 
Set objClipboard = Nothing
End Sub

Sub ATestCopyTextToClipBoard()
    Dim objItem As Object
    Dim objInsp As Outlook.inspector
    Dim strSmallTask As String
    Dim objWord As Word.Application
    Dim objDoc As Word.Document
    Dim objSel As Word.Selection
    On Error Resume Next
    
    ' Reference the current Outlook item
    Set objItem = Application.ActiveInspector.CurrentItem
    strSmallTask = "Requester: " + objItem.Sender + vbCrLf + "Subject: "
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Set objInsp = objItem.GetInspector
            If objInsp.EditorType = olEditorWord Then
                Set objDoc = objInsp.WordEditor
                Set objWord = objDoc.Application
                Set objSel = objWord.Selection
                On Error GoTo NotText
                strSmallTask = strSmallTask + objSel.Text + vbCrLf + "Request date: " + CStr(Now()) + vbCrLf + "Status: Pending"
                With New MSForms.DataObject
                    .SetText strSmallTask
                    .PutInClipboard
                End With
                On Error Resume Next

            End If
        End If
    End If

    Set objItem = Nothing
    Set objWord = Nothing
    Set objSel = Nothing
    Set objInsp = Nothing
NotText:
    If Err <> 0 Then
        MsgBox "Data on clipboard is not text."
    End If
End Sub
