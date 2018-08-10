Attribute VB_Name = "Tools"
'roadmap
Option Explicit
Public arrFoldersToCreate, arrDomains, arrCities
Public WorkingFolder As String
Sub subSetFolders()
'initialization of variables needed for all other scripts
    arrFoldersToCreate = Array("\ISSUES")
    arrDomains = Array("VEON", "SYST", "USMs", "VTBs")
    arrCities = Array("MSK", "VLG", "RoD", "KRD")
    WorkingFolder = "C:\OneDrive - Nokia\!wrk\"
    'Debug.Print "Working folder is " & WorkingFolder
End Sub

Sub CheckFolderExist(strFld2Chck)
'check folder if exist under inbox
'if does not then create it under inbox and then new folder will be created under ths one
'stub so far
End Sub
'updated 2018-02-12
Sub NewTaskFromEmail()
On Error Resume Next
'initialize folders to be used
subSetFolders
'uses global variable WorkingFolder
    Dim app As New Outlook.Application
    Dim Item As Object
    Set Item = app.ActiveInspector.CurrentItem
        If Item.Class <> olMail Then Exit Sub
    Dim email As MailItem
    Set email = Item
    Dim taskItem As taskItem
    Set taskItem = app.CreateItem(olTaskItem)
    Dim dtRecvdDate, srtMnth, strDay, strSubject
        dtRecvdDate = email.ReceivedTime
        'set received date as part of the subject
        'add leading zero to numbers below 10
            srtMnth = strAddLdZero(CStr(Month(dtRecvdDate)))
            strDay = strAddLdZero(CStr(Day(dtRecvdDate)))
        'create subject that contains original subject and received date as prefix
        'ask user to chak the string and then use this string in all further calculations
            strSubject = "RU " + Join(arrDomains, " ") + Join(arrCities, " ") + CStr(Year(dtRecvdDate)) + srtMnth + strDay + " " + subjCleaner(email.Subject)
        'confirm task name and folder to create for the task
        strSubject = Trim(InputBox("I'm going to create this folder and subject for a task", "Confirm the name please", strSubject))
        taskItem.Categories = email.Categories
        taskItem.Body = email.Body
        taskItem.Subject = strSubject
        taskItem.StartDate = dtRecvdDate
        taskItem.DueDate = dtRecvdDate + 14 'by default set due date to Received + 14 days
'create folder in working diirectory
    Dim fileSys, newFolder
    If strSubject = "" Then
        strSubject = "temp"
    End If
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    newFolder = fileSys.createFolder(WorkingFolder + strSubject)
Dim i
'define array of subfolders
'uses global variable arrFoldersToCreate
    'For i = 0 To UBound(arrFoldersToCreate)
        'MsgBox (arrFoldersToCreate(i))
        'newFolder = fileSys.createFolder(WorkingFolder + strSubject + arrFoldersToCreate(i))
    'Next
    'destroy FS object
    'Set fileSys = Nothing
'create subfolder from strSubject
CreateSubFolder (strSubject)

'add attachments from original mail
    Dim attachment As attachment
    For Each attachment In email.Attachments
        CopyAttachment attachment, taskItem.Attachments, WorkingFolder + strSubject
    Next attachment
    
    Dim inspector As inspector
    
    Set inspector = taskItem.GetInspector
    inspector.Display
End Sub

Private Sub RecipientToParticipant(recipient As recipient, participants As Recipients)
subSetFolders
    Dim participant As recipient
    If LCase(recipient.Address) <> LCase(Session.CurrentUser.Address) Then
        Set participant = participants.Add(recipient.Address)
        Select Case recipient.Type
        Case olBCC:
            participant.Type = olOptional
        Case olCC:
            participant.Type = olOptional
        Case olOriginator:
            participant.Type = olRequired
        Case olTo:
            participant.Type = olRequired
        End Select
        participant.Resolve
    End If

End Sub
'Copies attachments from original message
Private Sub CopyAttachment(source As attachment, destination As Attachments, destDir As String)
    On Error GoTo HandleError
    subSetFolders
    Dim filename As String
    
    'filename = Environ("temp") & "\" & source.filename
    filename = destDir + "\" + source.filename
    
    source.SaveAsFile (filename)
    
    destination.Add (filename)
    
    Exit Sub
    
HandleError:
    Debug.Print Err.Description
End Sub
'removes different prefixed added by mailing software like Re, FWD etc
Private Function subjCleaner(strSubj As String)
subSetFolders
'here we remove prefixes added by various mail programs when reply and forward a message
Dim arrSubj
arrSubj = Split(strSubj, "Hа:")
strSubj = Join(arrSubj)

arrSubj = Split(strSubj, "RE:")
strSubj = Join(arrSubj)

arrSubj = Split(strSubj, "Re:")
strSubj = Join(arrSubj)

arrSubj = Split(strSubj, "FW:")
strSubj = Join(arrSubj)


subjCleaner = Trim(strSubj)

End Function
'adds leading zero for a string
Private Function strAddLdZero(strToCheck As String) As String
    'adding leading zero to numbers that contain 1 digit to have beautiful string like 2014-01-01
    If Len(strToCheck) = 1 Then
        strAddLdZero = "0" + strToCheck
    Else
        strAddLdZero = strToCheck
    End If
End Function

Private Sub CreateSubFolder(strFolder As String)
On Error Resume Next
'init folders
subSetFolders
' assumes folder doesn't exist, so only call if calling sub knows that
' the folder doesn't exist; returns a folder object to calling sub
Dim olApp As Outlook.Application
Dim olNS As Outlook.NameSpace
Dim olInbox As Outlook.MAPIFolder
Dim olsubFolder As Outlook.MAPIFolder

Set olApp = Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")

Set olInbox = olNS.GetDefaultFolder(olFolderInbox)

olInbox.Folders.Add (strFolder)

Set olsubFolder = olInbox.Folders.Item(strFolder)

'Debug.Print olSubfolder.Name
'Dim i As Integer

'For i = 0 To UBound(arrFoldersToCreate)
'    olSubFolder.Folders.Add (Replace(arrFoldersToCreate(i), "\", ""))
'Next

ExitProc:
Set olInbox = Nothing
Set olNS = Nothing
Set olApp = Nothing
Set olsubFolder = Nothing
End Sub

Private Sub fnRemoveUnderscores()
'subSetFolders
'removes underscores and commas from directory names
On Error Resume Next
Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
Dim olFolderUnderInbox As Outlook.MAPIFolder
Dim InboxItem As Object

Set olApp = Outlook.Application
Set objNS = olApp.GetNamespace("MAPI")
Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
For Each olFolderUnderInbox In olFolder.Folders
    olFolderUnderInbox.Name = Replace(olFolderUnderInbox.Name, "_", " ")
    olFolderUnderInbox.Name = Replace(olFolderUnderInbox.Name, ", ", " ")
Next

End Sub

Sub displayCurrentFolderTask()

Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim olTasksFolder As Outlook.MAPIFolder
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Set olTasksFolder = objNS.GetDefaultFolder(olFolderTasks)
Dim olTask As Outlook.taskItem
'define regexp object
subSetFolders
'don't need these for IT tasks
'Dim re
'Set re = CreateObject("VBScript.Regexp")
'    re.IgnoreCase = True
'    re.Global = True
'    re.Pattern = "....-..-.. OM*"
'###
Dim strCurrentFolder As String
    
'If re.Test(Application.ActiveExplorer.CurrentFolder.Name) Then
            strCurrentFolder = Application.ActiveExplorer.CurrentFolder.Name
            
'            ElseIf re.Test(Application.ActiveExplorer.CurrentFolder.Parent) Then
'                strCurrentFolder = Application.ActiveExplorer.CurrentFolder.Parent
'                Else
'                    strCurrentFolder = ""
'End If
    If strCurrentFolder = "Inbox" Then
    openYearlyTask
    End If
    
    For Each olTask In olTasksFolder.Items
        If olTask.Subject = strCurrentFolder Then
            olTask.Body = Date & " " & Time & vbCrLf & olTask.Body
            olTask.Display
            Exit For
            'olApp.ActiveInspector.Activate
    End If
Next
Set olApp = Nothing
Set objNS = Nothing
Set olTasksFolder = Nothing
Set olTask = Nothing
End Sub

Sub fnCalStubFromMAPIFolder()
'Takes current MAPI folder name from 1st level (under inbox) and creates appointment in the calendar with the subject that is equal to the current folder name
'this is needed to remember tasks I was working and create stub in the calendar to avoid meeting invitations
On Error Resume Next
subSetFolders
'defining variables
Dim myNamespace As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
'to store currently selected folder name
Dim strCurrentFolder
Dim objAppoint As Object
'init working folders
Set myNamespace = Application.GetNamespace("MAPI")
    strCurrentFolder = Application.ActiveExplorer.CurrentFolder.Name
    'remove possible exclamation marks and spaces in the beginning
    strCurrentFolder = RTrim(LTrim(Replace(strCurrentFolder, "!", "")))
    'strCurrentFolder = Right(strCurrentFolder, Len(strCurrentFolder) - 11)
'now, create appontment in calendar with subject named as current folder in the MAPI tree
Set objAppoint = CreateItem(olAppointmentItem)
objAppoint.Subject = strCurrentFolder
objAppoint.Categories = "NOKIA"
Dim inspector As inspector
    Set inspector = objAppoint.GetInspector
    inspector.Display
    Set myNamespace = Nothing
    Set olFolder = Nothing
    Set objAppoint = Nothing
End Sub
Sub fnMailNameFromMAPIFolder()
subSetFolders
'Takes current MAPI folder name from 1st level (under inbox) and creates appointment in the calendar with the subject that is equal to the current folder name
'this is needed to remember tasks I was working and create stub in the calendar to avoid meeting invitations
On Error Resume Next
'defining variables
Dim myNamespace As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
'to store currently selected folder name
Dim strCurrentFolder
Dim objMailItem As Object
Dim strHeader, strFooter
strHeader = "<html><body><div style='font-family:Arial Unicode MS;font-size:10pt'>"
strFooter = "</div></body></html>"

'init working folders
Set myNamespace = Application.GetNamespace("MAPI")
    'define currently selected folder in a tree
     strCurrentFolder = Application.ActiveExplorer.CurrentFolder.Name
    'remove possible exclamation marks and spaces in the beginning
    strCurrentFolder = RTrim(LTrim(Replace(strCurrentFolder, "!", "")))
    'strCurrentFolder = Right(strCurrentFolder, Len(strCurrentFolder) - 11)
    'set full path to be openned
'now, create appontment in calendar with subject named as current folder in the MAPI tree
Set objMailItem = CreateItem(olMailItem)
frmPrefixes.Show
If frmPrefixes.lstBLPrefixes.Value <> "" Then
    If frmPrefixes.lstBLPrefixes.Value = "SITE VISIT" Then
        objMailItem.HTMLBody = strHeader & "Hello gents<p>Could you please approve [and organize] access to server room located at the basement of Stanislavskogo 21-18 on DD-MM-YYYY for following HPE engineers?<p>Visitors: <ul><li>Владимир Атаманов<li>Валерий Аникин</ul> <p>Engineers visited our site plenty of times and have all needed certificates.<p>Task:" & strCurrentFolder & "<br> Scope: %SCOPE%<p>Thank you in advance<p>With best regards<p>Egor" & strFooter
        strCurrentFolder = "[APPROVAL] [SITE VISIT] " & strCurrentFolder
        Else
            strCurrentFolder = "[" & frmPrefixes.lstBLPrefixes.Value & "] " & strCurrentFolder
    End If
    
    frmPrefixes.lstBLPrefixes.Value = ""
    objMailItem.Subject = strCurrentFolder
        Else
    objMailItem.Subject = strCurrentFolder
End If

Dim inspector As inspector
    Set inspector = objMailItem.GetInspector
    inspector.Display
End Sub
Sub subFindFolderInMess()
subSetFolders
    frmSelectFolder.Show
End Sub
'auto decline meeting
Sub AutoDeclineMeetings(oRequest As MeetingItem)
If oRequest.MessageClass <> "IPM.Schedule.Meeting.Request" And oRequest.MessageClass <> "IPM.Appointment" Then
  Exit Sub
End If
Dim oAppt As AppointmentItem
Set oAppt = oRequest.GetAssociatedAppointment(True)

Dim oResponse
 'to accept automatically
 'Set oResponse = oAppt.Respond(olMeetingAccepted, True)
 Set oResponse = oAppt.Respond(olMeetingDeclined, True)
 'to send a response
 'oResponse.Display '.Send
 ' to decline without sending a response
 'oResponse.Close (olSave)
 oRequest.Delete
 
End Sub


'this one is to rename selected folder its corresponding task and file system folder
Sub subFldrsRename()
're-initialize working folders just in case
subSetFolders
'name of selected folder in the outlook application
Dim strCurrentFolder, strNewFolderName As String
strCurrentFolder = ""
strNewFolderName = ""
'define namespace
Dim myNamespace As Outlook.NameSpace
'define selected folder as object
Dim olFolder As Outlook.MAPIFolder
'Define file system variable anf filesystem folder
Dim fileSys, fldFSFolder
'init the namespace and FS
    

Dim olApp As Outlook.Application
Dim olTasksFolder As Outlook.MAPIFolder
Dim olTask As Outlook.taskItem
    
    Set olApp = Outlook.Application
    Set myNamespace = Application.GetNamespace("MAPI")
    Set olTasksFolder = myNamespace.GetDefaultFolder(olFolderTasks)
Set fileSys = CreateObject("Scripting.FileSystemObject")
'get currently selected folder name as string
strCurrentFolder = Application.ActiveExplorer.CurrentFolder.Name
Debug.Print strCurrentFolder
If MsgBox("Do you want to rename current folder?", vbYesNo, "Folder rename confirmation") = vbYes Then
    Debug.Print "Will rename"
    If fileSys.FolderExists(WorkingFolder + strCurrentFolder) Then
        Debug.Print "Folder exists, will perform the rename"
        'set new folder name from input box dialogue
        strNewFolderName = Trim(InputBox("Please enter new name", "Please enter new name", strCurrentFolder))
        'rename folder and task only in case FS folder was renamed successfully
        fileSys.MoveFolder WorkingFolder & strCurrentFolder, WorkingFolder & strNewFolderName
            If fileSys.FolderExists(WorkingFolder + strNewFolderName) Then
                Application.ActiveExplorer.CurrentFolder.Name = strNewFolderName
                Debug.Print CStr(strCurrentFolder) & " to " & CStr(strNewFolderName)
                subRenameTask CStr(strCurrentFolder), CStr(strNewFolderName)
                Else
                    MsgBox "can't rename FS folder, please check is all files are closed in this folder" & WorkingFolder & strCurrentFolder
                Exit Sub
            End If
        
            Else
        Exit Sub
        End If 'if folder exists
    Else
    Exit Sub
End If 'rename or not
WayOut:
End Sub
Sub subRenameTask(strOldName As String, strNewName As String)
'init folders just in case
subSetFolders
Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim olTasksFolder As Outlook.MAPIFolder
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Set olTasksFolder = objNS.GetDefaultFolder(olFolderTasks)
Dim olTask As Outlook.taskItem
Dim olTaskNeeded As Outlook.taskItem
'if we can find task with needed name then rename it
    For Each olTask In olTasksFolder.Items
        ' Debug.Print "found" & olTask.Subject
            If olTask.Subject = strOldName Then
                Set olTaskNeeded = olTask
                'Debug.Print "found: " & olTask.Subject
                Exit For
    End If
    Next 'for
    olTaskNeeded.Subject = strNewName
    olTaskNeeded.Save
    'Debug.Print olTaskNeeded.Subject
    'olTaskNeeded.Display
    
'destroy objects
Set olApp = Nothing
Set objNS = Nothing
Set olTasksFolder = Nothing
Set olTask = Nothing
Set olTaskNeeded = Nothing

End Sub
Sub openYearlyTask()
'kind of notebook for daily activities
subSetFolders
Dim olApp As Outlook.Application
Dim blTaskExists As Boolean
Dim objNS As Outlook.NameSpace
Dim olTasksFolder As Outlook.MAPIFolder
Dim strTaskSubj, strDate As String
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Set olTasksFolder = objNS.GetDefaultFolder(olFolderTasks)
Dim olTask As Outlook.taskItem
Dim olTaskNeeded As Outlook.taskItem
    blTaskExists = False
    strDate = Date
    strDate = DatePart("yyyy", strDate)
    strTaskSubj = "[" & strDate & "] " & "Daily tasks"
'if we can find task with needed name then rename it
    For Each olTask In olTasksFolder.Items
        ' Debug.Print "found" & olTask.Subject
            If olTask.Subject = strTaskSubj Then
                Set olTaskNeeded = olTask
                olTaskNeeded.Body = Date & " " & Time & vbCrLf & "====" & vbCrLf & olTaskNeeded.Body
                olTaskNeeded.Display
                blTaskExists = True
                Exit For
    End If
    Next 'for
    If blTaskExists = False Then
        Set olTaskNeeded = olApp.CreateItem(olTaskItem)
        olTaskNeeded.Subject = strTaskSubj
        olTaskNeeded.StartDate = "01.01." & strDate
        olTaskNeeded.DueDate = "31.12." & strDate
        olTaskNeeded.Save
        olTaskNeeded.Display
    End If
Set olApp = Nothing
Set objNS = Nothing
Set olTasksFolder = Nothing
Set olTaskNeeded = Nothing
End Sub
Private Function strInArray(strToFind As String, arrStrArray() As Variant) As Boolean
    Dim strFromArray
    strFromArray = Join(arrStrArray)
    If InStr(strFromArray, strToFind) <> 0 Then
        strInArray = True
    Else
        strInArray = False
    End If
End Function
Sub testStrInArray()
Dim arrDomains()
Dim strChoice As String
strChoice = "HELL"
arrDomains = Array("VEON", "SYST", "MGFN", "RUBS")
    If strInArray(strChoice, arrDomains) Then
        Debug.Print "gotcha"
    Else
        Debug.Print "non-gotcha"
    End If
End Sub
