Attribute VB_Name = "foldersTools"
Private Sub getDestFolderID()
'Dim olApp As Outlook.Application
Dim olNS As Outlook.NameSpace
Dim folder As Outlook.MAPIFolder
Dim subfolder As Outlook.MAPIFolder
' get local namespace
Set olApp = Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")

mailboxCount = olNS.Folders.Count

For Each folder In olNS.Folders
    Debug.Print folder.Name
    Debug.Print folder.FolderPath
        For Each subfolder In folder.Folders
        Debug.Print subfolder.FolderPath + " ID " + subfolder.EntryID
        Next subfolder
    Next folder
End Sub
Sub fsSetConfig()
    'define environment variable to build cfg file path
    Dim strFullFolderPath As String
    Dim envVar As String
        envVar = CStr(Environ("APPDATA"))
    Dim objFSys 'define file system object
    Dim objFile ' define file object
    Dim strFileName As String
    Dim objShell 'define shell scripting object
    Dim strDestFolderID As String ' define var for dest folder ID, should be string
    Dim OutApp As Outlook.Application 'define outlook app object
    Dim oNS As Outlook.NameSpace 'define outlook namespace
    Dim objDestFolder As Outlook.MAPIFolder 'define destination backup folder
    Dim strdestFldrID As String
    'Objects init
    Set OutApp = Application 'init outlook app obj
    Set oNS = OutApp.GetNamespace("MAPI") 'init namespace
    strFileName = "backup_dest.cfg" 'file name for config setting
    Set objFSys = CreateObject("Scripting.FilesystemObject") 'init file system object
    strFullFolderPath = envVar + "\OutlookHelpers" 'hardcoded is the subject for refactoring
    If objFSys.FolderExists(strFullFolderPath) Then
        'if exists then open it
        'objShell.ShellExecute strFullFolderPath, strFullFolderPath, strFullFolderPath, "open", 1
        'Debug.Print "folder exist"
        strFileName = strFullFolderPath + "\" + strFileName
            Else
        objFSys.createFolder (strFullFolderPath)
        strFileName = strFullFolderPath + "\" + strFileName
        'Debug.Print strFileName
        Set objFile = objFSys.CreateTextFile(strFileName)
        Set objFile = objFSys.getFile(strFileName)
        'Debug.Print "folder didn't exist, created new one: " + strFullFolderPath
        'Debug.Print "file created " + objFile.Path
        Set objFile = Nothing
    End If
    'Set objShell = CreateObject("Shell.Application") 'shell object needed temporary for debugging
    'objShell.ShellExecute strFullFolderPath, strFullFolderPath, strFullFolderPath, "open", 1
    MsgBox "Press OK and then please select folder for backup", vbOKOnly, "Attention!"
    Set objDestFolder = oNS.PickFolder 'select destination folder for backup
        'Debug.Print objDestFolder.FolderPath + " ID " + objDestFolder.EntryID
        strDestFolderID = CStr(objDestFolder.EntryID)
        'Debug.Print "folder ID to write to file: " + strDestFolderID
    Set objFile = objFSys.OpenTextFile(strFileName, 2, 1, -2)
        objFile.Write (strDestFolderID & vbCrLf)
        objFile.Close
    
    'Kill all the objects
    Set objFile = Nothing
    Set objFSys = Nothing
    Set objShell = Nothing
    Set objDestFolder = Nothing
    Set oNS = Nothing
    Set OutApp = Nothing
    
End Sub
Function fsGetConfig() As String
    Dim objFSys 'define file system object
    Dim objFile ' define file object
    Dim srtFileName As String 'cfg file name will be stored here
    Dim strEnvVar As String 'storage for environment variable
    Dim strDestFolder As String
        strEnvVar = CStr(Environ("APPDATA")) 'init string for env variable
        strFileName = strEnvVar + "\" + "OutlookHelpers\backup_dest.cfg" 'file name for config setting
        'Debug.Print strFileName
    Set objFSys = CreateObject("Scripting.FilesystemObject") 'init file system object
        If Not objFSys.FileExists(strFileName) Then
            fsSetConfig
        End If
    Set objFile = objFSys.OpenTextFile(strFileName, 1, 1, -2)
        strDestFolder = objFile.ReadLine
        objFile.Close
    'destructor
    Set objFile = Nothing
    Set objFSys = Nothing
        fsGetConfig = strDestFolder
End Function
Sub backupFolders()
    Dim OutApp As Outlook.Application
    Dim oNS As Outlook.NameSpace
    Dim objInboxFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objSubfolder As Outlook.MAPIFolder 'subfolders to copy from current folder
    Dim curFolder As Outlook.MAPIFolder
    Dim strDestFolder ' string from cfg file
    'init objects
    
    Set OutApp = Application
    Set oNS = OutApp.GetNamespace("MAPI")

'use the selected folder
    Set curFolder = OutApp.ActiveExplorer.CurrentFolder
    Set objInboxFolder = oNS.GetDefaultFolder(olFolderInbox)

Set objDestFolder = oNS.GetFolderFromID(fsGetConfig)
    'Debug.Print objDestFolder.Name
'On Error Resume Next
    ' copy folder
    backupProgress.Show
    For Each objSubfolder In curFolder.Folders
    'add record to log file
    backupProgress.ProgressBox.Text = "Copying " + objSubfolder + " to " + objDestFolder.Name
    backupProgress.Repaint
        objSubfolder.CopyTo objDestFolder
    Next objSubfolder

'destruct objects
backupProgress.Hide

Set OutApp = Nothing
Set oNS = Nothing
Set curFolder = Nothing
Set objInboxFolder = Nothing
Set objSubfolder = Nothing
Set objDestFolder = Nothing
End Sub
Function buildFSPathFromNS(olCurFolder As Outlook.MAPIFolder) As String
Dim olFolder As Outlook.MAPIFolder
Dim olNS As Outlook.NameSpace
Dim strBuiltPath
Dim strPathTail
'setting folders for the session
subSetFolders

Set olNS = Application.GetNamespace("MAPI")
strBuiltPath = Tools.WorkingFolder
strPathTail = ""
    If olCurFolder <> olNS.GetDefaultFolder(olFolderInbox) Then
            Set olFolder = olCurFolder
                Do While olFolder <> olNS.GetDefaultFolder(olFolderInbox)
                    strPathTail = olFolder.Name + "\" + strPathTail
                    Set olFolder = olFolder.Parent
                Loop
    End If
    strBuiltPath = strBuiltPath + strPathTail
    buildFSPathFromNS = strBuiltPath
End Function
'provides test interface for Function buildFSPathFromNS
Sub testDrivebuildFSPathFromNS()
    Debug.Print buildFSPathFromNS(Application.ActiveExplorer.CurrentFolder)
End Sub
Sub fsCheckAndBuildPath(strPath As String)
    Dim objFSys 'define file system variable
    Dim strFullPath ' variable to store path for checking
    Dim arrPath ' array to keep path parts
    
    Set objFSys = CreateObject("Scripting.FilesystemObject") 'init file system object
    arrPath = Split(strPath, "\") ' split path by \ and ignore drive letter in further loops!
    'we are sure that the disk exists
    strFullPath = arrPath(0) + "\"
    
    For i = LBound(arrPath) + 1 To UBound(arrPath) - 1
        If arrPath(i) <> "" Then
            strFullPath = strFullPath + arrPath(i) + "\"
        End If
        Debug.Print "Sub fsCheckAndBuildPath: " + strFullPath
        
        If Not objFSys.FolderExists(strFullPath) Then
            objFSys.createFolder (strFullPath)
        End If
    Next
    
    Set objFSys = Nothing
End Sub
Sub testDrivefsCheckPath()
    fsCheckAndBuildPath (buildFSPathFromNS(Application.ActiveExplorer.CurrentFolder))
End Sub
'Opens working folder on HDD for a folder selected in MAPI tree
Sub fnOpenWorkingFolder()
subSetFolders
On Error Resume Next
'defining variables
'to store currently selected folder name
Dim objShell
Dim objFS
Dim strResult As String
Dim strDialogResult

Set myNamespace = Application.GetNamespace("MAPI")
Set olInbox = myNamespace.GetDefaultFolder(olFolderInbox)

'init working folders
subSetFolders
'========================================

'define currently selected folder in a tree using regexp mask defined above
Dim strFullPath As String

    strFullPath = buildFSPathFromNS(Application.ActiveExplorer.CurrentFolder)
    'create shell application
    Set objShell = CreateObject("Shell.Application")
    Set objFS = CreateObject("Scripting.FilesystemObject")
    
    'open needed folder



    'check if needed folder exists on file system
    If objFS.FolderExists(strFullPath) Then
        'if exists then open it
        objShell.ShellExecute strFullPath, strFullPath, strFullPath, "open", 1
        Else
            ' ask user to create needed folder on file system if it does not exist
            strDialogResult = MsgBox("The folder does not exist, do you want me to create it?", vbYesNo)
                ' if Yes is pressed then create folder
                If strDialogResult = vbYes Then
                    objFS.createFolder (strFullPath)
                End If
    End If
    'Destroy objects
    Set objShell = Nothing
    Set objFS = Nothing
    Set myNamespace = Nothing
End Sub

