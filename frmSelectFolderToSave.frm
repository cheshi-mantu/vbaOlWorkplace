VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectFolderToSave 
   Caption         =   "Please select folder to save attachments"
   ClientHeight    =   4650
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9552.001
   OleObjectBlob   =   "frmSelectFolderToSave.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectFolderToSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    frmSelectFolderToSave.Hide
End Sub
Private Sub UserForm_Activate()
    're init global variables
subSetFolders

Dim myNamespace As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
'to store currently selected folder name
Dim strCurrentFolder
Dim strResult As String
Dim strDialogResult
Dim fsFolder
Dim fsSubFolder
Dim strSubfolder As String
Dim strFullPath As String
Dim olApp As New Outlook.Application
Dim olExp As Outlook.Explorer
Dim olSel As Outlook.Selection
    
    Set olExp = olApp.ActiveExplorer
    Set olSel = olExp.Selection
        
        If olSel.Count > 1 Then
            frmSelectFolderToSave.Hide
            MsgBox ("I can process only 1 item at once")
            Else
                lstbxSelect.Clear
                    'set full path to be opened
                    strFullPath = foldersTools.buildFSPathFromNS(Application.ActiveExplorer.CurrentFolder)
                    foldersTools.fsCheckAndBuildPath (strFullPath)
                    'create shell application
                    Set objShell = CreateObject("Shell.Application")
                    Set objFS = CreateObject("Scripting.FilesystemObject")
                    'open needed folder
                    'check if needed folder exists on file system
                        If objFS.FolderExists(strFullPath) Then
                            Set fsFolder = objFS.getfolder(strFullPath)
                                For Each fsSubFolder In fsFolder.subfolders
                                    lstbxSelect.AddItem (Replace(fsSubFolder, strFullPath & "\", ""))
                                Next fsSubFolder
                        End If
           frmSelectFolderToSave.Caption = "Please select subfolder in " & strFullPath
    End If
End Sub

Private Sub btnSaveAttach_Click()

'select current item in application
'detect current folder
'to store currently selected folder name
Dim strCurrentFolder
Dim olApp As New Outlook.Application
Dim olExp As Outlook.Explorer
Dim olSel As Outlook.Selection
Dim olEmail As MailItem
Dim emailAtt As attachment

Dim strNewSubfolder As String
Dim arrName
Dim arrFilename
Dim strNewFolderPath As String

subSetFolders
    
    Set olExp = olApp.ActiveExplorer
    Set olSel = olExp.Selection
        If olSel.Item(1).Class = olMail Then
            Set olEmail = olSel.Item(1)
            strNewSubfolder = Year(olEmail.ReceivedTime)
            'needs to be refactored
                If Len(Month(olEmail.ReceivedTime)) = 1 Then
                    strNewSubfolder = strNewSubfolder & "0" & Month(olEmail.ReceivedTime)
                Else
                    strNewSubfolder = strNewSubfolder & "" & Month(olEmail.ReceivedTime)
                End If 'Len month = 1
                    If Len(Day(olEmail.ReceivedTime)) = 1 Then
                        strNewSubfolder = strNewSubfolder & "0" & Day(olEmail.ReceivedTime)
                    Else
                        strNewSubfolder = strNewSubfolder & "" & Day(olEmail.ReceivedTime)
                    End If 'len day =1
        
                arrName = Split(olEmail.Sender, " (")
                arrName = Split(arrName(0), ",")
                strNewSubfolder = strNewSubfolder & " " & arrName(0)
                Set objShell = CreateObject("Shell.Application")
                Set objFS = CreateObject("Scripting.FilesystemObject")

                'Debug.Print "btnSaveAttach:" + buildFSPathFromNS(Application.ActiveExplorer.CurrentFolder)
                'Debug.Print "btnSaveAttach:" + lstbxSelect.Value
                'Debug.Print "btnSaveAttach:" + strNewSubfolder
    
                strNewFolderPath = buildFSPathFromNS(Application.ActiveExplorer.CurrentFolder) & lstbxSelect.Value & strNewSubfolder
    
                fsCheckAndBuildPath (strNewFolderPath)
                
                'Debug.Print strNewFolderPath
    
                For Each emailAtt In olEmail.Attachments
                    If emailAtt.Type <> olOLE Then
                        emailAtt.SaveAsFile (strNewFolderPath & "\" & emailAtt.filename)
                    End If
                Next emailAtt
            End If

Set objFS = Nothing
frmSelectFolderToSave.Hide
End Sub

