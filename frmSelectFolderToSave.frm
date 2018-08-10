VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectFolderToSave 
   Caption         =   "Please select folder to save attachments"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
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
    
lstbxSelect.Clear
    'set full path to be openned
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
End Sub

Private Sub btnSaveAttach_Click()
subSetFolders
'select current item in application
'detect current folder
Dim myNamespace As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
'to store currently selected folder name
Dim strCurrentFolder
Dim olApp As New Outlook.Application
Dim olExp As Outlook.Explorer
Dim olSel As Outlook.Selection
Dim olEmail As Object
Dim emailAtt As attachment
Dim strNewSubfolder As String
Dim arrName

    Set myNamespace = Application.GetNamespace("MAPI")
'define currently selected folder in a tree
'check if current item has attachments if there is none then notify via msg box
'frmSelectFolderToSave.Hide
Set olExp = olApp.ActiveExplorer
Set olSel = olExp.Selection
If olSel.Count > 1 Then
    MsgBox ("I can process only 1 item at once")
    Else
        Set olEmail = olSel.Item(1)
        If olEmail.Class = olMail Then
        strNewSubfolder = Year(olEmail.ReceivedTime)
'needs to be refactored
        If Len(Month(olEmail.ReceivedTime)) = 1 Then
                strNewSubfolder = strNewSubfolder & "0" & Month(olEmail.ReceivedTime)
            Else
                strNewSubfolder = strNewSubfolder & "" & Month(olEmail.ReceivedTime)
        End If
        
        If Len(Day(olEmail.ReceivedTime)) = 1 Then
                strNewSubfolder = strNewSubfolder & "0" & Day(olEmail.ReceivedTime)
            Else
                strNewSubfolder = strNewSubfolder & "" & Day(olEmail.ReceivedTime)
        End If
        
        arrName = Split(olEmail.Sender, " (")
        arrName = Split(arrName(0), ",")
        
        strNewSubfolder = strNewSubfolder & " " & arrName(0)
        
    Set objShell = CreateObject("Shell.Application")
    Set objFS = CreateObject("Scripting.FilesystemObject")
Dim strNewFolderPath
    strNewFolderPath = WorkingFolder & strCurrentFolder & "\" & lstbxSelect.Value & "\" & strNewSubfolder
    MsgBox (strNewFolderPath)
    'check if needed folder does not exist
    If Not objFS.FolderExists(strNewFolderPath) Then
    'if it doesn't then create it
    Debug.Print strNewFolderPath
        objFS.createFolder (strNewFolderPath)
    End If
    'destroy FS object
    Set objFS = Nothing
Dim arrFilename
    
    For Each emailAtt In olEmail.Attachments
        arrFilename = Split(emailAtt.filename, ".")
        emailAtt.SaveAsFile (strNewFolderPath & "\" & emailAtt.filename)
    Next emailAtt
      End If
End If
    frmSelectFolderToSave.Hide
End Sub


