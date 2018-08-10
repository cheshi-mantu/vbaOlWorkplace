VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectFolder 
   Caption         =   "Search for a folder"
   ClientHeight    =   9900.001
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12012
   OleObjectBlob   =   "frmSelectFolder.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub frm1ListBox1_Click()
Dim myNamespace As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
Set myNamespace = Application.GetNamespace("MAPI")
Set olFolder = myNamespace.GetDefaultFolder(olFolderInbox)
Set Application.ActiveExplorer.CurrentFolder = olFolder.Folders.Item(frm1ListBox1.Value)
subSetFolders
End Sub

Private Sub objBtnCancel_Click()
frmSelectFolder.Hide
End Sub

Private Sub UserForm_Activate()
On Error Resume Next
Dim olApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim olFolder As Outlook.MAPIFolder
Dim olFolderUnderInbox As Outlook.MAPIFolder
Dim InboxItem As Object
subSetFolders
Set olApp = Outlook.Application
Set objNS = olApp.GetNamespace("MAPI")
Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
For Each olFolderUnderInbox In olFolder.Folders
    If Len(olFolderUnderInbox.Name) > 2 Then
        frm1ListBox1.AddItem (olFolderUnderInbox.Name)
        'add here the scan of subfolders fro folders that length is 2
    End If
    
Next
End Sub
Private Sub inpFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim arrListItems()
ReDim Preserve arrListItems(frm1ListBox1.ListCount)
Dim arrListItems1Dim()
Dim arrFiltered
subSetFolders
'get list of all folders
            For i = 0 To frm1ListBox1.ListCount - 1
                arrListItems(i) = frm1ListBox1.List(i)
            Next i
    
    If KeyCode = 13 Then
        If inpFilter.Value = "" Then
            frm1ListBox1.Clear
            For i = 0 To UBound(arrListItems)
                frm1ListBox1.AddItem (arrListItems(i))
            Next i
        Else
            frm1ListBox1.Clear
            
            arrFiltered = Filter(arrListItems, inpFilter.Value, True, vbTextCompare)
                For i = 0 To UBound(arrFiltered)
                    frm1ListBox1.AddItem (arrFiltered(i))
                Next i
        End If
    End If
End Sub

