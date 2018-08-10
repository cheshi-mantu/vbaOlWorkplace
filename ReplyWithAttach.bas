Attribute VB_Name = "ReplyWithAttach"
Sub ReplyWithAttachments()
    Dim oReply As Outlook.mailItem
    Dim oItem As Object
     
    Set oItem = GetCurrentItem()
    If Not oItem Is Nothing Then
        Set oReply = oItem.Reply
        CopyAttachments oItem, oReply
        oReply.Display
        oItem.UnRead = False
    End If
     
    Set oReply = Nothing
    Set oItem = Nothing
End Sub
Sub ReplyAllWithAttachments()
    Dim oReply As Outlook.mailItem
    Dim oItem As Object
     
    Set oItem = GetCurrentItem()
    If Not oItem Is Nothing Then
        Set oReply = oItem.ReplyAll
        CopyAttachments oItem, oReply
        oReply.Display
        oItem.UnRead = False
    End If
     
    Set oReply = Nothing
    Set oItem = Nothing
End Sub
Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
         
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
     
    Set objApp = Nothing
End Function
 
Sub CopyAttachments(objSourceItem, objTargetItem)
   Dim fso, fldTemp, strPath, strFile, objAtt
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set fldTemp = fso.GetSpecialFolder(2) ' TemporaryFolder
   strPath = fldTemp.Path & "\"
   For Each objAtt In objSourceItem.Attachments
      strFile = strPath & objAtt.filename
      objAtt.SaveAsFile strFile
      objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
      fso.DeleteFile strFile
   Next
 
   Set fldTemp = Nothing
   Set fso = Nothing
End Sub
