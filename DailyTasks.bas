Attribute VB_Name = "DailyTasks"
Sub CopyTextToClipBoard()
    Dim objItem As Object
    Dim objInsp As Outlook.inspector
    Dim strSmallTask As String
    Dim objWord As Word.Application
    Dim objDoc As Word.Document
    Dim objSel As Word.Selection
    On Error Resume Next
    
    ' Reference the current Outlook item
    Set objItem = Application.ActiveInspector.CurrentItem
    strSmallTask = "Request date: " + CStr(objItem.ReceivedTime) + vbCrLf + "Requester: " + objItem.Sender + vbCrLf + "Subject: "
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Set objInsp = objItem.GetInspector
            If objInsp.EditorType = olEditorWord Then
                Set objDoc = objInsp.WordEditor
                Set objWord = objDoc.Application
                Set objSel = objWord.Selection
                On Error GoTo NotText
                strSmallTask = strSmallTask + objSel.Text + vbCrLf + "Solution: ###TBD###" + vbCrLf + "Status: Pending" + vbCrLf
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

