Attribute VB_Name = "reports"
Sub LoopThroughTasks()
Dim objTask As Outlook.taskItem, objTaskFolder As Outlook.MAPIFolder
Dim objTaskItems As Outlook.Items, objNS As Outlook.NameSpace

Set objNS = Application.GetNamespace("MAPI")
Set objTaskFolder = objNS.GetDefaultFolder(olFolderTasks)
Set objTaskItems = objTaskFolder.Items

For Each objTask In objTaskItems
If objTask.Status <> olTaskComplete Then
Debug.Print objTask.Subject
'Returns value of olTaskStatus constants:
'olTaskComplete 2
'olTaskDeferred 4
'olTaskInProgress 1
'olTaskNotStarted 0
'olTaskWaiting 3
End If
Next

Set objTask = Nothing
Set objTaskItems = Nothing
Set objTaskFolder = Nothing
Set objNS = Nothing
End Sub
'===============
Sub taskreport()

Dim strReport As String
  Dim olnameSpace As Outlook.NameSpace
  Dim taskFolder As Outlook.MAPIFolder
  Dim tasks As Outlook.Items
  Dim tsk As Outlook.taskItem
  Dim objExcel As Object
  Dim exWb As Object
  Dim sht As Object
  
Set objExcel = CreateObject("Excel.Application")
'Set xlWB = objExcel.Workbooks.Open("C:\temp\report.xlsx")
Set xlWb = objExcel.Workbooks.Add()
Set xlSheet = xlWb.Sheets(1)
  

  Dim strMyName As String
  Dim x As Integer
  Dim y As Integer

  

  Set olnameSpace = Application.GetNamespace("MAPI")
  Set taskFolder = olnameSpace.GetDefaultFolder(olFolderTasks)

  Set tasks = taskFolder.Items

  strReport = ""

  'Create Header
  xlWb.Sheets(1).Cells(1, 1) = "CRM#"
  xlWb.Sheets(1).Cells(1, 2) = "Name"
  xlWb.Sheets(1).Cells(1, 3) = "CBT"
  xlWb.Sheets(1).Cells(1, 4) = "Status"
  xlWb.Sheets(1).Cells(1, 5) = "Comment"
  xlWb.Sheets(1).Cells(1, 6) = "Uploaded to IMS"
  xlWb.Sheets(1).Cells(1, 7) = "Acc Mgr"
y = 2

  For x = 1 To tasks.Count

       Set tsk = tasks.Item(x)

       strReport = strReport + tsk.Subject + "; "

       'Fill in Data
       If Not tsk.Complete Then

        xlWb.Sheets(1).Cells(y, 1) = extractCRM(tsk.Subject)
        xlWb.Sheets(1).Cells(y, 2) = extractName(tsk.Subject)
        xlWb.Sheets(1).Cells(y, 3) = "CBT SISTEMA"
        xlWb.Sheets(1).Cells(y, 4) = convStatus(tsk.Status)
        xlWb.Sheets(1).Cells(y, 5) = ""
        xlWb.Sheets(1).Cells(y, 6) = "yes"
        xlWb.Sheets(1).Cells(y, 7) = tsk.Categories
        y = y + 1

       End If

  Next x
  
objExcel.Visible = True

For Each xlSheet In xlWb.Worksheets
    xlSheet.Columns("A").EntireColumn.AutoFit
    xlSheet.Columns("B").EntireColumn.AutoFit
    xlSheet.Columns("C").EntireColumn.AutoFit
    xlSheet.Columns("D").EntireColumn.AutoFit
    xlSheet.Columns("E").EntireColumn.AutoFit
    xlSheet.Columns("F").EntireColumn.AutoFit
    xlSheet.Columns("G").EntireColumn.AutoFit
Next xlSheet


'xlWB.Close 1

Set exWb = Nothing
  
End Sub
Function extractCRM(strSubj As String)
    Dim pos As Integer
    pos = InStr(1, strSubj, "OM")
    extractCRM = Mid(strSubj, pos + 2, 8)

End Function
Function extractName(strSubj As String)
    Dim pos As Integer
    pos = InStr(1, strSubj, "OM")
    extractName = Mid(strSubj, pos + 10, Len(strSubj) - pos - 9)
End Function
Function convStatus(intStatus As Integer)
    
'olTaskComplete 2
'olTaskDeferred 4
'olTaskInProgress 1
'olTaskNotStarted 0
'olTaskWaiting 3
Select Case intStatus
Case 0
    convStatus = "Not started"
Case 1
    convStatus = "In progress"
Case 2
    convStatus = "Completed"
Case 3
    convStatus = "Waiting on decision"
Case 4
    convStatus = "Task is on hold"
    
End Select

End Function
